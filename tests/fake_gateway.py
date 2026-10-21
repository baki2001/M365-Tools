"""In-memory SpGateway used by the spmig tests. Files are bytes; every call is recorded in ``calls``."""

from __future__ import annotations

import asyncio
import itertools
from collections.abc import AsyncIterator, Callable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from m365tools.core.errors import AccessDeniedError, NotFoundError, ValidationError
from m365tools.tools.spmig.gateway import FieldInfo, RoleAssignment, SiteUser, UploadResult, UserNotFoundError
from m365tools.tools.spmig.models import Endpoint, Item, LibraryRef

_ids = itertools.count(1)


@dataclass
class Entry:
    id: int
    url: str  # server-relative
    is_folder: bool
    content: bytes = b""
    created: datetime | None = None
    modified: datetime | None = None
    author_email: str | None = None
    author_name: str | None = None
    editor_email: str | None = None
    editor_name: str | None = None
    version: str | None = "1.0"
    checked_out_to: str | None = None
    prog_id: str | None = None
    metadata_calls: int = 0


@dataclass
class FakeLibrary:
    ref: LibraryRef
    entries: dict[str, Entry] = field(default_factory=dict)  # lower url -> entry
    unique_ids: set[int] = field(default_factory=set)
    fields: list[FieldInfo] = field(default_factory=list)
    roles: list[RoleAssignment] = field(default_factory=list)
    has_unique_roles: bool = False


class FakeGateway:
    def __init__(self, site_url: str = "https://contoso.sharepoint.com/sites/src") -> None:
        self.site_url = site_url.rstrip("/")
        self.web = "/" + self.site_url.split("://", 1)[1].split("/", 1)[1] if self.site_url.count("/") > 2 else ""
        self.libraries: dict[str, FakeLibrary] = {}
        self.default_library: str | None = None
        self.users: dict[str, SiteUser] = {}
        self.usage: tuple[int, int] | Exception = (0, 0)
        self.write_denied = False
        self.owner_group_id: int | None = None
        self.calls: list[tuple[str, Any]] = []
        self.download_delay = 0.0
        self.download_failures: dict[str, list[BaseException]] = {}
        self.break_count = 0
        self.reader_set: list[int] = []

    # ---------------------------------------------------------------- setup helpers

    def add_library(self, title: str, root_name: str | None = None, *, default: bool = False, **settings: Any
                    ) -> FakeLibrary:
        root = f"{self.web}/{root_name or title}"
        ref = LibraryRef(site_url=self.site_url, web_server_relative_url=self.web, list_id=f"guid-{title}",
                         title=title, root_server_relative_url=root, item_count=0, **settings)
        lib = FakeLibrary(ref)
        self.libraries[title] = lib
        if default or self.default_library is None:
            self.default_library = title
        return lib

    def lib(self, title: str | None = None) -> FakeLibrary:
        return self.libraries[title or self.default_library or ""]

    def _lib_for_url(self, url: str) -> FakeLibrary:
        low = url.lower().rstrip("/")
        for lib in self.libraries.values():
            root = lib.ref.root_server_relative_url.lower()
            if low == root or low.startswith(root + "/"):
                return lib
        raise NotFoundError(f"no library for {url}")

    def _add_entry(self, lib: FakeLibrary, url: str, is_folder: bool, **kw: Any) -> Entry:
        e = Entry(id=next(_ids), url=url, is_folder=is_folder, **kw)
        lib.entries[url.lower()] = e
        lib.ref.item_count = len(lib.entries)
        return e

    def add_folder(self, rel_path: str, library: str | None = None) -> Entry:
        lib = self.lib(library)
        url = lib.ref.root_server_relative_url
        e: Entry | None = None
        for seg in rel_path.strip("/").split("/"):
            url = f"{url}/{seg}"
            e = lib.entries.get(url.lower()) or self._add_entry(lib, url, True)
        assert e is not None
        return e

    def add_file(self, rel_path: str, content: bytes = b"x", library: str | None = None, **kw: Any) -> Entry:
        lib = self.lib(library)
        if "/" in rel_path:
            self.add_folder(rel_path.rsplit("/", 1)[0], library)
        kw.setdefault("modified", datetime(2024, 1, 1, 12, 0, 0))
        kw.setdefault("created", datetime(2023, 1, 1, 12, 0, 0))
        return self._add_entry(lib, f"{lib.ref.root_server_relative_url}/{rel_path}", False, content=content, **kw)

    def add_user(self, email: str, title: str = "") -> None:
        self.users[email.lower()] = SiteUser(email=email, title=title or email.split("@")[0])

    def file(self, rel_path: str, library: str | None = None) -> Entry | None:
        lib = self.lib(library)
        return lib.entries.get(f"{lib.ref.root_server_relative_url}/{rel_path}".lower())

    def write_calls(self) -> list[tuple[str, Any]]:
        writes = {"ensure_folder", "upload_file", "set_item_metadata", "break_inheritance", "set_reader_only",
                  "download_file"}
        return [c for c in self.calls if c[0] in writes]

    # ---------------------------------------------------------------- SpGateway

    async def get_default_library(self) -> LibraryRef:
        self.calls.append(("get_default_library", None))
        if self.default_library is None:
            raise NotFoundError("no default library")
        return self.libraries[self.default_library].ref

    async def get_library(self, name: str) -> LibraryRef:
        self.calls.append(("get_library", name))
        for lib in self.libraries.values():
            if lib.ref.title.lower() == name.lower() or lib.ref.root_name.lower() == name.lower():
                return lib.ref
        raise NotFoundError(f"library {name!r} not found", available=list(self.libraries))

    async def iter_items(self, lib: LibraryRef) -> AsyncIterator[dict[str, Any]]:
        self.calls.append(("iter_items", lib.title))
        for e in list(self.libraries[lib.title].entries.values()):
            yield {"entry": e}

    def build_item(self, raw: dict[str, Any], lib: LibraryRef, base_url: str) -> Item | None:
        e: Entry = raw["entry"]
        prefix = base_url.rstrip("/") + "/"
        if not e.url.lower().startswith(prefix.lower()):
            return None
        rel = e.url[len(prefix):]
        name = rel.rsplit("/", 1)[-1]
        ext = ("." + name.rsplit(".", 1)[-1].lower()) if "." in name and not e.is_folder else ""
        return Item(id=e.id, is_folder=e.is_folder, rel_path=rel, name=name, ext=ext,
                    size=0 if e.is_folder else len(e.content), created=e.created, modified=e.modified,
                    author_email=e.author_email, author_name=e.author_name, editor_email=e.editor_email,
                    editor_name=e.editor_name, version=e.version, checked_out_to=e.checked_out_to,
                    prog_id=e.prog_id, server_relative_url=e.url)

    async def iter_unique_permission_ids(
        self, lib: LibraryRef, on_page: Callable[[int], None] | None = None
    ) -> AsyncIterator[int]:
        self.calls.append(("iter_unique_permission_ids", lib.title))
        if on_page is not None:
            on_page(len(self.libraries[lib.title].entries))
        for i in sorted(self.libraries[lib.title].unique_ids):
            yield i

    async def ensure_folder(self, server_relative_folder: str) -> None:
        self.calls.append(("ensure_folder", server_relative_folder))
        lib = self._lib_for_url(server_relative_folder)
        root = lib.ref.root_server_relative_url
        rel = server_relative_folder[len(root):].strip("/")
        if rel:
            url = root
            for seg in rel.split("/"):
                url = f"{url}/{seg}"
                if url.lower() not in lib.entries:
                    self._add_entry(lib, url, True)

    async def download_file(self, server_relative_url: str, dest_path: Path) -> None:
        self.calls.append(("download_file", server_relative_url))
        fails = self.download_failures.get(server_relative_url)
        if fails:
            raise fails.pop(0)
        if self.download_delay:
            await asyncio.sleep(self.download_delay)
        lib = self._lib_for_url(server_relative_url)
        e = lib.entries.get(server_relative_url.lower())
        if e is None or e.is_folder:
            raise NotFoundError(server_relative_url)
        Path(dest_path).write_bytes(e.content)

    async def upload_file(self, folder_server_relative_url: str, name: str, src_path: Path,
                          overwrite: bool = True) -> UploadResult:
        self.calls.append(("upload_file", f"{folder_server_relative_url}/{name}"))
        lib = self._lib_for_url(folder_server_relative_url)
        folder = folder_server_relative_url.rstrip("/")
        if folder.lower() != lib.ref.root_server_relative_url.lower() and folder.lower() not in lib.entries:
            raise NotFoundError(f"folder {folder} does not exist")
        url = f"{folder}/{name}"
        existing = lib.entries.get(url.lower())
        if existing is not None and not overwrite:
            raise ValidationError("exists")
        data = Path(src_path).read_bytes()
        now = datetime(2026, 10, 21, 10, 0, 0)
        if existing is not None:
            existing.content, existing.modified = data, now
            e = existing
        else:
            e = self._add_entry(lib, url, False, content=data, created=now, modified=now,
                                author_email="migrator@contoso.com", editor_email="migrator@contoso.com")
        return UploadResult(unique_id=f"uid-{e.id}", item_id=e.id)

    async def set_item_metadata(self, lib: LibraryRef, item_id: int, created: datetime | None,
                                modified: datetime | None, author_email: str | None,
                                editor_email: str | None) -> None:
        self.calls.append(("set_item_metadata", (item_id, author_email, editor_email)))
        for email in (author_email, editor_email):
            if email and email.lower() not in self.users:
                raise UserNotFoundError(f"user {email} not found")
        e = next(x for x in self.libraries[lib.title].entries.values() if x.id == item_id)
        e.metadata_calls += 1
        if created:
            e.created = created
        if modified:
            e.modified = modified
        if author_email:
            e.author_email = author_email
        if editor_email:
            e.editor_email = editor_email

    async def get_site_usage(self) -> tuple[int, int]:
        self.calls.append(("get_site_usage", None))
        if isinstance(self.usage, Exception):
            raise self.usage
        return self.usage

    async def probe_write_access(self, folder_url: str) -> None:
        self.calls.append(("probe_write_access", folder_url))
        if self.write_denied:
            raise AccessDeniedError("403 Forbidden")

    async def get_role_assignments(self, lib: LibraryRef) -> list[RoleAssignment]:
        self.calls.append(("get_role_assignments", lib.title))
        return list(self.libraries[lib.title].roles)

    async def get_associated_owner_group_id(self) -> int | None:
        return self.owner_group_id

    async def has_unique_role_assignments(self, lib: LibraryRef) -> bool:
        return self.libraries[lib.title].has_unique_roles

    async def break_inheritance(self, lib: LibraryRef, copy: bool = True) -> None:
        self.calls.append(("break_inheritance", copy))
        self.break_count += 1
        self.libraries[lib.title].has_unique_roles = True

    async def set_reader_only(self, lib: LibraryRef, principal_id: int) -> None:
        self.calls.append(("set_reader_only", principal_id))
        fl = self.libraries[lib.title]
        if not fl.has_unique_roles:
            raise ValidationError("inherits")
        ra = next((r for r in fl.roles if r.principal_id == principal_id), None)
        if ra is None or not ra.can_write:
            raise ValidationError("never widen")
        ra.role_names, ra.can_write = ["Read"], False
        self.reader_set.append(principal_id)

    async def get_fields(self, lib: LibraryRef) -> list[FieldInfo]:
        return list(self.libraries[lib.title].fields)

    async def find_site_user(self, email: str) -> SiteUser | None:
        self.calls.append(("find_site_user", email))
        return self.users.get(email.lower())


class FakeGatewayWithLocalPart(FakeGateway):
    async def find_users_by_local_part(self, local: str) -> list[SiteUser]:
        return [u for u in self.users.values() if u.email.split("@")[0].lower() == local.lower()]


def factory_for(gateways: dict[str, FakeGateway]):
    """GatewayFactory keyed on endpoint.site_url."""

    async def make(ep: Endpoint) -> FakeGateway:
        return gateways[ep.site_url.rstrip("/")]

    return make


# ---------------------------------------------------------------- pure-model helpers


def make_lib(title: str = "Docs", root: str = "/sites/a/Docs", **kw: Any) -> LibraryRef:
    return LibraryRef(site_url="https://contoso.sharepoint.com/sites/a", web_server_relative_url="/sites/a",
                      list_id=f"guid-{title}", title=title, root_server_relative_url=root, item_count=0, **kw)


def make_item(rel_path: str, *, folder: bool = False, size: int = 10, modified: datetime | None = None,
              base: str = "/sites/a/Docs", **kw: Any) -> Item:
    name = rel_path.rsplit("/", 1)[-1]
    ext = ("." + name.rsplit(".", 1)[-1].lower()) if "." in name and not folder else ""
    defaults: dict[str, Any] = {
        "id": next(_ids), "is_folder": folder, "rel_path": rel_path, "name": name, "ext": ext,
        "size": 0 if folder else size, "created": datetime(2023, 1, 1),
        "modified": modified or datetime(2024, 1, 1), "author_email": None, "author_name": None,
        "editor_email": None, "editor_name": None, "version": "1.0", "checked_out_to": None, "prog_id": None,
        "server_relative_url": f"{base}/{rel_path}",
    }
    defaults.update(kw)
    return Item(**defaults)


def make_inv(items: list[Item], lib: LibraryRef | None = None, base: str | None = None):
    from m365tools.tools.spmig.models import Inventory

    lib = lib or make_lib()
    return Inventory(library=lib, base_url=base or lib.root_server_relative_url, items=items)
