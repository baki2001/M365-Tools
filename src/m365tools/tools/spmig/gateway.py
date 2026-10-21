"""The only door from spmig to SharePoint.

Every spmig module talks to a site through the small :class:`SpGateway` protocol below. The production adapter
:class:`RestGateway` delegates to :mod:`m365tools.core.sharepoint` (imported lazily so this module, and the tests
using an in-memory fake, never need that module or a network). Methods mirror the ``core.sharepoint`` functions
but take no client argument: a gateway is bound to one ``SpClient`` and one site.
"""

from __future__ import annotations

from collections.abc import AsyncIterator, Awaitable, Callable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Protocol, runtime_checkable

from m365tools.core.errors import AccessDeniedError, NotFoundError

from .models import Endpoint, Item, LibraryRef


class UserNotFoundError(NotFoundError):
    """Author/editor could not be resolved at the destination while setting metadata."""


@dataclass(slots=True)
class UploadResult:
    unique_id: str | None
    item_id: int | None


@dataclass(slots=True)
class RoleAssignment:
    principal_id: int
    principal_title: str
    principal_type: str
    role_names: list[str] = field(default_factory=list)
    can_write: bool = False


@dataclass(slots=True)
class FieldInfo:
    internal_name: str
    title: str = ""
    type_name: str = ""


@dataclass(slots=True)
class SiteUser:
    email: str
    title: str = ""
    id: int | None = None


@runtime_checkable
class SpGateway(Protocol):
    """SharePoint operations spmig needs, bound to one site (and one authenticated client)."""

    site_url: str

    async def get_default_library(self) -> LibraryRef: ...
    async def get_library(self, name: str) -> LibraryRef: ...
    def iter_items(self, lib: LibraryRef) -> AsyncIterator[dict[str, Any]]: ...
    def build_item(self, raw: dict[str, Any], lib: LibraryRef, base_url: str) -> Item | None: ...
    def iter_unique_permission_ids(
        self, lib: LibraryRef, on_page: Callable[[int], None] | None = None
    ) -> AsyncIterator[int]: ...
    async def ensure_folder(self, server_relative_folder: str) -> None: ...
    async def download_file(self, server_relative_url: str, dest_path: Path) -> None: ...
    async def upload_file(
        self, folder_server_relative_url: str, name: str, src_path: Path, overwrite: bool = True
    ) -> UploadResult: ...
    async def set_item_metadata(
        self,
        lib: LibraryRef,
        item_id: int,
        created: datetime | None,
        modified: datetime | None,
        author_email: str | None,
        editor_email: str | None,
    ) -> None: ...
    async def get_site_usage(self) -> tuple[int, int]: ...
    async def probe_write_access(self, folder_url: str) -> None: ...
    async def get_role_assignments(self, lib: LibraryRef) -> list[RoleAssignment]: ...
    async def get_associated_owner_group_id(self) -> int | None: ...
    async def has_unique_role_assignments(self, lib: LibraryRef) -> bool: ...
    async def break_inheritance(self, lib: LibraryRef, copy: bool = True) -> None: ...
    async def set_reader_only(self, lib: LibraryRef, principal_id: int) -> None: ...
    async def get_fields(self, lib: LibraryRef) -> list[FieldInfo]: ...
    async def find_site_user(self, email: str) -> SiteUser | None: ...


class GatewayFactory(Protocol):
    """Builds (and authenticates) a gateway for an endpoint. Supplied by the CLI/TUI; tests pass a fake."""

    def __call__(self, endpoint: Endpoint) -> Awaitable[SpGateway]: ...


# ---------------------------------------------------------------------------------------------------------------
# Production adapter
# ---------------------------------------------------------------------------------------------------------------

def _get(obj: Any, *names: str, default: Any = None) -> Any:
    """Read the first present attribute/key of ``obj`` (tolerates dataclass, dict or tuple-ish returns)."""
    for n in names:
        if isinstance(obj, dict):
            if n in obj:
                return obj[n]
        elif hasattr(obj, n):
            return getattr(obj, n)
    return default


def _is_user_not_found(exc: BaseException) -> bool:
    return any(c.__name__ == "UserNotFoundError" for c in type(exc).__mro__)


class RestGateway:
    """:class:`SpGateway` backed by ``m365tools.core.sharepoint`` + an ``SpClient``."""

    def __init__(self, client: Any, site_url: str) -> None:
        self.client = client
        self.site_url = site_url.rstrip("/")

    @staticmethod
    def _sp() -> Any:
        from m365tools.core import sharepoint  # lazy: keeps tests independent of the HTTP layer

        return sharepoint

    async def get_default_library(self) -> LibraryRef:
        return await self._sp().get_default_library(self.client, self.site_url)

    async def get_library(self, name: str) -> LibraryRef:
        return await self._sp().get_library(self.client, self.site_url, name)

    async def iter_items(self, lib: LibraryRef) -> AsyncIterator[dict[str, Any]]:
        async for raw in self._sp().iter_items(self.client, lib):
            yield raw

    def build_item(self, raw: dict[str, Any], lib: LibraryRef, base_url: str) -> Item | None:
        return self._sp().build_item(raw, base_url)

    async def iter_unique_permission_ids(
        self, lib: LibraryRef, on_page: Callable[[int], None] | None = None
    ) -> AsyncIterator[int]:
        async for item_id in self._sp().iter_unique_permission_ids(self.client, lib, on_page=on_page):
            yield int(item_id)

    async def ensure_folder(self, server_relative_folder: str) -> None:
        await self._sp().ensure_folder(self.client, self.site_url, server_relative_folder)

    async def download_file(self, server_relative_url: str, dest_path: Path) -> None:
        await self._sp().download_file(self.client, self.site_url, server_relative_url, dest_path)

    async def upload_file(
        self, folder_server_relative_url: str, name: str, src_path: Path, overwrite: bool = True
    ) -> UploadResult:
        res = await self._sp().upload_file(
            self.client, self.site_url, folder_server_relative_url, name, src_path, overwrite
        )
        if isinstance(res, tuple):
            uid, iid = (list(res) + [None, None])[:2]
        else:
            uid = _get(res, "unique_id", "UniqueId")
            iid = _get(res, "ItemId", "item_id", "id", "Id")
            if iid is None:
                laf = _get(res, "ListItemAllFields")
                iid = _get(laf, "Id", "ID") if laf is not None else None
        return UploadResult(unique_id=str(uid) if uid else None, item_id=int(iid) if iid is not None else None)

    async def set_item_metadata(
        self,
        lib: LibraryRef,
        item_id: int,
        created: datetime | None,
        modified: datetime | None,
        author_email: str | None,
        editor_email: str | None,
    ) -> None:
        try:
            await self._sp().set_item_metadata(
                self.client, lib, item_id, created, modified, author_email, editor_email
            )
        except Exception as exc:
            if _is_user_not_found(exc) and not isinstance(exc, UserNotFoundError):
                raise UserNotFoundError(str(exc)) from exc
            raise

    async def get_site_usage(self) -> tuple[int, int]:
        used, quota = await self._sp().get_site_usage(self.client, self.site_url)
        return int(used or 0), int(quota or 0)

    async def probe_write_access(self, folder_url: str) -> None:
        if not await self._sp().probe_write_access(self.client, self.site_url, folder_url):
            raise AccessDeniedError(f"Access denied creating a test folder in {folder_url}")

    async def get_role_assignments(self, lib: LibraryRef) -> list[RoleAssignment]:
        out: list[RoleAssignment] = []
        for ra in await self._sp().get_role_assignments(self.client, lib):
            out.append(
                RoleAssignment(
                    principal_id=int(_get(ra, "principal_id")),
                    principal_title=str(_get(ra, "principal_title", default="")),
                    principal_type=str(_get(ra, "principal_type", default="")),
                    role_names=list(_get(ra, "role_names", default=[]) or []),
                    can_write=bool(_get(ra, "can_write", default=False)),
                )
            )
        return out

    async def get_associated_owner_group_id(self) -> int | None:
        return await self._sp().get_associated_owner_group_id(self.client, self.site_url)

    async def has_unique_role_assignments(self, lib: LibraryRef) -> bool:
        fn = getattr(self._sp(), "has_unique_role_assignments", None)
        if fn is not None:
            return bool(await fn(self.client, lib))
        url = f"{lib.site_url.rstrip('/')}/_api/web/lists(guid'{lib.list_id}')?$select=HasUniqueRoleAssignments"
        data = await self.client.get_json(url)
        if isinstance(data, dict) and isinstance(data.get("d"), dict):
            data = data["d"]
        v = data.get("HasUniqueRoleAssignments") if isinstance(data, dict) else None
        return v is True or str(v).lower() == "true"

    async def break_inheritance(self, lib: LibraryRef, copy: bool = True) -> None:
        await self._sp().break_inheritance(self.client, lib, copy_role_assignments=copy)

    async def set_reader_only(self, lib: LibraryRef, principal_id: int) -> None:
        await self._sp().set_reader_only(self.client, lib, principal_id)

    async def get_fields(self, lib: LibraryRef) -> list[FieldInfo]:
        out: list[FieldInfo] = []
        for f in await self._sp().get_fields(self.client, lib):
            if isinstance(f, str):
                out.append(FieldInfo(internal_name=f))
                continue
            out.append(
                FieldInfo(
                    internal_name=str(_get(f, "internal_name", "InternalName", "name")),
                    title=str(_get(f, "title", "Title", default="") or ""),
                    type_name=str(_get(f, "type_name", "type_as_string", "TypeAsString", default="") or ""),
                )
            )
        return out

    async def find_site_user(self, email: str) -> SiteUser | None:
        u = await self._sp().find_site_user(self.client, self.site_url, email)
        if not u:
            return None
        if isinstance(u, str):
            return SiteUser(email=u)
        if u is True:
            return SiteUser(email=email)
        return SiteUser(
            email=str(_get(u, "email", "Email", default=email) or email),
            title=str(_get(u, "title", "Title", default="") or ""),
            id=_get(u, "id", "Id"),
        )
