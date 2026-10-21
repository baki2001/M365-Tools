"""SharePoint REST domain calls (see docs/ARCHITECTURE.md, "SharePoint API surface").

Every function is async and takes the SharePoint client first. The client is the ``SpClient`` from
``core/http.py``; this module only relies on the small surface described by ``_SpClientLike`` so it can be
tested with a fake.

Conventions:

* Responses are requested as ``application/json;odata=nometadata`` (the client default). Parsers also accept the
  ``odata=verbose`` shapes (``{"d": {...}}`` / ``{"d": {"results": [...]}}``) defensively.
* Server-relative paths are always passed through ``GetFileByServerRelativePath(decodedurl=...)`` /
  ``GetFolderByServerRelativePath(decodedurl=...)`` / ``AddUsingPath(DecodedUrl=...)`` so names containing
  ``'``, ``#``, ``%`` or spaces ("DAP's", "0. Archief") work. String literals have single quotes doubled
  (OData escaping) and are then percent-encoded with ``urllib.parse.quote(safe="/")``.
* Query strings are built here (spaces as ``%20``, never ``+``) and the complete URL is handed to the client.
"""

from __future__ import annotations

import asyncio
import contextlib
import uuid
from collections.abc import AsyncIterator, Callable, Iterable
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path, PurePosixPath
from typing import Any, Protocol
from urllib.parse import quote, urlencode, urlsplit

import httpx

from m365tools.core.errors import AccessDeniedError, M365Error, NotFoundError, ValidationError
from m365tools.tools.spmig.models import Item, LibraryRef

__all__ = [
    "ADD_LIST_ITEMS",
    "CHUNK_SIZE",
    "DEFAULT_ITEM_FIELDS",
    "DELETE_LIST_ITEMS",
    "EDIT_LIST_ITEMS",
    "MANAGE_LISTS",
    "SMALL_UPLOAD_LIMIT",
    "WRITE_MASK",
    "FieldInfo",
    "RoleAssignment",
    "SiteHit",
    "SiteUser",
    "UserNotFoundError",
    "WebInfo",
    "break_inheritance",
    "build_item",
    "download_file",
    "ensure_folder",
    "ensure_user",
    "find_site_user",
    "format_sp_datetime",
    "get_associated_owner_group_id",
    "get_default_library",
    "get_fields",
    "get_library",
    "get_role_assignments",
    "get_site_usage",
    "get_web",
    "has_unique_role_assignments",
    "iter_items",
    "iter_unique_permission_ids",
    "probe_write_access",
    "resolve_user",
    "search_sites",
    "set_item_metadata",
    "set_reader_only",
    "upload_file",
]

# ---------------------------------------------------------------------------------------------------------------
# Client protocol (http.py's SpClient satisfies this structurally)
# ---------------------------------------------------------------------------------------------------------------


class _SpClientLike(Protocol):
    async def request(
        self,
        method: str,
        url: str,
        *,
        json: Any = None,
        data: Any = None,
        content: Any = None,
        headers: dict[str, str] | None = None,
        params: dict[str, Any] | None = None,
        stream: bool = False,
    ) -> httpx.Response: ...

    async def get_json(self, url: str, **kw: Any) -> dict: ...

    async def post_json(self, url: str, json: Any = None, **kw: Any) -> dict: ...

    def iter_odata(self, url: str) -> AsyncIterator[dict]: ...

    async def download_to(self, url: str, dest: Path) -> Any: ...

    async def upload_bytes(self, url: str, data: bytes | Path, headers: dict[str, str] | None = None) -> Any:
        """SpClient returns parsed JSON; an ``httpx.Response`` is accepted too (see ``_body``)."""
        ...


# ---------------------------------------------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------------------------------------------

SMALL_UPLOAD_LIMIT = 10 * 1024 * 1024  # files <= this go through a single Files/AddUsingPath call
CHUNK_SIZE = 10 * 1024 * 1024

# SP.BasePermissions, Low DWORD
ADD_LIST_ITEMS = 0x2
EDIT_LIST_ITEMS = 0x4
DELETE_LIST_ITEMS = 0x8
MANAGE_LISTS = 0x800
WRITE_MASK = ADD_LIST_ITEMS | EDIT_LIST_ITEMS | DELETE_LIST_ITEMS | MANAGE_LISTS

READER_ROLE_TYPE_KIND = 2  # SP.RoleType.Reader

NOMETA = "application/json;odata=nometadata"
JSON_HEADERS = {"Accept": NOMETA, "Content-Type": NOMETA}

# Fields requested per list item (REST /items). User fields are projected through $expand.
# `OData__UIVersionString` is how the `_UIVersionString` field is exposed on the REST entity.
DEFAULT_ITEM_FIELDS: tuple[str, ...] = (
    "Id",
    "FileLeafRef",
    "FileRef",
    "FileDirRef",
    "FSObjType",
    "File_x0020_Size",
    "Created",
    "Modified",
    "Author/EMail",
    "Author/Title",
    "Editor/EMail",
    "Editor/Title",
    "OData__UIVersionString",
    "CheckoutUser/Title",
    "CheckoutUser/EMail",
    "ProgId",
)

_LIBRARY_SELECT = (
    "Id,Title,ItemCount,BaseTemplate,Hidden,ParentWebUrl,EnableVersioning,MajorVersionLimit,"
    "EnableMinorVersions,ForceCheckout,EnableModeration,RootFolder/ServerRelativeUrl,RootFolder/Name"
)


# ---------------------------------------------------------------------------------------------------------------
# Result types
# ---------------------------------------------------------------------------------------------------------------


@dataclass(slots=True)
class SiteHit:
    title: str
    url: str
    description: str = ""


@dataclass(slots=True)
class WebInfo:
    server_relative_url: str
    title: str
    language: int


@dataclass(slots=True)
class SiteUser:
    id: int
    title: str
    email: str
    login_name: str


@dataclass(slots=True)
class FieldInfo:
    internal_name: str
    title: str
    type_as_string: str
    read_only: bool = False
    group: str = ""


@dataclass(slots=True)
class RoleAssignment:
    principal_id: int
    principal_title: str
    principal_type: int  # SP.PrincipalType: 1 user, 2 distribution list, 4 security group, 8 SharePoint group
    role_names: list[str]
    can_write: bool
    login_name: str = ""
    role_ids: list[int] = field(default_factory=list)


class UserNotFoundError(M365Error):
    """ValidateUpdateListItem could not resolve the Author/Editor user. Caller should retry with dates only."""

    def __init__(self, field: str, email: str | None, message: str) -> None:
        super().__init__(f"{field}: user {email!r} could not be set: {message}")
        self.field = field
        self.email = email
        self.message = message


# ---------------------------------------------------------------------------------------------------------------
# URL / response helpers
# ---------------------------------------------------------------------------------------------------------------


def _api(site_url: str, path: str) -> str:
    return f"{site_url.rstrip('/')}/_api/{path.lstrip('/')}"


def _list_path(lib: LibraryRef, rest: str = "") -> str:
    base = f"web/lists(guid'{lib.list_id.strip('{}')}')"
    return f"{base}/{rest}" if rest else base


def _lit(value: str) -> str:
    """OData string literal for use inside a URL *path* segment: quote-doubled then percent-encoded."""
    return "'" + quote(value.replace("'", "''"), safe="/") + "'"


def _q(value: str) -> str:
    """OData string literal for use inside a query option ($filter); encoding happens in _url()."""
    return "'" + value.replace("'", "''") + "'"


def _url(base: str, params: dict[str, Any] | None = None) -> str:
    if not params:
        return base
    qs = urlencode({k: str(v) for k, v in params.items()}, quote_via=quote, safe="$,/'()")
    return f"{base}{'&' if '?' in base else '?'}{qs}"


def _unwrap(data: Any) -> Any:
    """Strip the odata=verbose ``{"d": ...}`` envelope if present."""
    if isinstance(data, dict) and "d" in data and len(data) == 1:
        return data["d"]
    return data


def _values(data: Any) -> list[dict]:
    data = _unwrap(data)
    if isinstance(data, list):
        return data
    if not isinstance(data, dict):
        return []
    if isinstance(data.get("value"), list):
        return data["value"]
    if isinstance(data.get("results"), list):
        return data["results"]
    return []


def _results(value: Any) -> list:
    """A collection-valued property: nometadata gives a list, verbose gives {"results": [...]}."""
    if isinstance(value, list):
        return value
    if isinstance(value, dict) and isinstance(value.get("results"), list):
        return value["results"]
    return []


def _next_link(data: Any) -> str | None:
    if not isinstance(data, dict):
        return None
    d = data.get("d") if isinstance(data.get("d"), dict) else {}
    return data.get("odata.nextLink") or data.get("@odata.nextLink") or d.get("__next")


def _server_relative_path(url: str) -> str:
    """'/sites/x' from 'https://t.sharepoint.com/sites/x/' ; '' for the root site."""
    return urlsplit(url).path.rstrip("/")


def _join(*parts: str) -> str:
    out = "/".join(p.strip("/") for p in parts if p and p.strip("/"))
    return "/" + out


def _bool(v: Any) -> bool:
    if isinstance(v, str):
        return v.strip().lower() in {"1", "true", "yes"}
    return bool(v)


def _int(v: Any, default: int = 0) -> int:
    if v is None or v == "":
        return default
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def _parse_dt(v: Any) -> datetime | None:
    """Parse SharePoint REST datetimes ('2024-01-31T10:00:00Z') to aware UTC. Naive values are taken as UTC."""
    if not v or not isinstance(v, str):
        return None
    s = v.strip()
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        return None
    if dt.tzinfo is None:
        return dt.replace(tzinfo=UTC)
    return dt.astimezone(UTC)


def format_sp_datetime(dt: datetime) -> str:
    """Format a datetime for ValidateUpdateListItem formValues: ISO-8601 UTC, e.g. '2024-01-31T10:00:00Z'.

    ASSUMPTION (verify on tenant): ValidateUpdateListItem accepts ISO-8601 with 'Z' independent of the web's
    regional settings. If it turns out to be parsed per regional settings, pass ``date_format=`` to
    ``set_item_metadata`` with a locale-specific formatter (e.g. '%d-%m-%Y %H:%M' in the web's time zone).
    """
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def _body(resp: Any) -> dict:
    """JSON body of an upload result, whether the client returned parsed JSON or an httpx.Response."""
    if isinstance(resp, httpx.Response):
        if not resp.content:
            return {}
        try:
            resp = resp.json()
        except ValueError:
            return {}
    data = _unwrap(resp)
    return data if isinstance(data, dict) else {}


def _error_text(exc: BaseException) -> str:
    return str(exc).lower()


# ---------------------------------------------------------------------------------------------------------------
# Sites / webs / libraries
# ---------------------------------------------------------------------------------------------------------------


async def search_sites(
    client: _SpClientLike,
    query: str,
    *,
    root_url: str,
    graph_client: Any | None = None,
    row_limit: int = 50,
) -> list[SiteHit]:
    """Search site collections via SharePoint search; fall back to Graph ``/sites?search=`` if given.

    ``root_url`` is the tenant root (``https://<tenant>.sharepoint.com``). ``graph_client`` is any object with
    ``get_json(url)`` authorised for Graph; it is used when SharePoint search fails or returns nothing.
    """
    q = query.strip()
    text = f"contentclass:STS_Site {q}*" if q else "contentclass:STS_Site"
    url = _url(
        _api(root_url, "search/query"),
        {
            "querytext": _q(text),
            "rowlimit": row_limit,
            "selectproperties": _q("Title,Path,Description"),
            "trimduplicates": "false",
        },
    )
    hits: list[SiteHit] = []
    sp_error: M365Error | None = None
    try:
        data = _unwrap(await client.get_json(url))
        query_data = data.get("query", data) if isinstance(data, dict) else {}
        table = (
            (query_data.get("PrimaryQueryResult") or {}).get("RelevantResults", {}).get("Table", {})
        )
        for row in _results(table.get("Rows")):
            cells = {c.get("Key"): c.get("Value") for c in _results(row.get("Cells"))}
            if cells.get("Path"):
                hits.append(
                    SiteHit(
                        title=cells.get("Title") or "",
                        url=str(cells["Path"]).rstrip("/"),
                        description=cells.get("Description") or "",
                    )
                )
    except M365Error as exc:
        sp_error = exc
    if hits or graph_client is None:
        if sp_error is not None and not hits:
            raise sp_error
        return hits
    gdata = await graph_client.get_json(
        _url("https://graph.microsoft.com/v1.0/sites", {"search": q or "*", "$select": "displayName,webUrl,description"})
    )
    return [
        SiteHit(title=s.get("displayName") or "", url=(s.get("webUrl") or "").rstrip("/"),
                description=s.get("description") or "")
        for s in _values(gdata)
        if s.get("webUrl")
    ]


async def get_web(client: _SpClientLike, site_url: str) -> WebInfo:
    data = _unwrap(await client.get_json(_url(_api(site_url, "web"), {"$select": "ServerRelativeUrl,Title,Language"})))
    return WebInfo(
        server_relative_url=data.get("ServerRelativeUrl") or _server_relative_path(site_url) or "/",
        title=data.get("Title") or "",
        language=_int(data.get("Language")),
    )


def _library_from_json(site_url: str, d: dict) -> LibraryRef:
    d = _unwrap(d)
    root = d.get("RootFolder") or {}
    web_url = d.get("ParentWebUrl") or _server_relative_path(site_url) or "/"
    return LibraryRef(
        site_url=site_url.rstrip("/"),
        web_server_relative_url=web_url,
        list_id=str(d.get("Id") or "").strip("{}"),
        title=d.get("Title") or "",
        root_server_relative_url=(root.get("ServerRelativeUrl") or "").rstrip("/"),
        item_count=_int(d.get("ItemCount")),
        enable_versioning=_bool(d.get("EnableVersioning")),
        major_version_limit=_int(d.get("MajorVersionLimit")),
        enable_minor_versions=_bool(d.get("EnableMinorVersions")),
        force_checkout=_bool(d.get("ForceCheckout")),
        enable_moderation=_bool(d.get("EnableModeration")),
    )


async def get_default_library(client: _SpClientLike, site_url: str) -> LibraryRef:
    """The site's default document library (language independent: 'Documents', 'Gedeelde documenten', ...)."""
    url = _url(_api(site_url, "web/DefaultDocumentLibrary"), {"$select": _LIBRARY_SELECT, "$expand": "RootFolder"})
    return _library_from_json(site_url, await client.get_json(url))


async def get_library(client: _SpClientLike, site_url: str, name: str | None) -> LibraryRef:
    """Resolve a document library by Title, then by RootFolder name (URL segment), case-insensitive.

    ``name`` None/empty returns the default library. Raises NotFoundError with ``available`` set to the titles of
    all visible document libraries (BaseTemplate 101, not hidden).
    """
    if not name or not name.strip():
        return await get_default_library(client, site_url)
    url = _url(_api(site_url, "web/lists"), {"$select": _LIBRARY_SELECT, "$expand": "RootFolder"})
    libs = [
        d async for d in client.iter_odata(url)
        if _int(d.get("BaseTemplate")) == 101 and not _bool(d.get("Hidden"))
    ]
    wanted = name.strip().strip("/").lower()
    for d in libs:
        if (d.get("Title") or "").lower() == wanted:
            return _library_from_json(site_url, d)
    for d in libs:
        root = d.get("RootFolder") or {}
        root_name = root.get("Name") or (root.get("ServerRelativeUrl") or "").rstrip("/").rsplit("/", 1)[-1]
        if root_name.lower() == wanted or (root.get("ServerRelativeUrl") or "").strip("/").lower() == wanted:
            return _library_from_json(site_url, d)
    available = sorted(d.get("Title") or "" for d in libs)
    raise NotFoundError(
        f"Library {name!r} not found on {site_url}. Available: {', '.join(available) or '(none)'}",
        available=available,
    )


# ---------------------------------------------------------------------------------------------------------------
# Items
# ---------------------------------------------------------------------------------------------------------------


async def _iter_keyset(
    client: _SpClientLike,
    base: str,
    select: Iterable[str],
    expand: Iterable[str] = (),
    page_size: int = 2000,
) -> AsyncIterator[dict]:
    """Page a list's /items with ``$orderby=ID&$filter=ID gt <last>&$top=N`` (keyset paging).

    ID is always indexed, so each page stays under the list view threshold regardless of library size, and a
    page that fails (throttling) can be retried from a known position.
    """
    select_s = ",".join(select)
    expand_s = ",".join(expand)
    last_id = 0
    while True:
        params: dict[str, Any] = {"$select": select_s}
        if expand_s:
            params["$expand"] = expand_s
        params.update({"$filter": f"ID gt {last_id}", "$orderby": "ID", "$top": page_size})
        data = await client.get_json(_url(base, params))
        rows = _values(data)
        if not rows:
            return
        for row in rows:
            yield row
        ids = [_int(r.get("Id", r.get("ID"))) for r in rows]
        new_last = max(ids) if ids else last_id
        if new_last <= last_id:  # defensive: never loop on a server that ignores $filter
            return
        last_id = new_last
        if len(rows) < page_size and not _next_link(data):
            return


async def iter_items(
    client: _SpClientLike,
    lib: LibraryRef,
    fields: Iterable[str] | None = None,
    *,
    page_size: int = 2000,
) -> AsyncIterator[dict]:
    """Yield every item (files and folders, all depths) of the library as raw REST rows.

    Uses ``/_api/web/lists(guid'..')/items`` with keyset paging on ID (>5000 safe) instead of GetItems/CAML,
    because /items can ``$expand`` the user lookups (Author/Editor/CheckoutUser) in the same request.
    ``Id`` is always selected. Convert rows with ``build_item``.
    """
    flds = list(fields or DEFAULT_ITEM_FIELDS)
    if "Id" not in flds and "ID" not in flds:
        flds.insert(0, "Id")
    expand = sorted({f.split("/", 1)[0] for f in flds if "/" in f})
    async for row in _iter_keyset(client, _api(lib.site_url, _list_path(lib, "items")), flds, expand, page_size):
        yield row


async def iter_unique_permission_ids(
    client: _SpClientLike, lib: LibraryRef, *, page_size: int = 5000
) -> AsyncIterator[int]:
    """Yield item ids that have unique (broken) role assignments."""
    base = _api(lib.site_url, _list_path(lib, "items"))
    async for row in _iter_keyset(client, base, ["Id", "HasUniqueRoleAssignments"], (), page_size):
        if _bool(row.get("HasUniqueRoleAssignments")):
            yield _int(row.get("Id", row.get("ID")))


def _user(row: dict, key: str) -> tuple[str | None, str | None]:
    u = row.get(key)
    if not isinstance(u, dict):
        return None, None
    email = u.get("EMail") or u.get("Email") or None
    title = u.get("Title") or None
    return email, title


def build_item(row: dict, base_url: str) -> Item | None:
    """Convert an /items row to an Item. ``rel_path`` is relative to ``base_url`` (server-relative, compared
    case-insensitively); returns None for the base folder itself or anything outside it."""
    file_ref = (row.get("FileRef") or "").rstrip("/")
    base = "/" + base_url.strip("/") if base_url.strip("/") else ""
    if not file_ref:
        return None
    prefix = base + "/"
    if not file_ref.lower().startswith(prefix.lower()):
        return None
    rel_path = file_ref[len(prefix):]
    if not rel_path:
        return None
    is_folder = _int(row.get("FSObjType")) == 1
    name = row.get("FileLeafRef") or rel_path.rsplit("/", 1)[-1]
    ext = "" if is_folder else PurePosixPath(name).suffix.lower()
    size = 0
    if not is_folder:
        file_obj = row.get("File") if isinstance(row.get("File"), dict) else {}
        size = _int(row.get("File_x0020_Size", row.get("SMTotalFileStreamSize", file_obj.get("Length"))))
    author_email, author_name = _user(row, "Author")
    editor_email, editor_name = _user(row, "Editor")
    _co_email, co_title = _user(row, "CheckoutUser")
    meta = row.get("__metadata") if isinstance(row.get("__metadata"), dict) else {}
    return Item(
        id=_int(row.get("Id", row.get("ID"))),
        is_folder=is_folder,
        rel_path=rel_path,
        name=name,
        ext=ext,
        size=size,
        created=_parse_dt(row.get("Created")),
        modified=_parse_dt(row.get("Modified")),
        author_email=author_email,
        author_name=author_name,
        editor_email=editor_email,
        editor_name=editor_name,
        version=row.get("OData__UIVersionString") or row.get("_UIVersionString") or None,
        checked_out_to=co_title or _co_email,
        prog_id=row.get("ProgId") or None,
        server_relative_url=file_ref,
        unique_permissions=_bool(row.get("HasUniqueRoleAssignments")),
        etag=row.get("odata.etag") or row.get("@odata.etag") or meta.get("etag"),
    )


# ---------------------------------------------------------------------------------------------------------------
# Folders / files
# ---------------------------------------------------------------------------------------------------------------


async def _folder_exists(client: _SpClientLike, site_url: str, server_relative_url: str) -> bool:
    url = _url(
        _api(site_url, f"web/GetFolderByServerRelativePath(decodedurl={_lit(server_relative_url)})"),
        {"$select": "Exists"},
    )
    try:
        data = _unwrap(await client.get_json(url))
    except NotFoundError:
        return False
    return _bool(data.get("Exists", True))


async def _post_new_folder(client: _SpClientLike, site_url: str, server_relative_url: str) -> None:
    """Create one folder. ``POST web/folders {"ServerRelativeUrl"}`` is a URL-string API that does not support
    ``#`` / ``%`` in names, so such paths go through the ResourcePath variant ``web/folders/AddUsingPath``."""
    if "#" in server_relative_url or "%" in server_relative_url:
        await client.request(
            "POST",
            _api(site_url, f"web/folders/AddUsingPath(DecodedUrl={_lit(server_relative_url)},Overwrite=false)"),
            headers={"Accept": NOMETA},
        )
        return
    await client.post_json(
        _api(site_url, "web/folders"), json={"ServerRelativeUrl": server_relative_url}, headers=JSON_HEADERS
    )


async def _create_folder(client: _SpClientLike, site_url: str, server_relative_url: str) -> None:
    try:
        await _post_new_folder(client, site_url, server_relative_url)
    except (NotFoundError, AccessDeniedError):
        raise
    except M365Error as exc:
        if "already exists" in _error_text(exc) or "-2130245363" in str(exc):
            return
        raise


async def ensure_folder(client: _SpClientLike, site_url: str, server_relative_folder: str) -> None:
    """Make sure the folder (and its parents) exist. Idempotent.

    Checks the full path first; if missing, ensures the parent recursively, then creates the leaf with
    ``POST /_api/web/folders {"ServerRelativeUrl": ...}``. Never creates the web itself or a folder directly
    under the web (that would be a library root): raises NotFoundError instead.
    """
    target = _join(server_relative_folder)
    web_path = _server_relative_path(site_url)
    if await _folder_exists(client, site_url, target):
        return
    parent = target.rsplit("/", 1)[0]
    if parent.rstrip("/").lower() == web_path.lower() or len(target) <= len(web_path) + 1:
        raise NotFoundError(f"Library root folder {target!r} does not exist on {site_url}")
    if not target.lower().startswith(web_path.lower() + "/"):
        raise ValidationError(f"Folder {target!r} is not inside web {web_path or '/'}")
    await ensure_folder(client, site_url, parent)
    await _create_folder(client, site_url, target)


def _file_path(site_url: str, server_relative_url: str, rest: str = "") -> str:
    p = f"web/GetFileByServerRelativePath(decodedurl={_lit(server_relative_url)})"
    return _api(site_url, f"{p}/{rest}" if rest else p)


async def download_file(client: _SpClientLike, site_url: str, server_relative_url: str, dest_path: Path) -> None:
    """Stream a file's content to ``dest_path``."""
    await client.download_to(_file_path(site_url, server_relative_url, "$value"), Path(dest_path))


async def _file_item_id(client: _SpClientLike, site_url: str, server_relative_url: str) -> int | None:
    data = _unwrap(
        await client.get_json(_url(_file_path(site_url, server_relative_url, "ListItemAllFields"), {"$select": "Id"}))
    )
    v = data.get("Id", data.get("ID"))
    return _int(v) if v is not None else None


def _read_chunk(path: Path, offset: int, size: int) -> bytes:
    with path.open("rb") as fh:
        fh.seek(offset)
        return fh.read(size)


async def upload_file(
    client: _SpClientLike,
    site_url: str,
    folder_server_relative_url: str,
    name: str,
    src_path: Path,
    overwrite: bool = True,
    *,
    chunk_size: int = CHUNK_SIZE,
    small_limit: int = SMALL_UPLOAD_LIMIT,
) -> dict[str, Any]:
    """Upload ``src_path`` as ``<folder>/<name>``.

    <= ``small_limit`` bytes: one ``Files/AddUsingPath`` call with the raw bytes. Larger: create an empty file
    with AddUsingPath, then ``StartUpload`` / ``ContinueUpload`` / ``FinishUpload`` in ``chunk_size`` chunks
    (``CancelUpload`` on failure). Returns ``{"UniqueId", "ServerRelativeUrl", "Length", "ItemId"}``.
    """
    src_path = Path(src_path)
    size = src_path.stat().st_size
    folder = _join(folder_server_relative_url)
    file_url = f"{folder}/{name}"
    ow = "true" if overwrite else "false"
    add_url = _api(
        site_url,
        f"web/GetFolderByServerRelativePath(decodedurl={_lit(folder)})"
        f"/Files/AddUsingPath(DecodedUrl={_lit(name)},Overwrite={ow})",
    )
    octet = {"Content-Type": "application/octet-stream", "Accept": NOMETA}

    if size <= small_limit:
        resp = await client.upload_bytes(add_url, src_path, headers=octet)
        info = _body(resp)
    else:
        await client.upload_bytes(add_url, b"", headers=octet)
        upload_id = str(uuid.uuid4())
        offset = 0
        info = {}
        try:
            while offset < size:
                chunk = await asyncio.to_thread(_read_chunk, src_path, offset, chunk_size)
                if not chunk:
                    raise M365Error(f"Unexpected end of file reading {src_path} at {offset}/{size}")
                if offset == 0:
                    op = f"StartUpload(uploadId=guid'{upload_id}')"
                elif offset + len(chunk) >= size:
                    op = f"FinishUpload(uploadId=guid'{upload_id}',fileOffset={offset})"
                else:
                    op = f"ContinueUpload(uploadId=guid'{upload_id}',fileOffset={offset})"
                resp = await client.upload_bytes(_file_path(site_url, file_url, op), chunk, headers=octet)
                offset += len(chunk)
                if op.startswith("FinishUpload"):
                    info = _body(resp)
        except BaseException:
            with contextlib.suppress(Exception):  # best effort cleanup, the original error wins
                await client.request(
                    "POST", _file_path(site_url, file_url, f"CancelUpload(uploadId=guid'{upload_id}')")
                )
            raise

    server_relative_url = info.get("ServerRelativeUrl") or file_url
    item_id = await _file_item_id(client, site_url, server_relative_url)
    return {
        "UniqueId": info.get("UniqueId"),
        "ServerRelativeUrl": server_relative_url,
        "Length": _int(info.get("Length"), size),
        "ItemId": item_id,
    }


# ---------------------------------------------------------------------------------------------------------------
# Metadata
# ---------------------------------------------------------------------------------------------------------------


def _claims(email: str) -> str:
    return f'[{{"Key":"i:0#.f|membership|{email}"}}]'


async def set_item_metadata(
    client: _SpClientLike,
    lib: LibraryRef,
    item_id: int,
    created: datetime | None = None,
    modified: datetime | None = None,
    author_email: str | None = None,
    editor_email: str | None = None,
    *,
    date_format: Callable[[datetime], str] = format_sp_datetime,
) -> list[dict]:
    """Set Created/Modified/Author/Editor without creating a new version (ValidateUpdateListItem,
    bNewDocumentUpdate=true). Only non-None values are sent.

    Raises UserNotFoundError when Author or Editor is rejected (deleted/unknown user) — the caller should retry
    with dates only and count a metadata warning. Raises M365Error for other per-field errors. Returns the
    per-field result list.
    """
    form: list[dict[str, str]] = []
    if created is not None:
        form.append({"FieldName": "Created", "FieldValue": date_format(created)})
    if modified is not None:
        form.append({"FieldName": "Modified", "FieldValue": date_format(modified)})
    if author_email:
        form.append({"FieldName": "Author", "FieldValue": _claims(author_email)})
    if editor_email:
        form.append({"FieldName": "Editor", "FieldValue": _claims(editor_email)})
    if not form:
        return []
    url = _api(lib.site_url, _list_path(lib, f"items({int(item_id)})/ValidateUpdateListItem"))
    data = await client.post_json(
        url, json={"formValues": form, "bNewDocumentUpdate": True}, headers=JSON_HEADERS
    )
    data = _unwrap(data)
    if isinstance(data, dict) and "ValidateUpdateListItem" in data:
        data = data["ValidateUpdateListItem"]
    results = _values(data)
    emails = {"Author": author_email, "Editor": editor_email}
    errors = [r for r in results if _bool(r.get("HasException")) or r.get("ErrorMessage")]
    for r in errors:
        fname = r.get("FieldName") or ""
        if fname in emails:
            raise UserNotFoundError(fname, emails[fname], r.get("ErrorMessage") or "unknown error")
    if errors:
        detail = "; ".join(f"{r.get('FieldName')}: {r.get('ErrorMessage')}" for r in errors)
        raise M365Error(f"Metadata update of item {item_id} failed: {detail}")
    return results


# ---------------------------------------------------------------------------------------------------------------
# Quota / access probe
# ---------------------------------------------------------------------------------------------------------------


async def get_site_usage(client: _SpClientLike, site_url: str) -> tuple[int, int | None]:
    """(used_bytes, quota_bytes). Quota is derived from Storage / StoragePercentageUsed (a 0..1 fraction);
    None when it cannot be derived (e.g. empty site)."""
    data = _unwrap(await client.get_json(_api(site_url, "site/Usage")))
    usage = data.get("Usage", data) if isinstance(data, dict) else {}
    used = _int(usage.get("Storage"))
    try:
        pct = float(usage.get("StoragePercentageUsed") or 0)
    except (TypeError, ValueError):
        pct = 0.0
    quota = round(used / pct) if pct > 0 and used > 0 else None
    return used, quota


async def probe_write_access(client: _SpClientLike, site_url: str, folder_url: str) -> bool:
    """Create and delete ``<folder>/__m365_probe_<uuid>``. True if allowed, False on access denied."""
    probe = f"{_join(folder_url)}/__m365_probe_{uuid.uuid4().hex}"
    try:
        await _post_new_folder(client, site_url, probe)
    except AccessDeniedError:
        return False
    await client.request(
        "POST",
        _api(site_url, f"web/GetFolderByServerRelativePath(decodedurl={_lit(probe)})"),
        headers={"X-HTTP-Method": "DELETE", "IF-MATCH": "*", "Accept": NOMETA},
    )
    return True


# ---------------------------------------------------------------------------------------------------------------
# Permissions
# ---------------------------------------------------------------------------------------------------------------


def _low_bits(perm: Any) -> int:
    if not isinstance(perm, dict):
        return 0
    return _int(perm.get("Low"))


async def get_role_assignments(client: _SpClientLike, lib: LibraryRef) -> list[RoleAssignment]:
    url = _url(_api(lib.site_url, _list_path(lib, "roleassignments")), {"$expand": "Member,RoleDefinitionBindings"})
    out: list[RoleAssignment] = []
    async for ra in client.iter_odata(url):
        member = ra.get("Member") or {}
        bindings = _results(ra.get("RoleDefinitionBindings"))
        low = 0
        for b in bindings:
            low |= _low_bits(b.get("BasePermissions"))
        out.append(
            RoleAssignment(
                principal_id=_int(ra.get("PrincipalId", member.get("Id"))),
                principal_title=member.get("Title") or "",
                principal_type=_int(member.get("PrincipalType")),
                role_names=[b.get("Name") or "" for b in bindings],
                can_write=bool(low & WRITE_MASK),
                login_name=member.get("LoginName") or "",
                role_ids=[_int(b.get("Id")) for b in bindings],
            )
        )
    return out


async def get_associated_owner_group_id(client: _SpClientLike, site_url: str) -> int | None:
    try:
        data = _unwrap(await client.get_json(_url(_api(site_url, "web/AssociatedOwnerGroup"), {"$select": "Id"})))
    except NotFoundError:
        return None
    v = data.get("Id") if isinstance(data, dict) else None
    return _int(v) if v is not None else None


async def has_unique_role_assignments(client: _SpClientLike, lib: LibraryRef) -> bool:
    data = _unwrap(
        await client.get_json(_url(_api(lib.site_url, _list_path(lib)), {"$select": "HasUniqueRoleAssignments"}))
    )
    return isinstance(data, dict) and _bool(data.get("HasUniqueRoleAssignments"))


async def break_inheritance(
    client: _SpClientLike, lib: LibraryRef, copy_role_assignments: bool = True, clear_subscopes: bool = False
) -> None:
    c = "true" if copy_role_assignments else "false"
    s = "true" if clear_subscopes else "false"
    await client.request(
        "POST",
        _api(lib.site_url, _list_path(lib, f"breakroleinheritance(copyRoleAssignments={c},clearSubscopes={s})")),
        headers={"Accept": NOMETA},
    )


async def _reader_role_id(client: _SpClientLike, site_url: str) -> int:
    url = _url(_api(site_url, "web/roledefinitions"), {"$select": "Id,Name,RoleTypeKind"})
    async for rd in client.iter_odata(url):
        if _int(rd.get("RoleTypeKind")) == READER_ROLE_TYPE_KIND:
            return _int(rd.get("Id"))
    raise NotFoundError(f"No Reader role definition (RoleTypeKind 2) on {site_url}")


async def set_reader_only(client: _SpClientLike, lib: LibraryRef, principal_id: int) -> None:
    """Downgrade a principal on the library to Read. Never widens access:

    * the library must already have unique permissions (call break_inheritance first), else ValidationError;
    * the principal must currently hold a write-capable binding on the library, else ValidationError.

    Reader is added first (if not present) and then every other binding is removed, so there is no window in
    which the principal has no access at all.
    """
    if not await has_unique_role_assignments(client, lib):
        raise ValidationError(f"Library {lib.title!r} inherits permissions; break inheritance before locking")
    current = next((ra for ra in await get_role_assignments(client, lib) if ra.principal_id == principal_id), None)
    if current is None or not current.can_write:
        raise ValidationError(
            f"Principal {principal_id} has no write access on {lib.title!r}; refusing to grant Read (never widen)"
        )
    reader_id = await _reader_role_id(client, lib.site_url)
    ra_base = _list_path(lib, "roleassignments")
    if reader_id not in current.role_ids:
        await client.request(
            "POST",
            _api(lib.site_url, f"{ra_base}/addroleassignment(principalid={principal_id},roledefid={reader_id})"),
            headers={"Accept": NOMETA},
        )
    for role_id in current.role_ids:
        if role_id == reader_id:
            continue
        await client.request(
            "POST",
            _api(lib.site_url, f"{ra_base}/removeroleassignment(principalid={principal_id},roledefid={role_id})"),
            headers={"Accept": NOMETA},
        )


# ---------------------------------------------------------------------------------------------------------------
# Fields / users
# ---------------------------------------------------------------------------------------------------------------


async def get_fields(client: _SpClientLike, lib: LibraryRef) -> list[FieldInfo]:
    """Custom columns: not hidden, not FromBaseType, CanBeDeleted."""
    url = _url(
        _api(lib.site_url, _list_path(lib, "fields")),
        {"$select": "InternalName,Title,TypeAsString,Hidden,FromBaseType,CanBeDeleted,ReadOnlyField,Group"},
    )
    out: list[FieldInfo] = []
    async for f in client.iter_odata(url):
        if _bool(f.get("Hidden")) or _bool(f.get("FromBaseType")) or not _bool(f.get("CanBeDeleted", True)):
            continue
        out.append(
            FieldInfo(
                internal_name=f.get("InternalName") or "",
                title=f.get("Title") or "",
                type_as_string=f.get("TypeAsString") or "",
                read_only=_bool(f.get("ReadOnlyField")),
                group=f.get("Group") or "",
            )
        )
    return out


def _site_user(d: dict) -> SiteUser:
    d = _unwrap(d)
    return SiteUser(
        id=_int(d.get("Id")),
        title=d.get("Title") or "",
        email=d.get("Email") or d.get("EMail") or "",
        login_name=d.get("LoginName") or "",
    )


async def find_site_user(client: _SpClientLike, site_url: str, email: str) -> SiteUser | None:
    """Read-only lookup in the web's User Information List by email (no side effects)."""
    url = _url(
        _api(site_url, "web/siteusers"),
        {"$filter": f"Email eq {_q(email.strip())}", "$select": "Id,Title,Email,LoginName"},
    )
    rows = _values(await client.get_json(url))
    return _site_user(rows[0]) if rows else None


async def ensure_user(client: _SpClientLike, site_url: str, email: str) -> SiteUser:
    """``POST /_api/web/ensureuser`` — adds the user to the site's user list if needed (a write)."""
    data = await client.post_json(
        _api(site_url, "web/ensureuser"), json={"logonName": email.strip()}, headers=JSON_HEADERS
    )
    return _site_user(data)


async def resolve_user(client: _SpClientLike, site_url: str, email: str, *, ensure: bool = False) -> bool:
    """True if the user exists on the site (siteusers lookup); with ``ensure`` try ensureuser as a fallback."""
    if await find_site_user(client, site_url, email) is not None:
        return True
    if not ensure:
        return False
    try:
        await ensure_user(client, site_url, email)
    except (NotFoundError, AccessDeniedError):
        return False
    except M365Error:
        return False
    return True
