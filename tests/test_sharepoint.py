from __future__ import annotations

import json
from datetime import UTC, datetime
from pathlib import Path

import httpx
import pytest
import respx

from m365tools.core import sharepoint as sp
from m365tools.core.errors import NotFoundError, ValidationError
from m365tools.tools.spmig.models import LibraryRef

SITE = "https://contoso.sharepoint.com/sites/Docs"
LIST_ID = "11111111-2222-3333-4444-555555555555"
LIST_API = f"{SITE}/_api/web/lists(guid'{LIST_ID}')"
LIB_ROOT = "/sites/Docs/Gedeelde documenten"


def lib() -> LibraryRef:
    return LibraryRef(
        site_url=SITE, web_server_relative_url="/sites/Docs", list_id=LIST_ID, title="Documenten",
        root_server_relative_url=LIB_ROOT, item_count=3,
    )


def raw(request: httpx.Request) -> str:
    """Path + query exactly as sent on the wire (percent-encoded)."""
    return request.url.raw_path.decode()


def lib_json(title: str, root: str, **extra) -> dict:
    return {
        "Id": LIST_ID, "Title": title, "ItemCount": 1210, "BaseTemplate": 101, "Hidden": False,
        "ParentWebUrl": "/sites/Docs", "EnableVersioning": True, "MajorVersionLimit": 500,
        "EnableMinorVersions": False, "ForceCheckout": False, "EnableModeration": False,
        "RootFolder": {"ServerRelativeUrl": root, "Name": root.rsplit("/", 1)[-1]}, **extra,
    }


# ------------------------------------------------------------------------------------------------ encoding


def test_literal_encoding() -> None:
    assert sp._lit("/sites/x/0. Archief/DAP's #1 100%") == "'/sites/x/0.%20Archief/DAP%27%27s%20%231%20100%25'"
    url = sp._url("https://h/_api/web/siteusers", {"$filter": f"Email eq {sp._q("o'brien@x.nl")}"})
    assert url == "https://h/_api/web/siteusers?$filter=Email%20eq%20'o''brien%40x.nl'"


# ------------------------------------------------------------------------------------------------ libraries


async def test_get_default_library(fake_client, respx_mock: respx.MockRouter) -> None:
    route = respx_mock.get(url__startswith=f"{SITE}/_api/web/DefaultDocumentLibrary").respond(
        json=lib_json("Documenten", LIB_ROOT)
    )
    ref = await sp.get_default_library(fake_client, SITE + "/")
    assert ref.title == "Documenten"
    assert ref.list_id == LIST_ID
    assert ref.root_server_relative_url == LIB_ROOT
    assert ref.root_name == "Gedeelde documenten"
    assert ref.web_server_relative_url == "/sites/Docs"
    assert ref.site_url == SITE
    assert ref.item_count == 1210 and ref.enable_versioning and ref.major_version_limit == 500
    q = route.calls.last.request.url.params
    assert q["$expand"] == "RootFolder" and "RootFolder/ServerRelativeUrl" in q["$select"]


async def test_get_library_by_title_then_root_name(fake_client, respx_mock: respx.MockRouter) -> None:
    respx_mock.get(url__startswith=f"{SITE}/_api/web/lists").respond(
        json={"value": [
            lib_json("Documenten", LIB_ROOT),
            lib_json("Service management", "/sites/Docs/Servicemanagement"),
            lib_json("Hidden lib", "/sites/Docs/Hidden", Hidden=True),
            lib_json("Tasks", "/sites/Docs/Lists/Tasks", BaseTemplate=171),
        ]}
    )
    by_title = await sp.get_library(fake_client, SITE, "service MANAGEMENT")
    assert by_title.root_server_relative_url == "/sites/Docs/Servicemanagement"
    by_root = await sp.get_library(fake_client, SITE, "Servicemanagement")
    assert by_root.title == "Service management"
    with pytest.raises(NotFoundError) as ei:
        await sp.get_library(fake_client, SITE, "Archief")
    assert ei.value.available == ["Documenten", "Service management"]
    assert "Available: Documenten, Service management" in str(ei.value)
    with pytest.raises(NotFoundError):
        await sp.get_library(fake_client, SITE, "Hidden lib")


# ------------------------------------------------------------------------------------------------ items


def row(i: int, ref: str, folder: bool = False, **extra) -> dict:
    return {
        "Id": i, "FileRef": ref, "FileLeafRef": ref.rsplit("/", 1)[-1], "FileDirRef": ref.rsplit("/", 1)[0],
        "FSObjType": 1 if folder else 0, "File_x0020_Size": None if folder else "1234",
        "Created": "2024-01-31T10:00:00Z", "Modified": "2024-02-01T09:30:00Z",
        "Author": {"EMail": "jan@contoso.nl", "Title": "Jan"}, "Editor": {"EMail": "piet@contoso.nl", "Title": "Piet"},
        "OData__UIVersionString": "3.0", "CheckoutUser": None, "ProgId": "", **extra,
    }


async def test_iter_items_keyset_paging(fake_client, respx_mock: respx.MockRouter) -> None:
    pages = {
        "ID gt 0": {"value": [row(1, f"{LIB_ROOT}/A", True), row(7, f"{LIB_ROOT}/A/b.docx")]},
        "ID gt 7": {"value": [row(9, f"{LIB_ROOT}/c.pdf")]},
    }

    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200, json=pages[request.url.params["$filter"]])

    route = respx_mock.get(url__startswith=f"{LIST_API}/items").mock(side_effect=handler)
    rows = [r async for r in sp.iter_items(fake_client, lib(), page_size=2)]
    assert [r["Id"] for r in rows] == [1, 7, 9]
    assert route.call_count == 2
    first = route.calls[0].request.url.params
    assert first["$orderby"] == "ID" and first["$top"] == "2"
    assert first["$expand"] == "Author,CheckoutUser,Editor"
    assert "Author/EMail" in first["$select"] and first["$select"].startswith("Id,")
    assert "+" not in raw(route.calls[0].request)  # spaces encoded as %20


async def test_iter_unique_permission_ids(fake_client, respx_mock: respx.MockRouter) -> None:
    respx_mock.get(url__startswith=f"{LIST_API}/items").respond(
        json={"value": [{"Id": 1, "HasUniqueRoleAssignments": False}, {"Id": 4, "HasUniqueRoleAssignments": True}]}
    )
    assert [i async for i in sp.iter_unique_permission_ids(fake_client, lib())] == [4]


def test_build_item_file_folder_outside() -> None:
    f = sp.build_item(row(7, f"{LIB_ROOT}/0. Archief/DAP's/Plan.DOCX", ProgId="Word.Document"), LIB_ROOT + "/")
    assert f is not None
    assert f.rel_path == "0. Archief/DAP's/Plan.DOCX" and f.name == "Plan.DOCX" and f.ext == ".docx"
    assert not f.is_folder and f.size == 1234 and f.id == 7
    assert f.created == datetime(2024, 1, 31, 10, 0, tzinfo=UTC) and f.created.tzinfo is not None
    assert f.modified == datetime(2024, 2, 1, 9, 30, tzinfo=UTC)
    assert (f.author_email, f.author_name, f.editor_email, f.editor_name) == (
        "jan@contoso.nl", "Jan", "piet@contoso.nl", "Piet")
    assert f.version == "3.0" and f.checked_out_to is None and f.prog_id == "Word.Document"
    assert f.server_relative_url == f"{LIB_ROOT}/0. Archief/DAP's/Plan.DOCX"

    d = sp.build_item(row(1, f"{LIB_ROOT}/0. Archief", folder=True, FSObjType="1"), LIB_ROOT)
    assert d is not None and d.is_folder and d.size == 0 and d.ext == "" and d.rel_path == "0. Archief"

    # relative to a sub-folder base, case-insensitive prefix
    sub = sp.build_item(row(7, f"{LIB_ROOT}/0. Archief/x.txt"), "/sites/docs/gedeelde documenten/0. archief")
    assert sub is not None and sub.rel_path == "x.txt"
    assert sp.build_item(row(8, f"{LIB_ROOT}/Other/x.txt"), f"{LIB_ROOT}/0. Archief") is None
    assert sp.build_item(row(9, f"{LIB_ROOT}/0. Archief", folder=True), f"{LIB_ROOT}/0. Archief") is None
    assert sp.build_item(row(10, f"{LIB_ROOT}/0. Archiefx/y.txt"), f"{LIB_ROOT}/0. Archief") is None

    checked = sp.build_item(row(3, f"{LIB_ROOT}/a.xlsx", Author=None, CheckoutUser={"Title": "Kees"}), LIB_ROOT)
    assert checked is not None and checked.author_email is None and checked.checked_out_to == "Kees"


# ------------------------------------------------------------------------------------------------ folders


async def test_ensure_folder_idempotent(fake_client, respx_mock: respx.MockRouter) -> None:
    existing = {LIB_ROOT, f"{LIB_ROOT}/0. Archief"}

    def exists(request: httpx.Request) -> httpx.Response:
        path = request.url.path  # decoded
        target = path.split("decodedurl='", 1)[1].rsplit("')", 1)[0].replace("''", "'")
        if target in existing:
            return httpx.Response(200, json={"Exists": True})
        return httpx.Response(404, json={"odata.error": {"message": {"value": "File Not Found."}}})

    def create(request: httpx.Request) -> httpx.Response:
        url = json.loads(request.content)["ServerRelativeUrl"]
        existing.add(url)
        return httpx.Response(200, json={"ServerRelativeUrl": url, "Exists": True})

    get = respx_mock.get(url__startswith=f"{SITE}/_api/web/GetFolderByServerRelativePath").mock(side_effect=exists)
    post = respx_mock.post(f"{SITE}/_api/web/folders").mock(side_effect=create)

    target = f"{LIB_ROOT}/0. Archief/DAP's/2024"
    await sp.ensure_folder(fake_client, SITE, target)
    created = [json.loads(c.request.content)["ServerRelativeUrl"] for c in post.calls]
    assert created == [f"{LIB_ROOT}/0. Archief/DAP's", target]
    assert post.calls[0].request.headers["Content-Type"] == "application/json;odata=nometadata"
    assert "DAP%27%27s" in raw(get.calls[0].request)

    calls_before = post.call_count
    await sp.ensure_folder(fake_client, SITE, target + "/")
    assert post.call_count == calls_before  # second run: exists check only

    with pytest.raises(NotFoundError):
        await sp.ensure_folder(fake_client, SITE, "/sites/Docs/NoSuchLibrary/x")


async def test_ensure_folder_already_exists_error_is_ok(fake_client, respx_mock: respx.MockRouter) -> None:
    respx_mock.get(url__startswith=f"{SITE}/_api/web/GetFolderByServerRelativePath").mock(
        side_effect=lambda r: httpx.Response(200, json={"Exists": r.url.path.endswith(f"{LIB_ROOT}')")})
    )
    respx_mock.post(f"{SITE}/_api/web/folders").respond(
        500, json={"odata.error": {"code": "-2130245363, Microsoft.SharePoint.SPException",
                                   "message": {"value": "A file or folder with the name X already exists."}}}
    )
    await sp.ensure_folder(fake_client, SITE, f"{LIB_ROOT}/X")


async def test_ensure_folder_with_hash_or_percent_uses_resource_path(fake_client, respx_mock: respx.MockRouter) -> None:
    existing = {LIB_ROOT}

    def exists(request: httpx.Request) -> httpx.Response:
        target = request.url.path.split("decodedurl='", 1)[1].rsplit("')", 1)[0]
        return httpx.Response(200, json={"Exists": target in existing})

    respx_mock.get(url__startswith=f"{SITE}/_api/web/GetFolderByServerRelativePath").mock(side_effect=exists)
    legacy = respx_mock.post(f"{SITE}/_api/web/folders").respond(json={})
    add = respx_mock.post(url__startswith=f"{SITE}/_api/web/folders/AddUsingPath").respond(json={})
    await sp.ensure_folder(fake_client, SITE, f"{LIB_ROOT}/#1 100%")
    assert legacy.call_count == 0 and add.call_count == 1
    assert "DecodedUrl='/sites/Docs/Gedeelde%20documenten/%231%20100%25',Overwrite=false" in raw(
        add.calls.last.request)


# ------------------------------------------------------------------------------------------------ upload


async def test_upload_small_uses_single_add(fake_client, respx_mock: respx.MockRouter, tmp_path: Path) -> None:
    src = tmp_path / "a.txt"
    src.write_bytes(b"hello")
    add = respx_mock.post(url__regex=r".*/Files/AddUsingPath\(.*").respond(
        json={"UniqueId": "u-1", "ServerRelativeUrl": f"{LIB_ROOT}/DAP's/a b.txt", "Length": "5"}
    )
    item = respx_mock.get(url__regex=r".*/ListItemAllFields.*").respond(json={"Id": 42})
    start = respx_mock.post(url__regex=r".*/StartUpload.*").respond(json={"value": "0"})
    res = await sp.upload_file(fake_client, SITE, f"{LIB_ROOT}/DAP's", "a b.txt", src, overwrite=True)
    assert res == {"UniqueId": "u-1", "ServerRelativeUrl": f"{LIB_ROOT}/DAP's/a b.txt", "Length": 5, "ItemId": 42}
    assert add.call_count == 1 and start.call_count == 0 and item.call_count == 1
    req = add.calls.last.request
    assert req.content == b"hello"
    assert "DecodedUrl='a%20b.txt',Overwrite=true" in raw(req)
    assert "DAP%27%27s" in raw(req)


async def test_upload_large_is_chunked(fake_client, respx_mock: respx.MockRouter, tmp_path: Path) -> None:
    src = tmp_path / "big.bin"
    data = bytes(range(25))
    src.write_bytes(data)
    add = respx_mock.post(url__regex=r".*/Files/AddUsingPath\(.*").respond(json={"UniqueId": "u-0"})
    start = respx_mock.post(url__regex=r".*/StartUpload\(.*").respond(json={"value": "10"})
    cont = respx_mock.post(url__regex=r".*/ContinueUpload\(.*").respond(json={"value": "20"})
    fin = respx_mock.post(url__regex=r".*/FinishUpload\(.*").respond(
        json={"UniqueId": "u-2", "ServerRelativeUrl": f"{LIB_ROOT}/big.bin", "Length": "25"}
    )
    respx_mock.get(url__regex=r".*/ListItemAllFields.*").respond(json={"Id": 7})
    res = await sp.upload_file(
        fake_client, SITE, LIB_ROOT, "big.bin", src, overwrite=False, chunk_size=10, small_limit=10
    )
    assert res["UniqueId"] == "u-2" and res["ItemId"] == 7 and res["Length"] == 25
    assert add.call_count == 1 and add.calls.last.request.content == b""
    assert "Overwrite=false" in raw(add.calls.last.request)
    assert start.calls.last.request.content == data[:10]
    assert cont.calls.last.request.content == data[10:20] and "fileOffset=10" in raw(cont.calls.last.request)
    assert fin.calls.last.request.content == data[20:] and "fileOffset=20" in raw(fin.calls.last.request)
    upload_id = raw(start.calls.last.request).split("uploadId=guid'")[1][:36]
    assert upload_id in raw(fin.calls.last.request)


# ------------------------------------------------------------------------------------------------ metadata


def _vul(results: list[dict]) -> dict:
    return {"value": results}


async def test_set_item_metadata_ok(fake_client, respx_mock: respx.MockRouter) -> None:
    route = respx_mock.post(f"{LIST_API}/items(5)/ValidateUpdateListItem").respond(json=_vul([
        {"FieldName": "Created", "ErrorMessage": None, "HasException": False, "ErrorCode": 0},
        {"FieldName": "Modified", "ErrorMessage": None, "HasException": False, "ErrorCode": 0},
    ]))
    await sp.set_item_metadata(
        fake_client, lib(), 5, created=datetime(2024, 1, 31, 10, 0, tzinfo=UTC),
        modified=datetime(2024, 2, 1, 11, 30, tzinfo=UTC),
    )
    body = json.loads(route.calls.last.request.content)
    assert body == {"formValues": [
        {"FieldName": "Created", "FieldValue": "2024-01-31T10:00:00Z"},
        {"FieldName": "Modified", "FieldValue": "2024-02-01T11:30:00Z"},
    ], "bNewDocumentUpdate": True}
    assert route.calls.last.request.headers["Content-Type"] == "application/json;odata=nometadata"


async def test_set_item_metadata_user_not_found(fake_client, respx_mock: respx.MockRouter) -> None:
    route = respx_mock.post(f"{LIST_API}/items(5)/ValidateUpdateListItem").respond(json=_vul([
        {"FieldName": "Created", "ErrorMessage": None, "HasException": False},
        {"FieldName": "Author", "ErrorMessage": "The user does not exist or is not unique.", "HasException": True,
         "ErrorCode": -2146232832},
    ]))
    with pytest.raises(sp.UserNotFoundError) as ei:
        await sp.set_item_metadata(
            fake_client, lib(), 5, created=datetime(2024, 1, 31, 10, 0, tzinfo=UTC),
            author_email="gone.user@contoso.com",
        )
    assert ei.value.field == "Author" and ei.value.email == "gone.user@contoso.com"
    assert "does not exist" in ei.value.message
    body = json.loads(route.calls.last.request.content)
    assert {"FieldName": "Author", "FieldValue": '[{"Key":"i:0#.f|membership|gone.user@contoso.com"}]'} in body[
        "formValues"]


# ------------------------------------------------------------------------------------------------ permissions

FULL = {"Id": 1073741829, "Name": "Full Control", "RoleTypeKind": 5,
        "BasePermissions": {"High": "2147483647", "Low": "4294967295"}}
CONTRIBUTE = {"Id": 1073741827, "Name": "Contribute", "RoleTypeKind": 3,
              "BasePermissions": {"High": "432", "Low": "1011030767"}}
READ = {"Id": 1073741826, "Name": "Read", "RoleTypeKind": 2,
        "BasePermissions": {"High": "176", "Low": "138612833"}}
LIMITED = {"Id": 1073741825, "Name": "Limited Access", "RoleTypeKind": 1,
           "BasePermissions": {"High": "48", "Low": "134287360"}}
ADD_ONLY = {"Id": 1073741900, "Name": "Add only", "RoleTypeKind": 0,
            "BasePermissions": {"High": "0", "Low": str(0x2)}}


def ra(pid: int, title: str, ptype: int, *bindings: dict) -> dict:
    return {"PrincipalId": pid, "Member": {"Id": pid, "Title": title, "PrincipalType": ptype, "LoginName": title},
            "RoleDefinitionBindings": list(bindings)}


async def test_role_assignments_can_write(fake_client, respx_mock: respx.MockRouter) -> None:
    respx_mock.get(url__startswith=f"{LIST_API}/roleassignments").respond(json={"value": [
        ra(3, "Owners", 8, FULL),
        ra(4, "Members", 8, CONTRIBUTE, LIMITED),
        ra(5, "Visitors", 8, READ),
        ra(6, "Guest link", 1, LIMITED),
        ra(7, "Uploader", 1, ADD_ONLY),
    ]})
    out = await sp.get_role_assignments(fake_client, lib())
    assert {r.principal_title: r.can_write for r in out} == {
        "Owners": True, "Members": True, "Visitors": False, "Guest link": False, "Uploader": True}
    members = out[1]
    assert members.principal_id == 4 and members.principal_type == 8
    assert members.role_names == ["Contribute", "Limited Access"]
    assert members.role_ids == [1073741827, 1073741825]


async def test_set_reader_only_downgrades_and_never_widens(fake_client, respx_mock: respx.MockRouter) -> None:
    respx_mock.get(url__startswith=f"{LIST_API}?").respond(json={"HasUniqueRoleAssignments": True})
    respx_mock.get(url__startswith=f"{LIST_API}/roleassignments").respond(json={"value": [
        ra(4, "Members", 8, CONTRIBUTE, LIMITED), ra(5, "Visitors", 8, READ)]})
    respx_mock.get(url__startswith=f"{SITE}/_api/web/roledefinitions").respond(
        json={"value": [FULL, CONTRIBUTE, READ, LIMITED]})
    add = respx_mock.post(url__regex=r".*/addroleassignment\(.*").respond(204)
    rem = respx_mock.post(url__regex=r".*/removeroleassignment\(.*").respond(204)

    await sp.set_reader_only(fake_client, lib(), 4)
    assert add.call_count == 1
    assert "principalid=4,roledefid=1073741826" in raw(add.calls.last.request)
    removed = sorted(raw(c.request).split("roledefid=")[1].rstrip(")") for c in rem.calls)
    assert removed == ["1073741825", "1073741827"]

    with pytest.raises(ValidationError):
        await sp.set_reader_only(fake_client, lib(), 5)  # read-only principal: nothing to downgrade
    with pytest.raises(ValidationError):
        await sp.set_reader_only(fake_client, lib(), 99)  # no assignment at all: would grant access
    assert add.call_count == 1


async def test_set_reader_only_requires_unique_permissions(fake_client, respx_mock: respx.MockRouter) -> None:
    respx_mock.get(url__startswith=f"{LIST_API}?").respond(json={"HasUniqueRoleAssignments": False})
    with pytest.raises(ValidationError):
        await sp.set_reader_only(fake_client, lib(), 4)


# ------------------------------------------------------------------------------------------------ misc


async def test_site_usage_fields_users(fake_client, respx_mock: respx.MockRouter) -> None:
    respx_mock.get(f"{SITE}/_api/site/Usage").respond(
        json={"Storage": 1_000_000, "StoragePercentageUsed": 0.001, "Bandwidth": 0, "Hits": 0, "Visits": 0}
    )
    assert await sp.get_site_usage(fake_client, SITE) == (1_000_000, 1_000_000_000)

    respx_mock.get(url__startswith=f"{LIST_API}/fields").respond(json={"value": [
        {"InternalName": "Klant", "Title": "Klant", "TypeAsString": "Text", "Hidden": False, "FromBaseType": False,
         "CanBeDeleted": True},
        {"InternalName": "Title", "Title": "Title", "TypeAsString": "Text", "Hidden": False, "FromBaseType": True,
         "CanBeDeleted": False},
        {"InternalName": "_Hidden", "Title": "H", "TypeAsString": "Text", "Hidden": True, "FromBaseType": False,
         "CanBeDeleted": True},
    ]})
    assert [f.internal_name for f in await sp.get_fields(fake_client, lib())] == ["Klant"]

    users = respx_mock.get(url__startswith=f"{SITE}/_api/web/siteusers").mock(
        side_effect=lambda r: httpx.Response(200, json={"value": [
            {"Id": 12, "Title": "Jan", "Email": "jan@contoso.nl", "LoginName": "i:0#.f|membership|jan@contoso.nl"}
        ] if "jan@" in r.url.params["$filter"] else []})
    )
    u = await sp.find_site_user(fake_client, SITE, "jan@contoso.nl")
    assert u is not None and u.id == 12 and u.email == "jan@contoso.nl"
    assert users.calls.last.request.url.params["$filter"] == "Email eq 'jan@contoso.nl'"
    assert await sp.find_site_user(fake_client, SITE, "gone.user@contoso.com") is None
    assert await sp.resolve_user(fake_client, SITE, "gone.user@contoso.com") is False


async def test_search_sites_parses_rows(fake_client, respx_mock: respx.MockRouter) -> None:
    route = respx_mock.get(url__startswith="https://contoso.sharepoint.com/_api/search/query").respond(json={
        "PrimaryQueryResult": {"RelevantResults": {"Table": {"Rows": [
            {"Cells": [{"Key": "Title", "Value": "Docs"}, {"Key": "Path", "Value": f"{SITE}/"},
                       {"Key": "Description", "Value": None}]},
        ]}}}
    })
    hits = await sp.search_sites(fake_client, "Doc", root_url="https://contoso.sharepoint.com")
    assert hits == [sp.SiteHit(title="Docs", url=SITE, description="")]
    assert route.calls.last.request.url.params["querytext"] == "'contentclass:STS_Site Doc*'"
