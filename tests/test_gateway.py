"""RestGateway delegates to core.sharepoint with arguments matching the real signatures."""

import inspect
from datetime import datetime
from pathlib import Path

import pytest

from m365tools.core.errors import AccessDeniedError
from m365tools.tools.spmig.gateway import RestGateway, UserNotFoundError

from .fake_gateway import make_lib

sp = pytest.importorskip("m365tools.core.sharepoint")

RETURNS = {
    "get_default_library": make_lib(),
    "get_library": make_lib(),
    "upload_file": {"UniqueId": "u1", "ServerRelativeUrl": "/x", "Length": 1, "ItemId": 42},
    "get_site_usage": (10, None),
    "probe_write_access": False,
    "get_role_assignments": [sp.RoleAssignment(3, "Owners", 8, ["Full Control"], True)],
    "get_associated_owner_group_id": 3,
    "get_fields": [sp.FieldInfo("Klant", "Klant", "Text")],
    "find_site_user": sp.SiteUser(1, "Jan", "jan@x.nl", "i:0#.f|membership|jan@x.nl"),
}


@pytest.fixture
def calls(monkeypatch):
    seen: list[str] = []
    for name in ["get_default_library", "get_library", "ensure_folder", "download_file", "upload_file",
                 "set_item_metadata", "get_site_usage", "probe_write_access", "get_role_assignments",
                 "get_associated_owner_group_id", "break_inheritance", "set_reader_only", "get_fields",
                 "find_site_user"]:
        real = getattr(sp, name)

        def make(real=real, name=name):
            async def fake(*a, **kw):
                inspect.signature(real).bind(*a, **kw)
                seen.append(name)
                if name == "set_item_metadata" and a[5]:
                    raise sp.UserNotFoundError("Author", a[5], "not found")
                return RETURNS.get(name)
            return fake

        monkeypatch.setattr(sp, name, make())
    return seen


async def test_rest_gateway_signatures(calls):
    gw = RestGateway(object(), "https://c.sharepoint.com/sites/a/")
    lib = make_lib()
    assert (await gw.get_default_library()).title == "Docs"
    await gw.get_library("Docs")
    await gw.ensure_folder("/sites/a/Docs/x")
    await gw.download_file("/sites/a/Docs/x/f", Path("/tmp/none"))
    res = await gw.upload_file("/sites/a/Docs/x", "f", Path("/tmp/none"), True)
    assert (res.unique_id, res.item_id) == ("u1", 42)
    await gw.set_item_metadata(lib, 42, datetime(2024, 1, 1), None, None, None)
    with pytest.raises(UserNotFoundError):
        await gw.set_item_metadata(lib, 42, None, None, "gone@x.nl", None)
    assert await gw.get_site_usage() == (10, 0)
    with pytest.raises(AccessDeniedError):
        await gw.probe_write_access("/sites/a/Docs")
    ra = (await gw.get_role_assignments(lib))[0]
    assert (ra.principal_id, ra.can_write, ra.role_names) == (3, True, ["Full Control"])
    assert await gw.get_associated_owner_group_id() == 3
    await gw.break_inheritance(lib)
    await gw.set_reader_only(lib, 3)
    assert (await gw.get_fields(lib))[0].type_name == "Text"
    assert (await gw.find_site_user("jan@x.nl")).email == "jan@x.nl"
    assert len(set(calls)) == 14


async def test_rest_gateway_dates_use_regional_format_and_fall_back_to_minutes(monkeypatch):
    rs = sp.RegionalSettings(date_order=1, date_separator="-", time_zone="(UTC+01:00) Amsterdam, Berlin",
                             bias_minutes=-60)
    sent: list[str] = []

    async def regional(client, site_url):
        return rs

    async def set_meta(client, lib, item_id, created, modified, author, editor, *, date_format):
        text = date_format(created)
        sent.append(text)
        if text.count(":") == 2:  # this site only takes dates to the minute
            raise sp.InvalidDateError("Created: invalid")

    monkeypatch.setattr(sp, "get_regional_settings", regional)
    monkeypatch.setattr(sp, "set_item_metadata", set_meta)
    gw = RestGateway(object(), "https://c.sharepoint.com/sites/a")
    when = datetime(2024, 1, 31, 10, 0, 5)
    await gw.set_item_metadata(make_lib(), 1, when, when, None, None)
    await gw.set_item_metadata(make_lib(), 2, when, when, None, None)
    assert sent == ["31-1-2024 11:00:05", "31-1-2024 11:00", "31-1-2024 11:00"]


async def test_rest_gateway_dates_fall_back_to_iso_without_regional_settings(monkeypatch):
    sent: list[str] = []

    async def regional(client, site_url):
        raise RuntimeError("no access")

    async def set_meta(client, lib, item_id, created, modified, author, editor, *, date_format):
        sent.append(date_format(created))
        raise sp.InvalidDateError("Created: invalid")

    monkeypatch.setattr(sp, "get_regional_settings", regional)
    monkeypatch.setattr(sp, "set_item_metadata", set_meta)
    gw = RestGateway(object(), "https://c.sharepoint.com/sites/a")
    with pytest.raises(sp.InvalidDateError):
        await gw.set_item_metadata(make_lib(), 1, datetime(2024, 1, 31, 10, 0, 5), None, None, None)
    assert sent == ["2024-01-31T10:00:05Z"]


async def test_rest_gateway_probe_metadata(monkeypatch):
    from datetime import UTC, timedelta

    got: dict[str, object] = {}
    shift = timedelta(0)

    async def upload(client, site_url, folder, name, src_path, overwrite):
        got["name"] = name
        return {"ItemId": 9}

    async def set_meta(client, lib, item_id, created, modified, author, editor, *, date_format):
        got["dates"] = (created, modified)

    async def dates(client, lib, item_id):
        c, m = got["dates"]
        return c + shift, m + shift

    async def delete(client, site_url, url):
        got["deleted"] = url

    monkeypatch.setattr(sp, "upload_file", upload)
    monkeypatch.setattr(sp, "set_item_metadata", set_meta)
    monkeypatch.setattr(sp, "get_item_dates", dates)
    monkeypatch.setattr(sp, "delete_file", delete)
    gw = RestGateway(object(), "https://c.sharepoint.com/sites/a")
    assert await gw.probe_metadata(make_lib(), "/sites/a/Docs") is None
    assert got["deleted"] == f"/sites/a/Docs/{got['name']}"
    assert got["dates"][0].tzinfo is UTC
    shift = timedelta(hours=1)
    assert "+1.0 h" in await gw.probe_metadata(make_lib(), "/sites/a/Docs")
