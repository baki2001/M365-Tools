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
