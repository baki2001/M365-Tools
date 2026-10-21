from m365tools.tools.spmig.inventory import build_inventory
from m365tools.tools.spmig.models import UserMapEntry
from m365tools.tools.spmig.usermap import UserMap, build_usermap, load_usermap, save_usermap

from .fake_gateway import FakeGateway, FakeGatewayWithLocalPart


def test_csv_roundtrip(tmp_path):
    entries = [
        UserMapEntry("a@old.nl", "a@new.nl", "Anne; de Vries", "mapped"),
        UserMapEntry("b@old.nl", None, "Bert", "unmapped"),
        UserMapEntry("c@x.nl", "c@x.nl", "Cé", "auto"),
    ]
    p = tmp_path / "usermap.csv"
    save_usermap(p, entries)
    raw = p.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")
    assert raw.decode("utf-8-sig").splitlines()[0] == "source_email;target_email;display_name;status"
    assert load_usermap(p) == entries
    um = UserMap.load(p)
    assert um.resolve("A@OLD.nl") == "a@new.nl"
    assert um.resolve("b@old.nl") is None
    assert um.resolve("nobody@x.nl") is None


def test_load_hand_edited_rows(tmp_path):
    p = tmp_path / "m.csv"
    p.write_text("source_email;target_email;display_name;status\nx@o.nl;x@n.nl;X;\ny@o.nl;;Y;mapped\n",
                 encoding="utf-8")
    e = {x.source_email: x for x in load_usermap(p)}
    assert e["x@o.nl"].status == "mapped"
    assert e["y@o.nl"].status == "unmapped"


async def _src_inv():
    src = FakeGateway("https://contoso.sharepoint.com/sites/src")
    src.add_library("Docs")
    src.add_file("a.txt", author_email="same@contoso.com", author_name="Same",
                 editor_email="jan@oldtenant.nl", editor_name="Jan")
    src.add_file("b.txt", author_email="gone@oldtenant.nl", author_name="Gone")
    return await build_inventory(src, src.lib().ref)


async def test_auto_map_same_email_and_unmapped():
    inv = await _src_inv()
    dst = FakeGateway("https://contoso.sharepoint.com/sites/dst")
    dst.add_user("same@contoso.com")
    m = {e.source_email: e for e in await build_usermap(inv, dst)}
    assert m["same@contoso.com"].status == "auto" and m["same@contoso.com"].target_email == "same@contoso.com"
    assert m["jan@oldtenant.nl"].status == "unmapped" and m["jan@oldtenant.nl"].target_email is None
    assert m["jan@oldtenant.nl"].display_name == "Jan"


async def test_auto_map_by_local_part_when_supported():
    inv = await _src_inv()
    dst = FakeGatewayWithLocalPart("https://contoso.sharepoint.com/sites/dst")
    dst.add_user("jan@contoso.com")
    m = {e.source_email: e for e in await build_usermap(inv, dst)}
    assert m["jan@oldtenant.nl"].status == "auto" and m["jan@oldtenant.nl"].target_email == "jan@contoso.com"
    assert m["gone@oldtenant.nl"].status == "unmapped"
