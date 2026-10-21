from datetime import datetime

from m365tools.core.events import CollectingSink
from m365tools.tools.spmig.checks import (
    CHECKS,
    SYSTEM_FIELDS,
    CheckContext,
    check_checked_out,
    check_columns,
    check_conflicts,
    check_destination_library,
    check_duration,
    check_large_files,
    check_library_settings,
    check_onenote,
    check_paths,
    check_quota,
    check_unique_permissions,
    check_users,
    check_versions,
    check_write_access,
    run_checks,
)
from m365tools.tools.spmig.gateway import FieldInfo
from m365tools.tools.spmig.inventory import build_inventory
from m365tools.tools.spmig.models import Endpoint, JobOptions, Severity
from m365tools.tools.spmig.plan import build_plan
from m365tools.tools.spmig.usermap import UserMap, UserMapEntry

from .fake_gateway import FakeGateway

MB = 1024 * 1024
SRC = Endpoint("t", "https://contoso.sharepoint.com/sites/src", "Docs")
DST = Endpoint("t", "https://contoso.sharepoint.com/sites/dst")


async def make_ctx(setup_src=None, setup_dst=None, *, dst_missing=False, with_plan=True, usermap=None, **opts):
    src = FakeGateway("https://contoso.sharepoint.com/sites/src")
    src.add_library("Docs")
    dst = FakeGateway("https://contoso.sharepoint.com/sites/dst")
    dst.add_library("Documenten", "Gedeelde documenten")
    if setup_src:
        setup_src(src)
    if setup_dst:
        setup_dst(dst)
    o = JobOptions(source=SRC, destination=DST, **opts)
    src_lib = src.lib().ref
    src_inv = await build_inventory(src, src_lib)
    dst_lib = None if dst_missing else dst.lib().ref
    dst_inv = None if dst_missing else await build_inventory(dst, dst_lib)
    plan = build_plan(src_inv, dst_inv, o, dst_inv.base_url, {}) if with_plan and dst_inv else None
    return CheckContext(src_inv=src_inv, dst_inv=dst_inv, src_lib=src_lib, dst_lib=dst_lib, opts=o, plan=plan,
                        src_gw=src, dst_gw=dst, usermap=usermap,
                        dst_error="library 'X' not found" if dst_missing else None)


def sev(findings, category):
    return [f.severity for f in findings if f.category == category]


async def test_destination_unresolvable_is_blocking():
    ctx = await make_ctx(dst_missing=True)
    f = await check_destination_library(ctx)
    assert [x.severity for x in f] == [Severity.BLOCKING]
    assert f[0].fix
    assert await check_destination_library(await make_ctx()) == []


async def test_write_probe_access_denied_is_blocking():
    ctx = await make_ctx(setup_dst=lambda d: setattr(d, "write_denied", True))
    f = await check_write_access(ctx)
    assert [x.severity for x in f] == [Severity.BLOCKING]
    assert await check_write_access(await make_ctx()) == []


async def test_quota():
    def src(g):
        g.add_file("big.bin", b"x" * 1000)

    ctx = await make_ctx(src, lambda d: setattr(d, "usage", (0, 1050)))
    assert sev(await check_quota(ctx), "quota") == [Severity.BLOCKING]  # needs 1100
    ctx = await make_ctx(src, lambda d: setattr(d, "usage", (0, 1100)))
    assert await check_quota(ctx) == []
    ctx = await make_ctx(src, lambda d: setattr(d, "usage", RuntimeError("nope")))
    assert sev(await check_quota(ctx), "quota") == [Severity.INFO]
    ctx = await make_ctx(src, lambda d: setattr(d, "usage", (10, 0)))
    assert sev(await check_quota(ctx), "quota") == [Severity.INFO]


async def test_paths_length_and_illegal_names():
    long_name = "a" * 250

    def src(g):
        g.add_file(f"{long_name}/{'b' * 160}.txt")  # > 400 at destination
        g.add_file(f"{'c' * 200}/{'d' * 40}.txt")  # > 260
        g.add_file("ok.txt")
        g.add_file("~$lock.docx")
        g.add_file("trailing.")
        g.add_file(" lead.txt")
        g.add_file('q"uote.txt')

    f = await check_paths(await make_ctx(src))
    by = {(x.category, x.path.rsplit("/", 1)[-1]): x.severity for x in f}
    assert by[("path_length", "b" * 160 + ".txt")] == Severity.BLOCKING
    assert by[("path_length", "d" * 40 + ".txt")] == Severity.INFO
    assert by[("illegal_name", "~$lock.docx")] == Severity.BLOCKING
    assert by[("illegal_name", "trailing.")] == Severity.BLOCKING
    assert by[("illegal_name", " lead.txt")] == Severity.BLOCKING
    assert by[("illegal_name", 'q"uote.txt')] == Severity.BLOCKING
    assert not any(x.path == "ok.txt" for x in f)


async def test_excluded_items_not_checked_for_names():
    f = await check_paths(await make_ctx(lambda g: g.add_file("~$lock.docx"), excludes=["**/~$*"]))
    assert f == []


async def test_unique_permissions_warn():
    def src(g):
        e = g.add_file("a.txt")
        g.lib().unique_ids.add(e.id)

    assert sev(await check_unique_permissions(await make_ctx(src)), "unique_permissions") == [Severity.WARN]


async def test_checked_out_versions_onenote_large():
    def src(g):
        g.add_file("co.docx", checked_out_to="Jan")
        g.add_file("v.docx", version="7.0")
        g.add_folder("Notes")
        g.lib().entries["/sites/src/docs/notes"].prog_id = "OneNote.Notebook"
        g.add_file("huge.iso", b"x" * (100 * MB))

    ctx = await make_ctx(src)
    assert sev(await check_checked_out(ctx), "checked_out") == [Severity.WARN]
    assert sev(await check_versions(ctx), "versions") == [Severity.INFO]
    assert sev(await check_onenote(ctx), "onenote") == [Severity.WARN]
    assert sev(await check_large_files(ctx), "large_file") == [Severity.INFO]


async def test_columns_missing_ignores_system_fields():
    def src(g):
        g.lib().fields = [FieldInfo("Klant", "Klant", "Text"), FieldInfo("MediaServiceOCR", "OCR"),
                          FieldInfo("_ExtendedDescription", "Description"), FieldInfo("Status", "Status")]

    def dst(g):
        g.lib().fields = [FieldInfo("Status0", "Status")]

    f = await check_columns(await make_ctx(src, dst))
    assert [(x.severity, x.path) for x in f] == [(Severity.WARN, "Klant")]
    assert "MediaServiceOCR" in SYSTEM_FIELDS


async def test_users_unresolvable_warn_per_user_with_count():
    def src(g):
        g.add_file("a.txt", author_email="gone.user@contoso.com", editor_email="gone.user@contoso.com")
        g.add_file("b.txt", author_email="gone.user@contoso.com", editor_email="ok@contoso.com")
        g.add_file("c.txt", author_email="mapped@old.nl")

    def dst(g):
        g.add_user("ok@contoso.com")

    um = UserMap([UserMapEntry("mapped@old.nl", "ok@contoso.com", "M", "mapped")])
    f = await check_users(await make_ctx(src, dst, usermap=um))
    assert [(x.severity, x.path) for x in f] == [(Severity.WARN, "gone.user@contoso.com")]
    assert "2 file(s)" in f[0].detail


async def test_conflicts_warn():
    def src(g):
        g.add_file("a.txt", modified=datetime(2025, 1, 1))

    def dst(g):
        g.add_file("a.txt", modified=datetime(2024, 1, 1))

    ctx = await make_ctx(src, dst)
    assert sev(await check_conflicts(ctx), "conflict") == [Severity.WARN]
    ctx = await make_ctx(src, dst, with_plan=False)
    assert sev(await check_conflicts(ctx), "conflict") == [Severity.WARN]


async def test_library_settings_and_duration_info():
    def dst(g):
        g.lib().ref.enable_versioning = True

    ctx = await make_ctx(lambda g: g.add_file("a.bin", b"x" * MB), dst)
    assert sev(await check_library_settings(ctx), "library_settings") == [Severity.INFO]
    d = await check_duration(ctx)
    assert sev(d, "duration") == [Severity.INFO]
    assert "0.7 MB/s" in d[0].detail


async def test_run_checks_all_have_fix_and_survive_crash():
    def src(g):
        g.add_file("~$x.docx", checked_out_to="Jan")

    ctx = await make_ctx(src, dst_missing=True)

    async def boom(_ctx):
        raise RuntimeError("kaput")

    CHECKS.append(boom)
    try:
        sink = CollectingSink()
        f = await run_checks(ctx, sink)
    finally:
        CHECKS.remove(boom)
    assert f[0].severity == Severity.BLOCKING  # sorted by severity
    assert any(x.category == "check_failed" for x in f)
    assert all(x.fix for x in f)
