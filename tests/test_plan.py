from datetime import UTC, datetime, timedelta

from m365tools.tools.spmig.models import Action, Endpoint, JobOptions, Route
from m365tools.tools.spmig.plan import (
    ROUTE_UNAVAILABLE,
    build_plan,
    glob_match,
    plan_from_json,
    plan_to_json,
)

from .fake_gateway import make_inv, make_item, make_lib

SRC = Endpoint("t", "https://contoso.sharepoint.com/sites/a", "Docs")
DST = Endpoint("t", "https://contoso.sharepoint.com/sites/b")
DST_BASE = "/sites/b/Gedeelde documenten"


def opts(**kw) -> JobOptions:
    return JobOptions(source=SRC, destination=DST, **kw)


def dst_inv(items):
    lib = make_lib("Documenten", DST_BASE)
    return make_inv(items, lib)


def by_path(plan):
    return {e.item.rel_path: e for e in plan.entries}


def test_glob_semantics():
    assert glob_match("0. Archief/**", "0. Archief")
    assert glob_match("0. Archief/**", "0. archief/2019/x.pdf")
    assert not glob_match("0. Archief/**", "0. Archief old/x.pdf")
    assert glob_match("**/*.tmp", "a.tmp")
    assert glob_match("**/*.tmp", "x/y/a.TMP")
    assert not glob_match("**/*.tmp", "a.tmpx")


def test_excludes_include_folder_itself_and_targets():
    src = make_inv([
        make_item("0. Archief", folder=True),
        make_item("0. Archief/old.pdf"),
        make_item("Klanten", folder=True),
        make_item("Klanten/a.docx"),
        make_item("Klanten/b.tmp"),
    ])
    plan = build_plan(src, dst_inv([]), opts(excludes=["0. Archief/**", "**/*.tmp"]), DST_BASE, {})
    p = by_path(plan)
    assert p["0. Archief"].action == Action.EXCLUDED
    assert p["0. Archief/old.pdf"].action == Action.EXCLUDED
    assert p["Klanten/b.tmp"].action == Action.EXCLUDED
    assert p["Klanten"].action == Action.CREATE_FOLDER
    assert p["Klanten/a.docx"].action == Action.COPY
    assert p["Klanten/a.docx"].target_server_relative_url == f"{DST_BASE}/Klanten/a.docx"
    assert plan.dst_base == DST_BASE


def test_routes_capture_prefix_and_override_prefix_exclude():
    src = make_inv([
        make_item("0. Archief", folder=True),
        make_item("0. Archief/2019", folder=True),
        make_item("0. Archief/2019/x.pdf"),
        make_item("0. Archief/2019/y.tmp"),
        make_item("Klanten/a.pdf"),
    ])
    o = opts(excludes=["0. Archief/**", "**/*.tmp"],
             routes=[Route("0. Archief", Endpoint("t", "https://contoso.sharepoint.com/sites/b", "Archief"))])
    plan = build_plan(src, dst_inv([]), o, DST_BASE, {"0. Archief": "/sites/b/Archief"})
    p = by_path(plan)
    assert p["0. Archief/2019/x.pdf"].action == Action.COPY
    assert p["0. Archief/2019/x.pdf"].route == "0. Archief"
    assert p["0. Archief/2019/x.pdf"].target_server_relative_url == "/sites/b/Archief/2019/x.pdf"
    assert p["0. Archief/2019"].action == Action.CREATE_FOLDER
    assert p["0. Archief/2019"].target_server_relative_url == "/sites/b/Archief/2019"
    assert p["0. Archief"].action == Action.SKIP  # route root = library root
    assert p["0. Archief/2019/y.tmp"].action == Action.EXCLUDED  # other excludes still apply
    assert p["Klanten/a.pdf"].route is None


def test_overwrite_policies():
    t = datetime(2024, 5, 1, 12, 0, 0)
    src = make_inv([
        make_item("same.pdf", modified=t),
        make_item("dst_newer.pdf", modified=t),
        make_item("within_tolerance.pdf", modified=t),
        make_item("src_newer.pdf", modified=t),
        make_item("new.pdf", modified=t),
    ])
    dst = dst_inv([
        make_item("same.pdf", modified=t, base=DST_BASE),
        make_item("dst_newer.pdf", modified=t + timedelta(days=1), base=DST_BASE),
        make_item("within_tolerance.pdf", modified=t - timedelta(seconds=1), base=DST_BASE),
        make_item("src_newer.pdf", modified=t - timedelta(hours=1), base=DST_BASE),
    ])
    p = by_path(build_plan(src, dst, opts(), DST_BASE, {}))
    assert p["same.pdf"].action == Action.SKIP
    assert p["dst_newer.pdf"].action == Action.SKIP
    assert p["within_tolerance.pdf"].action == Action.SKIP
    assert p["src_newer.pdf"].action == Action.CONFLICT
    assert p["new.pdf"].action == Action.COPY

    p = by_path(build_plan(src, dst, opts(overwrite="never"), DST_BASE, {}))
    assert p["src_newer.pdf"].action == Action.SKIP
    assert p["new.pdf"].action == Action.COPY

    p = by_path(build_plan(src, dst, opts(overwrite="always"), DST_BASE, {}))
    assert p["same.pdf"].action == Action.COPY
    assert p["src_newer.pdf"].action == Action.COPY


def test_folder_ordering_and_existing_folders():
    src = make_inv([
        make_item("a/b/c/f.txt"),
        make_item("a/b/c", folder=True),
        make_item("z.txt"),
        make_item("a", folder=True),
        make_item("a/b", folder=True),
        make_item("x", folder=True),
    ])
    dst = dst_inv([make_item("x", folder=True, base=DST_BASE)])
    plan = build_plan(src, dst, opts(), DST_BASE, {})
    kinds = [e.item.is_folder for e in plan.entries]
    assert kinds == sorted(kinds, reverse=True)  # all folders before files
    folder_paths = [e.item.rel_path for e in plan.entries if e.item.is_folder]
    depths = [p.count("/") for p in folder_paths]
    assert depths == sorted(depths)
    p = by_path(plan)
    assert p["x"].action == Action.SKIP
    assert [e.item.rel_path for e in plan.folders_to_create()] == ["a", "a/b", "a/b/c"]


def test_plan_json_roundtrip():
    src = make_inv([make_item("a", folder=True), make_item("a/f.txt", size=5)])
    plan = build_plan(src, dst_inv([]), opts(), DST_BASE, {})
    again = plan_from_json(plan_to_json(plan))
    assert again == plan


def test_exclude_inside_route_still_applies():
    src = make_inv([
        make_item("0. Archief", folder=True),
        make_item("0. Archief/Concepten", folder=True),
        make_item("0. Archief/Concepten/x.pdf"),
        make_item("0. Archief/y.pdf"),
    ])
    o = opts(excludes=["0. Archief/**", "0. Archief/Concepten/**"],
             routes=[Route("0. Archief", Endpoint("t", "https://contoso.sharepoint.com/sites/b", "Archief"))])
    p = by_path(build_plan(src, dst_inv([]), o, DST_BASE, {"0. Archief": "/sites/b/Archief"}))
    assert p["0. Archief/y.pdf"].action == Action.COPY
    assert p["0. Archief/Concepten"].action == Action.EXCLUDED
    assert p["0. Archief/Concepten/x.pdf"].action == Action.EXCLUDED


def test_route_without_destination_never_falls_back_to_main_destination():
    src = make_inv([make_item("0. Archief", folder=True), make_item("0. Archief/x.pdf"), make_item("a.pdf")])
    o = opts(routes=[Route("0. Archief", Endpoint("t", "https://contoso.sharepoint.com/sites/b", "Nope"))])
    p = by_path(build_plan(src, dst_inv([]), o, DST_BASE, {}))  # route destination did not resolve
    assert p["0. Archief/x.pdf"].action == Action.EXCLUDED
    assert p["0. Archief/x.pdf"].reason.startswith(ROUTE_UNAVAILABLE)
    assert not p["0. Archief/x.pdf"].target_server_relative_url.startswith(DST_BASE)
    assert p["0. Archief"].action == Action.EXCLUDED
    assert p["a.pdf"].action == Action.COPY


def test_naive_and_aware_dates_compare_and_empty_destination_is_recopied():
    t = datetime(2024, 5, 1, 12, 0, 0)
    src = make_inv([make_item("a.pdf", modified=t), make_item("b.pdf", modified=t, size=10)])
    dst = dst_inv([
        make_item("a.pdf", modified=t.replace(tzinfo=UTC) - timedelta(hours=1), base=DST_BASE),
        make_item("b.pdf", modified=datetime(2026, 1, 1, tzinfo=UTC), size=0, base=DST_BASE),
    ])
    p = by_path(build_plan(src, dst, opts(), DST_BASE, {}))
    assert p["a.pdf"].action == Action.CONFLICT  # source newer, no TypeError on naive vs aware
    assert p["b.pdf"].action == Action.CONFLICT  # empty placeholder from an interrupted upload
