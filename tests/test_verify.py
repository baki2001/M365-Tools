from datetime import datetime, timedelta

from m365tools.tools.spmig.verify import compare

from .fake_gateway import make_inv, make_item

T = datetime(2024, 1, 1, 12)


def statuses(result):
    return {r.rel_path: r.status for r in result.rows}


def test_compare_rules_and_office_tolerance():
    src = make_inv([
        make_item("Map", folder=True),
        make_item("Map/ok.pdf", size=10, modified=T),
        make_item("Map/word.docx", size=100, modified=T),
        make_item("Map/data.csv", size=10, modified=T),
        make_item("Map/newer.pdf", size=10, modified=T + timedelta(hours=1)),
        make_item("Map/tolerance.pdf", size=10, modified=T + timedelta(seconds=1)),
        make_item("Map/missing.pdf"),
    ])
    dst = make_inv([
        make_item("map", folder=True),
        make_item("map/OK.pdf", size=10, modified=T),
        make_item("Map/word.docx", size=180, modified=T),  # SharePoint grew the Office file
        make_item("Map/data.csv", size=12, modified=T),
        make_item("Map/newer.pdf", size=10, modified=T),
        make_item("Map/tolerance.pdf", size=10, modified=T),
        make_item("Map/extra.pdf"),
    ])
    s = statuses(compare(src, dst))
    assert s == {
        "Map": "ok",
        "Map/ok.pdf": "ok",
        "Map/word.docx": "ok",
        "Map/data.csv": "size_differs",
        "Map/newer.pdf": "source_newer",
        "Map/tolerance.pdf": "ok",
        "Map/missing.pdf": "missing",
        "Map/extra.pdf": "only_in_destination",
    }
    assert compare(src, dst).counts()["ok"] == 4
