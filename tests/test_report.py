import csv
import json
import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

from m365tools.core.report import (
    Report,
    Section,
    audit_report,
    dryrun_report,
    fmt_duration,
    fmt_size,
    lock_report,
    migrate_report,
    render_html,
    render_text,
    verify_report,
    write_all,
    write_csv,
    write_findings_csv,
    write_json,
)
from m365tools.tools.spmig.models import (
    Action,
    CompareResult,
    CompareRow,
    Finding,
    Inventory,
    Item,
    LibraryRef,
    Plan,
    PlanEntry,
    RunSummary,
    Severity,
)

T0 = datetime(2026, 10, 1, 10, 0, 0)
LONG = "Folder/" + "/".join(["sub" * 10] * 5) + "/final-document-name.docx"


def mk_report():
    return Report(
        "Audit <x>", "audit", "src & co", "dst", T0, T0 + timedelta(seconds=90), [("Files", "<1>")],
        [Section("Alpha", "kv", rows=[["k", "v<b>"]]),
         Section("Beta", "table", ["Severity", "Path"], [["blocking", LONG], ["warn", "b"], ["info", "c"]],
                 severity_col=0),
         Section("Gamma", "text", text="hello"), Section("Delta", "table", ["a"], [["z"]], collapsed=True)],
        not_migrated=["versions"], footer="foot")


def test_html():
    h = render_html(mk_report())
    assert "Audit &lt;x&gt;" in h and "v&lt;b&gt;" in h and "src &amp; co" in h
    assert "v<b>" not in h
    for t in ("Alpha", "Beta", "Gamma", "Delta"):
        assert t in h
    assert not re.search(r"(src|href)=[\"']?https?://|url\(\s*https?://", h)
    assert "sev-blocking" in h and "sev-warn" in h and "sev-info" in h
    assert "What was NOT migrated" in h and "m365tools spmig" in h
    assert "<details>" in h and h.count("<details open>") == 3


def test_text():
    t = render_text(mk_report())
    assert all(len(line) <= 100 for line in t.splitlines())
    assert "…" in t and LONG not in t
    assert "Alpha\n-----" in t and "Audit <x>\n=========" in t


def test_text_wide_values():
    r = mk_report()
    r.source = "s" * 300
    assert all(len(x) <= 100 for x in render_text(r).splitlines())


def test_csv(tmp_path: Path):
    p = tmp_path / "a.csv"
    write_csv(p, ["a", "b"], [["x;y", "é"]])
    raw = p.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")
    rows = list(csv.reader(p.read_text("utf-8-sig").splitlines(), delimiter=";"))
    assert rows == [["a", "b"], ["x;y", "é"]]
    f = tmp_path / "f.csv"
    write_findings_csv(f, [Finding(Severity.WARN, "cat", "p", "d", "fix")])
    assert f.read_text("utf-8-sig").splitlines()[0] == "Severity;Category;Path;Detail;Fix"
    assert "warn;cat;p;d;fix" in f.read_text("utf-8-sig")


def test_json(tmp_path: Path):
    @dataclass
    class D:
        x: int

    p = tmp_path / "a.json"
    write_json(p, {"t": T0, "p": Path("/a"), "s": Severity.INFO, "d": D(1)})
    assert json.loads(p.read_text()) == {"t": T0.isoformat(), "p": "/a", "s": "info", "d": {"x": 1}}


def test_formats():
    assert fmt_size(2.4 * 1024**3) == "2.40 GB"
    assert fmt_size(10) == "10 B"
    assert fmt_size(1536) == "1.50 KB"
    assert fmt_duration(3437) == "00:57:17"
    assert fmt_duration(86400 + 2 * 3600 + 3 * 60 + 4) == "1d 02:03:04"


LIB = LibraryRef("https://x.sharepoint.com/sites/a", "/sites/a", "id", "Docs", "/sites/a/Docs", 3,
                 enable_versioning=True)
LIB2 = LibraryRef("https://x.sharepoint.com/sites/b", "/sites/b", "id", "Docs", "/sites/b/Docs", 0)


def item(path, size, folder=False, ext=".pdf"):
    return Item(1, folder, path, path.rsplit("/", 1)[-1], "" if folder else ext, size, T0, T0, None, None,
                None, None, None, None, None, "/sites/a/Docs/" + path)


def summ(mode="audit"):
    return RunSummary(mode, T0, T0 + timedelta(seconds=100), Path("."), copied=2, failed=1)


def titles(r):
    return [s.title for s in r.sections]


def test_audit_builder():
    inv = Inventory(LIB, "/sites/a/Docs", [item("A", 0, True), item("A/x.pdf", 100), item("y.docx", 5, ext=".docx")])
    fs = [Finding(Severity.INFO, "c", "p", "d"), Finding(Severity.BLOCKING, "c", "p2", "d")]
    r = audit_report(summ(), inv, None, LIB, LIB2, fs, {"Excludes": "none"})
    assert titles(r)[:3] == ["Source overview", "Library settings", "Findings"]
    assert any("File types" in t for t in titles(r)) and "Top folders" in titles(r)
    f = next(s for s in r.sections if s.title == "Findings")
    assert f.severity_col == 0 and f.rows[0][0] == "blocking"
    ls = next(s for s in r.sections if s.title == "Library settings")
    assert ["Versioning", "yes", "no", "differs"] in ls.rows
    assert len(r.not_migrated) == 4
    render_html(r), render_text(r)


def test_dryrun_migrate_verify_lock(tmp_path: Path):
    a = item("a.pdf", 10)
    plan = Plan([PlanEntry(Action.COPY, a, "/t"), PlanEntry(Action.CONFLICT, a, "/t", "dup"),
                 PlanEntry(Action.EXCLUDED, a, "", "ex", route="Arch")], LIB, LIB2, "/s", "/d")
    r = dryrun_report(plan, summ("dryrun"))
    assert titles(r) == ["Counts", "Conflicts", "Excluded", "Routed to other destinations"]
    m = migrate_report(summ("migrate"), plan, [{"status": "failed", "path": "a", "error": "boom"}])
    assert ("Failed", "1") in m.summary and m.sections[0].rows == [["a", "boom"]]
    cr = CompareResult([CompareRow("missing", False, "a", 1, None, None, None),
                        CompareRow("only_in_destination", False, "b", None, 2, None, None)])
    v = verify_report(cr, summ("verify"))
    assert next(s for s in v.sections if s.title == "Only in destination").collapsed
    lk = lock_report([{"principal": "bob", "before": "Edit", "after": "Read"}], applied=False)
    assert lk.sections[0].rows == [["bob", "Edit", "Read"]]
    write_all(tmp_path, m)
    assert (tmp_path / "report.html").exists() and (tmp_path / "report.txt").exists()
