"""Run folder summaries: counted once, recounted when a result file changes, only for the runs shown."""

from __future__ import annotations

import os
from pathlib import Path

from m365tools.core.paths import runs_dir
from m365tools.tui.sections import runs


def _run(name: str, rows: list[str]) -> Path:
    run = runs_dir() / "svc" / name
    run.mkdir(parents=True)
    (run / "results.csv").write_text("rel_path;result;detail\n" + "".join(f"a/{i};{r};\n" for i, r in
                                                                          enumerate(rows)), encoding="utf-8-sig")
    return run


def test_summary_is_cached_until_the_results_change(monkeypatch):
    run = _run("20261023-100000-dryrun", ["Copied", "Copied", "Skipped"])
    calls = []
    real = runs.summarise_run
    monkeypatch.setattr(runs, "summarise_run", lambda r, m: calls.append(r) or real(r, m))

    assert runs.describe_run(run, "dryrun") == "2 copied · 1 skipped"
    assert runs.describe_run(run, "dryrun") == "2 copied · 1 skipped"
    assert len(calls) == 1

    csv_path = run / "results.csv"
    csv_path.write_text(csv_path.read_text(encoding="utf-8-sig") + "a/9;Failed;\n", encoding="utf-8-sig")
    st = csv_path.stat()
    os.utime(csv_path, ns=(st.st_atime_ns, st.st_mtime_ns + 1_000_000_000))
    assert runs.describe_run(run, "dryrun") == "2 copied · 1 skipped · 1 failed"
    assert len(calls) == 2


def test_list_all_runs_limit_reads_only_the_newest(monkeypatch):
    for i in range(4):
        _run(f"2026102{i}-100000-dryrun", ["Copied"])
    seen = []
    monkeypatch.setattr(runs, "describe_run", lambda r, m: seen.append(r.name) or "")

    out = runs.list_all_runs(limit=2)

    assert [r.path.name for r in out] == ["20261023-100000-dryrun", "20261022-100000-dryrun"]
    assert seen == ["20261023-100000-dryrun", "20261022-100000-dryrun"]
    assert len(runs.list_all_runs()) == 4


def test_csv_column_without_the_column_is_empty(tmp_path):
    path = tmp_path / "x.csv"
    path.write_text("a;b\n1;2\n", encoding="utf-8")
    assert runs._csv_column(path, "result") == {}
    assert runs._csv_column(path, "b") == {"2": 1}
