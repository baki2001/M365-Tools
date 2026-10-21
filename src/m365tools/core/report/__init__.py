"""Report rendering: model, HTML/text/CSV/JSON writers and spmig builders."""

from __future__ import annotations

from pathlib import Path

from .builders import audit_report, dryrun_report, lock_report, migrate_report, verify_report
from .csv_ import write_csv, write_findings_csv
from .format import fmt_duration, fmt_size
from .html import render_html
from .json_ import write_json
from .model import Report, Section
from .text import render_text


def write_all(run_dir: Path, report: Report) -> None:
    run_dir = Path(run_dir)
    run_dir.mkdir(parents=True, exist_ok=True)
    (run_dir / "report.html").write_text(render_html(report), encoding="utf-8")
    (run_dir / "report.txt").write_text(render_text(report), encoding="utf-8")


__all__ = [
    "Report", "Section", "audit_report", "dryrun_report", "fmt_duration", "fmt_size", "lock_report",
    "migrate_report", "render_html", "render_text", "verify_report", "write_all", "write_csv",
    "write_findings_csv", "write_json",
]
