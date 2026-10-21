"""Small reusable widgets: findings table, key/value panel, formatting helpers."""

from __future__ import annotations

from collections.abc import Iterable

from rich.text import Text
from textual.widgets import DataTable, Static

from m365tools.core.events import LogLine
from m365tools.core.report.format import fmt_duration, fmt_size
from m365tools.services import SEVERITY_ORDER, sorted_findings
from m365tools.tools.spmig.models import Finding, RunSummary, Severity

SEVERITY_STYLE = {Severity.BLOCKING: "bold red", Severity.WARN: "yellow", Severity.INFO: "cyan"}
LEVEL_STYLE = {"debug": "dim", "info": "", "ok": "green", "warn": "yellow", "error": "bold red"}
MAX_TABLE_ROWS = 5000


def log_text(ev: LogLine) -> Text:
    style = LEVEL_STYLE.get(ev.level, "")
    return Text(f"{ev.at:%H:%M:%S} ", style="dim") + Text(f"{ev.level:<5} ", style=style or "dim") + Text(
        ev.text, style=style)


def severity_counts(findings: Iterable[Finding]) -> Text:
    findings = list(findings)
    t = Text()
    for sev in SEVERITY_ORDER:
        n = sum(1 for f in findings if f.severity == sev)
        t.append(f"{sev}: {n}   ", style=SEVERITY_STYLE[sev] if n else "dim")
    return t


class FindingsTable(DataTable):
    """Findings sorted blocking -> warn -> info, severity coloured."""

    def on_mount(self) -> None:
        self.cursor_type = "row"
        self.zebra_stripes = True
        self.add_columns("Severity", "Category", "Path", "Detail", "Fix")

    def load(self, findings: list[Finding]) -> None:
        self.clear()
        for f in sorted_findings(findings)[:MAX_TABLE_ROWS]:
            self.add_row(Text(str(f.severity).upper(), style=SEVERITY_STYLE.get(f.severity, "")),
                         Text(f.category), Text(f.path), Text(f.detail), Text(f.fix))


class KeyValue(Static):
    """A two-column key/value block."""

    def set_rows(self, rows: list[tuple[str, str]]) -> None:
        width = max((len(k) for k, _ in rows), default=0)
        t = Text()
        for i, (k, v) in enumerate(rows):
            if i:
                t.append("\n")
            t.append(f"{k:<{width}}  ", style="bold")
            t.append(v)
        self.update(t)


def summary_rows(s: RunSummary) -> list[tuple[str, str]]:
    rows = [("Mode", s.mode), ("Duration", fmt_duration(s.duration_s))]
    if s.mode in ("dryrun", "migrate"):
        rows += [("Copied" if s.mode == "migrate" else "Would copy", f"{s.copied:,}"),
                 ("Skipped", f"{s.skipped:,}"), ("Failed", f"{s.failed:,}"),
                 ("Metadata warnings", f"{s.metadata_warnings:,}")]
        if s.mode == "migrate":
            rows.append(("Bytes copied", fmt_size(s.bytes_copied)))
        if s.stopped_early:
            rows.append(("Stopped early", "yes"))
    elif s.mode == "verify":
        rows += [("OK", f"{s.copied:,}"), ("Problems", f"{s.failed:,}")]
    if s.findings:
        rows.append(("Findings", severity_counts(s.findings).plain))
    rows += [("Run folder", str(s.run_dir)), ("Report", str(s.run_dir / "report.html"))]
    return rows
