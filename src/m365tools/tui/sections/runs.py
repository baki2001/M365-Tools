"""Section 7 - Runs: every run folder (newest first) with a one-line summary; enter opens its report."""

from __future__ import annotations

import csv
import json
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import ClassVar

from rich.text import Text
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.widgets import DataTable, Static

from m365tools.core.paths import runs_dir

from .. import theme
from ..modals import TextModal
from ..theme import ACCENT, DIM, FG
from ..widgets import VimTable
from .base import Empty, Section

MODE_STYLE = {"audit": ACCENT, "dryrun": theme.CYAN, "migrate": theme.GREEN, "verify": theme.MAGENTA,
              "lock": theme.YELLOW}


@dataclass(slots=True)
class RunInfo:
    path: Path
    profile: str
    mode: str
    started: datetime | None
    summary: str


def _csv_column(path: Path, column: str) -> Counter[str]:
    try:
        with path.open(encoding="utf-8-sig", newline="") as fh:
            return Counter(row.get(column, "") for row in csv.DictReader(fh, delimiter=";"))
    except OSError:
        return Counter()


def describe_run(run: Path, mode: str) -> str:
    """Best-effort one-liner from the files a run leaves behind (status.json, findings.csv, ...)."""
    status = run / "status.json"
    if status.exists():
        try:
            st = json.loads(status.read_text(encoding="utf-8"))
            return (f"{st.get('done', '?')}/{st.get('total', '?')} files · copied {st.get('copied', 0)} · "
                    f"failed {st.get('failed', 0)}")
        except (OSError, ValueError, AttributeError):
            pass
    if mode == "audit" and (run / "findings.csv").exists():
        c = _csv_column(run / "findings.csv", "severity")
        return f"{c['blocking']} blocking · {c['warn']} warn · {c['info']} info"
    if mode == "verify" and (run / "verify.csv").exists():
        c = _csv_column(run / "verify.csv", "status")
        problems = c["missing"] + c["size_differs"] + c["source_newer"]
        return f"{c['ok']} ok · {problems} problems"
    if (run / "results.csv").exists():
        c = _csv_column(run / "results.csv", "result")
        return " · ".join(f"{n} {k.lower()}" for k, n in c.most_common()) or "no rows"
    if mode == "lock" and (run / "lock.csv").exists():
        c = _csv_column(run / "lock.csv", "status")
        changed = sum(n for k, n in c.items() if not k.startswith("skipped"))
        return f"{changed} principal(s) to Read" + (" (applied)" if c["applied"] else " (preview)")
    return "no results (interrupted?)" if not (run / "report.txt").exists() else ""


def list_all_runs() -> list[RunInfo]:
    root = runs_dir()
    if not root.exists():
        return []
    out = []
    for p in root.glob("*/*"):
        if not p.is_dir():
            continue
        parts = p.name.split("-")  # YYYYmmdd-HHMMSS-mode[-n]
        try:
            started = datetime.strptime("-".join(parts[:2]), "%Y%m%d-%H%M%S")
        except ValueError:
            started = None
        mode = parts[2] if len(parts) > 2 else "?"
        out.append(RunInfo(p, p.parent.name, mode, started, describe_run(p, mode)))
    return sorted(out, key=lambda r: r.path.name, reverse=True)


class RunsSection(Section):
    KEY, TITLE, NUMBER = "runs", "Runs", 7
    PRIMARY = "#runs-table"
    HELP = ("Every run folder under the runs directory, newest first. enter opens report.txt; o shows the "
            "path of report.html; r re-reads the folder.")
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("enter", "open", "open report", priority=True),
        Binding("o", "report", "report path"),
        Binding("r", "reload", "reload"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.runs: list[RunInfo] = []

    def compose(self) -> ComposeResult:
        yield Static("", id="runs-head", classes="counts")
        yield VimTable(id="runs-table")
        yield Empty("No runs yet. Every audit, dry run, migration, verify and lock leaves a folder here.",
                    id="runs-empty")

    def on_mount(self) -> None:
        table = self.query_one(VimTable)
        table.add_column("when", width=16)
        table.add_column("mode", width=8)
        table.add_column("profile", width=20)
        table.add_column("summary")
        self.action_reload()

    def on_show(self) -> None:
        self.action_reload()

    def action_reload(self) -> None:
        self.runs = list_all_runs()
        table = self.query_one(VimTable)
        row = table.cursor_row
        table.clear()
        for r in self.runs:
            table.add_row(Text(f"{r.started:%Y-%m-%d %H:%M}" if r.started else r.path.name, style=FG),
                          Text(r.mode, style=MODE_STYLE.get(r.mode, FG)), Text(r.profile, style=FG),
                          Text(r.summary, style=DIM))
        if self.runs:
            table.move_cursor(row=min(row, len(self.runs) - 1))
        self.query_one(VimTable).display = bool(self.runs)
        self.query_one("#runs-empty").display = not self.runs
        self.query_one("#runs-head", Static).update(
            Text.assemble((f"{len(self.runs)} run(s) in ", DIM), (str(runs_dir()), FG)))
        self.status = theme.DONE if self.runs else theme.TODO
        self.summary = f"{len(self.runs)} run(s)" if self.runs else "none yet"

    def _current(self) -> RunInfo | None:
        row = self.query_one(VimTable).cursor_row
        return self.runs[row] if 0 <= row < len(self.runs) else None

    def action_open(self) -> None:
        run = self._current()
        if run is None:
            return
        report = run.path / "report.txt"
        try:
            text = report.read_text(encoding="utf-8")
        except OSError:
            self.m365.notify(f"No report.txt in {run.path}", severity="warning", markup=False)
            return
        self.app.push_screen(TextModal(f"{run.mode} · {run.profile} · {run.path.name}", Text(text)))

    def on_data_table_row_selected(self, event: DataTable.RowSelected) -> None:
        event.stop()
        self.action_open()

    def action_report(self) -> None:
        run = self._current()
        self.m365.show_report_path(run.path if run else None)
