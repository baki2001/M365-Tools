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
from ..theme import FG, MODE_STYLE, SUB
from ..widgets import VimTable
from .base import Empty, Section


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
            rows = csv.reader(fh, delimiter=";")
            header = next(rows, [])
            if column not in header:
                return Counter()
            i = header.index(column)
            return Counter(row[i] if i < len(row) else "" for row in rows)
    except OSError:
        return Counter()


# A finished run folder does not change, but a dry run of 150k files leaves a results.csv that takes most of a
# second to count. Summaries are kept per folder and only recomputed when a file in it changes.
_SUMMARIES: dict[Path, tuple[tuple[tuple[str, int, int], ...], str]] = {}


def _signature(run: Path) -> tuple[tuple[str, int, int], ...]:
    out = []
    for name in ("status.json", "findings.csv", "verify.csv", "results.csv", "lock.csv", "report.txt"):
        try:
            st = (run / name).stat()
        except OSError:
            continue
        out.append((name, st.st_size, st.st_mtime_ns))
    return tuple(out)


def describe_run(run: Path, mode: str) -> str:
    """:func:`summarise_run`, cached until one of the run's result files changes."""
    sig = _signature(run)
    hit = _SUMMARIES.get(run)
    if hit is not None and hit[0] == sig:
        return hit[1]
    summary = summarise_run(run, mode)
    _SUMMARIES[run] = (sig, summary)
    return summary


def summarise_run(run: Path, mode: str) -> str:
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


def list_all_runs(limit: int | None = None) -> list[RunInfo]:
    """Run folders, newest first; only the first ``limit`` are read for their summary."""
    root = runs_dir()
    if not root.exists():
        return []
    folders = sorted((p for p in root.glob("*/*") if p.is_dir()), key=lambda p: p.name, reverse=True)
    out = []
    for p in folders[:limit]:
        parts = p.name.split("-")  # YYYYmmdd-HHMMSS-mode[-n]
        try:
            started = datetime.strptime("-".join(parts[:2]), "%Y%m%d-%H%M%S")
        except ValueError:
            started = None
        mode = parts[2] if len(parts) > 2 else "?"
        out.append(RunInfo(p, p.parent.name, mode, started, describe_run(p, mode)))
    return out


class RunsSection(Section):
    KEY, TITLE, NUMBER = "runs", "Runs", 7
    TAGLINE, IDLE = "every run folder, newest first", "none yet · r reloads"
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
        yield Empty(theme.empty_text("No runs yet", "Every audit, dry run, migration, verify and lock leaves a "
                                     "folder here.", [("a", "audit"), ("r", "reload")]), id="runs-empty")

    def on_mount(self) -> None:
        self.query_one(VimTable).add_fit_columns(("when", 16), ("mode", 10), ("profile", 20), ("summary", None))
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
                          theme.glyph_word(r.mode, MODE_STYLE.get(r.mode, SUB)),
                          Text(r.profile, style=FG, no_wrap=True, overflow="ellipsis"),
                          Text(r.summary, style=SUB, no_wrap=True, overflow="ellipsis"))
        if self.runs:
            table.move_cursor(row=min(row, len(self.runs) - 1))
        self.query_one(VimTable).display = bool(self.runs)
        self.query_one("#runs-empty").display = not self.runs
        self.query_one("#runs-head", Static).update(Text.assemble(
            theme.pill(f"{len(self.runs)} run(s)", theme.BLUE), ("  in ", SUB), (str(runs_dir()), FG)))
        self.status = theme.DONE if self.runs else theme.TODO
        self.summary = f"{len(self.runs)} run(s)" if self.runs else ""

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
