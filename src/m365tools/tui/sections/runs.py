"""Section 7 - Runs: every run folder (newest first) with a one-line summary.

enter loads a run back into the workspace: the job options it was started with, plus the findings of its
audit and the plan of its dry run (or the newest ones with the same options), so you can carry on from there.
t shows report.txt."""

from __future__ import annotations

import csv
import json
from collections import Counter
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path
from typing import ClassVar

from rich.text import Text
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.widgets import DataTable, Static

from m365tools.core.config import AppConfig
from m365tools.core.paths import runs_dir
from m365tools.services import load_run_options
from m365tools.tools.spmig.models import Finding, JobOptions, Plan, RunSummary, Severity
from m365tools.tools.spmig.plan import plan_from_json

from .. import theme
from ..modals import TextModal
from ..theme import DIM, FG, MODE_STYLE, SUB
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
        return f"{theme.plural(changed, 'principal')} to Read" + (" (applied)" if c["applied"] else " (preview)")
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


SPMIG_MODES = ("audit", "dryrun", "migrate", "verify", "lock")


def read_findings(run: Path) -> list[Finding]:
    """findings.csv of an audit run (raises OSError/ValueError when missing or unreadable)."""
    with (run / "findings.csv").open(encoding="utf-8-sig", newline="") as fh:
        return [Finding(Severity(row["severity"]), row["category"], row["path"], row["detail"], row.get("fix") or "")
                for row in csv.DictReader(fh, delimiter=";")]


def read_plan(run: Path) -> Plan:
    """plan.json of a dry run (raises OSError/ValueError when missing or unreadable)."""
    return plan_from_json(json.loads((run / "plan.json").read_text(encoding="utf-8")))


def _same_job(a: JobOptions, b: JobOptions) -> bool:
    return replace(a, profile=None) == replace(b, profile=None)


def latest_run_with(runs: list[RunInfo], mode: str, opts: JobOptions, cfg: AppConfig) -> RunInfo | None:
    """Newest run of ``mode`` that was started with the same job options as ``opts``."""
    for r in runs:
        if r.mode != mode:
            continue
        try:
            if _same_job(load_run_options(r.path, cfg), opts):
                return r
        except Exception:  # noqa: BLE001, S112 - a run without readable options just doesn't match
            continue
    return None


class RunsSection(Section):
    KEY, TITLE, NUMBER = "runs", "Runs", 7
    TAGLINE, IDLE = "every run folder, newest first", "none yet · r reloads"
    PRIMARY = "#runs-table"
    HELP = ("Every run folder under the runs directory, newest first. enter loads a run back: its job "
            "(source, destination, options), its audit findings into 2 Checks and its dry run into 3 Plan, so you "
            "can continue with m. When you load one of the two, the newest other one with the same job comes along. "
            "t shows report.txt; o shows the path of report.html; r re-reads the folder.")
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("enter", "load", "load run", priority=True),
        Binding("t", "open", "text report"),
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
            (f"{len(self.runs):,}", f"bold {FG}"), (" run" if len(self.runs) == 1 else " runs", SUB),
            ("   in ", DIM), (str(runs_dir()), SUB), no_wrap=True, overflow="ellipsis"))
        self.status = theme.DONE if self.runs else theme.TODO
        self.summary = theme.plural(len(self.runs), "run") if self.runs else ""

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
        self.action_load()

    def action_load(self) -> None:
        """Restore the run's job options and the audit + dry run results that go with them."""
        run = self._current()
        if run is None:
            return
        app = self.m365
        if run.mode not in SPMIG_MODES:  # e.g. groups runs: nothing to load, show the report
            self.action_open()
            return
        if app.busy:
            app.notify(f"{app.busy} is running: load a run after it finishes.", severity="warning")
            return
        try:
            opts = load_run_options(run.path, app.cfg)
        except Exception as exc:  # noqa: BLE001 - shown to the user
            app.report_error(exc, "load run")
            return
        app.set_options(opts)
        ws = app.workspace
        loaded: list[str] = []
        audit = run if run.mode == "audit" else latest_run_with(self.runs, "audit", opts, app.cfg)
        if audit is not None:
            try:
                findings = read_findings(audit.path)
            except (OSError, ValueError, KeyError) as exc:
                app.notify(f"Audit {audit.path.name}: findings not readable ({exc})", severity="warning",
                           markup=False)
            else:
                when = audit.started or datetime.now()
                ws.section("checks").show_findings(
                    findings, audit.path, RunSummary("audit", when, when, audit.path, findings=findings))
                loaded.append(f"audit {audit.path.name}")
        dryrun = run if run.mode == "dryrun" else latest_run_with(self.runs, "dryrun", opts, app.cfg)
        if dryrun is not None:
            try:
                plan = read_plan(dryrun.path)
            except (OSError, ValueError, KeyError, TypeError) as exc:
                app.notify(f"Dry run {dryrun.path.name}: plan not readable ({exc})", severity="warning",
                           markup=False)
            else:
                ws.section("plan").show_plan(plan, dryrun.path)
                loaded.append(f"dry run {dryrun.path.name}")
        app.notify("Job loaded" + (": " + ", ".join(loaded) if loaded else " (no audit or dry run to show)"),
                   markup=False)
        ws.show_section("checks" if run.mode == "audit" else "plan" if run.mode == "dryrun" and loaded
                        else "job")

    def action_report(self) -> None:
        run = self._current()
        self.m365.show_report_path(run.path if run else None)
