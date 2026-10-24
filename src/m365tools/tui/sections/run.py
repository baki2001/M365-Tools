"""Section 4 - Run: the migration dashboard, and the app-wide log tail.

Lifecycle of one migration (``run_migration``):
  confirm (y/n, or type ``force`` when the audit has blocking findings)
  -> leg: ``job.migrate(resume_run_dir, force)``  (the Job audits again first, then copies)
  -> p pauses: ``request_stop()`` lets in-flight files finish; p again starts a new leg in the same run folder,
     which skips everything the journal lists as done. x stops for good.
Before the transfer phases (audit/inventory) the Job cannot stop gracefully, so pause/stop cancel the worker.
"""

from __future__ import annotations

import asyncio
import json
import time
from datetime import datetime
from pathlib import Path
from typing import Any, ClassVar

from rich.table import Table
from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.widgets import Label, ProgressBar, Static
from textual.worker import Worker

from m365tools.core.events import (
    ItemStarted,
    ItemStep,
    LogLine,
    PhaseChanged,
    ProgressEvent,
    Stats,
    Throttled,
    phase_label,
)
from m365tools.services import list_runs, load_run_options
from m365tools.tools.spmig.job import BlockingFindingsError
from m365tools.tools.spmig.models import RunSummary

from .. import theme
from ..modals import Choice, ConfirmModal, PickerModal, TypedConfirmModal
from ..theme import ACCENT, BLUE, DIM, FG, GREEN, RED, SUB, SURFACE1, YELLOW
from ..widgets import VimLog
from .base import Section

STEP_COLOR = {"folder": theme.CYAN, "download": BLUE, "upload": theme.PURPLE, "metadata": theme.CYAN, "verify": GREEN}
TRANSFER_PHASES = ("folders", "files", "finalize")


def stat_grid(cells: list[tuple[str, str, str]], columns: int = 4) -> Table:
    """Stats as aligned ``label  value`` pairs, ``columns`` per row: label secondary, value bold."""
    grid = Table.grid(expand=True, padding=(0, 2, 0, 0))
    for _ in range(columns):
        grid.add_column(no_wrap=True, style=SUB)
        grid.add_column(no_wrap=True, ratio=1)
    for i in range(0, len(cells), columns):
        row: list[Text | str] = []
        for label, value, style in cells[i:i + columns]:
            row += [label, Text(value, style=f"bold {style}" if value != "–" else DIM)]
        grid.add_row(*row)
    return grid


def state_line(word: str, color: str, detail: str = "", *, fg: str = theme.ON_FILL) -> Text:
    """``[ migrating ]  auditing both sides first``: a state pill, then the detail in subtext."""
    t = theme.pill(word, color, fg=fg)
    if detail:
        t.append("  ").append(detail, style=SUB)
    return t


def run_status_line(run: Path) -> str:
    """``340/1210 files · failed 2`` from a run folder's status.json (empty when unreadable)."""
    try:
        st = json.loads((run / "status.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return ""
    if not isinstance(st, dict):
        return ""
    return f"{st.get('done', '?')}/{st.get('total', '?')} files · failed {st.get('failed', 0)}"


class RunSection(Section):
    KEY, TITLE, NUMBER = "run", "Run", 4
    TAGLINE, IDLE = "migration dashboard · log", "m to migrate"
    PRIMARY = "#log"
    HELP = ("m starts the migration (the audit runs again first). p pauses after the files in progress and "
            "resumes in the same run folder; x stops; r resumes an earlier run. L toggles a full-height log.")
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("p", "pause", "pause/resume"),
        Binding("x", "stop", "stop"),
        Binding("r", "resume", "resume run"),
        Binding("L", "toggle_log", "big log"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.running = False
        self.paused = False
        self.force = False
        self.phase = ""
        self.run_dir: Path | None = None
        self.result: RunSummary | None = None  # adds up the legs of a paused/resumed run
        self.stats: Stats | None = None
        self.started_at: float | None = None
        self._pause_requested = False
        self._stop_requested = False
        self._worker: Worker[Any] | None = None
        self._throttle_until = 0.0

    def compose(self) -> ComposeResult:
        yield Static("", id="run-state", classes="run-state")
        yield Static("", id="throttle", classes="banner")
        with Vertical(id="dash"):
            with Horizontal(classes="bar-row"):
                yield Label("files", classes="bar-label")
                yield ProgressBar(id="files-bar", show_eta=False)
                yield Static("", id="files-count", classes="bar-count")
            with Horizontal(classes="bar-row"):
                yield Label("bytes", classes="bar-label")
                yield ProgressBar(id="bytes-bar", show_eta=False)
                yield Static("", id="bytes-count", classes="bar-count")
            yield Static("", id="run-stats", classes="stats")
            yield Static("", id="run-now", classes="now")
        log = VimLog(id="log")
        log.border_title = "log"
        yield log

    def on_mount(self) -> None:
        self.query_one("#throttle").display = False
        idle = state_line("idle", SURFACE1, fg=FG)
        idle.append("  ").append_text(theme.key_hints([("m", "migrate the job"), ("r", "resume an earlier run")]))
        self._state(idle)
        self._now(Text.assemble(theme.pill("now", SURFACE1, fg=SUB, bold=False), ("  nothing running", DIM)))
        self._redraw_stats()
        self.set_interval(1, self._tick)

    # -- log tail (every LogLine of every operation lands here) -------------------------------------

    def write_log(self, ev: LogLine) -> None:
        if ev.level != "debug":
            self.query_one(VimLog).write(theme.log_text(ev))

    # -- starting -----------------------------------------------------------------------------------

    def start(self, resume_run_dir: Path | None = None, *, confirm: bool = True) -> None:
        if self.running:
            self.m365.notify("A migration is already running.", severity="warning")
            return
        self._worker = self.run_migration(resume_run_dir, confirm)

    @work(group="migrate", exit_on_error=False)
    async def run_migration(self, resume_run_dir: Path | None, confirm: bool = True,
                            force: bool = False) -> None:
        app = self.m365
        if app.busy:
            app.notify(f"Busy: {app.busy} is running.", severity="warning")
            return
        if not app.endpoints_ready():
            app.notify("Set source and destination first (1 Job, e).", severity="warning")
            return
        blocking = app.last_audit.blocking if app.last_audit is not None else []
        if confirm and not force:
            if blocking:
                force = await self._ask_force(len(blocking))
                if not force:
                    return
            elif not await app.push_screen_wait(ConfirmModal("Start the migration?", self._confirm_body(
                    resume_run_dir))):
                return
        while True:
            outcome = await self._leg(resume_run_dir, force)
            if outcome != "blocked" or force:
                return
            force = await self._ask_force(len(app.last_audit.blocking) if app.last_audit else 0)
            if not force:
                return

    def _confirm_body(self, resume_run_dir: Path | None) -> Text:
        o = self.m365.opts
        t = Text()
        t.append("from  ", style=SUB).append(o.source.describe(), style=FG)
        t.append("\nto    ", style=SUB).append(o.destination.describe(), style=FG)
        if resume_run_dir is not None:
            t.append("\nrun   ", style=SUB).append(str(resume_run_dir), style=FG)
        t.append("\n\nThe audit runs again first; files are then copied to the destination.", style=SUB)
        return t

    async def _ask_force(self, n: int) -> bool:
        body = Text.assemble((f"The audit found {theme.plural(n, 'blocking finding')}.", f"bold {RED}"),
                             (("\nSee section 2 (Checks) for what and how to fix. "
                               "Migrating anyway can fail files or lose data."), FG))
        return bool(await self.m365.push_screen_wait(TypedConfirmModal("Blocking findings", body, "force")))

    async def _leg(self, resume_run_dir: Path | None, force: bool) -> str:
        """One call of job.migrate. Returns done | paused | blocked | failed | busy."""
        app = self.m365
        job = app.begin("migrate", self)
        if job is None:
            return "busy"
        self.running, self.paused, self.force = True, False, force
        self._pause_requested = self._stop_requested = False
        self.phase = ""
        self.started_at = self.started_at if resume_run_dir is not None and self.started_at else time.monotonic()
        if resume_run_dir is None:
            self.result = None
        self.status, self.summary = theme.RUNNING, "starting"
        self._state(state_line("◐ migrating", YELLOW, "auditing both sides first"))
        self.refresh_bindings()
        self._redraw_stats()
        try:
            summary: RunSummary = await job.migrate(resume_run_dir=resume_run_dir, force=force)
        except asyncio.CancelledError:
            self.run_dir = getattr(job, "run_dir", None) or self.run_dir
            if self._stop_requested:
                self.status, self.summary = theme.TODO, "stopped"
                self._state(state_line("■ stopped", YELLOW, "before any file was copied"))
            else:
                self.paused = True
                self.status, self.summary = theme.RUNNING, "paused"
                self._state(state_line("❚❚ paused", YELLOW, "before the transfer started · p resumes · x stops"))
            raise
        except BlockingFindingsError as exc:
            self.run_dir = exc.run_dir
            app.workspace.section("checks").show_findings(exc.findings, exc.run_dir, RunSummary(
                "audit", datetime.now(), datetime.now(), exc.run_dir, findings=exc.findings))
            self.status, self.summary = theme.FAILED, f"refused: {len(exc.blocking)} blocking"
            self._state(state_line("✗ refused", RED, str(exc)))
            return "blocked"
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.status, self.summary = theme.FAILED, "failed"
            self._state(state_line("✗ failed", RED, f"{type(exc).__name__}: {exc}"))
            app.report_error(exc, "migrate")
            return "failed"
        finally:
            self.running = False
            app.end()
            self.refresh_bindings()
        self.result = summary = self._accumulate(summary)
        self.run_dir = summary.run_dir
        if summary.stopped_early and self._pause_requested and not self._stop_requested:
            self.paused = True
            self.status, self.summary = theme.RUNNING, "paused"
            self._state(state_line("❚❚ paused", YELLOW, "p resumes from the journal · x stops"))
            return "paused"
        self._finished(summary)
        return "done"

    def _accumulate(self, summary: RunSummary) -> RunSummary:
        """After a resume, count files copied by the earlier legs (the engine reports them as skipped)."""
        prev = self.result
        if prev is None or prev.run_dir != summary.run_dir:
            return summary
        summary.started = prev.started
        summary.skipped = max(0, summary.skipped - prev.copied)
        summary.copied += prev.copied
        summary.bytes_copied += prev.bytes_copied
        summary.metadata_warnings += prev.metadata_warnings
        return summary

    def _finished(self, s: RunSummary) -> None:
        self.paused = False
        bad = s.failed > 0
        self.status = theme.FAILED if bad else theme.DONE
        self.summary = f"{s.copied:,} copied" + (f" · {s.failed} failed" if s.failed else "") + (
            " · stopped" if s.stopped_early else "")
        t = state_line("✗ finished with failures" if bad else ("■ stopped" if s.stopped_early else "● done"),
                       RED if bad else YELLOW if s.stopped_early else GREEN)
        t.append("  ")
        for i, (label, n) in enumerate((("copied", s.copied), ("skipped", s.skipped), ("failed", s.failed),
                                        ("metadata warnings", s.metadata_warnings))):
            t.append(" · " if i else "", style=DIM).append(f"{n:,}", style=f"bold {FG}").append(f" {label}", style=SUB)
        t.append(" · ", style=DIM).append(theme.size(s.bytes_copied), style=f"bold {FG}")
        t.append("\nreport  ", style=SUB).append(str(Path(s.run_dir) / "report.html"), style=f"underline {BLUE}")
        self._state(t)
        self.m365.notify(f"Migration {'finished with failures' if bad else 'done'}: {s.copied:,} copied, "
                         f"{s.failed} failed.", severity="error" if bad else "information", markup=False)

    # -- keys ---------------------------------------------------------------------------------------

    @property
    def transferring(self) -> bool:
        """True once the engine runs (only then request_stop() works); before that we cancel the worker."""
        return self.phase in TRANSFER_PHASES

    def check_action(self, action: str, parameters: tuple[object, ...]) -> bool | None:
        if action == "pause":
            return self.running or self.paused
        if action == "stop":
            return self.running or self.paused
        if action == "resume":
            return not self.running
        return True

    def _halt(self, stop: bool) -> None:
        if stop:
            self._stop_requested = True
        else:
            self._pause_requested = True
        if self.transferring:
            job = self.m365.current_job
            if job is not None:
                job.request_stop()
            self._state(state_line("stopping" if stop else "pausing", YELLOW, "finishing the files in progress ..."))
        elif self._worker is not None:
            self._worker.cancel()

    def action_pause(self) -> None:
        if self.running:
            self._halt(stop=False)
        elif self.paused and self.run_dir is not None:
            self._worker = self.run_migration(self.run_dir, False, self.force)

    def action_stop(self) -> None:
        if self.running:
            self._halt(stop=True)
        elif self.paused:
            self.paused = False
            if self.result is not None:
                self.result.stopped_early = True
                self._finished(self.result)
            else:
                self.status, self.summary = theme.TODO, "stopped"
                self._state(state_line("■ stopped", YELLOW))
            self.refresh_bindings()

    @work(group="pick", exclusive=True, exit_on_error=False)
    async def action_resume(self) -> None:
        app = self.m365
        if self.running or app.busy:
            app.notify("Something is running already.", severity="warning")
            return
        runs = list_runs("migrate")
        if not runs:
            app.notify("No earlier migrate runs found.", severity="warning")
            return
        choices = [Choice(f"{r.parent.name} / {r.name}", r, run_status_line(r)) for r in runs]
        run: Path | None = await app.push_screen_wait(PickerModal("Resume which run?", choices,
                                                                  note="Finished files are skipped (journal)."))
        if run is None:
            return
        try:
            opts = load_run_options(run, app.cfg)
        except Exception as exc:  # noqa: BLE001 - shown to the user
            app.report_error(exc, "resume")
            return
        app.set_options(opts)
        self.result, self.started_at = None, None
        if await app.push_screen_wait(ConfirmModal("Resume this run?", self._confirm_body(run))):
            self._worker = self.run_migration(run, False)

    def action_toggle_log(self) -> None:
        self.toggle_class("-biglog")

    # -- progress -----------------------------------------------------------------------------------

    def on_job_event(self, ev: ProgressEvent) -> None:
        if isinstance(ev, PhaseChanged):
            self.phase = ev.phase
            self.summary = phase_label(ev.phase)
            t = state_line("◐ migrating", YELLOW)
            t.append("  ").append(phase_label(ev.phase), style=f"bold {FG}")
            if ev.detail:
                t.append(f"  {ev.detail}", style=SUB)
            self._state(t)
        elif isinstance(ev, Stats):
            if not self.transferring:  # inventory scan, not transfer progress
                unit = {"users": "users", "checks": "checks"}.get(self.phase, "items")
                self._now(Text.assemble(theme.pill("scan", SURFACE1, fg=FG), "  ", (phase_label(self.phase), FG),
                                        (f"  {ev.done:,}/{ev.total:,} {unit}", SUB)))
                return
            self.stats = ev
            self.summary = f"{100 * ev.done // max(ev.total, 1)}% · {ev.done:,}/{ev.total:,}"
            self._redraw_stats()
        elif isinstance(ev, ItemStarted):
            self._now(Text.assemble(theme.pill("start", BLUE), "  ", (f"{ev.index:,}/{ev.total:,}  ", SUB),
                                    (ev.rel_path, FG), (f"  {theme.size(ev.size)}", SUB)))
        elif isinstance(ev, ItemStep):
            self._now(Text.assemble(theme.pill(ev.step, STEP_COLOR.get(ev.step, ACCENT)), "  ", (ev.rel_path, FG)))
        elif isinstance(ev, Throttled):
            self._throttle_until = time.monotonic() + ev.seconds
            banner = self.query_one("#throttle", Static)
            banner.update(Text(f" ⚠ throttled by SharePoint · waiting {ev.seconds:.0f}s · attempt {ev.attempt} "
                               f"· {ev.what}"))
            banner.display = True

    def _tick(self) -> None:
        if self.running:
            self._redraw_stats()
        banner = self.query_one("#throttle")
        if banner.display and time.monotonic() >= self._throttle_until:
            banner.display = False

    def _redraw_stats(self) -> None:
        s = self.stats
        files_bar = self.query_one("#files-bar", ProgressBar)
        bytes_bar = self.query_one("#bytes-bar", ProgressBar)
        if s is None:
            files_bar.update(total=None, progress=0)
            bytes_bar.update(total=None, progress=0)
        else:
            files_bar.update(total=max(s.total, 1), progress=s.done)
            bytes_bar.update(total=max(s.bytes_total, 1), progress=s.bytes_done)
        self.query_one("#files-count", Static).update(
            Text(f"{s.done:,}/{s.total:,}" if s else "", style=FG))
        self.query_one("#bytes-count", Static).update(
            Text(f"{theme.size(s.bytes_done)} / {theme.size(s.bytes_total)}" if s else "", style=FG))
        elapsed = time.monotonic() - self.started_at if self.started_at else None
        cells = [
            ("files", f"{s.done:,}/{s.total:,}" if s else "–", FG),
            ("copied", theme.size(s.bytes_done) if s else "–", FG),
            ("speed", f"{s.mbps:.2f} MB/s" if s else "–", BLUE if s else FG),
            ("ETA", theme.duration(s.eta_s) if s else "–", FG),
            ("elapsed", theme.duration(elapsed) if elapsed is not None else "–", FG),
            ("failed", f"{s.failed:,}" if s else "–", RED if s and s.failed else FG),
            ("warnings", f"{s.warnings:,}" if s else "–", YELLOW if s and s.warnings else FG),
            ("workers", str(self.m365.opts.workers), FG),
        ]
        self.query_one("#run-stats", Static).update(stat_grid(cells))
        self.query_one("#dash").display = s is not None or self.running

    def _state(self, text: Text) -> None:
        self.query_one("#run-state", Static).update(text)

    def _now(self, text: Text) -> None:
        self.query_one("#run-now", Static).update(text)

    def escape(self) -> bool:
        if self.has_class("-biglog"):
            self.remove_class("-biglog")
            return True
        return False
