"""Live migration dashboard with pause/resume and stop."""

from __future__ import annotations

import asyncio
from pathlib import Path
from typing import Any, ClassVar

from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.widgets import Footer, Header, Label, ProgressBar, RichLog, Static

from m365tools.core.events import ItemStarted, ItemStep, ProgressEvent, Stats
from m365tools.core.report.format import fmt_duration, fmt_size
from m365tools.tools.spmig.job import BlockingFindingsError
from m365tools.tools.spmig.models import RunSummary

from .base import JobScreen

TRANSFER_PHASES = ("folders", "files", "finalize")


class MigrateScreen(JobScreen):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("p", "pause", "Pause/resume"),
        Binding("q", "stop", "Stop"),
        Binding("escape", "back", "Back"),
    ]

    def __init__(self, resume_run_dir: Path | None = None) -> None:
        super().__init__()
        self.run_dir: Path | None = resume_run_dir
        self.paused = False
        self._pause_requested = False
        self._stop_requested = False
        self.summary: RunSummary | None = None
        self._worker: Any = None

    def compose(self) -> ComposeResult:
        yield Header()
        with Vertical():
            yield Label("Migration", classes="section-title")
            yield Static("starting (audit first) ...", id="phase")
            yield ProgressBar(id="progress", show_eta=False)
            yield Static("", id="stats")
            yield Static("", id="throttle")
            yield Static("", id="current")
            with Horizontal(classes="row"):
                yield Label("", id="state")
            yield RichLog(id="log", wrap=True, markup=False, max_lines=2000)
        yield Footer()

    def on_mount(self) -> None:
        self.start(self.run_dir)

    # ------------------------------------------------------------------------------------------ worker

    def start(self, resume_run_dir: Path | None) -> None:
        self.paused = self._pause_requested = self._stop_requested = False
        self.phase = ""
        self._set("#state", "")
        self._worker = self.run_migrate(resume_run_dir)

    @property
    def transferring(self) -> bool:
        """True once the engine runs (it alone honours request_stop); before that we cancel instead."""
        return self.phase in TRANSFER_PHASES

    @work(exclusive=True, group="job", exit_on_error=False)
    async def run_migrate(self, resume_run_dir: Path | None) -> None:
        self.running = True
        try:
            summary = await self.state.job.migrate(resume_run_dir=resume_run_dir, force=self.state.force)
        except asyncio.CancelledError:
            self.running = False
            self.run_dir = getattr(self.state.job, "run_dir", None) or self.run_dir
            if self._stop_requested:
                self._set("#state", Text("Stopped before any file was copied. Esc to go back.", style="yellow"))
            else:
                self.paused = True
                self._set("#state", Text("Paused before the transfer started. Press p to resume.",
                                         style="bold yellow"))
            raise
        except BlockingFindingsError as exc:
            self.running = False
            self.run_dir = exc.run_dir
            log = self.query_one("#log", RichLog)
            for f in exc.blocking:
                log.write(Text(f"BLOCKING {f.category}: {f.path} - {f.detail}", style="bold red"))
            self._set("#phase", Text(str(exc), style="bold red"))
            self._set("#state", "Migration refused: fix the blocking findings or tick Force on the audit page.")
            return
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.running = False
            self._set("#phase", "migration failed")
            self.error(exc)
            return
        self.running = False
        summary = self._accumulate(summary)
        self.summary = summary
        self.run_dir = summary.run_dir
        if summary.stopped_early and self._pause_requested and not self._stop_requested:
            self.paused = True
            self._set("#state", Text("Paused. Press p to resume (continues from the journal), q to finish.",
                                     style="bold yellow"))
            return
        self.finish()

    def _accumulate(self, summary: RunSummary) -> RunSummary:
        """After a resume, count files copied by the earlier legs (the engine reports them as skipped)."""
        prev = self.summary
        if prev is None or prev.run_dir != summary.run_dir:
            return summary
        summary.started = prev.started
        summary.skipped = max(0, summary.skipped - prev.copied)
        summary.copied += prev.copied
        summary.bytes_copied += prev.bytes_copied
        summary.metadata_warnings += prev.metadata_warnings
        return summary

    def finish(self) -> None:
        from .done import ReportDoneScreen

        if self.summary is not None:
            self.app.switch_screen(ReportDoneScreen(self.summary))

    # ------------------------------------------------------------------------------------------ events

    def on_progress(self, ev: ProgressEvent) -> None:
        if isinstance(ev, Stats):
            bar = self.query_one("#progress", ProgressBar)
            if ev.bytes_total:
                bar.update(total=ev.bytes_total, progress=ev.bytes_done)
            else:
                bar.update(total=max(ev.total, 1), progress=ev.done)
            eta = fmt_duration(ev.eta_s) if ev.eta_s is not None else "--:--:--"
            t = Text()
            t.append("files ", style="bold").append(f"{ev.done:,}/{ev.total:,}   ")
            t.append("bytes ", style="bold").append(f"{fmt_size(ev.bytes_done)}/{fmt_size(ev.bytes_total)}   ")
            t.append(f"{ev.mbps:.2f} MB/s   ETA {eta}   ")
            t.append(f"failed {ev.failed}   ", style="bold red" if ev.failed else "")
            t.append(f"warnings {ev.warnings}", style="yellow" if ev.warnings else "")
            self._set("#stats", t)
        elif isinstance(ev, ItemStarted):
            self._set("#current", Text(f"{ev.index}/{ev.total}  {ev.rel_path}  ({fmt_size(ev.size)})"))
        elif isinstance(ev, ItemStep):
            self._set("#current", Text.assemble((f"{ev.step:<9}", "bold cyan"), ev.rel_path))

    # ------------------------------------------------------------------------------------------ keys

    def action_pause(self) -> None:
        if self.running:
            self._pause_requested = True
            if self.transferring:
                self.state.job.request_stop()
                self._set("#state", Text("Pausing: finishing files in progress ...", style="yellow"))
            else:
                self._worker.cancel()
        elif self.paused:
            self.start(self.run_dir)

    def action_stop(self) -> None:
        if self.running:
            self._stop_requested = True
            if self.transferring:
                self.state.job.request_stop()
                self._set("#state", Text("Stopping: finishing files in progress ...", style="yellow"))
            else:
                self._worker.cancel()
        elif self.summary is not None:
            self.finish()
        else:
            self.action_back()

    def action_back(self) -> None:
        if self.running:
            self.notify("A migration is running: press q to stop it first.", severity="warning")
            return
        super().action_back()
