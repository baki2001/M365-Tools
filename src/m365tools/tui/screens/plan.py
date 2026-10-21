"""Dry run: simulate, then show the plan filterable by action."""

from __future__ import annotations

from typing import ClassVar

from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.widgets import Button, DataTable, Footer, Header, Label, ProgressBar, RichLog, Select, Static

from m365tools.core.events import ProgressEvent, Stats
from m365tools.core.report.format import fmt_size
from m365tools.tools.spmig.models import Action, Plan

from ..widgets import MAX_TABLE_ROWS
from .base import JobScreen

ACTION_STYLE = {Action.COPY: "green", Action.CREATE_FOLDER: "cyan", Action.SKIP: "dim",
                Action.CONFLICT: "yellow", Action.EXCLUDED: "magenta"}


class PlanScreen(JobScreen):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("escape", "back", "Back"),
        Binding("m", "migrate", "Migrate"),
        Binding("o", "open_report", "Open report"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.plan: Plan | None = None
        self.run_dir = None

    def compose(self) -> ComposeResult:
        yield Header()
        with Vertical():
            yield Label("Dry run (nothing is written)", classes="section-title")
            yield Static("starting ...", id="phase")
            yield Static("", id="throttle")
            yield Static("", id="current")
            yield ProgressBar(id="progress", show_eta=False)
            yield Static("", id="counts")
            with Horizontal(classes="row"):
                yield Label("Show", classes="field-label")
                yield Select([("all actions", "all")] + [(a.value, a.value) for a in Action], value="all",
                             allow_blank=False, id="filter")
            yield DataTable(id="entries", cursor_type="row", zebra_stripes=True)
            yield RichLog(id="log", classes="log-small", wrap=True, markup=False)
            with Horizontal(classes="buttons"):
                yield Button("Back", id="back")
                yield Button("Open report path", id="report", disabled=True)
                yield Button("Migrate", id="migrate", variant="primary", disabled=True)
        yield Footer()

    def on_mount(self) -> None:
        self.query_one("#entries", DataTable).add_columns("Action", "Path", "Size", "Reason / route")
        self.run_dryrun()

    def on_progress(self, ev: ProgressEvent) -> None:
        if isinstance(ev, Stats):  # dry run: count files, never bytes
            self.query_one("#progress", ProgressBar).update(total=max(ev.total, 1), progress=ev.done)

    @work(exclusive=True, group="job", exit_on_error=False)
    async def run_dryrun(self) -> None:
        self.running = True
        try:
            plan, summary = await self.state.job.dryrun()
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self._set("#phase", "dry run failed")
            self.error(exc)
            return
        finally:
            self.running = False
        self.plan, self.run_dir = plan, summary.run_dir
        counts = plan.counts()
        t = Text()
        for a in Action:
            t.append(f"{a.value}: {counts.get(a, 0):,}   ", style=ACTION_STYLE.get(a, ""))
        t.append(f"to copy: {fmt_size(plan.bytes_to_copy())}", style="bold")
        self._set("#counts", t)
        self._set("#phase", f"dry run done - run folder {summary.run_dir}")
        self.fill_table()
        self.query_one("#report", Button).disabled = False
        blocked = bool(self.state.audit and self.state.audit.blocking) and not self.state.force
        self.query_one("#migrate", Button).disabled = blocked

    def fill_table(self) -> None:
        table = self.query_one("#entries", DataTable)
        table.clear()
        if self.plan is None:
            return
        flt = self.query_one("#filter", Select).value
        rows = [e for e in self.plan.entries if flt == "all" or e.action.value == flt]
        for e in rows[:MAX_TABLE_ROWS]:
            reason = e.reason + (f" [route {e.route}]" if e.route else "")
            table.add_row(Text(e.action.value, style=ACTION_STYLE.get(e.action, "")), Text(e.item.rel_path),
                          "" if e.item.is_folder else fmt_size(e.item.size), Text(reason))
        if len(rows) > MAX_TABLE_ROWS:
            table.add_row("...", f"{len(rows) - MAX_TABLE_ROWS:,} more (see plan.json)", "", "")

    def on_select_changed(self, event: Select.Changed) -> None:
        if event.select.id == "filter":
            self.fill_table()

    def on_button_pressed(self, event: Button.Pressed) -> None:
        {"back": self.action_back, "report": self.action_open_report, "migrate": self.action_migrate}.get(
            event.button.id or "", lambda: None)()

    def action_open_report(self) -> None:
        self.open_report(self.run_dir)

    def action_migrate(self) -> None:
        if self.query_one("#migrate", Button).disabled:
            return
        from .migrate import MigrateScreen

        self.app.push_screen(MigrateScreen())
