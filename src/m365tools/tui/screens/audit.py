"""Audit: read-only checks, findings table, gate for migrate."""

from __future__ import annotations

from typing import ClassVar

from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.widgets import Button, Checkbox, Footer, Header, Label, RichLog, Static

from m365tools.tools.spmig.models import RunSummary

from ..widgets import FindingsTable, severity_counts
from .base import JobScreen


class AuditScreen(JobScreen):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("escape", "back", "Back"),
        Binding("d", "dryrun", "Dry run"),
        Binding("m", "migrate", "Migrate"),
        Binding("o", "open_report", "Open report"),
        Binding("r", "rerun", "Re-run audit"),
        Binding("f", "toggle_force", "Force", show=False),
    ]

    def compose(self) -> ComposeResult:
        yield Header()
        with Vertical():
            yield Label("Audit (read-only)", classes="section-title")
            yield Static("starting ...", id="phase")
            yield Static("", id="throttle")
            yield Static("", id="current")
            yield Static("", id="counts")
            yield FindingsTable(id="findings")
            yield RichLog(id="log", classes="log-small", wrap=True, markup=False)
            with Horizontal(classes="buttons"):
                yield Button("Back", id="back")
                yield Button("Open report path", id="report", disabled=True)
                yield Checkbox("Force (ignore blocking)", value=self.state.force, id="force")
                yield Button("Dry run", id="dryrun", disabled=True)
                yield Button("Migrate", id="migrate", variant="primary", disabled=True)
        yield Footer()

    def on_mount(self) -> None:
        self.run_audit()

    @property
    def summary(self) -> RunSummary | None:
        return self.state.audit

    @work(exclusive=True, group="job", exit_on_error=False)
    async def run_audit(self) -> None:
        self.running = True
        self.update_buttons()
        try:
            summary = await self.state.job.audit()
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self._set("#phase", "audit failed")
            self.error(exc)
            return
        finally:
            self.running = False
        self.state.audit = summary
        self.query_one("#findings", FindingsTable).load(summary.findings)
        self._set("#counts", severity_counts(summary.findings))
        self._set("#phase", f"audit done - run folder {summary.run_dir}")
        if summary.blocking:
            self.notify(f"{len(summary.blocking)} blocking finding(s): migrate is disabled unless you force it.",
                        severity="warning", timeout=8)
        self.update_buttons()

    def update_buttons(self) -> None:
        done = self.summary is not None and not self.running
        blocked = bool(self.summary and self.summary.blocking) and not self.state.force
        self.query_one("#dryrun", Button).disabled = not done
        self.query_one("#report", Button).disabled = not done
        self.query_one("#migrate", Button).disabled = not done or blocked

    def on_checkbox_changed(self, event: Checkbox.Changed) -> None:
        self.state.force = event.value
        self.update_buttons()

    def on_button_pressed(self, event: Button.Pressed) -> None:
        {"back": self.action_back, "report": self.action_open_report, "dryrun": self.action_dryrun,
         "migrate": self.action_migrate}.get(event.button.id or "", lambda: None)()

    def action_toggle_force(self) -> None:
        cb = self.query_one("#force", Checkbox)
        cb.value = not cb.value

    def action_open_report(self) -> None:
        self.open_report(self.summary.run_dir if self.summary else None)

    def action_rerun(self) -> None:
        if not self.running:
            self.state.audit = None
            self.run_audit()

    def action_dryrun(self) -> None:
        if self.query_one("#dryrun", Button).disabled:
            return
        from .plan import PlanScreen

        self.app.push_screen(PlanScreen())

    def action_migrate(self) -> None:
        if self.query_one("#migrate", Button).disabled:
            if self.summary and self.summary.blocking:
                self.notify("Blocking findings: tick 'Force' to migrate anyway.", severity="warning")
            return
        from .migrate import MigrateScreen

        self.app.push_screen(MigrateScreen())
