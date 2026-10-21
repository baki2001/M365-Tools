"""Lock the source library read-only: preview, then apply after confirmation."""

from __future__ import annotations

from typing import ClassVar

from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.widgets import Button, Checkbox, DataTable, Footer, Header, Label, RichLog, Static

from m365tools.tools.spmig.lock import LockChange

from ..modals import ConfirmModal
from .base import JobScreen


class LockScreen(JobScreen):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("escape", "back", "Back"),
        Binding("a", "apply", "Apply lock"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.changes: list[LockChange] = []
        self.applied = False

    def compose(self) -> ComposeResult:
        yield Header()
        with Vertical():
            yield Label("Lock source library (read-only)", classes="section-title")
            yield Static(Text(f"Source: {self.state.opts.source.describe()}"), id="target")
            yield Static("previewing ...", id="phase")
            yield Static("", id="throttle")
            yield Static("", id="current")
            yield DataTable(id="changes", cursor_type="row", zebra_stripes=True)
            yield Label("Only principals that can write are set to Read; nothing is ever widened.",
                        classes="hint")
            yield RichLog(id="log", classes="log-small", wrap=True, markup=False)
            with Horizontal(classes="buttons"):
                yield Button("Back", id="back")
                yield Checkbox("Include site owners group", id="owners")
                yield Button("Apply lock", id="apply", variant="error", disabled=True)
        yield Footer()

    def on_mount(self) -> None:
        self.query_one("#changes", DataTable).add_columns("Principal", "Current roles", "After")
        self.run_lock(apply=False)

    @work(exclusive=True, group="job", exit_on_error=False)
    async def run_lock(self, apply: bool) -> None:
        self.running = True
        self.query_one("#apply", Button).disabled = True
        owners = self.query_one("#owners", Checkbox).value
        try:
            changes = await self.state.job.lock(apply=apply, include_owners=owners)
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self._set("#phase", "lock failed" if apply else "preview failed")
            self.error(exc)
            return
        finally:
            self.running = False
        self.changes, self.applied = changes, apply
        self.fill()
        n = sum(1 for c in changes if c.will_change)
        run_dir = getattr(self.state.job, "run_dir", None)
        if apply:
            self._set("#phase", Text(f"applied: {n} principal(s) set to Read - run folder {run_dir}",
                                     style="bold green"))
            self.notify(f"Source locked: {n} principal(s) set to Read.")
        else:
            self._set("#phase", f"preview: {n} principal(s) would be set to Read")
            self.query_one("#apply", Button).disabled = n == 0

    def fill(self) -> None:
        table = self.query_one("#changes", DataTable)
        table.clear()
        for c in self.changes:
            after = (Text(c.to_role, style="yellow") if c.will_change
                     else Text(f"unchanged ({c.skipped_reason})", style="dim"))
            table.add_row(Text(c.title), Text(", ".join(c.from_roles)), after)

    def on_checkbox_changed(self, event: Checkbox.Changed) -> None:
        if event.checkbox.id == "owners" and not self.running and not self.applied:
            self.run_lock(apply=False)

    def on_button_pressed(self, event: Button.Pressed) -> None:
        {"back": self.action_back, "apply": self.action_apply}.get(event.button.id or "", lambda: None)()

    def action_apply(self) -> None:
        if self.query_one("#apply", Button).disabled:
            return
        n = sum(1 for c in self.changes if c.will_change)
        names = ", ".join(c.title for c in self.changes if c.will_change)

        def done(ok: bool | None) -> None:
            if ok:
                self.run_lock(apply=True)

        self.app.push_screen(
            ConfirmModal("Apply lock?", f"Set {n} principal(s) to Read on {self.state.opts.source.describe()}:\n"
                         f"{names}", confirm_label="Apply lock"),
            done,
        )
