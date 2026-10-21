"""Run finished: summary, run folder and report paths, next steps."""

from __future__ import annotations

from typing import ClassVar

from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.widgets import Button, Footer, Header, Label

from m365tools.tools.spmig.models import RunSummary

from ..widgets import KeyValue, summary_rows
from .base import BaseScreen


class ReportDoneScreen(BaseScreen):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("v", "verify", "Verify"),
        Binding("l", "lock", "Lock source"),
        Binding("o", "open_report", "Open report"),
        Binding("h", "home", "Home"),
        Binding("escape", "home", "Home", show=False),
    ]

    def __init__(self, summary: RunSummary) -> None:
        super().__init__()
        self.summary = summary

    def compose(self) -> ComposeResult:
        s = self.summary
        yield Header()
        with Vertical():
            title = f"{s.mode} finished" + (" (stopped early)" if s.stopped_early else "")
            yield Label(title, classes="section-title")
            kv = KeyValue(id="summary")
            kv.set_rows(summary_rows(s) + [("Text report", str(s.run_dir / "report.txt"))])
            yield kv
            if s.stopped_early:
                yield Label("Resume later from Home > Resume run (continues from the journal).", classes="hint")
            with Horizontal(classes="buttons"):
                yield Button("Verify", id="verify", variant="primary")
                yield Button("Lock source", id="lock")
                yield Button("Open report path", id="report")
                yield Button("Home", id="home")
        yield Footer()

    def on_button_pressed(self, event: Button.Pressed) -> None:
        {"verify": self.action_verify, "lock": self.action_lock, "report": self.action_open_report,
         "home": self.action_home}.get(event.button.id or "", lambda: None)()

    def action_open_report(self) -> None:
        self.open_report(self.summary.run_dir)

    def action_verify(self) -> None:
        from .verify import VerifyScreen

        self.app.push_screen(VerifyScreen())

    def action_lock(self) -> None:
        from .lock import LockScreen

        self.app.push_screen(LockScreen())
