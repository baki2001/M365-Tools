"""Pick an earlier migrate run to continue."""

from __future__ import annotations

import json
from pathlib import Path
from typing import ClassVar

from rich.text import Text
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.widgets import Button, Footer, Header, Label, ListItem, ListView

from m365tools.services import list_runs, load_run_options

from .base import BaseScreen


def _status(run: Path) -> str:
    try:
        st = json.loads((run / "status.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return ""
    if not isinstance(st, dict):
        return ""
    return f"{st.get('done', '?')}/{st.get('total', '?')} files, failed {st.get('failed', 0)}"


class RunItem(ListItem):
    def __init__(self, run: Path) -> None:
        super().__init__(Label(Text.assemble((run.parent.name, "bold"), f" / {run.name}  ",
                                             (_status(run), "dim"))))
        self.run = run


class ResumeScreen(BaseScreen):
    BINDINGS: ClassVar[list[BindingType]] = [Binding("escape", "back", "Back")]

    def compose(self) -> ComposeResult:
        yield Header()
        with Vertical():
            yield Label("Resume a migration run", classes="section-title")
            yield ListView(*[RunItem(r) for r in list_runs("migrate")], id="runs")
            yield Label("Resuming re-checks both sides, then copies what the journal does not list as done.",
                        classes="hint")
            with Horizontal(classes="buttons"):
                yield Button("Back", id="back")
        yield Footer()

    def on_mount(self) -> None:
        if not list_runs("migrate"):
            self.notify("No migrate runs found.", severity="warning")
        self.query_one("#runs", ListView).focus()

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "back":
            self.action_back()

    def on_list_view_selected(self, event: ListView.Selected) -> None:
        if not isinstance(event.item, RunItem):
            return
        try:
            opts = load_run_options(event.item.run, self.m365.cfg)
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.notify(str(exc), severity="error", timeout=10)
            return
        from .migrate import MigrateScreen

        self.m365.new_job(opts)
        self.app.push_screen(MigrateScreen(resume_run_dir=event.item.run))
