"""Migration options + optional save as profile."""

from __future__ import annotations

from pathlib import Path
from typing import ClassVar

from rich.text import Text
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, VerticalScroll
from textual.widgets import (
    Button,
    Footer,
    Header,
    Input,
    Label,
    RadioButton,
    RadioSet,
    Static,
    Switch,
    TextArea,
)

from m365tools.core.config import save_config
from m365tools.services import opts_to_profile

from .base import BaseScreen

OVERWRITE = [
    ("skip_if_same_or_newer", "Skip if destination is same or newer (default)"),
    ("always", "Always overwrite"),
    ("never", "Never overwrite existing files"),
]


class OptionsScreen(BaseScreen):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("escape", "back", "Back"),
        Binding("ctrl+n", "audit", "Audit"),
        Binding("ctrl+e", "edit_endpoints", "Edit source/destination"),
    ]

    def compose(self) -> ComposeResult:
        o = self.state.opts
        yield Header()
        with VerticalScroll(classes="options"):
            yield Label("Options", classes="section-title")
            yield Static(id="endpoints")
            yield Label("Overwrite", classes="field-title")
            with RadioSet(id="overwrite"):
                for value, label in OVERWRITE:
                    yield RadioButton(label, value=o.overwrite == value, name=value)
            with Horizontal(classes="row"):
                yield Label("Restore metadata (dates, author, editor)", classes="field-label wide")
                yield Switch(value=o.restore_metadata, id="metadata")
            with Horizontal(classes="row"):
                yield Label("Lock source (read-only) after a clean migration", classes="field-label wide")
                yield Switch(value=o.lock_source_after, id="lock-after")
            with Horizontal(classes="row"):
                yield Label("Workers (1-16)", classes="field-label wide")
                yield Input(value=str(o.workers), type="integer", id="workers", classes="short")
            yield Label("Excludes (one glob per line, e.g. 0. Archief/**)", classes="field-title")
            yield TextArea("\n".join(o.excludes), id="excludes")
            with Horizontal(classes="row"):
                yield Label("User map CSV (optional)", classes="field-label wide")
                yield Input(value=str(o.usermap_path or ""), id="usermap")
            with Horizontal(classes="row"):
                yield Label("Save as profile (optional)", classes="field-label wide")
                yield Input(value=o.profile or "", placeholder="profile name", id="profile")
            with Horizontal(classes="buttons"):
                yield Button("Back", id="back")
                yield Button("Edit source/destination", id="edit")
                yield Button("Audit", id="audit", variant="primary")
        yield Footer()

    def on_mount(self) -> None:
        self.on_screen_resume()

    def on_screen_resume(self) -> None:
        o = self.state.opts
        t = Text()
        t.append("Source       ", style="bold").append(f"[{o.source.tenant}] {o.source.describe()}\n")
        t.append("Destination  ", style="bold").append(f"[{o.destination.tenant}] {o.destination.describe()}")
        for r in o.routes:
            t.append("\nRoute        ", style="bold").append(f"{r.prefix} -> {r.destination.describe()}")
        self.query_one("#endpoints", Static).update(t)

    def on_button_pressed(self, event: Button.Pressed) -> None:
        {"back": self.action_back, "edit": self.action_edit_endpoints, "audit": self.action_audit}.get(
            event.button.id or "", lambda: None)()

    def action_edit_endpoints(self) -> None:
        from .endpoint import SourceScreen

        self.app.push_screen(SourceScreen())

    def collect(self) -> bool:
        o = self.state.opts
        if not (o.source.tenant and o.source.site_url and o.destination.tenant and o.destination.site_url):
            self.notify("Source and destination are not complete.", severity="warning")
            return False
        try:
            workers = int(self.query_one("#workers", Input).value or "2")
        except ValueError:
            workers = 0
        if not 1 <= workers <= 16:
            self.notify("Workers must be between 1 and 16.", severity="warning")
            return False
        pressed = self.query_one("#overwrite", RadioSet).pressed_button
        o.overwrite = (pressed.name if pressed is not None and pressed.name else  # type: ignore[assignment]
                       "skip_if_same_or_newer")
        o.restore_metadata = self.query_one("#metadata", Switch).value
        o.lock_source_after = self.query_one("#lock-after", Switch).value
        o.workers = workers
        o.excludes = [x.strip() for x in self.query_one("#excludes", TextArea).text.splitlines() if x.strip()]
        usermap = self.query_one("#usermap", Input).value.strip()
        o.usermap_path = Path(usermap).expanduser() if usermap else None
        profile = self.query_one("#profile", Input).value.strip()
        if profile:
            cfg = self.m365.cfg
            o.profile = profile
            try:
                cfg.profiles[profile] = opts_to_profile(o)
                cfg.profiles[profile].name = profile
                save_config(cfg)
            except Exception as exc:  # noqa: BLE001 - shown to the user
                self.notify(f"Could not save profile: {exc}", severity="error", timeout=10)
                return False
            self.notify(f"Saved profile '{profile}'.")
        return True

    def action_audit(self) -> None:
        if not self.collect():
            return
        from .audit import AuditScreen

        self.m365.new_job()
        self.app.push_screen(AuditScreen())
