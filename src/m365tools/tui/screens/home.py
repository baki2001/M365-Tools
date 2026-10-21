"""Home: stored profiles plus the main actions."""

from __future__ import annotations

from typing import ClassVar

from rich.text import Text
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.widgets import Button, Footer, Header, Label, ListItem, ListView

from m365tools.core.config import profile_to_job_options
from m365tools.core.errors import ConfigError

from .base import BaseScreen


class ProfileItem(ListItem):
    def __init__(self, name: str, desc: str) -> None:
        super().__init__(Label(Text.assemble((name, "bold"), "\n", (desc, "dim"))))
        self.profile_name = name


class HomeScreen(BaseScreen):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("n", "new", "New migration"),
        Binding("r", "resume", "Resume run"),
        Binding("t", "tenants", "Tenants"),
        Binding("q", "app.quit", "Quit"),
    ]

    def compose(self) -> ComposeResult:
        yield Header()
        with Vertical(id="home-body"):
            yield Label("Profiles", classes="section-title")
            yield ListView(id="profiles")
            yield Label("", id="home-hint", classes="hint")
            with Horizontal(classes="buttons"):
                yield Button("New migration", id="new", variant="primary")
                yield Button("Resume run", id="resume")
                yield Button("Tenants", id="tenants")
                yield Button("Quit", id="quit", variant="error")
        yield Footer()

    def on_mount(self) -> None:
        self.refresh_profiles()

    def on_screen_resume(self) -> None:
        self.refresh_profiles()

    def refresh_profiles(self) -> None:
        cfg = self.m365.cfg
        lv = self.query_one("#profiles", ListView)
        lv.clear()
        for name, p in sorted(cfg.profiles.items()):
            desc = f"{p.source.to_endpoint().describe()}  ->  {p.destination.to_endpoint().describe()}"
            lv.append(ProfileItem(name, desc))
        hint = ""
        if not cfg.tenants:
            hint = "No tenants configured yet: press t to add one."
        elif not cfg.profiles:
            hint = "No profiles yet: start a new migration (n); you can save it as a profile on the options page."
        self.query_one("#home-hint", Label).update(hint)
        if cfg.profiles:
            lv.focus()

    def on_list_view_selected(self, event: ListView.Selected) -> None:
        item = event.item
        if not isinstance(item, ProfileItem):
            return
        try:
            self.state.opts = profile_to_job_options(self.m365.cfg, item.profile_name)
        except ConfigError as exc:
            self.notify(str(exc), severity="error", timeout=10)
            return
        self.state.force = False
        from .options import OptionsScreen

        self.app.push_screen(OptionsScreen())

    def on_button_pressed(self, event: Button.Pressed) -> None:
        action = {"new": self.action_new, "resume": self.action_resume, "tenants": self.action_tenants,
                  "quit": self.app.exit}.get(event.button.id or "")
        if action:
            action()

    def action_new(self) -> None:
        from m365tools.tui.app import blank_options

        from .endpoint import SourceScreen

        self.state.opts = blank_options()
        self.state.force = False
        self.app.push_screen(SourceScreen())

    def action_resume(self) -> None:
        from .resume import ResumeScreen

        self.app.push_screen(ResumeScreen())

    def action_tenants(self) -> None:
        from .tenants import TenantsScreen

        self.app.push_screen(TenantsScreen())
