"""Tenants: list, add/update, sign in."""

from __future__ import annotations

from typing import ClassVar

from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Grid, Horizontal, Vertical
from textual.widgets import Button, DataTable, Footer, Header, Input, Label

from m365tools.core.config import add_tenant, save_config
from m365tools.core.errors import M365Error

from ..modals import TextModal
from .base import BaseScreen


class TenantsScreen(BaseScreen):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("escape", "back", "Back"),
        Binding("l", "login", "Login selected"),
        Binding("h", "help", "Setup help"),
        Binding("ctrl+s", "save", "Save tenant"),
    ]

    def compose(self) -> ComposeResult:
        yield Header()
        with Vertical():
            yield Label("Tenants", classes="section-title")
            yield DataTable(id="tenant-table", cursor_type="row", zebra_stripes=True)
            yield Label("Add / update tenant", classes="section-title")
            with Grid(id="tenant-form"):
                yield Label("Name")
                yield Input(placeholder="contoso", id="t-name")
                yield Label("Tenant id")
                yield Input(placeholder="contoso.onmicrosoft.com or GUID", id="t-tenant-id")
                yield Label("Client id")
                yield Input(placeholder="Application (client) id", id="t-client-id")
                yield Label("SharePoint host")
                yield Input(placeholder="contoso.sharepoint.com", id="t-host")
            with Horizontal(classes="buttons"):
                yield Button("Save tenant", id="save", variant="primary")
                yield Button("Login selected", id="login")
                yield Button("Setup help", id="help")
                yield Button("Back", id="back")
        yield Footer()

    def on_mount(self) -> None:
        table = self.query_one("#tenant-table", DataTable)
        table.add_columns("Name", "Tenant id", "Client id", "SharePoint host", "Signed in")
        self.refresh_table()

    def refresh_table(self) -> None:
        from m365tools.core.auth import cache_file_for

        table = self.query_one("#tenant-table", DataTable)
        table.clear()
        for name, t in sorted(self.m365.cfg.tenants.items()):
            table.add_row(Text(name), Text(t.tenant_id), Text(t.client_id), Text(t.sharepoint_host),
                          "cached" if cache_file_for(t).exists() else "no", key=name)

    def selected(self) -> str | None:
        table = self.query_one("#tenant-table", DataTable)
        if table.row_count == 0:
            return None
        row = table.coordinate_to_cell_key(table.cursor_coordinate).row_key
        return str(row.value) if row.value is not None else None

    def on_data_table_row_highlighted(self, event: DataTable.RowHighlighted) -> None:
        name = event.row_key.value
        t = self.m365.cfg.tenants.get(str(name)) if name is not None else None
        if t is None:
            return
        self.query_one("#t-name", Input).value = t.name
        self.query_one("#t-tenant-id", Input).value = t.tenant_id
        self.query_one("#t-client-id", Input).value = t.client_id
        self.query_one("#t-host", Input).value = t.sharepoint_host

    def on_button_pressed(self, event: Button.Pressed) -> None:
        {"save": self.action_save, "login": self.action_login, "help": self.action_help,
         "back": self.action_back}.get(event.button.id or "", lambda: None)()

    def action_save(self) -> None:
        cfg = self.m365.cfg
        values = {k: self.query_one(f"#t-{k}", Input).value.strip()
                  for k in ("name", "tenant-id", "client-id", "host")}
        if not all(values.values()):
            self.notify("Fill in all four fields.", severity="warning")
            return
        try:
            add_tenant(cfg, values["name"], tenant_id=values["tenant-id"], client_id=values["client-id"],
                       sharepoint_host=values["host"])
            path = save_config(cfg)
        except (M365Error, OSError) as exc:
            self.notify(str(exc), severity="error", timeout=10)
            return
        forget = getattr(self.m365.clients, "forget", None)
        if forget is not None:
            forget(values["name"])
        self.notify(f"Saved tenant '{values['name']}' to {path}")
        self.refresh_table()

    def action_help(self) -> None:
        from m365tools.core.auth import tenant_setup_instructions

        tid = self.query_one("#t-tenant-id", Input).value.strip() or "<yourtenant>.onmicrosoft.com"
        self.app.push_screen(TextModal("App registration setup", tenant_setup_instructions(tid)))

    def action_login(self) -> None:
        name = self.selected()
        if not name:
            self.notify("Select a tenant first.", severity="warning")
            return
        self.login(name)

    @work(exclusive=True, group="login", exit_on_error=False)
    async def login(self, name: str) -> None:
        self.notify(f"Signing in to '{name}' ...")
        try:
            accounts = await self.m365.clients.login(name)
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.m365.close_auth_prompt()
            self.notify(f"{type(exc).__name__}: {exc}", title="Sign-in failed", severity="error", timeout=10)
            return
        self.m365.close_auth_prompt()
        self.notify(f"Signed in to '{name}' as {', '.join(accounts) or '?'}", title="Signed in")
        self.refresh_table()
