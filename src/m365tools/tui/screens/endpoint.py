"""Source / destination pickers: tenant, debounced site search, library, folder."""

from __future__ import annotations

from typing import Any, ClassVar

from rich.text import Text
from textual import on, work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.timer import Timer
from textual.widgets import Button, DataTable, Footer, Header, Input, Label, ListItem, ListView, Select

from m365tools.core.config import tenant_for_url
from m365tools.services import DEFAULT_LIBRARY_LABEL
from m365tools.tools.spmig.models import Endpoint

from .base import BaseScreen

_DEFAULT_KEY = "\x00default"


class SiteItem(ListItem):
    def __init__(self, title: str, url: str) -> None:
        super().__init__(Label(Text.assemble((title or url, "bold"), "  ", (url, "dim"))))
        self.url = url


class EndpointScreen(BaseScreen):
    """Shared picker. Subclasses set ``side`` ("source" / "destination") and ``next_screen``."""

    side = "source"
    heading = "Source"
    DEBOUNCE_S = 0.4
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("escape", "back", "Back"),
        Binding("ctrl+n", "next", "Next"),
        Binding("ctrl+l", "load_libraries", "Load libraries"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self._search_timer: Timer | None = None
        self.library: str | None = None

    @property
    def endpoint(self) -> Endpoint:
        return getattr(self.state.opts, self.side)

    def compose(self) -> ComposeResult:
        tenants = sorted(self.m365.cfg.tenants)
        ep = self.endpoint
        yield Header()
        with Vertical(classes="endpoint"):
            yield Label(self.heading, classes="section-title")
            with Horizontal(classes="row"):
                yield Label("Tenant", classes="field-label")
                yield Select([(t, t) for t in tenants], id="tenant", allow_blank=True, prompt="Select tenant",
                             value=ep.tenant if ep.tenant in tenants else Select.NULL)
            with Horizontal(classes="row"):
                yield Label("Search sites", classes="field-label")
                yield Input(placeholder="type part of a site name ...", id="search")
            yield ListView(id="hits")
            with Horizontal(classes="row"):
                yield Label("Site URL", classes="field-label")
                yield Input(value=ep.site_url, placeholder="https://contoso.sharepoint.com/sites/...",
                            id="site")
            with Horizontal(classes="row"):
                yield Label("Library", classes="field-label")
                yield Label(Text(ep.library or DEFAULT_LIBRARY_LABEL), id="library-label")
            yield DataTable(id="libraries", cursor_type="row")
            with Horizontal(classes="row"):
                yield Label("Folder", classes="field-label")
                yield Input(value=ep.folder, placeholder="(library root) or e.g. Klanten/2024", id="folder")
            with Horizontal(classes="buttons"):
                yield Button("Back", id="back")
                yield Button("Load libraries", id="load")
                yield Button("Next", id="next", variant="primary")
        yield Footer()

    def on_mount(self) -> None:
        self.library = self.endpoint.library
        table = self.query_one("#libraries", DataTable)
        table.add_column("Library")
        tenants = self.m365.cfg.tenants
        sel = self.query_one("#tenant", Select)
        if sel.value is Select.NULL:
            if self.endpoint.site_url and (t := tenant_for_url(self.m365.cfg, self.endpoint.site_url)):
                sel.value = t.name
            elif self.side == "destination" and self.state.opts.source.tenant in tenants:
                sel.value = self.state.opts.source.tenant
            elif len(tenants) == 1:
                sel.value = next(iter(tenants))
        if not tenants:
            self.notify("No tenants configured: add one on the Tenants screen first.", severity="warning")
        self.query_one("#search", Input).focus()

    # ------------------------------------------------------------------------------------------ helpers

    def tenant(self) -> str | None:
        v = self.query_one("#tenant", Select).value
        return None if v is Select.NULL else str(v)

    # ------------------------------------------------------------------------------------------ search

    @on(Input.Changed, "#search")
    def _search_changed(self, event: Input.Changed) -> None:
        if self._search_timer is not None:
            self._search_timer.stop()
        query = event.value.strip()
        if not query:
            self.query_one("#hits", ListView).clear()
            return
        self._search_timer = self.set_timer(self.DEBOUNCE_S, lambda: self.search(query))

    @on(Input.Submitted, "#search")
    def _search_submitted(self, event: Input.Submitted) -> None:
        if self._search_timer is not None:
            self._search_timer.stop()
        if event.value.strip():
            self.search(event.value.strip())

    @work(exclusive=True, group="search", exit_on_error=False)
    async def search(self, query: str) -> None:
        tenant = self.tenant()
        if not tenant:
            self.notify("Select a tenant first.", severity="warning")
            return
        try:
            hits: list[Any] = await self.m365.clients.search_sites(tenant, query)
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.notify(f"Site search failed: {exc}", severity="error", timeout=8)
            return
        self.m365.close_auth_prompt()
        lv = self.query_one("#hits", ListView)
        await lv.clear()
        await lv.extend([SiteItem(h.title, h.url) for h in hits])
        if not hits:
            self.notify(f"No sites match '{query}'.", severity="warning")

    @on(ListView.Selected, "#hits")
    def _hit_selected(self, event: ListView.Selected) -> None:
        if isinstance(event.item, SiteItem):
            self.query_one("#site", Input).value = event.item.url
            self.load_libraries(event.item.url)

    @on(Input.Submitted, "#site")
    def _site_submitted(self, event: Input.Submitted) -> None:
        self.action_load_libraries()

    # ------------------------------------------------------------------------------------------ libraries

    def action_load_libraries(self) -> None:
        url = self.query_one("#site", Input).value.strip()
        if url:
            self.load_libraries(url)
        else:
            self.notify("Enter or pick a site URL first.", severity="warning")

    @work(exclusive=True, group="libraries", exit_on_error=False)
    async def load_libraries(self, site_url: str) -> None:
        tenant = self.tenant() or (t.name if (t := tenant_for_url(self.m365.cfg, site_url)) else None)
        if not tenant:
            self.notify("Select a tenant first.", severity="warning")
            return
        table = self.query_one("#libraries", DataTable)
        table.loading = True
        try:
            names = await self.m365.clients.list_libraries(tenant, site_url)
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.notify(f"Could not list libraries: {exc}", severity="error", timeout=8)
            return
        finally:
            table.loading = False
        self.m365.close_auth_prompt()
        table.clear()
        table.add_row(DEFAULT_LIBRARY_LABEL, key=_DEFAULT_KEY)
        for n in names:
            table.add_row(Text(n), key=n)
        table.focus()

    @on(DataTable.RowSelected, "#libraries")
    def _library_selected(self, event: DataTable.RowSelected) -> None:
        key = event.row_key.value
        self.library = None if key in (None, _DEFAULT_KEY) else str(key)
        self.query_one("#library-label", Label).update(Text(self.library or DEFAULT_LIBRARY_LABEL))

    # ------------------------------------------------------------------------------------------ nav

    def on_button_pressed(self, event: Button.Pressed) -> None:
        {"back": self.action_back, "load": self.action_load_libraries, "next": self.action_next}.get(
            event.button.id or "", lambda: None)()

    def action_next(self) -> None:
        tenant = self.tenant()
        site = self.query_one("#site", Input).value.strip().rstrip("/")
        if not tenant or not site:
            self.notify("Choose a tenant and a site.", severity="warning")
            return
        ep = self.endpoint
        ep.tenant, ep.site_url, ep.library = tenant, site, self.library
        ep.folder = self.query_one("#folder", Input).value.strip().strip("/")
        from .options import OptionsScreen

        if self.side == "destination" and any(isinstance(s, OptionsScreen) for s in self.app.screen_stack):
            while not isinstance(self.app.screen, OptionsScreen):  # editing from the options page
                self.app.pop_screen()
            return
        self.app.push_screen(self.next_screen())

    def next_screen(self) -> BaseScreen:  # pragma: no cover - overridden
        raise NotImplementedError


class SourceScreen(EndpointScreen):
    side = "source"
    heading = "Source: where the files are now"

    def next_screen(self) -> BaseScreen:
        return DestinationScreen()


class DestinationScreen(EndpointScreen):
    side = "destination"
    heading = "Destination: where the files go"

    def next_screen(self) -> BaseScreen:
        from .options import OptionsScreen

        return OptionsScreen()
