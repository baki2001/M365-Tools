"""Section 1 - Job: where from, where to, and how.

Three stacked cards (Source, Destination, Options). ``e`` edits the focused card: for an endpoint that is
tenant -> site search -> library -> folder, each a small picker; for the options card the options editor.
"""

from __future__ import annotations

from dataclasses import replace
from typing import Any, ClassVar

from rich.table import Table
from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.widgets import Static

from m365tools.core.config import save_config
from m365tools.core.errors import M365Error
from m365tools.services import DEFAULT_LIBRARY_LABEL, opts_to_profile
from m365tools.tools.spmig.models import Endpoint, JobOptions

from .. import theme
from ..modals import Choice, OptionsModal, PickerModal, PromptModal
from ..theme import ACCENT, DIM, FG, YELLOW
from .base import Section


class Card(Static, can_focus=True):
    """A focusable bordered box; the job section renders its content."""


def endpoint_table(ep: Endpoint, facts: str) -> Table | Text:
    if not ep.site_url:
        return Text("not set  ·  press e to choose tenant, site and library", style=YELLOW)
    t = Table.grid(padding=(0, 2))
    t.add_column(style=DIM, no_wrap=True)
    t.add_column(style=FG)
    t.add_row("tenant", ep.tenant)
    t.add_row("site", ep.site_url)
    t.add_row("library", ep.library or DEFAULT_LIBRARY_LABEL)
    t.add_row("folder", ep.folder or "(library root)")
    if facts:
        t.add_row("found", Text(facts, style=ACCENT))
    return t


def options_text(o: JobOptions) -> Text:
    t = Text()
    parts = [("overwrite", o.overwrite), ("metadata", "on" if o.restore_metadata else "off"),
             ("workers", str(o.workers)), ("excludes", ", ".join(o.excludes) or "none"),
             ("lock after", "on" if o.lock_source_after else "off")]
    if o.usermap_path:
        parts.append(("user map", str(o.usermap_path)))
    for i, (k, v) in enumerate(parts):
        if i:
            t.append("  ·  ", style=DIM)
        t.append(f"{k} ", style=DIM).append(v, style=FG)
    for r in o.routes:
        t.append("\nroute ", style=DIM).append(r.prefix, style=FG).append("  →  ", style=DIM)
        t.append(r.destination.describe(), style=FG)
    return t


class JobSection(Section):
    KEY, TITLE, NUMBER = "job", "Job", 1
    PRIMARY = "#card-source"
    HELP = ("Choose the source and destination library. e edits the focused card; o edits the options; "
            "s saves everything as a profile, p loads one.")
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("j,down", "card(1)", "move", key_display="j/k"),
        Binding("k,up", "card(-1)", "up", show=False),
        Binding("e,enter", "edit", "edit"),
        Binding("o", "options", "options"),
        Binding("s", "save_profile", "save profile"),
        Binding("p", "load_profile", "load profile"),
    ]
    CARDS = ("card-source", "card-destination", "card-options")

    def compose(self) -> ComposeResult:
        for card_id, title in zip(self.CARDS, ("Source", "Destination", "Options"), strict=True):
            card = Card(id=card_id, classes="card")
            card.border_title = title
            yield card
        yield Static("", id="job-hint", classes="hint")

    def on_mount(self) -> None:
        self.refresh_cards()

    # -- rendering ----------------------------------------------------------------------------------

    def refresh_cards(self) -> None:
        app = self.m365
        o = app.opts
        src_facts, dst_facts = self._facts()
        self.query_one("#card-source", Card).update(endpoint_table(o.source, src_facts))
        self.query_one("#card-destination", Card).update(endpoint_table(o.destination, dst_facts))
        self.query_one("#card-options", Card).update(options_text(o))
        if not app.cfg.tenants:
            hint = Text("No tenants configured. Quit and run: m365 tenant add <name> --tenant-id ... "
                        "--client-id ... --host <tenant>.sharepoint.com", style=YELLOW)
        elif not app.endpoints_ready():
            hint = Text("Pick a source and a destination (e), or load a saved profile (p).", style=DIM)
        else:
            hint = Text("Ready. a audits both sides (nothing is written), d dry-runs, m migrates.", style=DIM)
        self.query_one("#job-hint", Static).update(hint)
        if app.endpoints_ready():
            self.status = theme.DONE
            self.summary = src_facts.split(" · folders")[0] if src_facts else (
                f"{theme.short_site(o.source.site_url)} → {theme.short_site(o.destination.site_url)}")
        else:
            self.status, self.summary = theme.TODO, "not set"

    def _facts(self) -> tuple[str, str]:
        """Inventory facts from the job's last operation (empty before the first audit)."""
        state = getattr(self.m365.current_job, "state", None)
        out = []
        for side in ("src", "dst"):
            inv = getattr(getattr(state, side, None), "inv", None)
            if inv is None:
                out.append("")
                continue
            files = inv.files()
            out.append(f"{len(files):,} files · {theme.size(inv.total_bytes())} · folders {len(inv.folders()):,}")
        return out[0], out[1]

    # -- keys ---------------------------------------------------------------------------------------

    def _focused_card(self) -> str:
        f = self.screen.focused
        return f.id if f is not None and f.id in self.CARDS else self.CARDS[0]

    def action_card(self, delta: int) -> None:
        i = self.CARDS.index(self._focused_card())
        self.query_one(f"#{self.CARDS[(i + delta) % len(self.CARDS)]}").focus()

    def action_edit(self) -> None:
        card = self._focused_card()
        if card == "card-options":
            self.action_options()
        else:
            self.edit_endpoint("source" if card == "card-source" else "destination")

    def action_options(self) -> None:
        def done(opts: JobOptions | None) -> None:
            if opts is not None:
                self.m365.set_options(opts)

        self.app.push_screen(OptionsModal(self.m365.opts), done)

    @work(exclusive=True, group="edit", exit_on_error=False)
    async def edit_endpoint(self, side: str) -> None:
        """tenant -> site -> library -> folder. Esc at any step leaves the endpoint unchanged."""
        app = self.m365
        tenants = sorted(app.cfg.tenants)
        if not tenants:
            app.notify("No tenants configured: run 'm365 tenant add' first.", severity="error")
            return
        current: Endpoint = getattr(app.opts, side)
        tenant = current.tenant if current.tenant in tenants else (app.tenant or tenants[0])
        if len(tenants) > 1:
            tenant = await app.push_screen_wait(PickerModal(
                f"{side.title()}: tenant", [Choice(t, t, app.cfg.tenants[t].sharepoint_host) for t in tenants],
                highlight=tenant))
            if tenant is None:
                return

        async def search(query: str) -> list[Choice]:
            hits = await app.clients.search_sites(tenant, query)
            return [Choice(h.title or theme.short_site(h.url), h.url, h.url) for h in hits]

        site = await app.push_screen_wait(PickerModal(
            f"{side.title()}: site in {tenant}", search=search, placeholder="type part of the site name, or paste a URL",
            accept_text=lambda s: s.rstrip("/") if s.lower().startswith("https://") else None,
            note=f"current: {current.site_url}" if current.site_url else ""))
        if site is None:
            return

        async def libraries() -> list[Choice]:
            names = await app.clients.list_libraries(tenant, site)
            return [Choice(DEFAULT_LIBRARY_LABEL, "", "the site's default document library"),
                    *[Choice(n, n) for n in sorted(names, key=str.lower)]]

        library = await app.push_screen_wait(PickerModal(f"{side.title()}: library", loader=libraries,
                                                         note=site))
        if library is None:
            return
        folder = await app.push_screen_wait(PromptModal(
            f"{side.title()}: folder (optional)", current.folder if current.site_url == site else "",
            placeholder="empty = library root, or e.g. Klanten/2024"))
        if folder is None:
            return
        ep = Endpoint(tenant=tenant, site_url=site, library=library or None, folder=folder.strip().strip("/"))
        app.set_options(replace(app.opts, **{side: ep}))
        app.notify(f"{side.title()} set to {ep.describe()}", markup=False)

    @work(exclusive=True, group="edit", exit_on_error=False)
    async def action_save_profile(self) -> None:
        app = self.m365
        if not app.endpoints_ready():
            app.notify("Set source and destination first.", severity="warning")
            return
        name = await app.push_screen_wait(PromptModal("Save profile as", app.opts.profile or "",
                                                      placeholder="e.g. finance-2026"))
        name = (name or "").strip()
        if not name:
            return
        try:
            prof = opts_to_profile(app.opts)
            prof.name = name
            app.cfg.profiles[name] = prof
            path = save_config(app.cfg)
        except (M365Error, OSError, ValueError) as exc:
            app.report_error(exc, "save profile")
            return
        app.set_options(replace(app.opts, profile=name))
        app.notify(f"Profile '{name}' saved to {path}", markup=False)

    @work(exclusive=True, group="edit", exit_on_error=False)
    async def action_load_profile(self) -> None:
        app = self.m365
        profiles = app.cfg.profiles
        if not profiles:
            app.notify("No profiles yet: set up a job and press s to save one.", severity="warning")
            return
        choices = [Choice(n, n, f"{theme.short_site(p.source.site_url)} → {theme.short_site(p.destination.site_url)}")
                   for n, p in sorted(profiles.items())]
        name: Any = await app.push_screen_wait(PickerModal("Load profile", choices, highlight=app.opts.profile))
        if name is not None:
            app.load_profile(name)
