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
from ..theme import DIM, FG, KEY, MODE_STYLE, PEACH, STATE_COLOR, SUB, SURFACE0
from .base import Section
from .runs import RunInfo, list_all_runs


class Card(Static, can_focus=True):
    """A focusable rounded box; the job section renders its content."""


def _pairs(pairs: list[tuple[str, str]], value_style: str = FG) -> Text:
    t = Text()
    for i, (k, v) in enumerate(pairs):
        if i:
            t.append(" · ", style=DIM)
        t.append(f"{k} ", style=SUB).append(v, style=value_style)
    return t


def endpoint_text(ep: Endpoint, facts: str) -> Text:
    """Line 1: site name (bold) and URL. Line 2: tenant · library · folder · inventory facts when known."""
    if not ep.site_url:
        t = Text("not set", style=f"bold {PEACH}")
        t.append("\n").append_text(theme.key_hint("e", "choose tenant, site, library and folder"))
        return t
    t = Text()
    t.append(theme.short_site(ep.site_url), style=f"bold {FG}").append("  ").append(ep.site_url, style=SUB)
    t.append("\n").append_text(_pairs([("tenant", ep.tenant), ("library", ep.library or DEFAULT_LIBRARY_LABEL),
                                       ("folder", ep.folder or "library root")]))
    if facts:
        t.append(" · ", style=DIM).append(facts, style=theme.TEAL)
    return t


def options_text(o: JobOptions) -> Text:
    t = _pairs([("overwrite", o.overwrite), ("metadata", "on" if o.restore_metadata else "off"),
                ("workers", str(o.workers)), ("lock after", "on" if o.lock_source_after else "off")], KEY)
    t.append("\n").append_text(_pairs([("excludes", ", ".join(o.excludes) or "none"),
                                       ("user map", str(o.usermap_path) if o.usermap_path else "none")], KEY))
    for r in o.routes:
        t.append("\nroute ", style=SUB).append(r.prefix, style=KEY).append("  →  ", style=DIM)
        t.append(r.destination.describe(), style=KEY)
    return t


# ------------------------------------------------------------------------------------------------ pipeline

STEPS = (("job", "Job"), ("checks", "Checks"), ("plan", "Plan"), ("run", "Run"), ("verify", "Verify"),
         ("lock", "Lock"))
NUMERALS = "①②③④⑤⑥"
DONE_WORD = {"job": "ready", "checks": "clear", "plan": "planned", "run": "done", "verify": "ok", "lock": "locked"}
FAILED_WORD = {"checks": "blocked", "verify": "problems"}


def step_word(key: str, status: str, summary: str) -> str:
    if status == theme.DONE:
        return "previewed" if key == "lock" and "locked" not in summary else DONE_WORD[key]
    if status == theme.FAILED:
        return FAILED_WORD.get(key, "failed")
    if status == theme.RUNNING:
        return "paused" if summary == "paused" else "running"
    return "not set" if key == "job" else "todo"


def pipeline(states: dict[str, tuple[str, str]], width: int) -> Table:
    """``① Job → ② Checks → ...`` as pills coloured by state, with a one-word status under each."""
    arrow = " → " if width >= 70 else "›"
    grid = Table.grid(padding=0)
    chips: list[Text] = []
    words: list[Text] = []
    for i, (key, title) in enumerate(STEPS):
        status, summary = states.get(key, (theme.TODO, ""))
        color = STATE_COLOR[status]
        if i:
            grid.add_column(no_wrap=True)
            chips.append(Text(arrow, style=DIM))
            words.append(Text(""))
        grid.add_column(no_wrap=True, justify="center")
        label = f"{NUMERALS[i]} {title}"
        chips.append(theme.pill(label, color) if status != theme.TODO else
                     theme.pill(label, SURFACE0, fg=SUB, bold=False))
        words.append(Text(step_word(key, status, summary), style=color if status != theme.TODO else DIM))
    grid.add_row(*chips)
    grid.add_row(*words)
    return grid


class Pipeline(Static):
    """The six steps of a migration and where this job stands (from the sidebar states)."""

    def __init__(self, *, id: str | None = None, classes: str | None = None) -> None:
        super().__init__(id=id, classes=classes)
        self.states: dict[str, tuple[str, str]] = {}

    def set_states(self, states: dict[str, tuple[str, str]]) -> None:
        self.states = states
        self.refresh()

    def render(self) -> Table:
        return pipeline(self.states, self.content_region.width or 80)


def recent_runs(runs: list[RunInfo]) -> Table | Text:
    if not runs:
        return Text.assemble(("no runs yet · ", SUB), theme.chip("a"),
                             (" audits both sides; every run leaves a folder here (7 Runs)", SUB))
    t = Table(box=None, show_edge=False, pad_edge=False, expand=True, padding=(0, 2, 0, 0),
              header_style=f"bold {SUB}")
    t.add_column("when", no_wrap=True)
    t.add_column("mode", no_wrap=True)
    t.add_column("profile", no_wrap=True, max_width=18, overflow="ellipsis")
    t.add_column("result", ratio=1, no_wrap=True, overflow="ellipsis")
    for r in runs:
        color = MODE_STYLE.get(r.mode, SUB)
        t.add_row(Text(f"{r.started:%d %b %H:%M}" if r.started else "?", style=FG),
                  theme.glyph_word(r.mode, color), Text(r.profile, style=SUB), Text(r.summary, style=SUB))
    return t


class JobSection(Section):
    KEY, TITLE, NUMBER = "job", "Job", 1
    TAGLINE, IDLE = "where from, where to, and how", "not set · e to choose"
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
        pipe = Pipeline(id="pipeline", classes="panel")
        pipe.border_title = "pipeline"
        yield pipe
        runs = Static("", id="recent-runs", classes="panel")
        runs.border_title = "recent runs"
        yield runs

    def on_mount(self) -> None:
        self.refresh_cards()

    def on_show(self) -> None:
        self.refresh_runs()

    def refresh_pipeline(self) -> None:
        try:
            ws = self.m365.workspace
            states = {key: (ws.section(key).status, ws.section(key).summary) for key, _ in STEPS}
            self.query_one(Pipeline).set_states(states)
        except Exception:  # noqa: BLE001 - not mounted yet
            return

    def refresh_runs(self) -> None:
        try:
            runs = list_all_runs(limit=5)
        except OSError:
            runs = []
        self.query_one("#recent-runs", Static).update(recent_runs(runs))

    # -- rendering ----------------------------------------------------------------------------------

    def refresh_cards(self) -> None:
        app = self.m365
        o = app.opts
        src_facts, dst_facts = self._facts()
        self.query_one("#card-source", Card).update(endpoint_text(o.source, src_facts))
        self.query_one("#card-destination", Card).update(endpoint_text(o.destination, dst_facts))
        self.query_one("#card-options", Card).update(options_text(o))
        if not app.cfg.tenants:
            cmd = "quit and run: m365 tenant add <name> --tenant-id ... --client-id ... --host <tenant>.sharepoint.com"
            hint = Text.assemble(theme.pill("no tenants", PEACH), "  ", (cmd, SUB))
        elif not app.endpoints_ready():
            hint = Text.assemble(("next  ", SUB), theme.key_hints(
                [("e", "edit the focused card"), ("p", "load a profile"), ("o", "options")]))
        else:
            hint = Text.assemble(("next  ", SUB), theme.key_hints(
                [("a", "audit (writes nothing)"), ("d", "dry run"), ("m", "migrate"), ("s", "save profile")]))
        self.query_one("#job-hint", Static).update(hint)
        self.refresh_runs()
        if app.endpoints_ready():
            self.status = theme.DONE
            self.summary = src_facts.split(" · folders")[0] if src_facts else (
                f"{theme.short_site(o.source.site_url)} → {theme.short_site(o.destination.site_url)}")
        else:
            self.status, self.summary = theme.TODO, ""

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
