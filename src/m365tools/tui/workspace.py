"""The one screen: header, sidebar (sections), main area (the selected section), key bar.

Navigation keys live here; each section adds its own keys while focus is inside it.

    1-7            show section            tab / shift+tab  toggle sidebar <-> main
    h  [           focus sidebar           l  ]  enter      focus main
    esc            close / back            a d m v          audit, dry run, migrate, verify (from anywhere)
    :              command palette         ?                help for the current section
    ctrl+l         sign in                 q                quit
    g              back to the toolbox home
"""

from __future__ import annotations

from functools import partial
from typing import TYPE_CHECKING, Any, ClassVar

from rich.console import Group
from rich.table import Table
from rich.text import Text
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.screen import Screen
from textual.widget import Widget
from textual.widgets import ContentSwitcher, OptionList

from m365tools.core.events import LogLine

from . import theme
from .modals import TextModal
from .sections import SECTIONS, Section
from .theme import FG, SUB
from .widgets import HeaderBar, KeyBar, NavStatus, Sidebar

if TYPE_CHECKING:
    from .app import M365App
    from .sections import RunSection


class MainTitle(Widget):
    """``1 · Job`` in bold on the left, the section's one-line purpose in subtext on the right."""

    DEFAULT_CSS = "MainTitle { height: 1; margin: 0 0 1 0; }"

    def __init__(self, *, id: str | None = None) -> None:
        super().__init__(id=id)
        self.section: type[Section] = SECTIONS[0]

    def show(self, cls: type[Section]) -> None:
        self.section = cls
        self.refresh()

    def render(self) -> Table:
        cls = self.section
        grid = Table.grid(expand=True)
        grid.add_column(no_wrap=True)
        grid.add_column(justify="right", ratio=1, no_wrap=True, overflow="ellipsis")
        grid.add_row(Text.assemble((f"{cls.NUMBER} · {cls.TITLE}", f"bold {FG}")), Text(cls.TAGLINE, style=SUB))
        return grid


class Workspace(Screen[None]):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("tab,shift+tab", "toggle_pane", "switch pane", show=False),
        Binding("h,left_square_bracket", "focus_sidebar", "sidebar", show=False),
        Binding("l,right_square_bracket", "focus_main", "main", show=False),
        Binding("escape", "back", "back", show=False),
        *[Binding(str(s.NUMBER), f"show({s.NUMBER})", s.TITLE, show=False) for s in SECTIONS],
        Binding("a", "audit", "audit", show=False),
        Binding("d", "dryrun", "dry run", show=False),
        Binding("m", "migrate", "migrate", show=False),
        Binding("v", "verify", "verify", show=False),
        Binding("colon", "app.command_palette", "commands", key_display=":", show=False),
        Binding("question_mark", "help", "help", key_display="?"),
        Binding("ctrl+l", "app.login", "sign in", show=False),
        Binding("g", "app.go_home", "home", priority=True, show=False),
        Binding("ctrl+1", "app.tool(1)", "spmig", show=False, priority=True),
        Binding("ctrl+2", "app.tool(2)", "groups", show=False, priority=True),
        Binding("ctrl+3", "app.tool(3)", "onboard", show=False, priority=True),
        Binding("q", "app.quit_app", "quit", show=False),
    ]

    @property
    def m365(self) -> M365App:
        return self.app  # type: ignore[return-value]

    def compose(self) -> ComposeResult:
        yield HeaderBar("spmig", id="header")
        with Horizontal(id="body"):
            with Vertical(id="nav"):
                yield Sidebar([(s.KEY, s.TITLE, s.IDLE) for s in SECTIONS], id="sidebar")
                yield NavStatus(id="nav-status")
            with Vertical(id="main"):
                yield MainTitle(id="main-title")
                with ContentSwitcher(id="switcher", initial=SECTIONS[0].KEY):
                    for cls in SECTIONS:
                        yield cls()
        yield KeyBar(id="keybar")

    def on_mount(self) -> None:
        self.sidebar.highlighted = 0
        self._title_main(SECTIONS[0])
        self.sidebar.focus()

    # -- lookups ------------------------------------------------------------------------------------

    @property
    def sidebar(self) -> Sidebar:
        return self.query_one(Sidebar)

    def section(self, key: str) -> Any:
        return self.query_one(f"#{key}", Section)

    @property
    def current(self) -> Section:
        return self.section(self.query_one(ContentSwitcher).current or SECTIONS[0].KEY)

    @property
    def run_section(self) -> RunSection:
        return self.section("run")

    def write_log(self, ev: LogLine) -> None:
        self.run_section.write_log(ev)

    # -- section switching --------------------------------------------------------------------------

    def show_section(self, key: str) -> None:
        switcher = self.query_one(ContentSwitcher)
        main_focused = self._main_focused()
        if switcher.current != key:
            switcher.current = key
            sec = self.section(key)
            self._title_main(type(sec))
            sec.on_show()
            idx = [s.KEY for s in SECTIONS].index(key)
            if self.sidebar.highlighted != idx:
                self.sidebar.highlighted = idx
        if main_focused:
            self.current.focus_primary()

    def _title_main(self, cls: type[Section]) -> None:
        self.query_one(MainTitle).show(cls)

    def _main_focused(self) -> bool:
        f = self.focused
        return f is not None and self.query_one("#main") in f.ancestors

    def on_option_list_option_highlighted(self, event: OptionList.OptionHighlighted) -> None:
        if event.option_list is self.sidebar and event.option.id:
            self.show_section(event.option.id)

    def on_option_list_option_selected(self, event: OptionList.OptionSelected) -> None:
        if event.option_list is self.sidebar:
            self.action_focus_main()

    def on_section_changed(self, event: Section.Changed) -> None:
        s = event.section
        self.sidebar.set_entry(s.KEY, s.status, s.summary)
        self.section("job").refresh_pipeline()

    # -- navigation actions -------------------------------------------------------------------------

    def action_show(self, number: int) -> None:
        self.show_section(SECTIONS[number - 1].KEY)

    def action_toggle_pane(self) -> None:
        if self._main_focused():
            self.action_focus_sidebar()
        else:
            self.action_focus_main()

    def action_focus_sidebar(self) -> None:
        self.sidebar.focus()

    def action_focus_main(self) -> None:
        self.current.focus_primary()

    def action_back(self) -> None:
        if self._main_focused():
            if not self.current.escape():
                self.action_focus_sidebar()
        else:
            self.m365.action_go_home()

    # -- job actions (from anywhere) ----------------------------------------------------------------

    def _go(self, key: str) -> Any:
        """Show section ``key`` and focus its pane, so the results can be browsed right away."""
        self.show_section(key)
        sec = self.section(key)
        sec.focus_primary()
        return sec

    def action_audit(self) -> None:
        self._go("checks").run_audit()

    def action_dryrun(self) -> None:
        self._go("plan").run_dryrun()

    def action_migrate(self) -> None:
        self._go("run").start()

    def action_verify(self) -> None:
        self._go("verify").run_verify()

    def action_lock_preview(self) -> None:
        self._go("lock").action_preview()

    def action_lock_apply(self) -> None:
        self._go("lock").action_apply()

    def action_resume(self) -> None:
        self._go("run").action_resume()

    # -- palette ------------------------------------------------------------------------------------

    def palette_commands(self) -> list[tuple[str, Any, str]]:
        app = self.m365
        cmds: list[tuple[str, Any, str]] = [
            ("Audit", self.action_audit, "inventory both sides and run every check; writes nothing  (a)"),
            ("Dry run", self.action_dryrun, "compute the full plan; writes nothing  (d)"),
            ("Migrate", self.action_migrate, "copy files to the destination, asks first  (m)"),
            ("Resume a migration run", self.action_resume, "continue an earlier run from its journal  (4, r)"),
            ("Verify", self.action_verify, "compare source and destination  (v)"),
            ("Lock: preview", self.action_lock_preview, "who would be set to Read on the source  (6, l)"),
            ("Lock: apply", self.action_lock_apply, "downgrade writers on the source to Read  (6, A)"),
            ("Help", self.action_help, "keys for the current section  (?)"),
        ]
        for name in sorted(app.cfg.profiles):
            cmds.append((f"Load profile {name}", partial(app.load_profile, name), "replace the job with it"))
        for i, cls in enumerate(SECTIONS, 1):
            cmds.append((f"Go to {cls.TITLE}", partial(self.action_show, i), f"({i})"))
        return cmds

    # -- help ---------------------------------------------------------------------------------------

    def action_help(self) -> None:
        sec = self.current
        own = [(b.key_display or self._keys(b.key), b.description) for b in type(sec).BINDINGS
               if isinstance(b, Binding)]
        lists = [("j / k", "move down / up"), ("home / G", "top / bottom"), ("ctrl+d / u", "page"),
                 ("/", "filter (where shown)"), ("enter", "open / details")]
        move = [("1 - 7", "show section"), ("tab", "switch sidebar / main"), ("h  [", "focus sidebar"),
                ("l  ]", "focus main (in Lock, l previews)"), ("esc", "close panel / back to sidebar")]
        anywhere = [("a", "audit: inventory + checks, writes nothing"), ("d", "dry run: full plan"),
                    ("m", "migrate (asks first)"), ("v", "verify"), (":", "command palette"),
                    ("ctrl+l", "sign in to a tenant"), ("?", "this help"), ("g", "toolbox home"), ("q", "quit")]
        left = Group(help_group(f"this section · {sec.NUMBER} {sec.TITLE}", own),
                     Text(""), help_group("lists and tables", lists))
        right = Group(help_group("global · move around", move), Text(""), help_group("global · anywhere", anywhere),
                      Text(""), help_group("back and switch", BACK_HELP))
        cols = Table.grid(expand=True, padding=(0, 4))
        cols.add_column(ratio=1)
        cols.add_column(ratio=1)
        cols.add_row(left, right)
        body = Group(Text(sec.HELP, style=SUB), Text(""), cols)
        self.app.push_screen(TextModal(f"Keys · {sec.NUMBER} {sec.TITLE}", body, fit=True))

    @staticmethod
    def _keys(key: str) -> str:
        names = {"slash": "/", "colon": ":", "question_mark": "?", "left_square_bracket": "[",
                 "right_square_bracket": "]", "plus": "+", "minus": "-"}
        return " ".join(names.get(k, k) for k in key.split(","))


BACK_HELP = [("esc", "back one step; from the top of a tool: home"),
             ("g  1/2/3", "home, then spmig / groups / onboard"),
             ("ctrl+1 2 3", "jump straight there (not every terminal sends it)"), ("q", "quit (asks first)")]


def help_group(title: str, rows: list[tuple[str, str]]) -> Group:
    """A heading in bold mauve, then one row per key: a chip and what it does."""
    t = Table.grid(padding=(0, 2))
    t.add_column(no_wrap=True, width=12)
    t.add_column(style=FG)
    for key, desc in rows:
        t.add_row(theme.chip(key), desc)
    return Group(Text(title, style=f"bold {theme.ACCENT}"), t)
