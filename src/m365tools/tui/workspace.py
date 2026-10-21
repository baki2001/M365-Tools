"""The one screen: header, sidebar (sections), main area (the selected section), key bar.

Navigation keys live here; each section adds its own keys while focus is inside it.

    1-7            show section            tab / shift+tab  toggle sidebar <-> main
    h  [           focus sidebar           l  ]  enter      focus main
    esc            close / back            a d m v          audit, dry run, migrate, verify (from anywhere)
    :              command palette         ?                help for the current section
    ctrl+l         sign in                 q                quit
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, ClassVar

from rich.table import Table
from rich.text import Text
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.screen import Screen
from textual.widgets import ContentSwitcher, OptionList

from m365tools.core.events import LogLine

from .modals import TextModal
from .sections import SECTIONS, Section
from .theme import ACCENT, DIM, FG
from .widgets import HeaderBar, KeyBar, Sidebar

if TYPE_CHECKING:
    from .app import M365App
    from .sections import RunSection


class Workspace(Screen[None]):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("tab,shift+tab", "toggle_pane", "switch pane", show=False),
        Binding("h,left_square_bracket", "focus_sidebar", "sidebar", show=False),
        Binding("l,right_square_bracket", "focus_main", "main", show=False),
        Binding("escape", "back", "back", show=False),
        *[Binding(str(s.NUMBER), f"show({s.NUMBER})", s.TITLE, show=False) for s in SECTIONS],
        Binding("a", "audit", "audit"),
        Binding("d", "dryrun", "dry run"),
        Binding("m", "migrate", "migrate"),
        Binding("v", "verify", "verify"),
        Binding("colon", "app.command_palette", "commands", key_display=":"),
        Binding("question_mark", "help", "help", key_display="?"),
        Binding("ctrl+l", "app.login", "sign in", show=False),
        Binding("q", "app.quit_app", "quit"),
    ]

    @property
    def m365(self) -> M365App:
        return self.app  # type: ignore[return-value]

    def compose(self) -> ComposeResult:
        yield HeaderBar(id="header")
        with Horizontal(id="body"):
            sidebar = Sidebar([(s.KEY, s.TITLE) for s in SECTIONS], id="sidebar")
            sidebar.border_title = "sections"
            yield sidebar
            with Vertical(id="main"), ContentSwitcher(id="switcher", initial=SECTIONS[0].KEY):
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
        self.query_one("#main").border_title = f"{cls.NUMBER} {cls.TITLE}"

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
        if self._main_focused() and not self.current.escape():
            self.action_focus_sidebar()

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

    # -- help ---------------------------------------------------------------------------------------

    def action_help(self) -> None:
        sec = self.current
        t = Table.grid(padding=(0, 2))
        t.add_column(style=f"bold {ACCENT}", no_wrap=True, min_width=14)
        t.add_column(style=FG)

        def heading(title: str) -> None:
            t.add_row("", "")
            t.add_row(Text(title, style="bold"), "")

        def keys(bindings: list[BindingType]) -> None:
            for b in bindings:
                if isinstance(b, Binding):
                    t.add_row(b.key_display or self._keys(b.key), b.description)

        t.add_row(Text(f"{sec.NUMBER} {sec.TITLE}", style=f"bold {ACCENT}"), Text(sec.HELP, style=DIM))
        heading("This section")
        keys(list(type(sec).BINDINGS))
        heading("Lists and tables")
        for k, d in (("j / k", "move down / up"), ("g / G", "top / bottom"), ("ctrl+d / ctrl+u", "page"),
                     ("/", "filter (where shown)"), ("enter", "open / details")):
            t.add_row(k, d)
        heading("Move around")
        for k, d in (("1 - 7", "show section"), ("tab / shift+tab", "switch between sidebar and main"),
                     ("h  [", "focus sidebar"), ("l  ]  enter", "focus main (in Lock, l previews)"),
                     ("esc", "close panel / back to sidebar")):
            t.add_row(k, d)
        heading("Anywhere")
        for k, d in (("a", "audit: inventory + checks, writes nothing"), ("d", "dry run: full plan"),
                     ("m", "migrate (asks first)"), ("v", "verify"), (":", "command palette"),
                     ("ctrl+l", "sign in to a tenant"), ("?", "this help"), ("q", "quit")):
            t.add_row(k, d)
        self.app.push_screen(TextModal(f"Keys · {sec.TITLE}", t))

    @staticmethod
    def _keys(key: str) -> str:
        names = {"slash": "/", "colon": ":", "question_mark": "?", "left_square_bracket": "[",
                 "right_square_bracket": "]", "plus": "+", "minus": "-"}
        return " ".join(names.get(k, k) for k in key.split(","))
