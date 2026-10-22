"""The toolbox home: the mark, a one-line tagline and the tools, each with its last run.

    1 2 3  enter   open a tool          j / k   move
    t              switch tenant        ctrl+l  sign in
    :              command palette      ?       help          q  quit
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, ClassVar

from rich.console import Group
from rich.table import Table
from rich.text import Text
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Vertical
from textual.screen import Screen
from textual.widgets import OptionList, Static
from textual.widgets.option_list import Option

from . import theme
from .modals import TextModal
from .theme import CRUST, DIM, FG, KEY, MAUVE, SUB
from .widgets import BRAND, MARK, BarList, BarOption, HeaderBar, KeyBar

if TYPE_CHECKING:
    from .app import M365App


@dataclass(frozen=True, slots=True)
class Tool:
    key: str
    number: int
    title: str


TOOLS = (Tool("spmig", 1, "SharePoint library migration"), Tool("groups", 2, "access groups, drift and the registry"),
         Tool("onboard", 3, "bring people in and out"))
TAGLINE = "Microsoft 365 admin from the terminal. Keyboard first; nothing is written until you confirm."


@dataclass(slots=True)
class LastRun:
    mode: str
    started: datetime | None
    where: str


def _parse(run: Path) -> LastRun:
    parts = run.name.split("-")  # YYYYmmdd-HHMMSS-mode[-n]
    try:
        started: datetime | None = datetime.strptime(f"{parts[0]}-{parts[1]}", "%Y%m%d-%H%M%S")
    except (IndexError, ValueError):
        started = None
    mode = parts[2] if len(parts) > 2 else "?"
    return LastRun(mode, started, run.parent.name)


def last_run(tool: str, tenant: str | None) -> LastRun | None:
    """Newest run folder of ``tool`` (spmig: any profile; groups: the current tenant)."""
    from m365tools.services import list_group_runs

    if tool in ("groups", "onboard"):
        runs = list_group_runs(tenant) if tenant else list_group_runs()
        people = ("onboard", "offboard")
        runs = [r for r in runs if (_parse(r).mode in people) == (tool == "onboard")]
        return _parse(runs[0]) if runs else None
    from m365tools.core.paths import runs_dir

    root = runs_dir()
    if not root.exists():
        return None
    runs = [p for p in root.glob("*/*-*") if p.is_dir() and not p.parent.name.startswith("groups-")]
    if not runs:
        return None
    return _parse(max(runs, key=lambda p: p.name))


def ago(when: datetime | None) -> str:
    if when is None:
        return "?"
    secs = (datetime.now() - when).total_seconds()
    if secs < 90:
        return "just now"
    if secs < 3600:
        return f"{secs / 60:.0f} min ago"
    if secs < 36 * 3600:
        return f"{secs / 3600:.0f} h ago"
    return f"{secs / 86400:.0f} days ago"


def tool_entry(tool: Tool, last: LastRun | None, selected: bool) -> Table:
    """``1  spmig   SharePoint library migration          audit · 2 h ago``."""
    grid = Table.grid(expand=True, padding=(0, 0))
    grid.add_column(width=5, no_wrap=True)
    grid.add_column(width=10, no_wrap=True)
    grid.add_column(ratio=1, no_wrap=True, overflow="ellipsis")
    grid.add_column(justify="right", no_wrap=True)
    if last is None:
        run = Text("no runs yet", style=CRUST if selected else DIM)
    else:
        run = Text.assemble(("last ", CRUST if selected else DIM), (last.mode, CRUST if selected else FG),
                            (f" · {ago(last.started)}", CRUST if selected else SUB))
    if selected:
        grid.add_row(Text(f" {tool.number}", style=f"bold {CRUST}"), Text(tool.key, style=f"bold {CRUST}"),
                     Text(tool.title, style=CRUST), run)
    else:
        grid.add_row(Text(f" {tool.number}", style=f"bold {KEY}"), Text(tool.key, style=f"bold {FG}"),
                     Text(tool.title, style=SUB), run)
    return grid


def mark() -> Table:
    """The mark (a diamond in a frame) beside the name and the tagline."""
    logo = Text.assemble(("╭───╮\n", MAUVE), ("│ ", MAUVE), (MARK, f"bold {theme.LAVENDER}"), (" │\n", MAUVE),
                         ("╰───╯", MAUVE))
    words = Text.assemble((BRAND, f"bold {FG}"), "\n", ("tools for Microsoft 365 admins", SUB), "\n",
                          ("spmig · groups · onboard", DIM))
    grid = Table.grid(padding=(0, 2))
    grid.add_column(no_wrap=True)
    grid.add_column()
    grid.add_row(logo, words)
    return grid


class ToolList(BarList):
    BINDINGS: ClassVar[list[BindingType]] = [Binding("enter", "select", "open")]


class HomeScreen(Screen[None]):
    BINDINGS: ClassVar[list[BindingType]] = [
        *[Binding(str(t.number), f"open({t.number})", t.key, show=False) for t in TOOLS],
        Binding("t", "app.pick_tenant", "tenant"),
        Binding("ctrl+l", "app.login", "sign in", key_display="ctrl+l"),
        Binding("colon", "app.command_palette", "commands", key_display=":"),
        Binding("question_mark", "help", "help", key_display="?"),
        Binding("q", "app.quit_app", "quit"),
    ]
    HELP = ("This is the toolbox home. esc in a tool goes back one step at a time and ends here. Each tool has its own screen with the same keys: j/k to move, enter to "
            "open, : for every command, ? for help. g in any tool brings you back here.")

    @property
    def m365(self) -> M365App:
        return self.app  # type: ignore[return-value]

    def compose(self) -> ComposeResult:
        yield HeaderBar("", id="header")
        with Vertical(id="home"):
            yield Static(mark(), id="home-mark")
            yield Static(Text(TAGLINE, style=SUB), id="home-tagline")
            yield Static(Text("TOOLS", style=f"bold {DIM}"), classes="home-label")
            yield ToolList(*self._options(), id="tools")
            yield Static("", id="home-foot")
        yield KeyBar(id="keybar")

    def _options(self) -> list[Option | None]:
        app = self.m365
        out: list[Option | None] = []
        for i, tool in enumerate(TOOLS):
            if i:
                out.append(None)
            last = last_run(tool.key, app.tenant)
            out.append(BarOption(_maker(tool, last), id=tool.key))
        return out

    def on_mount(self) -> None:
        tools = self.query_one(ToolList)
        tools.highlighted = 0
        tools.focus()
        self._foot()

    def on_screen_resume(self) -> None:
        tools = self.query_one(ToolList)
        idx = tools.highlighted
        tools.set_options(self._options())
        tools.highlighted = idx if idx is not None else 0
        tools.focus()
        self._foot()

    def _foot(self) -> None:
        app = self.m365
        t = Text()
        if app.account:
            t.append("● ", style=theme.GREEN).append("signed in as ", style=SUB).append(app.account, style=FG)
        else:
            t.append("○ ", style=DIM).append("not signed in  ", style=SUB)
            t.append_text(theme.key_hint("ctrl+l", "sign in"))
        t.append("    ")
        if app.cfg.tenants:
            t.append("tenant ", style=SUB).append(app.tenant or "none", style=FG)
            if len(app.cfg.tenants) > 1:
                t.append("  ").append_text(theme.key_hint("t", "switch"))
        else:
            t.append("no tenant yet: run ", style=theme.PEACH).append("m365 tenant add NAME", style=f"bold {FG}")
        self.query_one("#home-foot", Static).update(t)

    def on_option_list_option_selected(self, event: OptionList.OptionSelected) -> None:
        if event.option.id:
            self.m365.open_tool(event.option.id)

    def action_open(self, number: int) -> None:
        self.m365.open_tool(TOOLS[number - 1].key)

    def action_help(self) -> None:
        from .workspace import help_group

        keys = [("1  2  3", "open spmig / groups / onboard"), ("enter", "open the highlighted tool"), ("j / k", "move"),
                ("t", "switch tenant"), ("ctrl+l", "sign in"), (":", "command palette"), ("q", "quit")]
        tools = [(str(t.number), f"{t.key}: {t.title}") for t in TOOLS]
        body = Group(Text(self.HELP, style=SUB), Text(""), help_group("keys", keys), Text(""),
                     help_group("tools", tools))
        self.app.push_screen(TextModal("Keys · home", body, fit=True))


def _maker(tool: Tool, last: LastRun | None) -> Callable[[bool], Any]:
    def make(selected: bool) -> Any:
        return tool_entry(tool, last, selected)

    return make
