"""Reusable widgets: vim-key tables/lists/logs, the header, the sidebar, the key bar and the activity line."""

from __future__ import annotations

import time
from collections.abc import Iterable
from datetime import datetime
from typing import TYPE_CHECKING, Any, ClassVar

from rich.table import Table
from rich.text import Text
from textual.binding import Binding, BindingType
from textual.containers import VerticalScroll
from textual.screen import Screen
from textual.timer import Timer
from textual.widget import Widget
from textual.widgets import DataTable, OptionList, RichLog
from textual.widgets.option_list import Option

from m365tools.core.events import ItemStarted, LogLine, PhaseChanged, ProgressEvent, Stats, Throttled

from . import theme
from .theme import ACCENT, DIM, FG, GLYPH, GLYPH_STYLE, SPINNER, YELLOW

if TYPE_CHECKING:
    from .app import M365App

# Bindings shared by every vertical list-like widget. Only "j/k move" is shown in the key bar.
VIM_MOVE: list[BindingType] = [
    Binding("j", "cursor_down", "move", key_display="j/k"),
    Binding("k", "cursor_up", "up", show=False),
    Binding("ctrl+d", "page_down", "page down", show=False),
    Binding("ctrl+u", "page_up", "page up", show=False),
]


class VimTable(DataTable):
    """DataTable with row cursor and vim movement (j/k, g/G, ctrl+d/u)."""

    BINDINGS: ClassVar[list[BindingType]] = [
        *VIM_MOVE,
        Binding("g", "scroll_top", "top", show=False),
        Binding("G", "scroll_bottom", "bottom", show=False),
    ]

    def __init__(self, *, id: str | None = None) -> None:
        super().__init__(id=id, cursor_type="row", zebra_stripes=False, cursor_foreground_priority="renderable")


class VimList(OptionList):
    """OptionList with vim movement."""

    BINDINGS: ClassVar[list[BindingType]] = [
        *VIM_MOVE,
        Binding("g", "first", "top", show=False),
        Binding("G", "last", "bottom", show=False),
    ]


class VimScroll(VerticalScroll, can_focus=True):
    """A focusable scrolling area (text viewers, detail panes)."""

    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("j", "scroll_down", "scroll", key_display="j/k"),
        Binding("k", "scroll_up", "up", show=False),
        Binding("ctrl+d,space", "page_down", "page down", show=False),
        Binding("ctrl+u", "page_up", "page up", show=False),
        Binding("g", "scroll_home", "top", show=False),
        Binding("G", "scroll_end", "bottom", show=False),
    ]


class VimLog(RichLog):
    """The log tail: keeps the last 200 lines, scrolls with j/k, follows new lines when at the bottom."""

    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("j", "scroll_down", "scroll", key_display="j/k"),
        Binding("k", "scroll_up", "up", show=False),
        Binding("g", "scroll_home", "top", show=False),
        Binding("G", "scroll_end", "bottom", show=False),
    ]

    def __init__(self, *, id: str | None = None, max_lines: int = 200) -> None:
        super().__init__(id=id, max_lines=max_lines, wrap=True, markup=False, highlight=False, auto_scroll=True)


# ------------------------------------------------------------------------------------------------ chrome


class HeaderBar(Widget):
    """Top line: app name, active tenant and account, loaded profile, clock."""

    @property
    def m365(self) -> M365App:
        return self.app  # type: ignore[return-value]

    def on_mount(self) -> None:
        self.set_interval(1, self.refresh)

    def render(self) -> Table:
        app = self.m365
        grid = Table.grid(expand=True)
        grid.add_column(justify="left", ratio=1)
        grid.add_column(justify="center", ratio=2)
        grid.add_column(justify="right", ratio=1)
        left = Text.assemble((" m365tools", f"bold {ACCENT}"), (" ▸ spmig", DIM))
        mid = Text()
        mid.append("tenant ", style=DIM).append(app.tenant or "none", style=FG if app.tenant else YELLOW)
        mid.append("  ·  ", style=DIM).append("account ", style=DIM)
        mid.append(app.account or "not signed in", style=FG if app.account else DIM)
        mid.append("  ·  ", style=DIM).append("profile ", style=DIM).append(app.opts.profile or "-", style=FG)
        if app.busy:
            mid.append("  ·  ", style=DIM).append(f"◐ {app.busy}", style=f"bold {YELLOW}")
        right = Text(datetime.now().strftime("%H:%M:%S "), style=DIM)
        grid.add_row(left, mid, right)
        return grid


class KeyBar(Widget):
    """Bottom line: the keys that work right now. Keys of the focused pane come first, then a ``│`` and the
    global keys. Rebuilt whenever focus or a binding's enabled state changes (``bindings_updated_signal``)."""

    def on_mount(self) -> None:
        self.screen.bindings_updated_signal.subscribe(self, self._bindings_changed)

    def on_unmount(self) -> None:
        self.screen.bindings_updated_signal.unsubscribe(self)

    def _bindings_changed(self, _screen: Screen[Any]) -> None:
        self.refresh()

    def render(self) -> Text:
        screen, app = self.screen, self.app
        local, glob = Text(), Text()
        seen: set[str] = set()
        for active in screen.active_bindings.values():
            b = active.binding
            if not b.show or not active.enabled or b.action in seen:
                continue
            seen.add(b.action)
            target = glob if active.node in (screen, app) else local
            target.append_text(theme.key_hint(b.key_display or app.get_key_display(b), b.description))
            target.append("  ")
        out = Text(" ")
        out.append_text(local)
        if local and glob:
            out.append("│  ", style=DIM)
        out.append_text(glob)
        out.no_wrap = True
        out.overflow = "ellipsis"
        return out


# ------------------------------------------------------------------------------------------------ sidebar


def sidebar_entry(number: int, title: str, status: str, summary: str) -> Text:
    t = Text()
    t.append(f" {GLYPH[status]} ", style=GLYPH_STYLE[status])
    t.append(f"{number} ", style=f"bold {ACCENT}")
    t.append(title, style="bold")
    t.append("\n     ")
    t.append(summary or " ", style=DIM)
    return t


class Sidebar(VimList):
    """Left navigation list: one entry per section with a state glyph and a one-line summary."""

    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("enter", "select", "open", key_display="enter"),
    ]

    def __init__(self, entries: Iterable[tuple[str, str]], *, id: str | None = None) -> None:
        self._titles = list(entries)  # (key, title)
        options = [Option(sidebar_entry(i, title, theme.TODO, ""), id=key)
                   for i, (key, title) in enumerate(self._titles, 1)]
        super().__init__(*options, id=id)

    def set_entry(self, key: str, status: str, summary: str) -> None:
        index = self.get_option_index(key)
        self.replace_option_prompt_at_index(index, sidebar_entry(index + 1, self._titles[index][1], status,
                                                                 summary))


# ------------------------------------------------------------------------------------------------ activity


class Activity(Widget):
    """One status line for a running operation: spinner, operation, phase and progress. Hidden when idle.

    ``feed(event)`` understands the core progress events, so sections just forward what they receive."""

    DEFAULT_CSS = "Activity { height: 1; display: none; }"

    def __init__(self, *, id: str | None = None) -> None:
        super().__init__(id=id)
        self.label = ""
        self.phase = ""
        self.progress = ""
        self.throttle_until = 0.0
        self.throttle_text = ""
        self._frame = 0
        self._timer: Timer | None = None

    @property
    def active(self) -> bool:
        return self._timer is not None

    def start(self, label: str) -> None:
        self.label, self.phase, self.progress = label, "starting", ""
        self.display = True
        if self._timer is None:
            self._timer = self.set_interval(0.1, self._tick)
        self.refresh()

    def stop(self) -> None:
        if self._timer is not None:
            self._timer.stop()
            self._timer = None
        self.display = False

    def _tick(self) -> None:
        self._frame += 1
        self.refresh()

    def feed(self, ev: ProgressEvent) -> None:
        if isinstance(ev, PhaseChanged):
            self.phase = ev.phase + (f": {ev.detail}" if ev.detail else "")
            self.progress = ""
        elif isinstance(ev, Stats):
            unit = "items" if self.phase.startswith("inventory") else "files"
            self.progress = f"{ev.done:,}/{ev.total:,} {unit}"
        elif isinstance(ev, ItemStarted):
            self.progress = f"{ev.index:,}/{ev.total:,} {ev.rel_path}"
        elif isinstance(ev, Throttled):
            self.throttle_until = time.monotonic() + ev.seconds
            self.throttle_text = f"throttled, waiting {ev.seconds:.0f}s ({ev.what})"
        elif isinstance(ev, LogLine) and ev.level in ("warn", "error"):
            self.progress = ev.text
        self.refresh()

    def render(self) -> Text:
        spin = SPINNER[self._frame % len(SPINNER)]
        t = Text.assemble((f" {spin} ", f"bold {ACCENT}"), (self.label, "bold"), ("  ", ""), (self.phase, FG))
        if self.progress:
            t.append("  ·  ", style=DIM).append(self.progress, style=DIM)
        if time.monotonic() < self.throttle_until:
            t.append("  ·  ", style=DIM).append(self.throttle_text, style=YELLOW)
        t.no_wrap = True
        t.overflow = "ellipsis"
        return t
