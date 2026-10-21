"""Reusable widgets: vim-key tables/lists/logs, the header, the sidebar, the key bar and the activity line."""

from __future__ import annotations

import time
from collections.abc import Callable, Iterable
from datetime import datetime
from typing import TYPE_CHECKING, Any, ClassVar

from rich.table import Table
from rich.text import Text
from textual import events
from textual.binding import Binding, BindingType
from textual.containers import VerticalScroll
from textual.screen import Screen
from textual.timer import Timer
from textual.visual import VisualType
from textual.widget import Widget
from textual.widgets import DataTable, OptionList, RichLog
from textual.widgets.option_list import Option

from m365tools.core.events import ItemStarted, LogLine, PhaseChanged, ProgressEvent, Stats, Throttled

from . import theme
from .theme import CRUST, DIM, FG, GLYPH, GLYPH_STYLE, KEY, MAUVE, SPINNER, SUB, YELLOW

if TYPE_CHECKING:
    from .app import M365App

# Bindings shared by every vertical list-like widget. Only "j/k move" is shown in the key bar.
VIM_MOVE: list[BindingType] = [
    Binding("j", "cursor_down", "move", key_display="j/k"),
    Binding("k", "cursor_up", "up", show=False),
    Binding("ctrl+d", "page_down", "page down", show=False),
    Binding("ctrl+u", "page_up", "page up", show=False),
]


def fit(text: str, width: int) -> str:
    """Cut ``text`` to ``width`` cells with an ellipsis."""
    return text if len(text) <= width else text[: max(width - 1, 0)] + "…"


class VimTable(DataTable):
    """DataTable with row cursor and vim movement (j/k, g/G, ctrl+d/u).

    Columns added with ``add_fit_columns`` keep their fixed widths; the flexible ones (width ``None``) share
    whatever is left of the table width, so long paths are cut with an ellipsis instead of scrolling sideways."""

    BINDINGS: ClassVar[list[BindingType]] = [
        *VIM_MOVE,
        Binding("g", "scroll_top", "top", show=False),
        Binding("G", "scroll_bottom", "bottom", show=False),
    ]
    FLEX_MIN = 10

    def __init__(self, *, id: str | None = None) -> None:
        super().__init__(id=id, cursor_type="row", zebra_stripes=False, cursor_foreground_priority="css")
        self._flex: list[str] = []

    def add_fit_columns(self, *columns: tuple[str, int | None]) -> None:
        for label, width in columns:
            key = label.strip() or f"col{len(self.columns)}"
            self.add_column(label, width=width or self.FLEX_MIN, key=key)
            if width is None:
                self._flex.append(key)
        self._fit()

    def on_resize(self, _event: events.Resize) -> None:
        self._fit()

    def _fit(self) -> None:
        if not self._flex or not self.size.width:
            return
        pad = 2 * self.cell_padding
        fixed = sum(c.get_render_width(self) for k, c in self.columns.items() if k.value not in self._flex)
        avail = self.size.width - 1 - fixed - pad * len(self._flex)  # 1 = vertical scrollbar
        each = max(self.FLEX_MIN, avail // len(self._flex))
        for i, key in enumerate(self._flex):
            col = self.columns[key]  # type: ignore[index]
            col.auto_width = False
            col.width = max(self.FLEX_MIN, avail - each * (len(self._flex) - 1)) if i == len(self._flex) - 1 \
                else each
        self._require_update_dimensions = True
        self._clear_caches()
        self.refresh()


class VimList(OptionList):
    """OptionList with vim movement."""

    BINDINGS: ClassVar[list[BindingType]] = [
        *VIM_MOVE,
        Binding("g", "first", "top", show=False),
        Binding("G", "last", "bottom", show=False),
    ]


class BarOption(Option):
    """An option that can render itself twice: normal (coloured) and selected (plain, for the filled bar)."""

    def __init__(self, make: Callable[[bool], VisualType], id: str | None = None) -> None:
        self.make = make
        super().__init__(make(False), id=id)


class BarList(VimList):
    """The highlighted entry is a filled bar: mauve with crust text while the list is active, surface1 when not.

    ``BarOption`` prompts are re-rendered when the highlight or focus changes, so the selected entry drops its
    value colours (which would not read on mauve)."""

    ALWAYS_ACTIVE: ClassVar[bool] = False  # pickers: the list never has focus but the bar is always live

    def _active(self) -> bool:
        # screen.focused is updated before Focus/Blur are delivered (has_focus only after our handlers ran)
        return self.ALWAYS_ACTIVE or (self.is_mounted and self.screen.focused is self)

    def repaint(self, *indices: int | None) -> None:
        for i in indices:
            if i is None or not 0 <= i < self.option_count:
                continue
            opt = self.get_option_at_index(i)
            if isinstance(opt, BarOption):
                self.replace_option_prompt_at_index(i, opt.make(i == self.highlighted and self._active()))
        self.refresh()

    def watch_highlighted(self, old: int | None, new: int | None) -> None:
        super().watch_highlighted(new)
        self.repaint(old, new)

    def on_focus(self, _event: events.Focus) -> None:
        self.repaint(self.highlighted)

    def on_blur(self, _event: events.Blur) -> None:
        self.repaint(self.highlighted)


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
    """Top line on mantle: ``m365tools ▸ spmig`` | tenant · account · profile breadcrumb | clock."""

    @property
    def m365(self) -> M365App:
        return self.app  # type: ignore[return-value]

    def on_mount(self) -> None:
        self.set_interval(1, self.refresh)

    def render(self) -> Table:
        app = self.m365
        grid = Table.grid(expand=True)
        grid.add_column(justify="left", no_wrap=True)
        grid.add_column(justify="center", ratio=1, no_wrap=True, overflow="ellipsis")
        grid.add_column(justify="right", no_wrap=True)
        left = Text(" m365tools ▸ spmig", style=f"bold {MAUVE}")
        mid = Text()
        if app.busy:
            mid.append_text(theme.pill(f"◐ {app.busy}", YELLOW)).append("   ")
        mid.append("tenant ", style=SUB).append(app.tenant or "none", style=FG if app.tenant else theme.PEACH)
        mid.append(" · ", style=DIM).append("account ", style=SUB)
        mid.append(app.account or "not signed in", style=FG if app.account else DIM)
        mid.append(" · ", style=DIM).append("profile ", style=SUB)
        mid.append(app.opts.profile or "unsaved", style=FG if app.opts.profile else DIM)
        right = Text(datetime.now().strftime("%H:%M:%S "), style=SUB)
        grid.add_row(left, mid, right)
        return grid


class KeyBar(Widget):
    """Bottom line: the keys that work right now, as chips. Keys of the focused pane come first, then a ``│`` and
    the global keys. Rebuilt whenever focus or a binding's enabled state changes (``bindings_updated_signal``)."""

    def on_mount(self) -> None:
        self.screen.bindings_updated_signal.subscribe(self, self._bindings_changed)

    def on_unmount(self) -> None:
        self.screen.bindings_updated_signal.unsubscribe(self)

    def _bindings_changed(self, _screen: Screen[Any]) -> None:
        self.refresh()

    def render(self) -> Text:
        screen, app = self.screen, self.app
        local: list[tuple[str, str]] = []
        glob: list[tuple[str, str]] = []
        seen: set[str] = set()
        for active in screen.active_bindings.values():
            b = active.binding
            if not b.show or not active.enabled or b.action in seen:
                continue
            seen.add(b.action)
            target = glob if active.node in (screen, app) else local
            target.append((b.key_display or app.get_key_display(b), b.description))
        out = Text(" ")
        out.append_text(theme.key_hints(local))
        if local and glob:
            out.append("  │  ", style=DIM)
        out.append_text(theme.key_hints(glob))
        out.no_wrap = True
        out.overflow = "ellipsis"
        return out


# ------------------------------------------------------------------------------------------------ sidebar

SIDEBAR_TEXT = 22  # 26 columns - 2 x 1 padding - 2 indent


def sidebar_entry(number: int, title: str, status: str, summary: str, *, hint: str = "", selected: bool = False,
                  frame: int = 0) -> Text:
    """``● 1  Job`` / ``  12 files · 4.1 MB``. Selected: plain crust text (it sits on the mauve bar)."""
    glyph = SPINNER[frame % len(SPINNER)] if status == theme.RUNNING else GLYPH[status]
    line2 = fit(summary or hint, SIDEBAR_TEXT)
    t = Text()
    if selected:
        t.append(f"{glyph} ", style=f"bold {CRUST}").append(f"{number}  ", style=CRUST)
        t.append(title, style=f"bold {CRUST}").append("\n  ").append(line2, style=CRUST)
        return t
    t.append(f"{glyph} ", style=GLYPH_STYLE[status]).append(f"{number}  ", style=KEY)
    t.append(title, style=f"bold {FG}").append("\n  ")
    t.append(line2, style=SUB if summary and status != theme.TODO else DIM)
    return t


class Sidebar(BarList):
    """Left navigation list: one entry per section with a state glyph and a one-line summary."""

    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("enter", "select", "open", key_display="enter"),
    ]

    def __init__(self, entries: Iterable[tuple[str, str, str]], *, id: str | None = None) -> None:
        self._entries = list(entries)  # (key, title, hint when there is nothing to summarise)
        self._state: dict[str, tuple[str, str]] = {key: (theme.TODO, "") for key, _t, _h in self._entries}
        self._frame = 0
        options: list[Option | None] = []
        for i, (key, _title, _hint) in enumerate(self._entries):
            if i:
                options.append(None)  # a separator line (invisible): one blank line between entries
            options.append(BarOption(self._maker(i), id=key))
        super().__init__(*options, id=id)

    def _maker(self, index: int) -> Callable[[bool], Text]:
        key, title, hint = self._entries[index]

        def make(selected: bool) -> Text:
            status, summary = self._state[key]
            return sidebar_entry(index + 1, title, status, summary, hint=hint, selected=selected,
                                 frame=self._frame)

        return make

    def on_mount(self) -> None:
        self.set_interval(0.15, self._spin)

    def _spin(self) -> None:
        running = [i for i, (key, _t, _h) in enumerate(self._entries) if self._state[key][0] == theme.RUNNING]
        if running:
            self._frame += 1
            self.repaint(*running)

    def set_entry(self, key: str, status: str, summary: str) -> None:
        self._state[key] = (status, summary)
        self.repaint(self.get_option_index(key))


class NavStatus(Widget):
    """Bottom of the sidebar: who is signed in, or how to sign in."""

    @property
    def m365(self) -> M365App:
        return self.app  # type: ignore[return-value]

    def on_mount(self) -> None:
        self.set_interval(1, self.refresh)

    def render(self) -> Text:
        app = self.m365
        t = Text()
        if app.account:
            t.append("● ", style=theme.GREEN).append("signed in\n", style=FG)
            t.append(fit(app.account, 24), style=SUB)
        else:
            t.append("○ ", style=DIM).append("not signed in\n", style=FG)
            t.append_text(theme.key_hint("ctrl+l", "to sign in"))
        if app.tenant:
            t.append("\n").append("tenant ", style=DIM).append(fit(app.tenant, 17), style=SUB)
        return t


# ------------------------------------------------------------------------------------------------ activity


class Activity(Widget):
    """One status line for a running operation: spinner, operation, phase and progress. Hidden when idle.

    ``feed(event)`` understands the core progress events, so sections just forward what they receive."""

    DEFAULT_CSS = "Activity { height: 1; display: none; margin: 0 0 1 0; }"

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
            self._timer = self.set_interval(0.15, self._tick)
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
        t = Text.assemble(theme.pill(f"{spin} {self.label}", YELLOW), ("  ", ""), (self.phase, FG))
        if self.progress:
            t.append(" · ", style=DIM).append(self.progress, style=SUB)
        if time.monotonic() < self.throttle_until:
            t.append(" · ", style=DIM).append(self.throttle_text, style=YELLOW)
        t.no_wrap = True
        t.overflow = "ellipsis"
        return t
