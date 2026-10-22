"""Reusable widgets: vim-key tables/lists/logs, the header, the sidebar, the key bar and the activity line."""

from __future__ import annotations

import time
from collections.abc import Callable, Iterable
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

from m365tools.core.events import (
    ItemStarted,
    LogLine,
    PhaseChanged,
    ProgressEvent,
    Stats,
    Throttled,
    phase_label,
)
from m365tools.core.report.format import fmt_elapsed

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


MARK = "◆"
BRAND = "m365 toolbox"


class HeaderBar(Widget):
    """Top line on mantle, the same on every screen: ``◆ m365 toolbox ▸ tool`` | busy pill and tool context |
    tenant · signed-in account."""

    def __init__(self, tool: str = "", *, id: str | None = None) -> None:
        super().__init__(id=id)
        self.tool = tool

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
        grid.add_column(justify="right", no_wrap=True, overflow="ellipsis")
        left = Text.assemble((f" {MARK} ", f"bold {MAUVE}"), (BRAND, f"bold {FG}"))
        if self.tool:
            left.append("  ▸ ", style=DIM).append(self.tool, style=f"bold {MAUVE}")
        mid = Text()
        if not self.tool:
            mid.append_text(theme.key_hints([(str(i), t) for i, t in enumerate(app.TOOLS, 1)], sep="  "))
        if app.busy:
            mid.append_text(theme.pill(f"◐ {app.busy}", YELLOW)).append("   ")
        if self.tool:
            others = [(f"^{i}", t) for i, t in enumerate(app.TOOLS, 1) if t != self.tool]
            mid.append_text(theme.key_hints(others, sep=" ")).append("   ")
        if self.tool == "spmig":
            mid.append("profile ", style=SUB).append(app.opts.profile or "unsaved",
                                                      style=FG if app.opts.profile else DIM)
        right = Text()
        right.append("tenant ", style=SUB).append(app.tenant or "none", style=FG if app.tenant else theme.PEACH)
        right.append(" · ", style=DIM)
        right.append(app.account or "not signed in", style=FG if app.account else DIM).append(" ")
        grid.add_row(left, mid, right)
        return grid


ESSENTIAL_KEYS = frozenset({":", "?", "q", "g"})


def fit_keys(local: list[tuple[str, str]], glob: list[tuple[str, str]], width: int
             ) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
    """Drop chips until the bar fits ``width``: movement chips first, then the last non-essential global ones,
    then the last local ones. ``:``, ``?`` and ``q`` always stay."""
    local, glob = list(local), list(glob)

    def size() -> int:
        n = 1 + theme.key_hints(local).cell_len + theme.key_hints(glob).cell_len
        return n + (5 if local and glob else 0)

    while size() > width:
        move = next((i for i, (k, _d) in enumerate(local) if k in ("j/k",)), None)
        if move is not None:
            local.pop(move)
            continue
        drop = next((i for i in range(len(glob) - 1, -1, -1) if glob[i][0] not in ESSENTIAL_KEYS), None)
        if drop is not None:
            glob.pop(drop)
            continue
        if local:
            local.pop()
            continue
        break
    return local, glob


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
        local, glob = fit_keys(local, glob, self.size.width or 10_000)
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
        self.phase_id = ""
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
            self.phase_id = ev.phase
            self.phase = phase_label(ev.phase) + (f": {ev.detail}" if ev.detail else "")
            self.progress = ""
        elif isinstance(ev, Stats):
            unit = STEP_UNITS.get(self.phase_id, "items" if self.phase_id.split(".")[0] in (
                "inventory", "permissions") else "files")
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


# ------------------------------------------------------------------------------------------------ step list

STEP_SPINNER = "◐◓◑◒"
STEP_UNITS = {"users": "users", "checks": "checks", "plan": "files"}
LEGACY_PHASES = {"inventory": "inventory.source", "permissions": "permissions.source"}


class _Step:
    __slots__ = ("bytes", "detail", "done", "ended", "phase", "started", "state", "total")

    def __init__(self, phase: str) -> None:
        self.phase = phase
        self.state = "todo"  # todo | running | done | failed
        self.detail = ""
        self.done = self.total = self.bytes = 0
        self.started: float | None = None
        self.ended: float | None = None

    def elapsed(self, now: float) -> float | None:
        if self.started is None:
            return None
        return (self.ended or now) - self.started


class StepList(Widget):
    """The steps of an operation's pre-phase (audit, dry run), one line each::

        ● inventory source      102,149 items · 2.4 GB      12.4s
        ◐ permissions scan      45,000/102,149  ━━━━━──     8.1s
        ○ inventory destination

    Fed with the core progress events (``PhaseChanged`` moves to a step, ``Stats`` is that step's progress);
    phases in ``aliases`` count towards another step, unknown phases are ignored. After :meth:`finish` it
    collapses to one line ("audit took 1m 42s") that stays above the results."""

    DEFAULT_CSS = "StepList { height: auto; display: none; margin: 0 0 1 0; }"
    LABEL_W, INFO_W, BAR_W = 22, 26, 16

    def __init__(self, phases: Iterable[str], *, aliases: dict[str, str] | None = None, id: str | None = None
                 ) -> None:
        super().__init__(id=id)
        self.phases = tuple(phases)
        self.aliases = {**LEGACY_PHASES, **(aliases or {})}
        self.label = ""
        self.steps: list[_Step] = [_Step(p) for p in self.phases]
        self.collapsed = False
        self.ok = True
        self.started = 0.0
        self.ended: float | None = None
        self.throttle_until = 0.0
        self.throttle_text = ""
        self._frame = 0
        self._timer: Timer | None = None

    # -- state ----------------------------------------------------------------------------------------

    @property
    def running(self) -> bool:
        return self._timer is not None

    @property
    def current(self) -> _Step | None:
        return next((st for st in self.steps if st.state == "running"), None)

    def start(self, label: str) -> None:
        self.label = label
        self.steps = [_Step(p) for p in self.phases]
        self.collapsed, self.ok, self.ended = False, True, None
        self.started = time.monotonic()
        self.display = True
        if self._timer is None:
            self._timer = self.set_interval(0.15, self._tick)
        self.refresh(layout=True)

    def finish(self, ok: bool = True) -> None:
        now = time.monotonic()
        cur = self.current
        if cur is not None:
            cur.state, cur.ended = ("done" if ok else "failed"), now
        self.ok, self.ended, self.collapsed = ok, now, True
        if self._timer is not None:
            self._timer.stop()
            self._timer = None
        self.refresh(layout=True)

    def clear(self) -> None:
        self.finish()
        self.display = False

    def _tick(self) -> None:
        self._frame += 1
        self.refresh()

    def feed(self, ev: ProgressEvent) -> None:
        if isinstance(ev, PhaseChanged):
            self._enter(self.aliases.get(ev.phase, ev.phase), ev.detail)
        elif isinstance(ev, Stats):
            cur = self.current
            if cur is not None:
                cur.done, cur.total, cur.bytes = ev.done, ev.total, ev.bytes_done
        elif isinstance(ev, Throttled):
            self.throttle_until = time.monotonic() + ev.seconds
            self.throttle_text = f"throttled, waiting {ev.seconds:.0f}s ({ev.what})"
        else:
            return
        self.refresh(layout=isinstance(ev, PhaseChanged))

    def _enter(self, phase: str, detail: str) -> None:
        if phase not in self.phases:
            return
        idx = self.phases.index(phase)
        step = self.steps[idx]
        if step.state == "running":  # a sub-step (the next check, the next route ...)
            step.detail = detail
            return
        now = time.monotonic()
        for st in self.steps[:idx]:
            if st.state == "running":
                st.state, st.ended = "done", now
        step.state, step.started, step.ended, step.detail = "running", now, None, detail
        step.done = step.total = step.bytes = 0

    # -- rendering ------------------------------------------------------------------------------------

    def _info(self, st: _Step) -> Text:
        unit = STEP_UNITS.get(st.phase, "items")
        t = Text(no_wrap=True, overflow="ellipsis")
        if st.state == "running":
            if st.total:
                t.append(f"{st.done:,}/{st.total:,}", style=FG)
                if st.phase == "checks" and st.detail:
                    t.append(f" {st.detail}", style=SUB)
            elif st.detail:
                t.append(st.detail, style=SUB)
        elif st.state in ("done", "failed"):
            if st.phase.startswith("inventory") and (st.total or st.done):
                t.append(f"{st.done:,} items", style=SUB)
                if st.bytes:
                    t.append(f" · {theme.size(st.bytes)}", style=SUB)
            elif st.total:
                t.append(f"{st.total:,} {unit}", style=SUB)
            elif st.detail:
                t.append(st.detail, style=SUB)
        return t

    def _bar(self, st: _Step) -> Text:
        if st.state != "running" or not st.total:
            return Text(" " * self.BAR_W)
        filled = round(self.BAR_W * min(st.done / st.total, 1.0))
        return Text.assemble(("━" * filled, YELLOW), ("─" * (self.BAR_W - filled), theme.SURFACE1))

    def render(self) -> Table | Text:
        now = time.monotonic()
        if self.collapsed:
            return self._summary_line(now)
        grid = Table.grid(padding=(0, 1))
        grid.add_column(width=1, no_wrap=True)
        grid.add_column(width=self.LABEL_W, no_wrap=True)
        grid.add_column(width=self.INFO_W, no_wrap=True)
        grid.add_column(width=self.BAR_W, no_wrap=True)
        grid.add_column(justify="right", no_wrap=True, min_width=7)
        for st in self.steps:
            if st.state == "done":
                glyph, gstyle, lstyle = "●", theme.GREEN, FG
            elif st.state == "failed":
                glyph, gstyle, lstyle = "✗", theme.RED, FG
            elif st.state == "running":
                glyph, gstyle, lstyle = STEP_SPINNER[self._frame % len(STEP_SPINNER)], YELLOW, f"bold {FG}"
            else:
                glyph, gstyle, lstyle = "○", DIM, DIM
            el = st.elapsed(now)
            grid.add_row(Text(glyph, style=gstyle), Text(phase_label(st.phase), style=lstyle), self._info(st),
                         self._bar(st), Text(fmt_elapsed(el) if el is not None else "", style=SUB))
        if now < self.throttle_until:
            grid.add_row("", Text(self.throttle_text, style=YELLOW), "", "", "")
        return grid

    def _summary_line(self, now: float) -> Text:
        total = (self.ended or now) - self.started
        t = Text(no_wrap=True, overflow="ellipsis")
        if self.ok:
            t.append("● ", style=theme.GREEN)
        else:
            t.append("✗ ", style=theme.RED)
        t.append(f"{self.label} {'took' if self.ok else 'failed after'} {fmt_elapsed(total)}", style=FG)
        timed = [(st.elapsed(now) or 0.0, st) for st in self.steps if st.started is not None]
        if timed:
            secs, slow = max(timed, key=lambda x: x[0])
            t.append(" · slowest ", style=DIM).append(f"{phase_label(slow.phase)} {fmt_elapsed(secs)}", style=SUB)
        return t
