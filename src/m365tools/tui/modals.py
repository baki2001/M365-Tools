"""Modal screens. Every one is keyboard-only: no buttons; the bottom border says which keys apply.

- ConfirmModal       y / n question                                  -> bool
- TypedConfirmModal  type a word (``force``, ``lock``) to confirm      -> bool
- PromptModal        one line of text (``/`` filters, names, folders) -> str | None
- PickerModal        filterable list; local, loaded once, or searched -> chosen value | None
- OptionsModal       job options editor                               -> JobOptions | None
- TextModal          scrollable read-only text (reports, help)        -> None
- AuthModal          sign-in URL / device code                        -> None
"""

from __future__ import annotations

from collections.abc import Awaitable, Callable
from dataclasses import dataclass, replace
from typing import Any, ClassVar

from rich.console import RenderableType
from rich.text import Text
from textual import on, work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.screen import ModalScreen
from textual.timer import Timer
from textual.widgets import Input, OptionList, Static
from textual.widgets.option_list import Option

from m365tools.core.events import AuthPrompt
from m365tools.tools.spmig.models import JobOptions

from . import theme
from .theme import ACCENT, BLUE, DIM, FG, KEY, RED, SUB, YELLOW
from .widgets import BarList, BarOption, VimScroll


def _hint(*pairs: tuple[str, str]) -> Text:
    """Key hints for the bottom border: key in bold blue, label in subtext."""
    t = Text()
    for i, (k, d) in enumerate(pairs):
        t.append(" · " if i else " ", style=DIM).append(k, style=f"bold {KEY}").append(f" {d}", style=SUB)
    return t.append(" ")


def prompt_row(input_widget: Input) -> Horizontal:
    """``› text``: an input on surface1 with a prompt glyph (mauve while focused)."""
    row = Horizontal(classes="prompt-row")
    row.compose_add_child(Static("›", classes="prompt-glyph"))
    row.compose_add_child(input_widget)
    return row


class Dialog(ModalScreen[Any]):
    """Base: a centred box with a title in the top border and key hints in the bottom border."""

    DIALOG_CLASSES = "dialog"

    def __init__(self, title: str, hint: Text) -> None:
        super().__init__()
        self.dialog_title = title
        self.dialog_hint = hint

    def box(self) -> Vertical:
        box = Vertical(classes=self.DIALOG_CLASSES)
        box.border_title = self.dialog_title
        box.border_subtitle = self.dialog_hint
        return box


# ------------------------------------------------------------------------------------------------ confirms


class ConfirmModal(Dialog):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("y,enter", "answer(True)", "yes"),
        Binding("n,escape", "answer(False)", "no"),
    ]

    def __init__(self, title: str, body: RenderableType) -> None:
        super().__init__(title, _hint(("y", "yes"), ("n/esc", "no")))
        self.body = body

    def compose(self) -> ComposeResult:
        with self.box():
            yield Static(self.body)

    def action_answer(self, yes: bool) -> None:
        self.dismiss(yes)


class TypedConfirmModal(Dialog):
    """Dangerous actions: the user must type ``word`` and press enter."""

    BINDINGS: ClassVar[list[BindingType]] = [Binding("escape", "cancel", "cancel")]

    def __init__(self, title: str, body: RenderableType, word: str) -> None:
        super().__init__(title, _hint(("type", word), ("enter", "confirm"), ("esc", "cancel")))
        self.body = body
        self.word = word

    def compose(self) -> ComposeResult:
        with self.box():
            yield Static(self.body)
            yield Static(Text.assemble(("type ", SUB), theme.pill(self.word, RED), (" to continue", SUB)),
                         classes="ask")
            yield prompt_row(Input(placeholder=self.word, id="confirm-word"))
            yield Static("", id="confirm-error", classes="error-line")

    @on(Input.Submitted)
    def _submitted(self, event: Input.Submitted) -> None:
        if event.value.strip().lower() == self.word:
            self.dismiss(True)
        else:
            self.query_one("#confirm-error", Static).update(
                Text.assemble((f"That is not '{self.word}'.", RED), ("  esc cancels", SUB)))

    def action_cancel(self) -> None:
        self.dismiss(False)


# ------------------------------------------------------------------------------------------------ prompt


class PromptModal(Dialog):
    """One line of text. ``on_change`` follows the typing (live filtering), once the keys pause for
    ``DEBOUNCE_S``; esc returns None."""

    DEBOUNCE_S = 0.15
    BINDINGS: ClassVar[list[BindingType]] = [Binding("escape", "cancel", "cancel")]

    def __init__(self, title: str, value: str = "", *, placeholder: str = "", note: str = "",
                 on_change: Callable[[str], None] | None = None, bottom: bool = False) -> None:
        super().__init__(title, _hint(("enter", "ok"), ("esc", "cancel")))
        self.value = value
        self.placeholder = placeholder
        self.note = note
        self.on_change = on_change
        self._timer: Timer | None = None
        if bottom:
            self.add_class("-bottom")

    def compose(self) -> ComposeResult:
        with self.box():
            if self.note:
                yield Static(Text(self.note, style=SUB))
            yield prompt_row(Input(value=self.value, placeholder=self.placeholder, id="prompt-input"))

    @on(Input.Changed)
    def _changed(self, event: Input.Changed) -> None:
        if self.on_change is None:
            return
        if self._timer is not None:
            self._timer.stop()
        on_change, value = self.on_change, event.value
        self._timer = self.set_timer(self.DEBOUNCE_S, lambda: on_change(value))

    def _stop_timer(self) -> None:
        if self._timer is not None:
            self._timer.stop()
            self._timer = None

    @on(Input.Submitted)
    def _submitted(self, event: Input.Submitted) -> None:
        self._stop_timer()
        self.dismiss(event.value)

    def action_cancel(self) -> None:
        self._stop_timer()
        self.dismiss(None)


# ------------------------------------------------------------------------------------------------ picker


@dataclass(slots=True)
class Choice:
    label: str
    value: Any
    detail: str = ""
    extra: Text | None = None  # shown after the detail as is (the theme picker's colour swatches)

    def prompt(self, selected: bool = False) -> Text:
        t = Text(self.label, style=f"bold {ACCENT if selected else FG}")
        if self.detail:
            t.append("  ").append(self.detail, style=SUB)
        if self.extra is not None:
            t.append("  ").append_text(self.extra)
        return t


class PickerList(BarList):
    """The picker's list: never focused (typing goes to the input), but its bar is always live."""

    ALWAYS_ACTIVE = True


ChoiceLoader = Callable[[], Awaitable[list[Choice]]]
ChoiceSearch = Callable[[str], Awaitable[list[Choice]]]


class PickerModal(Dialog):
    """Type to narrow a list, arrows (or ctrl+j/k) to move, enter to pick.

    Exactly one source of choices:
      ``choices``  a fixed list, filtered locally;
      ``loader``   fetched once when the picker opens (e.g. libraries), then filtered locally;
      ``search``   called with the typed text after ``DEBOUNCE_S`` of no typing (e.g. SharePoint site search).
    ``accept_text`` turns typed text into a value when nothing in the list matches (e.g. a pasted URL).
    ``on_highlight`` is called with the value under the bar whenever it moves (live previews).
    """

    DEBOUNCE_S = 0.3
    DIALOG_CLASSES = "dialog picker"
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("escape", "cancel", "cancel"),
        Binding("down,ctrl+j,ctrl+n", "move(1)", "down", show=False),
        Binding("up,ctrl+k,ctrl+p", "move(-1)", "up", show=False),
        Binding("pagedown", "move(10)", "page down", show=False),
        Binding("pageup", "move(-10)", "page up", show=False),
    ]

    def __init__(self, title: str, choices: list[Choice] | None = None, *, loader: ChoiceLoader | None = None,
                 search: ChoiceSearch | None = None, accept_text: Callable[[str], Any] | None = None,
                 placeholder: str = "type to filter", note: str = "", highlight: Any = None,
                 on_highlight: Callable[[Any], None] | None = None) -> None:
        super().__init__(title, _hint(("type", "filter"), ("↑↓", "move"), ("enter", "pick"), ("esc", "cancel")))
        self.all_choices = list(choices or [])
        self.shown: list[Choice] = []
        self.loader = loader
        self.searcher = search
        self.accept_text = accept_text
        self.placeholder = placeholder
        self.note = note
        self.highlight_value = highlight
        self.on_highlight = on_highlight
        self._timer: Timer | None = None

    def compose(self) -> ComposeResult:
        with self.box():
            if self.note:
                yield Static(Text(self.note, style=SUB))
            yield prompt_row(Input(placeholder=self.placeholder, id="picker-input"))
            yield PickerList(id="picker-list")
            yield Static("", id="picker-status")

    def on_mount(self) -> None:
        self.query_one(OptionList).can_focus = False  # typing always goes to the input
        if self.loader is not None:
            self._status("loading ...")
            self._load()
        elif self.searcher is not None:
            self._status("type part of the name; results appear as you type")
        else:
            self._show(self.all_choices)

    # -- data ---------------------------------------------------------------------------------------

    @work(exclusive=True, group="picker", exit_on_error=False)
    async def _load(self) -> None:
        assert self.loader is not None
        try:
            self.all_choices = await self.loader()
        except Exception as exc:  # noqa: BLE001 - shown in the picker
            self._status(f"failed: {exc}", RED)
            return
        self._show(self._filtered(self.query_one(Input).value))

    @work(exclusive=True, group="picker", exit_on_error=False)
    async def search(self, query: str) -> None:
        assert self.searcher is not None
        self._status(f"searching '{query}' ...")
        try:
            found = await self.searcher(query)
        except Exception as exc:  # noqa: BLE001 - shown in the picker
            self._status(f"search failed: {exc}", RED)
            return
        self._show(found, empty=f"no match for '{query}'")

    def _filtered(self, text: str) -> list[Choice]:
        words = text.lower().split()
        return [c for c in self.all_choices if all(w in f"{c.label} {c.detail}".lower() for w in words)]

    def _show(self, choices: list[Choice], empty: str = "nothing matches") -> None:
        self.shown = choices
        ol = self.query_one(OptionList)
        ol.set_options([BarOption(c.prompt) for c in choices])
        if choices:
            idx = next((i for i, c in enumerate(choices) if c.value == self.highlight_value), 0)
            ol.highlighted = idx
            self._status(theme.plural(len(choices), "result"), SUB)
        else:
            self._status(empty, YELLOW)

    def _status(self, text: str, style: str = SUB) -> None:
        self.query_one("#picker-status", Static).update(Text(text, style=style))

    # -- keys ---------------------------------------------------------------------------------------

    @on(Input.Changed)
    def _changed(self, event: Input.Changed) -> None:
        if self.searcher is None:
            self._show(self._filtered(event.value))
            return
        if self._timer is not None:
            self._timer.stop()
        query = event.value.strip()
        if query and not query.lower().startswith("http"):
            self._timer = self.set_timer(self.DEBOUNCE_S, lambda: self.search(query))

    @on(Input.Submitted)
    def _submitted(self, event: Input.Submitted) -> None:
        idx = self.query_one(OptionList).highlighted
        if self.shown and idx is not None:
            self.dismiss(self.shown[idx].value)
            return
        if self.accept_text is not None and event.value.strip():
            value = self.accept_text(event.value.strip())
            if value is not None:
                self.dismiss(value)
                return
        self._status("nothing to pick: type to search, or esc to cancel", YELLOW)

    @on(OptionList.OptionHighlighted)
    def _highlighted(self, event: OptionList.OptionHighlighted) -> None:
        if self.on_highlight is not None and 0 <= event.option_index < len(self.shown):
            self.on_highlight(self.shown[event.option_index].value)

    @on(OptionList.OptionSelected)
    def _clicked(self, event: OptionList.OptionSelected) -> None:
        self.dismiss(self.shown[event.option_index].value)

    def action_move(self, delta: int) -> None:
        ol = self.query_one(OptionList)
        if not self.shown:
            return
        cur = ol.highlighted or 0
        ol.highlighted = max(0, min(len(self.shown) - 1, cur + delta))

    def action_cancel(self) -> None:
        self.dismiss(None)


# ------------------------------------------------------------------------------------------------ options


OVERWRITE_CYCLE = ("skip_if_same_or_newer", "always", "never")
OVERWRITE_HELP = {
    "skip_if_same_or_newer": "copy new files; skip files whose destination copy is as new or newer",
    "always": "copy everything, replacing existing files",
    "never": "only copy files that do not exist yet",
}


class OptionsModal(Dialog):
    """Edit the job options on a copy; ``s`` saves, esc discards."""

    DIALOG_CLASSES = "dialog wide"
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("enter,space", "change", "change"),
        Binding("plus,equals_sign", "workers(1)", "workers +", show=False),
        Binding("minus", "workers(-1)", "workers -", show=False),
        Binding("s", "save", "save"),
        Binding("escape", "cancel", "cancel"),
    ]
    ROWS = ("overwrite", "restore_metadata", "workers", "excludes", "lock_source_after", "usermap_path")

    def __init__(self, opts: JobOptions) -> None:
        super().__init__("Options", _hint(("j/k", "move"), ("enter", "change"), ("+/-", "workers"),
                                          ("s", "save"), ("esc", "cancel")))
        self.opts = replace(opts, excludes=list(opts.excludes), routes=list(opts.routes))

    def compose(self) -> ComposeResult:
        with self.box():
            yield BarList(*self._options(), id="options-list")
            yield Static("", id="options-help")

    def on_mount(self) -> None:
        self.query_one(BarList).focus()
        self._help()

    def _options(self) -> list[Option]:
        o = self.opts
        values = {
            "overwrite": o.overwrite,
            "restore_metadata": "on" if o.restore_metadata else "off",
            "workers": str(o.workers),
            "excludes": ", ".join(o.excludes) or "(none)",
            "lock_source_after": "on" if o.lock_source_after else "off",
            "usermap_path": str(o.usermap_path) if o.usermap_path else "(none: proposal written by audit)",
        }
        labels = {"overwrite": "Overwrite", "restore_metadata": "Restore metadata", "workers": "Workers",
                  "excludes": "Excludes", "lock_source_after": "Lock source after", "usermap_path": "User map"}

        def maker(key: str) -> Any:
            def make(selected: bool) -> Text:
                return Text.assemble((f"{labels[key]:<20}", FG if selected else SUB), (values[key], f"bold {BLUE}"))
            return make

        return [BarOption(maker(k), id=k) for k in self.ROWS]

    def _refresh(self) -> None:
        ol = self.query_one(BarList)
        idx = ol.highlighted
        ol.set_options(self._options())
        ol.highlighted = idx
        self._help()

    def _help(self) -> None:
        key = self._row()
        text = {
            "overwrite": OVERWRITE_HELP[self.opts.overwrite],
            "restore_metadata": "set created/modified dates and author/editor at the destination",
            "workers": "files copied in parallel (1-16); SharePoint throttles above ~4",
            "excludes": "glob patterns, comma separated, e.g. 0. Archief/**, **/*.tmp",
            "lock_source_after": "after a clean migration, downgrade writers on the source library to Read",
            "usermap_path": "CSV source_email;target_email used to map authors and editors",
        }.get(key or "", "")
        self.query_one("#options-help", Static).update(Text(text, style=SUB))

    def _row(self) -> str | None:
        ol = self.query_one(BarList)
        return self.ROWS[ol.highlighted] if ol.highlighted is not None else None

    def on_option_list_option_highlighted(self, _event: OptionList.OptionHighlighted) -> None:
        self._help()

    def on_option_list_option_selected(self, event: OptionList.OptionSelected) -> None:
        event.stop()
        self.action_change()

    def action_change(self) -> None:
        o = self.opts
        row = self._row()
        if row == "overwrite":
            o.overwrite = OVERWRITE_CYCLE[(OVERWRITE_CYCLE.index(o.overwrite) + 1) % len(OVERWRITE_CYCLE)]
        elif row == "restore_metadata":
            o.restore_metadata = not o.restore_metadata
        elif row == "workers":
            self.action_workers(1)
            return
        elif row == "lock_source_after":
            o.lock_source_after = not o.lock_source_after
        elif row == "excludes":
            self.app.push_screen(PromptModal("Excludes", ", ".join(o.excludes),
                                             placeholder="0. Archief/**, **/*.tmp",
                                             note="Glob patterns, comma separated. Empty = exclude nothing."),
                                 self._set_excludes)
            return
        elif row == "usermap_path":
            self.app.push_screen(PromptModal("User map CSV", str(o.usermap_path or ""),
                                             placeholder="~/usermap.csv (empty = none)"), self._set_usermap)
            return
        self._refresh()

    def _set_excludes(self, value: str | None) -> None:
        if value is not None:
            self.opts.excludes = [p.strip() for p in value.split(",") if p.strip()]
            self._refresh()

    def _set_usermap(self, value: str | None) -> None:
        if value is not None:
            from pathlib import Path

            self.opts.usermap_path = Path(value.strip()).expanduser() if value.strip() else None
            self._refresh()

    def action_workers(self, delta: int) -> None:
        n = self.opts.workers + delta
        self.opts.workers = 1 if n > 16 else 16 if n < 1 else n
        self._refresh()

    def action_save(self) -> None:
        self.dismiss(self.opts)

    def action_cancel(self) -> None:
        self.dismiss(None)


# ------------------------------------------------------------------------------------------------ text


class TextModal(Dialog):
    """Read-only scrollable content (report.txt, help)."""

    DIALOG_CLASSES = "dialog viewer"
    BINDINGS: ClassVar[list[BindingType]] = [Binding("escape,q", "close", "close")]

    def __init__(self, title: str, content: RenderableType, *, fit: bool = False) -> None:
        super().__init__(title, _hint(("j/k", "scroll"), ("g/G", "top/bottom"), ("esc", "close")))
        self.content = content
        if fit:  # as tall as the content (help), instead of a fixed large viewer (reports)
            self.DIALOG_CLASSES = "dialog viewer fit"

    def compose(self) -> ComposeResult:
        with self.box(), VimScroll(id="viewer"):
            yield Static(self.content, id="viewer-body")

    def on_mount(self) -> None:
        self.query_one(VimScroll).focus()

    def action_close(self) -> None:
        self.dismiss(None)


def auth_text(ev: AuthPrompt) -> Text:
    t = Text(ev.message, style=FG)
    if ev.url:
        t.append("\n\nurl   ", style=SUB).append(ev.url, style=f"bold underline {BLUE}")
    if ev.code:
        t.append("\ncode  ", style=SUB).append_text(theme.pill(ev.code, ACCENT))
        t.append("\n\n").append_text(theme.key_hint("c", "copies the code"))
    return t


class AuthModal(Dialog):
    """Sign-in instructions. The app closes it as soon as the operation makes progress."""

    BINDINGS: ClassVar[list[BindingType]] = [Binding("escape", "close", "close"), Binding("c", "copy", "copy")]

    def __init__(self, ev: AuthPrompt) -> None:
        super().__init__(f"Sign in to tenant '{ev.tenant}'", _hint(("c", "copy code"), ("esc", "close")))
        self.ev = ev

    def compose(self) -> ComposeResult:
        with self.box():
            yield Static(auth_text(self.ev), id="auth-body")

    def show(self, ev: AuthPrompt) -> None:
        self.ev = ev
        self.query_one("#auth-body", Static).update(auth_text(ev))

    def action_copy(self) -> None:
        if self.ev.code:
            self.app.copy_to_clipboard(self.ev.code)
            self.notify("Code copied.")

    def action_close(self) -> None:
        self.dismiss(None)
