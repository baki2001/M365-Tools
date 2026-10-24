"""Section 6 - Lock: make the source library read-only after the move. Only ever downgrades writers to Read."""

from __future__ import annotations

from typing import ClassVar

from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.widgets import Static

from m365tools.tools.spmig.lock import LockChange

from .. import theme
from ..modals import TypedConfirmModal
from ..theme import DIM, FG, GREEN, SUB, YELLOW
from ..widgets import Activity, VimTable
from .base import Empty, Section


class LockSection(Section):
    KEY, TITLE, NUMBER = "lock", "Lock", 6
    TAGLINE, IDLE = "source library to read-only", "l to preview"
    PRIMARY = "#lock-table"
    HELP = ("l previews which principals of the SOURCE library would be downgraded to Read (nothing changes). "
            "A applies it after you type 'lock'. Access is never widened; the site owners group is kept.")
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("l", "preview", "preview"),
        Binding("A", "apply", "apply lock"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.changes: list[LockChange] | None = None
        self.applied = False

    def compose(self) -> ComposeResult:
        yield Activity()
        yield Static("", id="lock-counts", classes="counts")
        yield VimTable(id="lock-table")
        yield Empty(theme.empty_text("No preview yet", "Shows who would be set to Read on the source library.\n"
                                     "Nothing changes until you press A and type 'lock'.", [("l", "preview")]),
                    id="lock-empty")

    def on_mount(self) -> None:
        self.query_one(VimTable).add_fit_columns(("principal", None), ("before", 28), (" ", 1), ("after", 12),
                                                 ("note", None))
        self._redraw()

    def check_action(self, action: str, parameters: tuple[object, ...]) -> bool | None:
        if action == "apply":
            return bool(self.changes) and not self.applied
        return True

    @work(exclusive=True, group="op", exit_on_error=False)
    async def action_preview(self) -> None:
        ok, changes = await self.run_job("lock", lambda job: job.lock(apply=False))
        if ok and changes is not None:
            self.show_changes(changes, applied=False)

    @work(exclusive=True, group="op", exit_on_error=False)
    async def action_apply(self) -> None:
        app = self.m365
        todo = [c for c in self.changes or [] if c.will_change]
        if not todo:
            app.notify("Nothing to lock: preview first (l), or nobody can write.", severity="warning")
            return
        body = Text.assemble((f"{theme.plural(len(todo), 'principal')} on the SOURCE library will be set to Read:\n", FG),
                             ("\n".join(f"  {c.title}" for c in todo[:12]), FG),
                             ("\n  ..." if len(todo) > 12 else "", DIM),
                             ("\n\nUsers keep read access. Undo needs a site owner.", DIM))
        if not await app.push_screen_wait(TypedConfirmModal("Lock the source library", body, "lock")):
            return
        ok, changes = await self.run_job("lock", lambda job: job.lock(apply=True))
        if ok and changes is not None:
            self.show_changes(changes, applied=True)

    def show_changes(self, changes: list[LockChange], applied: bool) -> None:
        self.changes, self.applied = changes, applied
        n = sum(1 for c in changes if c.will_change)
        self.status = theme.DONE
        self.summary = f"locked · {n} → Read" if applied else f"{n} to downgrade"
        self.refresh_bindings()
        self._redraw()

    def reset(self) -> None:
        super().reset()
        self.changes, self.applied = None, False
        self._redraw()

    def _redraw(self) -> None:
        has = self.changes is not None
        for sel in ("#lock-counts", "#lock-table"):
            self.query_one(sel).display = has
        self.query_one("#lock-empty").display = not has
        if self.changes is None:
            return
        n = sum(1 for c in self.changes if c.will_change)
        head = theme.counts_line({"downgrade": n, "untouched": len(self.changes) - n},
                                 {"downgrade": YELLOW, "untouched": SUB},
                                 {"downgrade": "to downgrade", "untouched": "left untouched"})
        head.append("\n")
        if self.applied:
            head.append_text(theme.notice(GREEN, ("applied", f"bold {GREEN}"), " · these principals now have Read"))
        else:
            head.append_text(theme.notice(YELLOW, ("preview only", f"bold {YELLOW}"), " · nothing changed yet · ",
                                          ("A", f"bold {theme.KEY}"), " applies after you type 'lock'"))
        self.query_one("#lock-counts", Static).update(head)
        table = self.query_one(VimTable)
        table.clear()
        for c in self.changes:
            after = theme.glyph_word(c.to_role, GREEN if self.applied else YELLOW) if c.will_change else \
                Text("unchanged", style=DIM)
            table.add_row(Text(c.title, style=FG, no_wrap=True, overflow="ellipsis"),
                          Text(", ".join(c.from_roles), style=SUB, no_wrap=True, overflow="ellipsis"),
                          Text("→", style=DIM), after, Text(c.skipped_reason or "", style=SUB))
