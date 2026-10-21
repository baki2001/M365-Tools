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
from ..theme import ACCENT, DIM, FG, GREEN, RED, YELLOW
from ..widgets import Activity, VimTable
from .base import Empty, Section


class LockSection(Section):
    KEY, TITLE, NUMBER = "lock", "Lock", 6
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
        yield Empty("No preview yet.\n\nPress l to see who would be set to Read on the source library.\n"
                    "Nothing changes until you press A and type 'lock'.", id="lock-empty")

    def on_mount(self) -> None:
        table = self.query_one(VimTable)
        table.add_column("principal", width=36)
        table.add_column("before", width=30)
        table.add_column(" ", width=1)
        table.add_column("after", width=10)
        table.add_column("note")
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
        body = Text.assemble((f"{len(todo)} principal(s) on the SOURCE library will be set to Read:\n", FG),
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
        head = Text()
        head.append("■ ", style=RED if n else DIM).append(f"{n} to downgrade", style=FG)
        head.append("   ■ ", style=DIM).append(f"{len(self.changes) - n} left untouched", style=FG)
        head.append("   ")
        head.append("APPLIED" if self.applied else "preview only: A applies",
                    style=f"bold {GREEN}" if self.applied else ACCENT)
        self.query_one("#lock-counts", Static).update(head)
        table = self.query_one(VimTable)
        table.clear()
        for c in self.changes:
            after = Text(c.to_role, style=f"bold {GREEN if self.applied else YELLOW}") if c.will_change else \
                Text("unchanged", style=DIM)
            table.add_row(Text(c.title, style=FG), Text(", ".join(c.from_roles), style=DIM),
                          Text("→", style=DIM), after, Text(c.skipped_reason or "", style=DIM))
