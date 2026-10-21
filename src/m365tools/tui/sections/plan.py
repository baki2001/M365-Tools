"""Section 3 - Plan: the dry run. What a migration would do with every item, without writing anything."""

from __future__ import annotations

from pathlib import Path
from typing import ClassVar

from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.widgets import Static

from m365tools.tools.spmig.models import Action, Plan, PlanEntry

from .. import theme
from ..modals import PromptModal
from ..theme import ACCENT, ACTION_LABEL, ACTION_STYLE, DIM, FG
from ..widgets import Activity, VimTable
from .base import Empty, Section

MAX_ROWS = 20_000
FILTERS: tuple[str | None, ...] = (None, *(a.value for a in Action))


class PlanSection(Section):
    KEY, TITLE, NUMBER = "plan", "Plan", 3
    PRIMARY = "#plan-table"
    HELP = ("d runs a dry run: the full plan (create / copy / skip / conflict / excluded per item) without "
            "copying anything. f cycles the action filter, / filters by text.")
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("f", "cycle_filter", "action filter"),
        Binding("slash", "filter", "filter", key_display="/"),
        Binding("o", "report", "report path"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.plan: Plan | None = None
        self.run_dir: Path | None = None
        self.kind_filter: str | None = None
        self.filter_text = ""

    def compose(self) -> ComposeResult:
        yield Activity()
        yield Static("", id="plan-counts", classes="counts")
        yield VimTable(id="plan-table")
        yield Static("", id="plan-totals", classes="totals")
        yield Empty("No dry run yet.\n\nPress d to compute the plan: what would be created, copied, skipped.\n"
                    "Nothing is written to SharePoint.", id="plan-empty")

    def on_mount(self) -> None:
        table = self.query_one(VimTable)
        table.add_column("action", width=9)
        table.add_column("path", width=60)
        table.add_column("size", width=10)
        table.add_column("modified", width=16)
        table.add_column("note")
        self._redraw()

    @work(exclusive=True, group="op", exit_on_error=False)
    async def run_dryrun(self) -> None:
        ok, result = await self.run_job("dryrun", lambda job: job.dryrun())
        if ok and result is not None:
            plan, summary = result
            self.show_plan(plan, summary.run_dir)

    def show_plan(self, plan: Plan, run_dir: Path | None) -> None:
        self.plan, self.run_dir = plan, run_dir
        c = plan.counts()
        self.status = theme.DONE
        self.summary = f"{c[Action.COPY]:,} copy" + (f" · {c[Action.CONFLICT]} conflict" if c[Action.CONFLICT]
                                                     else "")
        self._redraw()

    def reset(self) -> None:
        super().reset()
        self.plan = None
        self._redraw()

    def _matches(self, e: PlanEntry, words: list[str]) -> bool:
        if self.kind_filter and e.action != self.kind_filter:
            return False
        return all(w in f"{e.item.rel_path} {e.reason}".lower() for w in words)

    def _redraw(self) -> None:
        plan = self.plan
        has = plan is not None
        for sel in ("#plan-counts", "#plan-table", "#plan-totals"):
            self.query_one(sel).display = has
        self.query_one("#plan-empty").display = not has
        if plan is None:
            return
        counts = {a.value: n for a, n in plan.counts().items()}
        line = theme.counts_line(counts, ACTION_STYLE, ACTION_LABEL)
        line.append("filter: ", style=DIM).append(ACTION_LABEL.get(self.kind_filter or "", "all"), style=ACCENT)
        if self.filter_text:
            line.append(f"  /{self.filter_text}", style=ACCENT)
        self.query_one("#plan-counts", Static).update(line)
        words = self.filter_text.lower().split()
        shown = [e for e in plan.entries if self._matches(e, words)]
        table = self.query_one(VimTable)
        table.clear()
        for e in shown[:MAX_ROWS]:
            it = e.item
            style = ACTION_STYLE.get(e.action, FG)
            note = e.reason + (f"  [route {e.route}]" if e.route else "")
            table.add_row(Text(ACTION_LABEL.get(e.action, e.action), style=style),
                          Text(it.rel_path + ("/" if it.is_folder else ""), overflow="ellipsis"),
                          Text("" if it.is_folder else theme.size(it.size), justify="right"),
                          Text(f"{it.modified:%Y-%m-%d %H:%M}" if it.modified else "", style=DIM),
                          Text(note, style=DIM))
        files = [e for e in shown if not e.item.is_folder]
        tot = Text()
        tot.append("Σ ", style=f"bold {ACCENT}").append(f"{len(shown):,} entries", style=FG)
        tot.append(f"  ·  {len(files):,} files  ·  {theme.size(sum(e.item.size for e in files))}", style=FG)
        tot.append(f"  ·  to copy {theme.size(plan.bytes_to_copy())}", style=ACTION_STYLE["copy"])
        if len(shown) > MAX_ROWS:
            tot.append(f"  ·  first {MAX_ROWS:,} rows shown (see plan.json)", style=theme.YELLOW)
        self.query_one("#plan-totals", Static).update(tot)

    def action_cycle_filter(self) -> None:
        self.kind_filter = FILTERS[(FILTERS.index(self.kind_filter) + 1) % len(FILTERS)]
        self._redraw()

    def action_filter(self) -> None:
        before = self.filter_text

        def live(text: str) -> None:
            self.filter_text = text.strip()
            self._redraw()

        def done(value: str | None) -> None:
            self.filter_text = before if value is None else value.strip()
            self._redraw()

        self.app.push_screen(PromptModal("Filter plan", self.filter_text, placeholder="path words",
                                         on_change=live, bottom=True), done)

    def action_report(self) -> None:
        self.m365.show_report_path(self.run_dir)
