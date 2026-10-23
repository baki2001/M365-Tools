"""Section 3 - Plan: the dry run. What a migration would do with every item, without writing anything."""

from __future__ import annotations

from pathlib import Path
from typing import ClassVar

from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.widgets import Static

from m365tools.core.events import DRYRUN_PHASES
from m365tools.tools.spmig.models import Action, Plan, PlanEntry

from .. import theme
from ..modals import PromptModal
from ..theme import ACTION_LABEL, ACTION_STYLE, FG, MAUVE, RED, SUB
from ..widgets import StepList, VimTable
from .base import Empty, Section

MAX_ROWS = 20_000
# The dry-run engine walk (folders/files) and the report are part of computing the plan.
DRYRUN_ALIASES = {"folders": "plan", "files": "plan", "finalize": "plan", "report": "plan"}
FILTERS: tuple[str | None, ...] = (None, *(a.value for a in Action))


class PlanSection(Section):
    KEY, TITLE, NUMBER = "plan", "Plan", 3
    TAGLINE, IDLE = "dry run · what a migration would do", "d for a dry run"
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
        yield StepList(DRYRUN_PHASES, aliases=DRYRUN_ALIASES)
        yield Static("", id="plan-counts", classes="counts")
        yield VimTable(id="plan-table")
        yield Static("", id="plan-totals", classes="totals")
        yield Empty(theme.empty_text("No dry run yet", "Computes the full plan: what would be created, copied, "
                                     "skipped.\nNothing is written to SharePoint.", [("d", "dry run")]),
                    id="plan-empty")

    def on_mount(self) -> None:
        self.query_one(VimTable).add_fit_columns(("action", 10), ("path", None), ("size", 10), ("modified", 16),
                                                 ("note", 24))
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

    def run_state_changed(self) -> None:
        self._redraw()

    def reset(self) -> None:
        super().reset()
        if self.steps is not None and not self.steps.running:
            self.steps.clear()
        self.plan = None
        self._redraw()

    def _blocked(self) -> dict[str, str]:
        """rel_path (lower) -> check name, for items the last audit flagged as blocking. The plan itself only
        says what a migration would do with each item; migrate refuses to start while these remain."""
        audit = self.m365.last_audit
        return {f.path.strip("/").lower(): f.category for f in audit.blocking} if audit is not None else {}

    def _matches(self, e: PlanEntry, words: list[str]) -> bool:
        if self.kind_filter and e.action != self.kind_filter:
            return False
        return all(w in f"{e.item.rel_path} {e.reason}".lower() for w in words)

    def _redraw(self) -> None:
        plan = self.plan
        has = plan is not None
        for sel in ("#plan-counts", "#plan-table", "#plan-totals"):
            self.query_one(sel).display = has
        steps = self.steps
        self.query_one("#plan-empty").display = not has and not (steps is not None and steps.running)
        if plan is None:
            return
        counts = {a.value: n for a, n in plan.counts().items()}
        line = theme.counts_line(counts, ACTION_STYLE, ACTION_LABEL)
        line.append(" f ", style=SUB).append(ACTION_LABEL.get(self.kind_filter or "", "all"), style=f"bold {MAUVE}")
        if self.filter_text:
            line.append("  / ", style=SUB).append(self.filter_text, style=f"bold {MAUVE}")
        blocked = self._blocked()
        n_blocked = sum(1 for e in plan.entries if e.item.rel_path.strip("/").lower() in blocked)
        if blocked:
            line.append(f"\n● {len(blocked):,} blocking audit finding(s)", style=f"bold {RED}")
            line.append(f" · {n_blocked:,} of these planned items are affected; migrate refuses until they are "
                        "fixed or excluded (2 checks)", style=SUB)
        self.query_one("#plan-counts", Static).update(line)
        words = self.filter_text.lower().split()
        shown = [e for e in plan.entries if self._matches(e, words)]
        table = self.query_one(VimTable)
        table.clear()
        for e in shown[:MAX_ROWS]:
            it = e.item
            color = ACTION_STYLE.get(e.action, SUB)
            check = blocked.get(it.rel_path.strip("/").lower())
            note = " · ".join(p for p in (f"BLOCKED: {check}" if check else "", e.reason,
                                          f"route {e.route}" if e.route else "") if p)
            table.add_row(theme.glyph_word(ACTION_LABEL.get(e.action, e.action), color, theme.BLOCK),
                          Text(it.rel_path + ("/" if it.is_folder else ""), style=FG, no_wrap=True,
                               overflow="ellipsis"),
                          Text("" if it.is_folder else theme.size(it.size), justify="right", style=FG),
                          Text(f"{it.modified:%Y-%m-%d %H:%M}" if it.modified else "", style=SUB),
                          Text(note, style=SUB, no_wrap=True, overflow="ellipsis"))
        files = [e for e in shown if not e.item.is_folder]
        tot = Text()
        tot.append("Σ ", style=f"bold {MAUVE}").append(f"{len(shown):,}", style=f"bold {FG}")
        tot.append(" entries · ", style=SUB).append(f"{len(files):,}", style=f"bold {FG}").append(" files · ", style=SUB)
        tot.append(theme.size(sum(e.item.size for e in files)), style=f"bold {FG}")
        tot.append(" · to copy ", style=SUB).append(theme.size(plan.bytes_to_copy()), style=f"bold {theme.GREEN}")
        if len(shown) > MAX_ROWS:
            tot.append(f" · first {MAX_ROWS:,} rows shown (see plan.json)", style=theme.YELLOW)
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
