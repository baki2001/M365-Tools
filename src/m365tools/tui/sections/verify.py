"""Section 5 - Verify: compare source and destination after a migration."""

from __future__ import annotations

from pathlib import Path
from typing import ClassVar

from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.widgets import Static

from m365tools.tools.spmig.models import CompareResult

from .. import theme
from ..modals import PromptModal
from ..theme import ACCENT, DIM, FG, VERIFY_STYLE
from ..widgets import Activity, VimTable
from .base import Empty, Section

MAX_ROWS = 20_000
PROBLEMS = ("missing", "size_differs", "source_newer")
STATUSES = ("ok", "missing", "size_differs", "source_newer", "only_in_destination")
FILTERS: tuple[str | None, ...] = (None, "problems", *STATUSES)


class VerifySection(Section):
    KEY, TITLE, NUMBER = "verify", "Verify", 5
    PRIMARY = "#verify-table"
    HELP = ("v compares source and destination: missing files, size differences (Office files are allowed to "
            "differ), source newer than destination. f cycles the status filter, / filters by text.")
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("f", "cycle_filter", "status filter"),
        Binding("slash", "filter", "filter", key_display="/"),
        Binding("o", "report", "report path"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.result: CompareResult | None = None
        self.run_dir: Path | None = None
        self.status_filter: str | None = None
        self.filter_text = ""

    def compose(self) -> ComposeResult:
        yield Activity()
        yield Static("", id="verify-counts", classes="counts")
        yield VimTable(id="verify-table")
        yield Empty("Not verified yet.\n\nPress v after a migration to compare source and destination.",
                    id="verify-empty")

    def on_mount(self) -> None:
        table = self.query_one(VimTable)
        table.add_column("status", width=20)
        table.add_column("path", width=60)
        table.add_column("source", width=10)
        table.add_column("destination", width=11)
        table.add_column("detail")
        self._redraw()

    @work(exclusive=True, group="op", exit_on_error=False)
    async def run_verify(self) -> None:
        ok, out = await self.run_job("verify", lambda job: job.verify())
        if ok and out is not None:
            result, summary = out
            self.show_result(result, summary.run_dir)

    def show_result(self, result: CompareResult, run_dir: Path | None) -> None:
        self.result, self.run_dir = result, run_dir
        counts = result.counts()
        problems = sum(counts.get(k, 0) for k in PROBLEMS)
        self.status = theme.FAILED if problems else theme.DONE
        self.summary = f"{counts.get('ok', 0):,} ok" + (f" · {problems} problems" if problems else "")
        self._redraw()

    def reset(self) -> None:
        super().reset()
        self.result = None
        self._redraw()

    def _redraw(self) -> None:
        res = self.result
        self.query_one("#verify-counts").display = res is not None
        self.query_one(VimTable).display = res is not None
        self.query_one("#verify-empty").display = res is None
        if res is None:
            return
        counts = res.counts()
        line = theme.counts_line({k: counts.get(k, 0) for k in STATUSES}, VERIFY_STYLE)
        line.append("filter: ", style=DIM).append(self.status_filter or "all", style=ACCENT)
        if self.filter_text:
            line.append(f"  /{self.filter_text}", style=ACCENT)
        self.query_one("#verify-counts", Static).update(line)
        words = self.filter_text.lower().split()
        flt = self.status_filter
        rows = [r for r in res.rows
                if (flt is None or (r.status in PROBLEMS if flt == "problems" else r.status == flt))
                and all(w in f"{r.rel_path} {r.detail}".lower() for w in words)]
        table = self.query_one(VimTable)
        table.clear()
        for r in rows[:MAX_ROWS]:
            table.add_row(Text(r.status, style=VERIFY_STYLE.get(r.status, FG)),
                          Text(r.rel_path + ("/" if r.is_folder else ""), overflow="ellipsis"),
                          Text(theme.size(r.src_size), justify="right"),
                          Text(theme.size(r.dst_size), justify="right"),
                          Text(r.detail, style=DIM))

    def action_cycle_filter(self) -> None:
        self.status_filter = FILTERS[(FILTERS.index(self.status_filter) + 1) % len(FILTERS)]
        self._redraw()

    def action_filter(self) -> None:
        before = self.filter_text

        def live(text: str) -> None:
            self.filter_text = text.strip()
            self._redraw()

        def done(value: str | None) -> None:
            self.filter_text = before if value is None else value.strip()
            self._redraw()

        self.app.push_screen(PromptModal("Filter verify rows", self.filter_text, placeholder="path words",
                                         on_change=live, bottom=True), done)

    def action_report(self) -> None:
        self.m365.show_report_path(self.run_dir)
