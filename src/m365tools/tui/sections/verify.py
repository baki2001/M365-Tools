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
from ..theme import FG, SUB, VERIFY_LABEL, VERIFY_STYLE
from ..widgets import Activity, VimTable
from .base import Empty, Section

# Rows put in the table. Each one takes about a quarter of a millisecond to build, so this keeps a redraw well
# under a second; f and / narrow the list, and the full set is in the run folder.
MAX_ROWS = 2_000
PROBLEMS = ("missing", "size_differs", "source_newer")
STATUSES = ("ok", "missing", "size_differs", "source_newer", "only_in_destination")
FILTERS: tuple[str | None, ...] = (None, "problems", *STATUSES)


class VerifySection(Section):
    KEY, TITLE, NUMBER = "verify", "Verify", 5
    TAGLINE, IDLE = "compare source and destination", "v after a migration"
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
        yield Empty(theme.empty_text("Not verified yet", "After a migration: compares source and destination file "
                                     "by file.", [("v", "verify")]), id="verify-empty")

    def on_mount(self) -> None:
        self.query_one(VimTable).add_fit_columns(("status", 21), ("path", None), ("source", 10),
                                                 ("destination", 11), ("detail", 24))
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
        line = theme.facets(theme.counts_line({k: counts.get(k, 0) for k in STATUSES}, VERIFY_STYLE, VERIFY_LABEL),
                            [("f", VERIFY_LABEL.get(self.status_filter or "", self.status_filter or "all")),
                             ("/", self.filter_text)])
        words = self.filter_text.lower().split()
        flt = self.status_filter
        rows = [r for r in res.rows
                if (flt is None or (r.status in PROBLEMS if flt == "problems" else r.status == flt))
                and all(w in f"{r.rel_path} {r.detail}".lower() for w in words)]
        if len(rows) > MAX_ROWS:
            line.append(f"  first {MAX_ROWS:,} of {len(rows):,} rows shown (see verify.csv)", style=theme.YELLOW)
        self.query_one("#verify-counts", Static).update(line)
        table = self.query_one(VimTable)
        table.clear()
        for r in rows[:MAX_ROWS]:
            table.add_row(theme.glyph_word(VERIFY_LABEL.get(r.status, r.status), VERIFY_STYLE.get(r.status, SUB)),
                          Text(r.rel_path + ("/" if r.is_folder else ""), style=FG, no_wrap=True,
                               overflow="ellipsis"),
                          Text(theme.size(r.src_size), justify="right", style=FG),
                          Text(theme.size(r.dst_size), justify="right", style=FG),
                          Text(r.detail, style=SUB, no_wrap=True, overflow="ellipsis"))

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
