"""Section 2 - Checks: the audit. Inventories both sides and lists what could go wrong, worst first."""

from __future__ import annotations

from pathlib import Path
from typing import ClassVar

from rich.table import Table
from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal
from textual.widgets import DataTable, Static

from m365tools.services import SEVERITY_ORDER, sorted_findings
from m365tools.tools.spmig.models import Finding, RunSummary

from .. import theme
from ..modals import PromptModal
from ..theme import DIM, FG, SEVERITY_STYLE, SUB
from ..widgets import Activity, VimScroll, VimTable
from .base import Empty, Section


def finding_detail(f: Finding) -> Table:
    t = Table.grid(padding=(0, 2))
    t.add_column(style=SUB, no_wrap=True)
    t.add_column(style=FG)
    t.add_row("severity", theme.pill(f"{theme.DOT} {f.severity}", SEVERITY_STYLE.get(str(f.severity), SUB)))
    t.add_row("check", Text(f.category, style=f"bold {FG}"))
    t.add_row("path", f.path or Text("library level", style=DIM))
    t.add_row("", "")
    t.add_row("what", f.detail)
    t.add_row("", "")
    t.add_row("fix", Text(f.fix or "no suggestion", style=theme.GREEN if f.fix else DIM))
    return t


class ChecksSection(Section):
    KEY, TITLE, NUMBER = "checks", "Checks", 2
    TAGLINE, IDLE = "audit · inventory + checks · writes nothing", "a to run the audit"
    PRIMARY = "#findings"
    HELP = ("a runs the audit: both libraries are inventoried and every check runs. Nothing is written to "
            "SharePoint. Blocking findings stop a migration unless you force it.")
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("enter", "detail", "details", priority=True),
        Binding("slash", "filter", "filter", key_display="/"),
        Binding("o", "report", "report path"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.findings: list[Finding] = []
        self.shown: list[Finding] = []
        self.run_dir: Path | None = None
        self.filter_text = ""
        self.detail_open = False

    def compose(self) -> ComposeResult:
        yield Activity()
        yield Static("", id="checks-counts", classes="counts")
        with Horizontal(classes="split"):
            yield VimTable(id="findings")
            detail = VimScroll(id="finding-detail", classes="detail")
            detail.border_title = "details"
            with detail:
                yield Static("", id="finding-detail-body")
        yield Empty(theme.empty_text("No audit yet", "Inventories source and destination and runs every check.\n"
                                     "Nothing is written to SharePoint.", [("a", "run the audit")]),
                    id="checks-empty")

    def on_mount(self) -> None:
        self.query_one(VimTable).add_fit_columns(("severity", 10), ("check", 18), ("path", None), ("detail", None))
        self.query_one("#finding-detail").display = False
        self._redraw()

    # -- operation ----------------------------------------------------------------------------------

    @work(exclusive=True, group="op", exit_on_error=False)
    async def run_audit(self) -> None:
        ok, summary = await self.run_job("audit", lambda job: job.audit())
        if ok and summary is not None:
            self.show_findings(summary.findings, summary.run_dir, summary)

    def show_findings(self, findings: list[Finding], run_dir: Path | None, summary: RunSummary | None = None
                      ) -> None:
        """Called after an audit, and by the Run section when a migration was refused."""
        self.findings = sorted_findings(findings)
        self.run_dir = run_dir
        if summary is not None:
            self.m365.last_audit = summary
        blocking = sum(1 for f in findings if f.severity == "blocking")
        warn = sum(1 for f in findings if f.severity == "warn")
        self.status = theme.FAILED if blocking else theme.DONE
        self.summary = f"{blocking} blocking" + (f" · {warn} warn" if warn else "") if blocking else (
            f"no blocking · {warn} warn" if warn else "all clear")
        self._redraw()

    def reset(self) -> None:
        super().reset()
        self.findings, self.run_dir = [], None
        self._redraw()

    # -- rendering ----------------------------------------------------------------------------------

    def _redraw(self) -> None:
        words = self.filter_text.lower().split()
        self.shown = [f for f in self.findings
                      if all(w in f"{f.severity} {f.category} {f.path} {f.detail}".lower() for w in words)]
        counts = {str(s): sum(1 for f in self.findings if f.severity == s) for s in SEVERITY_ORDER}
        line = theme.counts_line(counts, SEVERITY_STYLE)
        if self.filter_text:
            line.append(" filter ", style=SUB).append(self.filter_text, style=f"bold {theme.MAUVE}")
            line.append(f" · {len(self.shown)} shown", style=SUB)
        self.query_one("#checks-counts", Static).update(line)
        table = self.query_one(VimTable)
        table.clear()
        for i, f in enumerate(self.shown):
            color = SEVERITY_STYLE.get(str(f.severity), SUB)
            table.add_row(theme.glyph_word(str(f.severity), color), Text(f.category, style=FG),
                          Text(f.path, style=FG, no_wrap=True, overflow="ellipsis"),
                          Text(f.detail, style=SUB, no_wrap=True, overflow="ellipsis"), key=str(i))
        has = bool(self.findings)
        self.query_one(".split").display = has
        self.query_one("#checks-counts").display = has
        self.query_one("#checks-empty").display = not has
        self._update_detail()

    def _update_detail(self) -> None:
        panel = self.query_one("#finding-detail")
        panel.display = self.detail_open and bool(self.shown)
        if panel.display:
            row = self.query_one(VimTable).cursor_row
            if 0 <= row < len(self.shown):
                self.query_one("#finding-detail-body", Static).update(finding_detail(self.shown[row]))

    def on_data_table_row_highlighted(self, _event: DataTable.RowHighlighted) -> None:
        self._update_detail()

    def on_data_table_row_selected(self, event: DataTable.RowSelected) -> None:
        event.stop()
        self.action_detail()

    # -- keys ---------------------------------------------------------------------------------------

    def action_detail(self) -> None:
        self.detail_open = not self.detail_open
        self._update_detail()

    def escape(self) -> bool:
        if self.detail_open:
            self.action_detail()
            return True
        return False

    def action_filter(self) -> None:
        before = self.filter_text

        def live(text: str) -> None:
            self.filter_text = text.strip()
            self._redraw()

        def done(value: str | None) -> None:
            self.filter_text = before if value is None else value.strip()
            self._redraw()

        self.app.push_screen(PromptModal("Filter findings", self.filter_text, placeholder="words to match",
                                         on_change=live, bottom=True), done)

    def action_report(self) -> None:
        self.m365.show_report_path(self.run_dir)
