"""Verify: compare source and destination."""

from __future__ import annotations

from pathlib import Path
from typing import ClassVar

from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.widgets import Button, DataTable, Footer, Header, Label, RichLog, Select, Static

from m365tools.core.report.format import fmt_size
from m365tools.tools.spmig.models import CompareResult

from ..widgets import MAX_TABLE_ROWS
from .base import JobScreen

PROBLEMS = ("missing", "size_differs", "source_newer")
STATUS_STYLE = {"ok": "green", "missing": "bold red", "size_differs": "red", "source_newer": "yellow",
                "only_in_destination": "cyan"}


class VerifyScreen(JobScreen):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("escape", "back", "Back"),
        Binding("o", "open_report", "Open report"),
        Binding("r", "rerun", "Re-run"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.result: CompareResult | None = None
        self.run_dir: Path | None = None

    def compose(self) -> ComposeResult:
        yield Header()
        with Vertical():
            yield Label("Verify", classes="section-title")
            yield Static("starting ...", id="phase")
            yield Static("", id="throttle")
            yield Static("", id="current")
            yield Static("", id="counts")
            with Horizontal(classes="row"):
                yield Label("Show", classes="field-label")
                yield Select([("problems", "problems"), ("all", "all")] + [(s, s) for s in STATUS_STYLE],
                             value="problems", allow_blank=False, id="filter")
            yield DataTable(id="rows", cursor_type="row", zebra_stripes=True)
            yield RichLog(id="log", classes="log-small", wrap=True, markup=False)
            with Horizontal(classes="buttons"):
                yield Button("Back", id="back")
                yield Button("Open report path", id="report", disabled=True)
        yield Footer()

    def on_mount(self) -> None:
        self.query_one("#rows", DataTable).add_columns("Status", "Path", "Source size", "Dest. size", "Detail")
        self.run_verify()

    @work(exclusive=True, group="job", exit_on_error=False)
    async def run_verify(self) -> None:
        self.running = True
        try:
            result, summary = await self.state.job.verify()
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self._set("#phase", "verify failed")
            self.error(exc)
            return
        finally:
            self.running = False
        self.result, self.run_dir = result, summary.run_dir
        counts = result.counts()
        t = Text()
        for k, style in STATUS_STYLE.items():
            n = counts.get(k, 0)
            t.append(f"{k}: {n:,}   ", style=style if n else "dim")
        self._set("#counts", t)
        problems = summary.failed
        self._set("#phase", Text(f"verify done: {problems} problem(s) - run folder {summary.run_dir}",
                                 style="bold red" if problems else "bold green"))
        self.query_one("#report", Button).disabled = False
        self.fill_table()

    def fill_table(self) -> None:
        table = self.query_one("#rows", DataTable)
        table.clear()
        if self.result is None:
            return
        flt = self.query_one("#filter", Select).value
        rows = [r for r in self.result.rows
                if flt == "all" or (flt == "problems" and r.status in PROBLEMS) or r.status == flt]
        for r in rows[:MAX_TABLE_ROWS]:
            table.add_row(Text(r.status, style=STATUS_STYLE.get(r.status, "")), Text(r.rel_path),
                          "" if r.src_size is None else fmt_size(r.src_size),
                          "" if r.dst_size is None else fmt_size(r.dst_size), Text(r.detail))

    def on_select_changed(self, event: Select.Changed) -> None:
        if event.select.id == "filter":
            self.fill_table()

    def on_button_pressed(self, event: Button.Pressed) -> None:
        {"back": self.action_back, "report": self.action_open_report}.get(event.button.id or "", lambda: None)()

    def action_open_report(self) -> None:
        self.open_report(self.run_dir)

    def action_rerun(self) -> None:
        if not self.running:
            self.run_verify()
