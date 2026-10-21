"""Base classes shared by the screens."""

from __future__ import annotations

import time
import webbrowser
from pathlib import Path
from typing import TYPE_CHECKING, Any, ClassVar

from rich.text import Text
from textual.binding import Binding, BindingType
from textual.css.query import NoMatches
from textual.screen import Screen
from textual.widgets import RichLog, Static

from m365tools.core.events import LogLine, PhaseChanged, ProgressEvent, Stats, Throttled

from ..widgets import log_text

if TYPE_CHECKING:
    from ..app import M365App, WizardState


class BaseScreen(Screen[Any]):
    BINDINGS: ClassVar[list[BindingType]] = [Binding("escape", "back", "Back")]

    @property
    def m365(self) -> M365App:
        return self.app  # type: ignore[return-value]

    @property
    def state(self) -> WizardState:
        return self.m365.state

    def notify(self, message: str, *, markup: bool = False, **kw: Any) -> None:  # type: ignore[override]
        """Notifications carry user data (paths, server errors): no markup by default."""
        super().notify(message, markup=markup, **kw)

    def action_back(self) -> None:
        if len(self.app.screen_stack) > 2:
            self.app.pop_screen()

    def action_home(self) -> None:
        from .home import HomeScreen

        while len(self.app.screen_stack) > 2 and not isinstance(self.app.screen, HomeScreen):
            self.app.pop_screen()

    def open_report(self, run_dir: Path | None) -> None:
        if run_dir is None:
            self.notify("No report yet.", severity="warning")
            return
        report = Path(run_dir) / "report.html"
        self.app.copy_to_clipboard(str(report))
        self.notify(f"{report}\n(path copied to clipboard)", title="Report", timeout=8)
        if report.exists():
            self.run_worker(_open_in_browser(report), group="browser", exit_on_error=False)


async def _open_in_browser(path: Path) -> None:
    import asyncio

    await asyncio.to_thread(webbrowser.open, path.as_uri())


class JobScreen(BaseScreen):
    """A screen that runs job operations and receives their progress events.

    Optional child widgets it updates by id: ``#phase``, ``#throttle``, ``#current`` (Static), ``#log`` (RichLog).
    """

    running = False
    phase = ""

    def handle_job_event(self, ev: ProgressEvent) -> None:
        if isinstance(ev, Stats) and self.phase == "inventory":  # item scan, not transfer progress
            self._set("#current", Text(f"scanning {ev.done:,}/{ev.total:,} items", style="dim"))
            return
        if isinstance(ev, LogLine):
            if ev.level == "debug":
                return
            try:
                self.query_one("#log", RichLog).write(log_text(ev))
            except NoMatches:
                pass
        elif isinstance(ev, PhaseChanged):
            self.phase = ev.phase
            self._set("#current", "")
            self._set("#phase", Text(f"{ev.phase}" + (f": {ev.detail}" if ev.detail else ""), style="bold"))
        elif isinstance(ev, Throttled):
            until = time.monotonic() + ev.seconds
            self._set("#throttle", Text(f"throttled, waiting {ev.seconds:.0f}s (attempt {ev.attempt}) "
                                        f"- {ev.what}", style="yellow"))
            self.set_timer(max(ev.seconds, 0.1), lambda: self._clear_throttle(until))
        self.on_progress(ev)

    def on_progress(self, ev: ProgressEvent) -> None:
        """Screen-specific handling (override)."""

    def _clear_throttle(self, until: float) -> None:
        if time.monotonic() >= until - 0.05:
            self._set("#throttle", "")

    def _set(self, selector: str, content: Any) -> None:
        if isinstance(content, str):
            content = Text(content)  # never interpret user data as markup
        try:
            self.query_one(selector, Static).update(content)
        except NoMatches:
            pass

    def error(self, exc: BaseException) -> None:
        msg = f"{type(exc).__name__}: {exc}"
        try:
            self.query_one("#log", RichLog).write(Text(msg, style="bold red"))
        except NoMatches:
            pass
        self.notify(msg, title="Error", severity="error", timeout=10)
