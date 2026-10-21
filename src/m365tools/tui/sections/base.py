"""What every section (the panes in the main area) has in common.

A section is a container with:
  * ``KEY`` / ``TITLE`` / ``NUMBER``  identity, shown in the sidebar as ``○ 2 Checks``;
  * ``status`` / ``summary``          reactive; changing them updates the sidebar entry (via ``Changed``);
  * ``BINDINGS``                      its own keys, active while focus is inside it (shown in the key bar);
  * ``PRIMARY``                       CSS selector of the widget that gets focus when the pane is entered;
  * ``on_job_event``                  progress events while one of *its* operations runs;
  * ``escape``                        handle esc locally (close a detail panel ...); False = leave the pane;
  * ``reset``                         forget results after the job options changed.

``run_job`` is the one way to call the Job: it refuses when another operation runs, shows the activity line,
routes progress events here and turns any exception into a red toast + a log line instead of a crash.
"""

from __future__ import annotations

import asyncio
from collections.abc import Awaitable, Callable
from typing import TYPE_CHECKING, Any, ClassVar, TypeVar

from rich.text import Text
from textual.containers import Vertical
from textual.css.query import NoMatches
from textual.message import Message
from textual.reactive import reactive
from textual.widgets import Static

from m365tools.core.events import ProgressEvent

from .. import theme
from ..widgets import Activity, StepList

if TYPE_CHECKING:
    from ..app import M365App

T = TypeVar("T")


class Section(Vertical):
    KEY: ClassVar[str] = ""
    TITLE: ClassVar[str] = ""
    NUMBER: ClassVar[int] = 0
    PRIMARY: ClassVar[str] = ""
    HELP: ClassVar[str] = ""  # one paragraph shown at the top of the ? help
    TAGLINE: ClassVar[str] = ""  # right side of the title row
    IDLE: ClassVar[str] = ""  # sidebar line 2 while there is nothing to summarise (a hint, never bare)

    status: reactive[str] = reactive(theme.TODO)
    summary: reactive[str] = reactive("")

    class Changed(Message):
        """Status or summary changed: the workspace updates the sidebar entry."""

        def __init__(self, section: Section) -> None:
            super().__init__()
            self.section = section

    def __init__(self) -> None:
        super().__init__(id=self.KEY, classes="section")

    @property
    def m365(self) -> M365App:
        return self.app  # type: ignore[return-value]

    def watch_status(self) -> None:
        self.post_message(self.Changed(self))

    def watch_summary(self) -> None:
        self.post_message(self.Changed(self))

    # -- focus / navigation -------------------------------------------------------------------------

    def focus_primary(self) -> None:
        try:
            self.query_one(self.PRIMARY).focus()
        except NoMatches:
            self.focus()

    def escape(self) -> bool:
        """Handle esc inside the pane. Return True when handled (otherwise focus goes to the sidebar)."""
        return False

    def reset(self) -> None:
        """Results no longer match the job options."""
        self.status, self.summary = theme.TODO, ""

    def on_show(self) -> None:
        """Called when the section becomes visible (override to refresh)."""

    # -- operations ---------------------------------------------------------------------------------

    @property
    def activity(self) -> Activity | None:
        try:
            return self.query_one(Activity)
        except NoMatches:
            return None

    @property
    def steps(self) -> StepList | None:
        """Sections with a :class:`StepList` show it instead of the one-line activity while running."""
        try:
            return self.query_one(StepList)
        except NoMatches:
            return None

    def on_job_event(self, ev: ProgressEvent) -> None:
        if self.steps is not None:
            self.steps.feed(ev)
        elif self.activity is not None:
            self.activity.feed(ev)

    async def run_job(self, name: str, op: Callable[[Any], Awaitable[T]]) -> tuple[bool, T | None]:
        """Run ``op(job)`` as operation ``name``. Returns ``(ok, result)``; never raises (except cancel)."""
        app = self.m365
        job = app.begin(name, self)
        if job is None:
            return False, None
        previous = self.status
        self.status = theme.RUNNING
        steps = self.steps
        if steps is not None:
            steps.start(name)
            self.run_state_changed()
        elif self.activity is not None:
            self.activity.start(name)
        ok = False
        try:
            result = await op(job)
        except asyncio.CancelledError:
            self.status = previous
            raise
        except Exception as exc:  # noqa: BLE001 - every failure becomes a toast + log line
            self.status = theme.FAILED
            self.summary = f"{name} failed"
            app.report_error(exc, name)
            return False, None
        else:
            ok = True
        finally:
            if steps is not None:
                # Progress events reach us as messages on the app queue: finish after the ones still queued.
                app.call_later(self._finish_steps, steps, ok)
            elif self.activity is not None:
                self.activity.stop()
            app.end()
        return True, result


    def _finish_steps(self, steps: StepList, ok: bool) -> None:
        steps.finish(ok)
        self.run_state_changed()

    def run_state_changed(self) -> None:
        """The step list started or finished (override to hide the empty hint while running)."""


def set_text(widget: Static, text: str | Text, style: str = "") -> None:
    """Update a Static with plain text (user data is never parsed as markup)."""
    widget.update(Text(text, style=style) if isinstance(text, str) else text)


class Empty(Static):
    """Centred hint shown instead of an empty table (``Press a to run the audit``)."""

    DEFAULT_CSS = "Empty { width: 1fr; height: 1fr; content-align: center middle; text-align: center; }"
