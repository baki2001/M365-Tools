"""The Textual application: shared wizard state, job construction and event routing.

Job operations run in ``@work`` async workers on the app's event loop. Core emits progress through an
``EventSink``; ours (:meth:`M365App.sink`) wraps every event in a :class:`JobEvent` message and posts it to the
app (thread-safe: auth prompts come from MSAL worker threads). The app then hands it to the top-most
:class:`JobScreen` on the app thread, so all UI updates happen there.
"""

from __future__ import annotations

import contextlib
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any, ClassVar

from textual.app import App
from textual.binding import Binding, BindingType
from textual.message import Message

from m365tools.core.config import AppConfig, load_config
from m365tools.core.events import AuthPrompt, LogLine, ProgressEvent, Throttled
from m365tools.services import TenantClients, make_gateway_factory, make_run_dir_factory
from m365tools.tools.spmig.job import Job
from m365tools.tools.spmig.models import Endpoint, JobOptions, RunSummary


class JobEvent(Message):
    """A core ProgressEvent delivered on the app thread."""

    def __init__(self, event: ProgressEvent) -> None:
        super().__init__()
        self.event = event


def blank_options() -> JobOptions:
    return JobOptions(source=Endpoint("", ""), destination=Endpoint("", ""))


@dataclass
class WizardState:
    """What the wizard screens fill in and the job screens consume."""

    opts: JobOptions = field(default_factory=blank_options)
    force: bool = False
    job: Any = None
    audit: RunSummary | None = None


class M365App(App[None]):
    TITLE = "m365tools"
    SUB_TITLE = "SharePoint library migration"
    CSS_PATH = "app.tcss"
    BINDINGS: ClassVar[list[BindingType]] = [Binding("ctrl+q", "quit", "Quit", priority=True)]

    def __init__(
        self,
        cfg: AppConfig | None = None,
        *,
        job_cls: Callable[..., Any] = Job,
        gateway_factory: Callable[[AppConfig, Any], Any] = make_gateway_factory,
        clients: TenantClients | None = None,
    ) -> None:
        super().__init__()
        self.cfg = cfg if cfg is not None else load_config()
        self.job_cls = job_cls
        self.gateway_factory = gateway_factory
        self.clients = clients if clients is not None else TenantClients(self.cfg, self.sink)
        self.state = WizardState()
        self._auth_modal: Any = None

    def on_mount(self) -> None:
        self.theme = "textual-dark"
        from .screens.home import HomeScreen

        self.push_screen(HomeScreen())

    # -------------------------------------------------------------------------------------------- jobs

    def sink(self, ev: ProgressEvent) -> None:
        """EventSink handed to core; may be called from any thread."""
        self.post_message(JobEvent(ev))

    def new_job(self, opts: JobOptions | None = None) -> Any:
        """Replace the current job with a fresh one for ``opts`` (default: the wizard's options)."""
        old = self.state.job
        if old is not None:
            self.run_worker(_aclose(old), group="close", exit_on_error=False)
        opts = opts or self.state.opts
        self.state.opts = opts
        self.state.job = self.job_cls(self.gateway_factory(self.cfg, self.sink), opts, self.sink,
                                      make_run_dir_factory(opts))
        self.state.audit = None
        return self.state.job

    def on_job_event(self, message: JobEvent) -> None:
        from .modals import AuthPromptModal
        from .screens.base import JobScreen

        ev = message.event
        if isinstance(ev, AuthPrompt):
            if self._auth_modal is not None and self._auth_modal.is_attached:
                self._auth_modal.show(ev)
            else:
                self._auth_modal = AuthPromptModal(ev)
                self.push_screen(self._auth_modal)
            return
        if not isinstance(ev, LogLine | Throttled):
            self.close_auth_prompt()
        for screen in reversed(self.screen_stack):
            if isinstance(screen, JobScreen):
                screen.handle_job_event(ev)
                break

    def close_auth_prompt(self) -> None:
        modal, self._auth_modal = self._auth_modal, None
        if modal is not None and modal.is_attached and modal is self.screen:
            modal.dismiss()

    async def on_unmount(self) -> None:
        if self.state.job is not None:
            await _aclose(self.state.job)
        await self.clients.aclose()


async def _aclose(job: Any) -> None:
    close = getattr(job, "aclose", None)
    if close is None:
        return
    with contextlib.suppress(Exception):  # closing is best effort
        await close()
