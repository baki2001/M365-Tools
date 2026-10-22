"""The Textual application: configuration, the current jobs, and the bridge from core events to the UI.

Screens are Textual *modes* that stay alive while you switch: ``home`` (the toolbox), ``spmig`` (the
:class:`Workspace`) and ``groups`` (the :class:`GroupsScreen`). ``g`` goes home when nothing runs.

How an operation flows (e.g. audit):
  1. a section calls ``Section.run_job("audit", lambda job: job.audit())`` inside an ``@work`` worker;
  2. ``run_job`` calls ``app.begin()``: refuses if something else runs, remembers the section as *listener*;
  3. core emits progress through ``app.sink`` (any thread) -> a ``JobEvent`` message -> ``on_job_event`` on the
     app thread -> LogLines to the log tail, everything to the listener section, AuthPrompts to a modal;
  4. ``app.end()`` clears the listener; the section renders the result.
"""

from __future__ import annotations

import contextlib
from collections.abc import Callable
from pathlib import Path
from typing import Any, ClassVar

from textual import work
from textual.app import App
from textual.binding import Binding, BindingType
from textual.message import Message

from m365tools.core.config import AppConfig, load_config, profile_to_job_options
from m365tools.core.events import AuthPrompt, LogLine, ProgressEvent, Throttled
from m365tools.services import TenantClients, make_gateway_factory, make_groups_job, make_run_dir_factory
from m365tools.tools.spmig.job import Job
from m365tools.tools.spmig.models import Endpoint, JobOptions, RunSummary

from .commands import M365Commands
from .groups import GroupsScreen
from .home import HomeScreen
from .modals import AuthModal, Choice, ConfirmModal, PickerModal
from .theme import THEME, THEME_VARIABLES
from .workspace import Workspace


class JobEvent(Message):
    """A core ProgressEvent, delivered on the app thread."""

    def __init__(self, event: ProgressEvent, listener: Any) -> None:
        super().__init__()
        self.event = event
        self.listener = listener  # the section whose operation emitted it (captured when posted)


def blank_options() -> JobOptions:
    return JobOptions(source=Endpoint("", ""), destination=Endpoint("", ""))


class M365App(App[None]):
    TITLE = "m365 toolbox"
    CSS_PATH = "app.tcss"
    COMMANDS: ClassVar[set[Any]] = {M365Commands}
    BINDINGS: ClassVar[list[BindingType]] = [Binding("ctrl+q", "quit_app", "quit", show=False, priority=True)]
    MODES: ClassVar[dict[str, Any]] = {"home": HomeScreen, "spmig": Workspace, "groups": GroupsScreen}
    TOOLS: ClassVar[tuple[str, ...]] = ("spmig", "groups")

    def __init__(
        self,
        cfg: AppConfig | None = None,
        *,
        job_cls: Callable[..., Any] = Job,
        gateway_factory: Callable[[AppConfig, Any], Any] = make_gateway_factory,
        clients: TenantClients | None = None,
        groups_job_factory: Callable[[AppConfig, str | None, Any], Any] | None = None,
        start: str = "home",
    ) -> None:
        super().__init__()
        self.cfg = cfg if cfg is not None else load_config()
        self.job_cls = job_cls
        self.gateway_factory = gateway_factory
        self.clients = clients if clients is not None else TenantClients(self.cfg, self.sink)
        self.opts: JobOptions = blank_options()
        self.tenant: str | None = next(iter(sorted(self.cfg.tenants)), None)
        self.account = ""
        self.busy: str | None = None  # name of the running operation
        self.last_audit: RunSummary | None = None  # last audit for the current options (migrate checks it)
        self.current_job: Any = None
        self.groups_job_factory = groups_job_factory or (lambda cfg, tenant, sink: make_groups_job(cfg, tenant, sink))
        self.groups_current: Any = None
        self.start = start
        self.busy_tool: str | None = None
        self._listener: Any = None
        self._auth_modal: AuthModal | None = None

    def get_theme_variable_defaults(self) -> dict[str, str]:
        """The palette tokens ($mauve, $surface1 ...) exist before the theme is applied, so app.tcss parses."""
        return {**super().get_theme_variable_defaults(), **THEME_VARIABLES}

    def on_mount(self) -> None:
        self.register_theme(THEME)
        self.theme = THEME.name
        self.switch_mode(self.start)

    @property
    def workspace(self) -> Workspace:
        """The spmig screen (raises LookupError when it was never opened)."""
        try:
            return next(s for s in self.get_screen_stack("spmig") if isinstance(s, Workspace))
        except (KeyError, StopIteration):
            raise LookupError("spmig is not open") from None

    @property
    def groups_screen(self) -> GroupsScreen | None:
        try:
            return next((s for s in self.get_screen_stack("groups") if isinstance(s, GroupsScreen)), None)
        except KeyError:
            return None

    # -------------------------------------------------------------------------------------------- tools

    def open_tool(self, tool: str) -> None:
        if self.current_mode == tool:
            return
        if self.busy:
            self.notify(f"{self.busy} is running: wait until it finishes.", severity="warning")
            return
        self.switch_mode(tool)

    def action_go_home(self) -> None:
        if self.current_mode == "home":
            return
        if self.busy:
            self.notify(f"{self.busy} is running: home is available when it finishes.", severity="warning")
            return
        self.switch_mode("home")

    def refresh_header(self) -> None:
        with contextlib.suppress(Exception):
            for header in self.screen.query("#header"):
                header.refresh()

    def notify(self, message: str, *, markup: bool = False, **kw: Any) -> None:  # type: ignore[override]
        """Notifications carry user data (paths, server errors): no markup by default."""
        super().notify(message, markup=markup, **kw)

    # -------------------------------------------------------------------------------------------- options

    def endpoints_ready(self) -> bool:
        return bool(self.opts.source.site_url and self.opts.destination.site_url)

    def set_options(self, opts: JobOptions) -> None:
        """Replace the job options. Results of earlier operations no longer apply and are cleared."""
        if self.busy:
            self.notify(f"{self.busy} is running: change the job after it finishes.", severity="warning")
            return
        self._close_job()
        changed = _comparable(opts) != _comparable(self.opts)
        self.opts = opts
        if opts.source.tenant:
            self.tenant = opts.source.tenant
        ws = self.workspace
        if changed:
            self.last_audit = None
            for key in ("checks", "plan", "verify", "lock"):
                ws.section(key).reset()
        ws.section("job").refresh_cards()
        self.refresh_header()

    def load_profile(self, name: str) -> None:
        try:
            opts = profile_to_job_options(self.cfg, name)
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.report_error(exc, "load profile")
            return
        self.set_options(opts)
        self.notify(f"Profile '{name}' loaded.")

    @work(group="tenant", exclusive=True, exit_on_error=False)
    async def action_pick_tenant(self) -> None:
        tenants = sorted(self.cfg.tenants)
        if not tenants:
            self.notify("No tenants configured: run 'm365 tenant add' first.", severity="error")
            return
        name = await self.push_screen_wait(PickerModal(
            "Switch tenant", [Choice(t, t, self.cfg.tenants[t].sharepoint_host) for t in tenants],
            highlight=self.tenant))
        if name is not None and name != self.tenant:
            self.switch_tenant(name)

    def switch_tenant(self, name: str) -> None:
        if self.busy:
            self.notify(f"{self.busy} is running: switch the tenant after it finishes.", severity="warning")
            return
        self.tenant, self.account = name, ""
        self._close_groups_job()
        self.refresh_header()
        gs = self.groups_screen
        if gs is not None:
            gs.tenant_changed()
        self.notify(f"Tenant: {name}")

    def groups_job(self) -> Any:
        """The GroupsJob of the current tenant (created on first use; raises when no tenant is configured)."""
        if self.groups_current is None:
            self.groups_current = self.groups_job_factory(self.cfg, self.tenant, self.sink)
        return self.groups_current

    def _close_groups_job(self) -> None:
        old, self.groups_current = self.groups_current, None
        if old is not None:
            self.run_worker(_aclose(old), group="close", exit_on_error=False)

    def job(self) -> Any:
        if self.current_job is None:
            self.current_job = self.job_cls(self.gateway_factory(self.cfg, self.sink), self.opts, self.sink,
                                            make_run_dir_factory(self.opts))
        return self.current_job

    def _close_job(self) -> None:
        old, self.current_job = self.current_job, None
        if old is not None:
            self.run_worker(_aclose(old), group="close", exit_on_error=False)

    # -------------------------------------------------------------------------------------------- operations

    def begin(self, name: str, listener: Any, *, tool: str = "spmig") -> Any:
        """Start operation ``name`` of ``tool``; returns its job, or None (with a toast) when it cannot start now.
        One operation at a time across all tools."""
        if self.busy:
            self.notify(f"Busy: {self.busy} is running.", severity="warning")
            return None
        if tool == "groups":
            try:
                job = self.groups_job()
            except Exception as exc:  # noqa: BLE001 - shown to the user
                self.report_error(exc, name)
                return None
        else:
            if not self.endpoints_ready():
                self.notify("Set source and destination first (1 Job, then e), or load a profile (p).",
                            severity="warning")
                return None
            job = self.job()
        self.busy, self.busy_tool, self._listener = name, tool, listener
        self.refresh_header()
        self.write_log(LogLine("info", f"── {name} ──"), listener)
        return job

    def end(self) -> None:
        tool = self.busy_tool
        self.busy, self.busy_tool, self._listener = None, None, None
        self.close_auth_prompt()
        self.refresh_header()
        if tool != "spmig":
            return
        with contextlib.suppress(Exception):
            ws = self.workspace
            ws.section("job").refresh_cards()
            ws.section("runs").action_reload()  # the operation left a run folder
            ws.refresh_bindings()

    def write_log(self, ev: LogLine, listener: Any = None) -> None:
        """Log lines go to the screen of the operation that emitted them, else to the current screen."""
        candidates = []
        with contextlib.suppress(Exception):
            if listener is not None:
                candidates.append(listener.screen)
        candidates.append(self.screen)
        for screen in candidates:
            fn = getattr(screen, "write_log", None)
            if fn is not None:
                fn(ev)
                return
        with contextlib.suppress(Exception):
            self.workspace.write_log(ev)

    def report_error(self, exc: BaseException, where: str) -> None:
        """A failure anywhere: red toast + log line, never a crash."""
        msg = f"{where}: {type(exc).__name__}: {exc}"
        self.notify(msg, title="Error", severity="error", timeout=10)
        with contextlib.suppress(Exception):
            self.write_log(LogLine("error", msg), self._listener)

    def show_report_path(self, run_dir: Path | None) -> None:
        if run_dir is None:
            self.notify("No report yet.", severity="warning")
            return
        report = Path(run_dir) / "report.html"
        self.copy_to_clipboard(str(report))
        self.notify(f"{report}\n(copied to clipboard)", title="Report", timeout=10)

    # -------------------------------------------------------------------------------------------- events

    def sink(self, ev: ProgressEvent) -> None:
        """EventSink handed to core; may be called from any thread."""
        self.post_message(JobEvent(ev, self._listener))

    def on_job_event(self, message: JobEvent) -> None:
        ev = message.event
        if isinstance(ev, AuthPrompt):
            if self._auth_modal is not None and self._auth_modal.is_attached:
                self._auth_modal.show(ev)
            else:
                self._auth_modal = AuthModal(ev)
                self.push_screen(self._auth_modal)
            return
        if not isinstance(ev, LogLine | Throttled):
            self.close_auth_prompt()  # progress means sign-in completed
        if isinstance(ev, LogLine):
            self.write_log(ev, message.listener)
        if message.listener is not None:
            message.listener.on_job_event(ev)

    def close_auth_prompt(self) -> None:
        modal, self._auth_modal = self._auth_modal, None
        if modal is not None and modal.is_attached and modal is self.screen:
            modal.dismiss()

    # -------------------------------------------------------------------------------------------- global actions

    @work(group="login", exclusive=True, exit_on_error=False)
    async def action_login(self) -> None:
        tenants = sorted(self.cfg.tenants)
        if not tenants:
            self.notify("No tenants configured: run 'm365 tenant add' first.", severity="error")
            return
        tenant = self.tenant or tenants[0]
        if len(tenants) > 1:
            tenant = await self.push_screen_wait(PickerModal(
                "Sign in to", [Choice(t, t, self.cfg.tenants[t].sharepoint_host) for t in tenants],
                highlight=tenant))
            if tenant is None:
                return
        self.notify(f"Signing in to {tenant} ...")
        try:
            accounts = await self.clients.login(tenant)
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.report_error(exc, "sign in")
            return
        finally:
            self.close_auth_prompt()
        if tenant != self.tenant:
            self._close_groups_job()
            gs = self.groups_screen
            self.tenant = tenant
            if gs is not None:
                gs.tenant_changed()
        self.tenant, self.account = tenant, ", ".join(accounts)
        self.refresh_header()
        self.notify(f"Signed in to {tenant} as {self.account or '?'}")

    def action_open_runs_dir(self) -> None:
        from m365tools.core.paths import runs_dir

        path = runs_dir()
        self.copy_to_clipboard(str(path))
        self.notify(f"{path}\n(copied to clipboard)", title="Runs folder", timeout=10)

    @work(group="quit", exclusive=True, exit_on_error=False)
    async def action_quit_app(self) -> None:
        if self.busy == "migrate":
            ok = await self.push_screen_wait(ConfirmModal(
                "Quit during a migration?",
                "Files in progress are finished, then the app quits.\nResume later with 4 Run -> r."))
            if not ok:
                return
            if self.current_job is not None:
                self.current_job.request_stop()
        elif self.busy:
            ok = await self.push_screen_wait(ConfirmModal(
                f"Quit while {self.busy} runs?", "The operation is cancelled. Writes already made stay made."))
            if not ok:
                return
        self.exit()

    async def on_unmount(self) -> None:
        if self.current_job is not None:
            await _aclose(self.current_job)
        if self.groups_current is not None:
            await _aclose(self.groups_current)
        with contextlib.suppress(Exception):
            await self.clients.aclose()


def _comparable(opts: JobOptions) -> JobOptions:
    from dataclasses import replace

    return replace(opts, profile=None)


async def _aclose(job: Any) -> None:
    close = getattr(job, "aclose", None)
    if close is None:
        return
    with contextlib.suppress(Exception):  # closing is best effort
        await close()
