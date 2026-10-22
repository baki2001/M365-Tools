"""``m365`` command line: tenants, profiles, the headless spmig runner and the groups tool (rich), plus the TUI
launcher.

Exit codes (spmig): 0 ok, 1 failures / blocking findings / errors, 2 migrate refused because of blocking
findings (BlockingFindingsError), 130 interrupted (Ctrl+C).
Exit codes (groups/onboard/offboard): 0 ok (previews too), 1 errors or a failed write; ``groups drift --strict``
and ``onboard check`` also exit 1 on warnings.
"""

import asyncio
import contextlib
import inspect
import json
import os
import signal
import threading
import time
from collections.abc import Awaitable, Callable
from dataclasses import dataclass, replace
from enum import Enum
from pathlib import Path
from typing import Annotated, Any

import typer
from rich.console import Console, Group
from rich.live import Live
from rich.markup import escape
from rich.panel import Panel
from rich.progress import BarColumn, Progress, TaskProgressColumn, TextColumn
from rich.table import Table
from rich.text import Text

from m365tools.core.config import (
    AppConfig,
    add_tenant,
    get_profile,
    get_tenant,
    load_config,
    profile_to_job_options,
    save_config,
)
from m365tools.core.errors import ConfigError, M365Error
from m365tools.core.events import (
    SCAN_PHASES,
    AuthPrompt,
    ItemStarted,
    ItemStep,
    LogLine,
    PhaseChanged,
    ProgressEvent,
    Stats,
    Throttled,
    phase_label,
)
from m365tools.core.report.format import fmt_duration, fmt_elapsed, fmt_size
from m365tools.services import (
    SEVERITY_ORDER,
    TenantClients,
    default_tenant,
    finding_to_json,
    load_run_options,
    lock_change_to_json,
    make_gateway_factory,
    make_groups_job,
    make_run_dir_factory,
    sorted_findings,
    summary_to_json,
    to_jsonable,
)
from m365tools.tools.spmig.job import BlockingFindingsError, Job
from m365tools.tools.spmig.models import Action, Endpoint, Finding, JobOptions, RunSummary, Severity

EXIT_OK, EXIT_FAIL, EXIT_BLOCKED, EXIT_INTERRUPTED = 0, 1, 2, 130
FINDINGS_PER_SEVERITY = 50

app = typer.Typer(
    name="m365", help="Microsoft 365 / SharePoint admin toolkit. Run without arguments for the TUI.",
    invoke_without_command=True, no_args_is_help=False, rich_markup_mode="rich",
)
tenant_app = typer.Typer(help="Manage tenants (Entra app registrations) and sign in.", no_args_is_help=True)
profile_app = typer.Typer(help="Show stored migration profiles.", no_args_is_help=True)
spmig_app = typer.Typer(help="SharePoint library migration: audit, dryrun, migrate, verify, lock.",
                        no_args_is_help=True)
app.add_typer(tenant_app, name="tenant")
app.add_typer(profile_app, name="profile")
groups_app = typer.Typer(help="SG-* access groups: inventory, suggest, drift, explain, sync, import.",
                         no_args_is_help=True)
app.add_typer(spmig_app, name="spmig")
app.add_typer(groups_app, name="groups")

LEVEL_STYLE = {"debug": "dim", "info": "", "ok": "green", "warn": "yellow", "error": "bold red"}
SEVERITY_STYLE = {Severity.BLOCKING: "bold red", Severity.WARN: "yellow", Severity.INFO: "cyan"}
ACCENT = "bold cyan"


def _console(stderr: bool = False) -> Console:
    return Console(stderr=stderr, highlight=False)


def _fail(console: Console, message: str, code: int = EXIT_FAIL) -> typer.Exit:
    console.print(f"[bold red]✗ error:[/] {escape(message)}")
    return typer.Exit(code)


def _load_cfg(console: Console) -> AppConfig:
    try:
        return load_config()
    except ConfigError as exc:
        raise _fail(console, str(exc)) from None


def launch_tui() -> None:
    from m365tools.tui.app import M365App

    M365App().run()


@app.callback()
def _root(ctx: typer.Context) -> None:
    if ctx.invoked_subcommand is None:
        launch_tui()


# ================================================================================================ rendering


class _StatusLine:
    """Phase / current file, replaced by a yellow 'throttled' line while a Retry-After wait is in progress."""

    def __init__(self) -> None:
        self.phase = ""
        self.current = ""
        self.throttle = ""
        self.throttle_until = 0.0

    def __rich__(self) -> Text:
        if self.throttle and time.monotonic() < self.throttle_until:
            left = self.throttle_until - time.monotonic()
            return Text(f"{self.throttle} ({left:.0f}s left)", style="yellow")
        t = Text()
        if self.phase:
            t.append(f"[{self.phase}] ", style="bold")
        t.append(self.current, style="dim")
        return t


class ConsoleRenderer:
    """EventSink rendering progress events with rich. Thread-safe (auth prompts arrive from worker threads)."""

    def __init__(self, console: Console, *, quiet: bool = False, progress: str | None = None) -> None:
        self.console = console
        self.quiet = quiet
        self.mode = progress  # "migrate" | "dryrun" | None
        self.lock = threading.RLock()
        self.status = _StatusLine()
        self.live: Live | None = None
        self.progress: Progress | None = None
        self._tasks: dict[str, Any] = {}
        self._last_plain = 0.0
        self._last_done = -1
        self._t0 = time.monotonic()

    # -------------------------------------------------------------------------------------------- live

    def start(self) -> None:
        if self.quiet or self.mode is None or not self.console.is_terminal:
            return
        self.progress = Progress(
            TextColumn("{task.description:>6}"), BarColumn(bar_width=None), TaskProgressColumn(),
            TextColumn("{task.fields[count]}"), TextColumn("[dim]{task.fields[extra]}"),
            console=self.console, auto_refresh=False, expand=True,
        )
        self._tasks["files"] = self.progress.add_task("files", total=None, count="", extra="")
        if self.mode == "migrate":
            self._tasks["bytes"] = self.progress.add_task("bytes", total=None, count="", extra="")
        self.live = Live(Group(self.progress, self.status), console=self.console, refresh_per_second=4,
                         transient=False)
        self.live.start()

    def stop(self) -> None:
        with self.lock:
            if self.live is not None:
                self.live.stop()
                self.live = None

    def print(self, *args: Any, **kw: Any) -> None:
        with self.lock:
            self.console.print(*args, **kw)

    # -------------------------------------------------------------------------------------------- events

    def __call__(self, ev: ProgressEvent) -> None:
        with self.lock:
            self._handle(ev)

    def _handle(self, ev: ProgressEvent) -> None:
        if isinstance(ev, AuthPrompt):
            body = Text(ev.message)
            if ev.url:
                body.append("\n\nURL:  ", style="bold").append(ev.url, style="bold underline cyan")
            if ev.code:
                body.append("\nCode: ", style="bold").append(ev.code, style="bold reverse green")
            self.console.print(Panel(body, title=f"[{ACCENT}]Sign in to {escape(ev.tenant)}[/]",
                                       border_style="cyan", expand=False))
        elif isinstance(ev, LogLine):
            if self.quiet and ev.level not in ("warn", "error"):
                return
            if ev.level == "debug" or ev.phase:  # phase-start lines: the ▸ line below already says it
                return
            self.console.print(Text(f"{ev.level:>5}  ", style=LEVEL_STYLE.get(ev.level, "") or "dim")
                               + Text(ev.text, style=LEVEL_STYLE.get(ev.level, "")))
        elif isinstance(ev, PhaseChanged):
            if ev.phase == self.status.phase:  # same phase, new sub-step (e.g. the next check)
                self.status.current = ev.detail
            else:
                self.status.phase, self.status.current = ev.phase, ""
                if not self.quiet:
                    line = Text(f"▸ {phase_label(ev.phase):<22}", style=ACCENT)
                    if ev.detail:
                        line.append(f" {ev.detail}")
                    line.append(f"  +{fmt_elapsed(time.monotonic() - self._t0)}", style="dim")
                    self.console.print(line)
        elif isinstance(ev, Throttled):
            self.status.throttle = f"throttled, waiting {ev.seconds:.0f}s (attempt {ev.attempt}) - {ev.what}"
            self.status.throttle_until = time.monotonic() + ev.seconds
            if self.live is None and not self.quiet:
                self.console.print(Text(self.status.throttle, style="yellow"))
        elif isinstance(ev, ItemStarted):
            self.status.current = f"{ev.index}/{ev.total} {ev.rel_path}"
        elif isinstance(ev, ItemStep):
            self.status.current = f"{ev.step:<8} {ev.rel_path}"
        elif isinstance(ev, Stats):
            self._stats(ev)
        if self.live is not None:
            self.live.refresh()

    def _stats(self, s: Stats) -> None:
        if self.status.phase in SCAN_PHASES:  # pre-phase progress, not transfer progress
            unit = {"users": "users", "checks": "checks"}.get(self.status.phase, "items")
            self.status.current = f"{s.done:,}/{s.total:,} {unit}"
            if self.live is None and not self.quiet and time.monotonic() - self._last_plain >= 10:
                self._last_plain = time.monotonic()
                self.console.print(Text(f"  {phase_label(self.status.phase)} {s.done:,}/{s.total:,} {unit}",
                                        style="dim"))
            return
        eta = fmt_duration(s.eta_s) if s.eta_s is not None else "--:--:--"
        if self.progress is not None:
            self.progress.update(self._tasks["files"], total=max(s.total, 1) if s.total else None,
                                 completed=s.done, count=f"{s.done:,}/{s.total:,}",
                                 extra=f"failed {s.failed}  warnings {s.warnings}")
            if "bytes" in self._tasks:
                self.progress.update(self._tasks["bytes"], total=max(s.bytes_total, 1), completed=s.bytes_done,
                                     count=f"{fmt_size(s.bytes_done)}/{fmt_size(s.bytes_total)}",
                                     extra=f"{s.mbps:.2f} MB/s  ETA {eta}")
            return
        if self.quiet or self.mode is None:
            return
        now = time.monotonic()
        final = bool(s.total) and s.done >= s.total and self._last_done != s.done
        if now - self._last_plain >= 10 or final:
            self._last_plain, self._last_done = now, s.done
            line = f"progress {s.done:,}/{s.total:,} files"
            if self.mode == "migrate":
                line += f", {fmt_size(s.bytes_done)}/{fmt_size(s.bytes_total)}, {s.mbps:.2f} MB/s, ETA {eta}"
            self.console.print(Text(line + f", failed {s.failed}", style="dim"))


def findings_table(findings: list[Finding], limit: int | None = FINDINGS_PER_SEVERITY) -> Table:
    t = Table(title="Findings", show_lines=False, expand=True, border_style="dim")
    t.add_column("Severity", no_wrap=True)
    t.add_column("Category", no_wrap=True)
    t.add_column("Path", overflow="fold", ratio=3)
    t.add_column("Detail", overflow="fold", ratio=4)
    t.add_column("Fix", overflow="fold", ratio=3)
    ordered = sorted_findings(findings)
    for sev in SEVERITY_ORDER:
        group = [f for f in ordered if f.severity == sev]
        if not group:
            continue
        shown = group if limit is None else group[:limit]
        for f in shown:
            t.add_row(Text(str(sev).upper(), style=SEVERITY_STYLE[sev]), Text(f.category), Text(f.path),
                      Text(f.detail), Text(f.fix))
        if len(group) > len(shown):
            t.add_row("", "", Text(f"... {len(group) - len(shown)} more {sev} finding(s), see findings.csv",
                                   style="dim"), "", "")
        t.add_section()
    return t


def _counts_line(findings: list[Finding]) -> Text:
    """``N blocking · N warn · N info`` with each count coloured by severity (dim when zero)."""
    t = Text()
    for i, sev in enumerate(SEVERITY_ORDER):
        n = sum(1 for f in findings if f.severity == sev)
        if i:
            t.append(" · ", style="dim")
        t.append(f"{n} {sev}", style=SEVERITY_STYLE[sev] if n else "dim")
    return t


def summary_table(s: RunSummary, extra: list[tuple[str, str]] | None = None) -> Table:
    t = Table(title=f"{s.mode} summary", show_header=False, border_style="dim")
    t.add_column(style="bold")
    t.add_column()
    rows: list[tuple[str, str]] = [("Duration", fmt_duration(s.duration_s))]
    if s.mode in ("dryrun", "migrate"):
        rows += [("Copied" if s.mode == "migrate" else "Would copy", f"{s.copied:,}"),
                 ("Skipped", f"{s.skipped:,}"), ("Failed", f"{s.failed:,}"),
                 ("Metadata warnings", f"{s.metadata_warnings:,}")]
        if s.mode == "migrate":
            rows.append(("Bytes copied", fmt_size(s.bytes_copied)))
        if s.stopped_early:
            rows.append(("Stopped early", "yes"))
    elif s.mode == "verify":
        rows += [("OK", f"{s.copied:,}"), ("Problems", f"{s.failed:,}")]
    if s.findings:
        rows.append(("Findings", ""))
    for k, v in rows + (extra or []):
        t.add_row(k, _counts_line(s.findings) if k == "Findings" else v)
    return t


def _print_paths(console: Console, run_dir: Path | None) -> None:
    if run_dir is None:
        return
    console.print(f"[dim]Run folder: {escape(str(run_dir))}[/]", soft_wrap=True)
    console.print(f"[bold]Report: {escape(str(Path(run_dir) / 'report.html'))}[/]", soft_wrap=True)


# ================================================================================================ options


class Overwrite(str, Enum):
    skip_if_same_or_newer = "skip_if_same_or_newer"
    always = "always"
    never = "never"


@dataclass
class CommonArgs:
    profile: str | None = None
    src_tenant: str | None = None
    src_site: str | None = None
    src_library: str | None = None
    src_folder: str | None = None
    dst_tenant: str | None = None
    dst_site: str | None = None
    dst_library: str | None = None
    dst_folder: str | None = None
    exclude: list[str] | None = None
    overwrite: Overwrite | None = None
    no_metadata: bool = False
    workers: int | None = None
    usermap: Path | None = None
    json_out: bool = False
    quiet: bool = False


_EP = "Endpoints"
_OPT = "Options"


def _common_params(
    profile: Annotated[str | None, typer.Option("--profile", "-p", help="Stored profile name.",
                                                rich_help_panel=_EP)] = None,
    src_tenant: Annotated[str | None, typer.Option(help="Source tenant (default: by site host).",
                                                   rich_help_panel=_EP)] = None,
    src_site: Annotated[str | None, typer.Option(help="Source site URL.", rich_help_panel=_EP)] = None,
    src_library: Annotated[str | None, typer.Option(help="Source library (default library if omitted).",
                                                    rich_help_panel=_EP)] = None,
    src_folder: Annotated[str | None, typer.Option(help="Source sub-folder.", rich_help_panel=_EP)] = None,
    dst_tenant: Annotated[str | None, typer.Option(help="Destination tenant (default: by site host).",
                                                   rich_help_panel=_EP)] = None,
    dst_site: Annotated[str | None, typer.Option(help="Destination site URL.", rich_help_panel=_EP)] = None,
    dst_library: Annotated[str | None, typer.Option(help="Destination library.", rich_help_panel=_EP)] = None,
    dst_folder: Annotated[str | None, typer.Option(help="Destination sub-folder.", rich_help_panel=_EP)] = None,
    exclude: Annotated[list[str] | None, typer.Option("--exclude", "-x", help="Glob on the relative path "
                                                      "(repeatable), e.g. '0. Archief/**'.",
                                                      rich_help_panel=_OPT)] = None,
    overwrite: Annotated[Overwrite | None, typer.Option(help="Overwrite policy.", rich_help_panel=_OPT)] = None,
    no_metadata: Annotated[bool, typer.Option("--no-metadata", help="Do not restore created/modified/author.",
                                              rich_help_panel=_OPT)] = False,
    workers: Annotated[int | None, typer.Option(min=1, max=16, help="Parallel file transfers.",
                                                rich_help_panel=_OPT)] = None,
    usermap: Annotated[Path | None, typer.Option(help="User map CSV (source;target;name;status).",
                                                 rich_help_panel=_OPT)] = None,
    json_out: Annotated[bool, typer.Option("--json", help="Print the run summary as JSON on stdout.",
                                           rich_help_panel=_OPT)] = False,
    quiet: Annotated[bool, typer.Option("--quiet", "-q", help="Only warnings, errors and the summary.",
                                        rich_help_panel=_OPT)] = False,
) -> None:  # pragma: no cover - signature donor only
    ...


_COMMON = list(inspect.signature(_common_params).parameters.values())
_COMMON_NAMES = [p.name for p in _COMMON]


def _spmig_command(name: str) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Register ``fn(common: CommonArgs, **own)`` as ``m365 spmig <name>`` with the shared endpoint options."""

    def deco(fn: Callable[..., Any]) -> Callable[..., Any]:
        own = [p for p in inspect.signature(fn).parameters.values() if p.name != "common"]

        def wrapper(**kw: Any) -> Any:
            common = CommonArgs(**{n: kw.pop(n) for n in _COMMON_NAMES})
            return fn(common, **kw)

        params = [p.replace(kind=inspect.Parameter.KEYWORD_ONLY) for p in [*own, *_COMMON]]
        wrapper.__signature__ = inspect.Signature(params)  # type: ignore[attr-defined]
        wrapper.__annotations__ = {p.name: p.annotation for p in params}
        wrapper.__name__ = fn.__name__
        wrapper.__doc__ = fn.__doc__
        spmig_app.command(name)(wrapper)
        return fn

    return deco


def _override(cfg: AppConfig, ep: Endpoint, tenant: str | None, site: str | None, library: str | None,
              folder: str | None) -> Endpoint:
    if not tenant and site:
        # A site on another tenant's host must not silently keep the profile's tenant (wrong token audience).
        tenant = default_tenant(cfg, site, fallback=ep.tenant)
    return replace(
        ep,
        tenant=tenant or ep.tenant,
        site_url=site.rstrip("/") if site else ep.site_url,
        library=library if library is not None else ep.library,
        folder=folder if folder is not None else ep.folder,
    )


def build_options(cfg: AppConfig, c: CommonArgs, *, need_dst: bool = True,
                  resume_dir: Path | None = None) -> JobOptions:
    """JobOptions from ``--profile`` (or a resumed run's options.json) plus explicit overrides."""
    if c.profile:
        opts = profile_to_job_options(cfg, c.profile)
    elif resume_dir is not None and not c.src_site:
        opts = load_run_options(resume_dir, cfg)
    else:
        if not c.src_site:
            raise ConfigError("Give --profile NAME, or the endpoints with --src-site URL"
                              + (" and --dst-site URL" if need_dst else "") + ".")
        if need_dst and not c.dst_site:
            raise ConfigError("--dst-site URL is required (or use --profile NAME).")
        src_t = c.src_tenant or default_tenant(cfg, c.src_site)
        get_tenant(cfg, src_t)
        src = Endpoint(src_t, c.src_site.rstrip("/"), c.src_library or None, c.src_folder or "")
        if c.dst_site:
            dst_t = c.dst_tenant or default_tenant(cfg, c.dst_site, fallback=src_t)
            get_tenant(cfg, dst_t)
            dst = Endpoint(dst_t, c.dst_site.rstrip("/"), c.dst_library or None, c.dst_folder or "")
        else:
            dst = src
        return _apply_flags(JobOptions(source=src, destination=dst), c)
    opts.source = _override(cfg, opts.source, c.src_tenant, c.src_site, c.src_library, c.src_folder)
    opts.destination = _override(cfg, opts.destination, c.dst_tenant, c.dst_site, c.dst_library, c.dst_folder)
    for ep in (opts.source, opts.destination):
        get_tenant(cfg, ep.tenant)
    return _apply_flags(opts, c)


def _apply_flags(opts: JobOptions, c: CommonArgs) -> JobOptions:
    if c.exclude:
        opts.excludes = [*opts.excludes, *(x for x in c.exclude if x not in opts.excludes)]
    if c.overwrite is not None:
        opts.overwrite = c.overwrite.value  # type: ignore[assignment]
    if c.no_metadata:
        opts.restore_metadata = False
    if c.workers:
        opts.workers = c.workers
    if c.usermap is not None:
        opts.usermap_path = c.usermap.expanduser()
    return opts


# ================================================================================================ job runner


def _run_interruptible(job: Any, op: Callable[[Any], Awaitable[Any]], renderer: ConsoleRenderer
                       ) -> tuple[Any, bool]:
    """Run ``op(job)`` on an event loop in a helper thread so SIGINT stays with us (the main thread).

    First Ctrl+C: ``job.request_stop()`` (graceful: files in flight finish, journal flushed); if no transfer
    is running yet the operation is cancelled. Second Ctrl+C: hard exit 130.
    Returns ``(result, interrupted)``; exceptions from ``op`` are re-raised here.
    """
    box: dict[str, Any] = {}
    ready = threading.Event()

    async def main() -> Any:
        box["loop"] = asyncio.get_running_loop()
        box["task"] = asyncio.current_task()
        ready.set()
        try:
            return await op(job)
        finally:
            close = getattr(job, "aclose", None)
            if close is not None:
                with contextlib.suppress(Exception):  # closing is best effort
                    await close()

    def target() -> None:
        try:
            box["result"] = asyncio.run(main())
        except BaseException as exc:  # noqa: BLE001 - re-raised in the main thread
            box["error"] = exc
        finally:
            ready.set()

    hits = 0

    def on_sigint(signum: int, frame: Any) -> None:
        nonlocal hits
        hits += 1
        if hits > 1:
            renderer.stop()
            renderer.print("[bold red]aborted[/]")
            os._exit(EXIT_INTERRUPTED)
        renderer.print("[yellow]Stopping: finishing files in progress (Ctrl+C again to abort immediately)[/]")
        with contextlib.suppress(Exception):
            job.request_stop()
        if getattr(job, "_migrator", None) is None and "task" in box:
            box["loop"].call_soon_threadsafe(box["task"].cancel)

    in_main = threading.current_thread() is threading.main_thread()
    previous = signal.signal(signal.SIGINT, on_sigint) if in_main else None
    thread = threading.Thread(target=target, name="m365-job", daemon=True)
    renderer.start()
    try:
        thread.start()
        while thread.is_alive():
            thread.join(0.2)
    finally:
        renderer.stop()
        if in_main:
            signal.signal(signal.SIGINT, previous)
    err = box.get("error")
    if isinstance(err, asyncio.CancelledError | KeyboardInterrupt):
        return None, True
    if err is not None:
        raise err
    return box.get("result"), hits > 0


def _emit_json(data: dict[str, Any]) -> None:
    typer.echo(json.dumps(data, indent=2, ensure_ascii=False, default=str))


def _execute(mode: str, c: CommonArgs, op: Callable[[Any], Awaitable[Any]], *, need_dst: bool = True,
             resume_dir: Path | None = None) -> tuple[Any, Any, ConsoleRenderer]:
    console = _console(stderr=c.json_out)
    cfg = _load_cfg(console)
    if resume_dir is not None and not Path(resume_dir).is_dir():
        # Job.migrate would create it and start a full migration there instead of resuming.
        raise _fail(console, f"--resume: run folder {resume_dir} does not exist")
    try:
        opts = build_options(cfg, c, need_dst=need_dst, resume_dir=resume_dir)
    except M365Error as exc:
        raise _fail(console, str(exc)) from None
    renderer = ConsoleRenderer(console, quiet=c.quiet, progress=mode if mode in ("migrate", "dryrun") else None)
    if not c.quiet:
        route = escape(opts.source.describe()) + (f" → {escape(opts.destination.describe())}" if need_dst else "")
        if resume_dir is not None:
            route += f"\n[dim]run folder: {escape(str(resume_dir))}[/]"
        console.print(Panel(route, title=f"[{ACCENT}]m365 spmig {mode}[/]", border_style="cyan", expand=False))
    job = Job(make_gateway_factory(cfg, renderer), opts, renderer, make_run_dir_factory(opts))
    try:
        result, interrupted = _run_interruptible(job, op, renderer)
    except BlockingFindingsError as exc:
        console.print(findings_table(exc.blocking, limit=None))
        console.print(Text("audit: ") + _counts_line(exc.findings))
        console.print(f"[bold red]✗ {escape(str(exc))}[/]")
        console.print("[yellow]![/] Fix the blocking findings, or re-run with --force to migrate anyway.")
        _print_paths(console, exc.run_dir)
        if c.json_out:
            _emit_json({"mode": mode, "error": "blocking_findings", "run_dir": str(exc.run_dir),
                        "findings": [finding_to_json(f) for f in exc.findings]})
        raise typer.Exit(EXIT_BLOCKED) from None
    except M365Error as exc:
        raise _fail(console, f"{type(exc).__name__}: {exc}") from None
    if interrupted and result is None:
        console.print("[yellow]! interrupted[/]")
        _print_paths(console, getattr(job, "run_dir", None))
        raise typer.Exit(EXIT_INTERRUPTED)
    return result, job, renderer


def _finish(console: Console, summary: RunSummary, c: CommonArgs, code: int, *,
            extra_rows: list[tuple[str, str]] | None = None, extra_json: dict[str, Any] | None = None) -> None:
    console.print(summary_table(summary, extra_rows))
    _print_paths(console, summary.run_dir)
    if c.json_out:
        _emit_json(summary_to_json(summary, **(extra_json or {})))
    raise typer.Exit(code)


# ================================================================================================ spmig


@_spmig_command("audit")
def audit(common: CommonArgs) -> None:
    """Read-only pre-flight: inventory both sides and run all checks. Exit 1 if blocking findings."""
    summary, _job, r = _execute("audit", common, lambda job: job.audit())
    if summary.findings:
        r.console.print(findings_table(summary.findings))
        r.console.print(Text("Findings: ") + _counts_line(summary.findings))
    else:
        r.console.print("[green]✓ No findings.[/]")
    _finish(r.console, summary, common, EXIT_FAIL if summary.blocking else EXIT_OK)


@_spmig_command("dryrun")
def dryrun(common: CommonArgs) -> None:
    """Build the plan and simulate the migration without writing anything."""
    res, _job, r = _execute("dryrun", common, lambda job: job.dryrun())
    plan, summary = res
    counts = plan.counts()
    t = Table(title="Plan", show_header=True, border_style="dim")
    t.add_column("Action")
    t.add_column("Count", justify="right")
    for a in Action:
        t.add_row(a.value, f"{counts.get(a, 0):,}")
    r.console.print(t)
    _finish(r.console, summary, common, EXIT_FAIL if summary.failed else EXIT_OK,
            extra_rows=[("Bytes to copy", fmt_size(plan.bytes_to_copy()))],
            extra_json={"plan": {a.value: n for a, n in counts.items()}, "bytes_to_copy": plan.bytes_to_copy()})


@_spmig_command("migrate")
def migrate(
    common: CommonArgs,
    resume: Annotated[Path | None, typer.Option("--resume", metavar="RUN_DIR",
                                                help="Continue an earlier (stopped) migrate run folder.")] = None,
    force: Annotated[bool, typer.Option("--force", help="Migrate even with blocking findings.")] = False,
) -> None:
    """Copy files + metadata (audits first; refuses on blocking findings unless --force)."""
    summary, _job, r = _execute("migrate", common, lambda job: job.migrate(resume_run_dir=resume, force=force),
                                resume_dir=resume)
    if summary.findings:
        r.console.print(Text("audit: ") + _counts_line(summary.findings))
    code = EXIT_OK
    if summary.failed:
        code = EXIT_FAIL
    if summary.stopped_early:
        r.console.print(f"[yellow]! stopped early.[/] resume with: m365 spmig migrate --resume "
                        f"{escape(str(summary.run_dir))}", soft_wrap=True)
        code = EXIT_INTERRUPTED
    _finish(r.console, summary, common, code)


@_spmig_command("verify")
def verify(common: CommonArgs) -> None:
    """Compare source and destination (Office files: size differences tolerated). Exit 1 on problems."""
    res, _job, r = _execute("verify", common, lambda job: job.verify())
    result, summary = res
    counts = result.counts()
    problems = [x for x in result.rows if x.status in ("missing", "size_differs", "source_newer")]
    if problems:
        t = Table(title="Problems", expand=True)
        t.add_column("Status", no_wrap=True)
        t.add_column("Path", overflow="fold")
        t.add_column("Source size", justify="right")
        t.add_column("Dest. size", justify="right")
        t.add_column("Detail", overflow="fold")
        for x in problems[:FINDINGS_PER_SEVERITY]:
            t.add_row(Text(x.status, style="red"), Text(x.rel_path),
                      "" if x.src_size is None else str(x.src_size),
                      "" if x.dst_size is None else str(x.dst_size), Text(x.detail))
        if len(problems) > FINDINGS_PER_SEVERITY:
            t.add_row("", f"... {len(problems) - FINDINGS_PER_SEVERITY} more, see verify.csv", "", "", "")
        r.console.print(t)
    _finish(r.console, summary, common, EXIT_FAIL if summary.failed else EXIT_OK,
            extra_rows=[(k, f"{v:,}") for k, v in sorted(counts.items()) if k != "ok"],
            extra_json={"counts": counts})


@_spmig_command("lock")
def lock(
    common: CommonArgs,
    apply: Annotated[bool, typer.Option("--apply", help="Apply the changes (default: preview only).")] = False,
    include_owners: Annotated[bool, typer.Option("--include-owners",
                                                 help="Also downgrade the site owners group.")] = False,
) -> None:
    """Make the source library read-only: downgrade principals that can write to Read. Never widens access."""
    changes, job, r = _execute("lock", common, lambda job: job.lock(apply=apply, include_owners=include_owners),
                               need_dst=False)
    t = Table(title="Lock " + ("(applied)" if apply else "preview"), expand=True)
    t.add_column("Principal", overflow="fold")
    t.add_column("Current roles", overflow="fold")
    t.add_column("After")
    for ch in changes:
        after = Text(ch.to_role, style="yellow") if ch.will_change else Text(f"unchanged ({ch.skipped_reason})",
                                                                               style="dim")
        t.add_row(Text(ch.title), Text(", ".join(ch.from_roles)), after)
    r.console.print(t)
    n = sum(1 for ch in changes if ch.will_change)
    if apply:
        r.console.print(f"[green]✓ {n} principal(s) set to Read.[/]")
    else:
        r.console.print(f"[yellow]![/] Preview only: {n} principal(s) would be set to Read. "
                        "Re-run with --apply to apply.")
    run_dir = getattr(job, "run_dir", None)
    _print_paths(r.console, run_dir)
    if common.json_out:
        _emit_json({"mode": "lock", "applied": apply, "run_dir": str(run_dir) if run_dir else None,
                    "changes": [lock_change_to_json(ch) for ch in changes]})
    raise typer.Exit(EXIT_OK)


@spmig_app.command("sites")
def sites(
    tenant: Annotated[str, typer.Argument(help="Tenant name.")],
    query: Annotated[str, typer.Argument(help="Search text (prefix match).")] = "",
    json_out: Annotated[bool, typer.Option("--json", help="Print JSON.")] = False,
) -> None:
    """Search site collections of a tenant."""
    console = _console(stderr=json_out)
    cfg = _load_cfg(console)
    renderer = ConsoleRenderer(console)
    clients = TenantClients(cfg, renderer)

    async def run() -> list[Any]:
        try:
            return await clients.search_sites(tenant, query)
        finally:
            await clients.aclose()

    try:
        hits = asyncio.run(run())
    except M365Error as exc:
        raise _fail(console, str(exc)) from None
    if json_out:
        _emit_json({"sites": [{"title": h.title, "url": h.url, "description": h.description} for h in hits]})
        return
    t = Table(title=f"Sites matching '{query}' ({len(hits)})", expand=True)
    t.add_column("Title")
    t.add_column("URL", overflow="fold")
    t.add_column("Description", overflow="fold", style="dim")
    for h in hits:
        t.add_row(Text(h.title), Text(h.url), Text(h.description))
    console.print(t)


@spmig_app.command("tui")
def tui() -> None:
    """Open the interactive terminal UI."""
    launch_tui()


# ================================================================================================ groups

TenantOpt = Annotated[str | None, typer.Option("--tenant", "-t", help="Tenant name (default: the only/first).")]
JsonOpt = Annotated[bool, typer.Option("--json", help="Print the result as JSON on stdout.")]
QuietOpt = Annotated[bool, typer.Option("--quiet", "-q", help="Only warnings, errors and the result.")]
ApplyOpt = Annotated[bool, typer.Option("--apply", help="Apply the changes (default: preview only).")]
STATUS_STYLE = {"planned": "yellow", "done": "green", "already": "dim", "not_created": "red", "failed": "bold red",
                "kept": "dim", "manual": "magenta", "ok": "green", "missing": "yellow", "unexplained": "yellow"}


def _groups_execute(title: str, tenant: str | None, json_out: bool, quiet: bool,
                    op: Callable[[Any], Awaitable[Any]]) -> tuple[Any, Any, ConsoleRenderer]:
    """Build a GroupsJob for the tenant and run ``op(job)`` with the shared renderer / Ctrl+C handling."""
    console = _console(stderr=json_out)
    cfg = _load_cfg(console)
    renderer = ConsoleRenderer(console, quiet=quiet)
    try:
        job = make_groups_job(cfg, tenant, renderer)
    except M365Error as exc:
        raise _fail(console, str(exc)) from None
    if not quiet:
        console.print(Panel(f"tenant {escape(job.tenant)}\n[dim]registry: {escape(str(job.registry_path))}[/]",
                            title=f"[{ACCENT}]m365 {title}[/]", border_style="cyan", expand=False))
    try:
        result, interrupted = _run_interruptible(job, op, renderer)
    except M365Error as exc:
        raise _fail(console, f"{type(exc).__name__}: {exc}") from None
    if interrupted and result is None:
        console.print("[yellow]! interrupted[/]")
        raise typer.Exit(EXIT_INTERRUPTED)
    return result, job, renderer


def _groups_finish(console: Console, job: Any, json_out: bool, data: dict[str, Any], code: int = EXIT_OK) -> None:
    _print_paths(console, job.run_dir)
    if json_out:
        _emit_json({"tenant": job.tenant, "registry": str(job.registry_path),
                    "run_dir": str(job.run_dir) if job.run_dir else None, **data})
    raise typer.Exit(code)


def _changes_table(title: str, rows: list[tuple[str, str, str]]) -> Table:
    t = Table(title=title, expand=True, border_style="dim")
    t.add_column("Group", overflow="fold")
    t.add_column("Detail", overflow="fold")
    t.add_column("Status", no_wrap=True)
    for group, detail, status in rows:
        t.add_row(Text(group), Text(detail, style="dim"), Text(status, style=STATUS_STYLE.get(status, "")))
    return t


def _preview_note(console: Console, apply: bool, n: int, what: str) -> None:
    if apply:
        return
    if n:
        console.print(f"[yellow]![/] Preview only: {n} {what}. Re-run with --apply to apply.")
    else:
        console.print("[green]✓ Nothing to do.[/]")


@groups_app.command("inventory")
def groups_inventory(tenant: TenantOpt = None, json_out: JsonOpt = False, quiet: QuietOpt = False) -> None:
    """Read SG-* groups, users, licences, enterprise apps and roles; write snapshot.json, matrix.csv, report."""
    snap, job, r = _groups_execute("groups inventory", tenant, json_out, quiet, lambda job: job.inventory())
    t = Table(title="inventory summary", show_header=False, border_style="dim")
    t.add_column(style="bold")
    t.add_column()
    for k, v in [("Groups", len(snap.groups)), ("People", len(snap.people)), ("Enterprise apps", len(snap.apps)),
                 ("Directory roles", len(snap.roles))]:
        t.add_row(k, f"{v:,}")
    t.add_row("Findings", _counts_line(job.findings))
    r.console.print(t)
    _groups_finish(r.console, job, json_out, {
        "mode": "inventory", "groups": len(snap.groups), "people": len(snap.people), "apps": len(snap.apps),
        "roles": len(snap.roles), "findings": [finding_to_json(f) for f in job.findings]})


def _parse_accept(values: list[str]) -> dict[str, str]:
    out: dict[str, str] = {}
    for v in values:
        key, _, name = v.partition("=")
        out[key.strip()] = name.strip()
    return out


@groups_app.command("suggest")
def groups_suggest(
    accept: Annotated[list[str] | None, typer.Option("--accept", metavar="KEY[=NAME]",
                                                     help="Accept candidate KEY (number or proposed name), "
                                                          "optionally renamed (repeatable).")] = None,
    no_near: Annotated[bool, typer.Option("--exact-only", help="Accept only people with exactly the candidate's "
                                                               "groups (not the nearest matches).")] = False,
    apply: ApplyOpt = False, tenant: TenantOpt = None, json_out: JsonOpt = False, quiet: QuietOpt = False,
) -> None:
    """Propose bundles from shared group sets; --accept KEY=NAME (+ --apply) writes them to the registry."""
    sug, job, r = _groups_execute("groups suggest", tenant, json_out, quiet, lambda job: job.suggest())
    names = {c.key: c.name for c in sug.candidates}
    t = Table(title="Candidate bundles", expand=True, border_style="dim")
    for col, kw in [("#", {"no_wrap": True}), ("Proposed name", {}), ("Groups", {"overflow": "fold", "ratio": 3}),
                    ("People", {"overflow": "fold", "ratio": 3})]:
        t.add_column(col, **kw)
    for c in sug.candidates:
        name = Text(c.name, style="bold") + (Text(" (existing)", style="dim") if c.existing else Text(""))
        t.add_row(c.key, name, Text(", ".join(c.groups)), Text(", ".join(c.people)))
    r.console.print(t)
    if sug.near:
        n = Table(title="Everyone else: nearest candidate", expand=True, border_style="dim")
        for col in ("Person", "Candidate", "Score", "+ addons", "- missing"):
            n.add_column(col, overflow="fold")
        for m in sug.near:
            n.add_row(Text(m.upn), Text(names.get(m.candidate, "-")), f"{m.score:.2f}",
                      Text(", ".join(m.addons), style="green"), Text(", ".join(m.missing), style="yellow"))
        r.console.print(n)
    changes: list[str] = []
    if accept:
        try:
            changes = job.accept(sug, _parse_accept(accept), apply=apply, include_near=not no_near)
        except KeyError as exc:
            raise _fail(r.console, f"--accept: no candidate {exc}") from None
        for line in changes:
            r.console.print(f"  [green]+[/] {escape(line)}")
        if apply:
            r.console.print(f"[green]✓ {len(changes)} registry change(s) written.[/]")
        _preview_note(r.console, apply, len(changes), "registry change(s)")
    _groups_finish(r.console, job, json_out, {
        "mode": "suggest", "candidates": to_jsonable(sug.candidates), "near": to_jsonable(sug.near),
        "accepted": changes, "applied": bool(apply and accept)})


@groups_app.command("drift")
def groups_drift(
    strict: Annotated[bool, typer.Option("--strict", help="Exit 1 when there are warnings.")] = False,
    tenant: TenantOpt = None, json_out: JsonOpt = False, quiet: QuietOpt = False,
) -> None:
    """Compare the registry with the tenant: unexplained, missing, unregistered, not created, unassigned..."""
    findings, job, r = _groups_execute("groups drift", tenant, json_out, quiet, lambda job: job.drift())
    if findings:
        r.console.print(findings_table(findings))
        r.console.print(Text("Findings: ") + _counts_line(findings))
    else:
        r.console.print("[green]✓ No drift.[/]")
    warn = any(f.severity == "warn" for f in findings)
    _groups_finish(r.console, job, json_out, {"mode": "drift", "findings": [finding_to_json(f) for f in findings]},
                   EXIT_FAIL if strict and warn else EXIT_OK)


def _kv_table(title: str, rows: list[tuple[str, str]]) -> Table:
    t = Table(title=title, show_header=False, box=None, padding=(0, 2))
    t.add_column(style="bold")
    t.add_column(overflow="fold")
    for k, v in rows:
        t.add_row(k, v or "-")
    return t


@groups_app.command("explain")
def groups_explain(
    subject: Annotated[str, typer.Argument(help="A UPN (or its unique local part) or a group name.")],
    tenant: TenantOpt = None, json_out: JsonOpt = False, quiet: QuietOpt = True,
) -> None:
    """Why does a person have what they have / who is in a group and why."""
    d, _job, r = _groups_execute("groups explain", tenant, json_out, quiet, lambda job: job.explain(subject))
    c = r.console
    if d["type"] == "person":
        c.print(_kv_table(d["upn"], [
            ("Name", d["name"]), ("Job title", d["job_title"]), ("Department", d["department"]),
            ("Bundles", ", ".join(d["bundles"]) or ("(not registered)" if not d["registered"] else "")),
            ("Add-ons", ", ".join(d["addons"])), ("Note", d["note"]), ("Offboarded", d["offboarded"]),
            ("Licences", ", ".join(d["licences"])),
            ("Apps", ", ".join(f"{a['app']} ({a['via']})" for a in d["apps"])), ("Roles", ", ".join(d["roles"]))]))
        t = Table(title="Groups", expand=True, border_style="dim")
        for col in ("Group", "Purpose", "Via", "Status"):
            t.add_column(col, overflow="fold")
        for g in d["groups"]:
            t.add_row(Text(g["group"]), Text(g["purpose"], style="dim"), Text(g["via"]),
                      Text(g["status"], style=STATUS_STYLE.get(g["status"], "")))
        c.print(t)
    else:
        c.print(_kv_table(d["name"], [
            ("Kind", d["kind"]), ("Purpose", d["purpose"] or ("(not registered)" if not d["registered"] else "")),
            ("Owner (registry)", d["owner"]), ("Owners (tenant)", ", ".join(d["tenant_owners"])),
            ("In tenant", "yes" if d["in_tenant"] else "no"), ("Bundles", ", ".join(d["bundles"])),
            ("Add-on of", ", ".join(d["addon_of"])), ("Apps", ", ".join(d["apps"])),
            ("Members", f"{len(d['members'])}: " + ", ".join(d["members"])),
            ("Unexplained", ", ".join(d["unexplained"]))]))
    if json_out:
        _emit_json(d)


@groups_app.command("sync")
def groups_sync(apply: ApplyOpt = False, tenant: TenantOpt = None, json_out: JsonOpt = False,
                quiet: QuietOpt = False) -> None:
    """Registry -> tenant: create missing groups (security, mail-disabled) and record their ids."""
    plan, job, r = _groups_execute("groups sync", tenant, json_out, quiet, lambda job: job.sync(apply=apply))
    created, failed = dict(plan.created), dict(plan.failed)
    rows = [(e.name, f"create: {e.purpose}",
             ("done" if e.name in created else "failed") if apply else "planned") for e in plan.create]
    rows += [(n, f"record id {gid}", "done" if apply else "planned") for n, gid in plan.fill_ids]
    r.console.print(_changes_table("Sync " + ("(applied)" if apply else "preview"), rows))
    for name, err in failed.items():
        r.console.print(f"[bold red]✗[/] {escape(name)}: {escape(err)}")
    _preview_note(r.console, apply, len(rows), "change(s)")
    _groups_finish(r.console, job, json_out, {"mode": "sync", "applied": apply, "plan": to_jsonable(plan)},
                   EXIT_FAIL if failed else EXIT_OK)


@groups_app.command("import")
def groups_import(apply: ApplyOpt = False, tenant: TenantOpt = None, json_out: JsonOpt = False,
                  quiet: QuietOpt = False) -> None:
    """Tenant -> registry: add unregistered prefix groups (purpose left empty for you to fill in)."""
    entries, job, r = _groups_execute("groups import", tenant, json_out, quiet,
                                      lambda job: job.import_(apply=apply))
    r.console.print(_changes_table("Import " + ("(applied)" if apply else "preview"),
                                   [(e.name, f"kind {e.kind}", "done" if apply else "planned") for e in entries]))
    if apply and entries:
        r.console.print(f"[green]✓ {len(entries)} group(s) added.[/] Fill in their purpose in "
                        f"{escape(str(job.registry_path))}", soft_wrap=True)
    _preview_note(r.console, apply, len(entries), "group(s) would be added to the registry")
    _groups_finish(r.console, job, json_out, {"mode": "import", "applied": apply, "groups": to_jsonable(entries)})


@app.command("onboard")
def onboard(
    target: Annotated[list[str], typer.Argument(metavar="UPN | check UPN",
                                                help="The person to onboard, or 'check UPN' for their drift.")],
    bundle: Annotated[list[str] | None, typer.Option("--bundle", "-b", help="Bundle (repeatable).")] = None,
    addon: Annotated[list[str] | None, typer.Option("--addon", "-a", help="Extra group (repeatable).")] = None,
    like: Annotated[str | None, typer.Option("--like", metavar="UPN", help="Copy this person's access.")] = None,
    apply: ApplyOpt = False, tenant: TenantOpt = None, json_out: JsonOpt = False, quiet: QuietOpt = False,
) -> None:
    """Add a person directly to every group of a bundle (+ add-ons), or like another person.

    Without --bundle/--like a registered person gets their registry entry re-applied (adds missing groups).
    'm365 onboard check UPN' shows the drift for one person (exit 1 on warnings)."""
    if target and target[0] == "check":
        if len(target) != 2:
            raise _fail(_console(), "usage: m365 onboard check UPN")
        upn = target[1]
        findings, job, r = _groups_execute("onboard check", tenant, json_out, quiet,
                                           lambda job: job.onboard_check(upn))
        if findings:
            r.console.print(findings_table(findings))
        else:
            r.console.print(f"[green]✓ {escape(upn)} matches the registry.[/]")
        if json_out:
            _emit_json({"mode": "onboard check", "upn": upn, "findings": [finding_to_json(f) for f in findings]})
        raise typer.Exit(EXIT_FAIL if any(f.severity == "warn" for f in findings) else EXIT_OK)
    if len(target) != 1:
        raise _fail(_console(), "give exactly one UPN (or: check UPN)")
    upn = target[0]
    plan, job, r = _groups_execute("onboard", tenant, json_out, quiet, lambda job: job.onboard(
        upn, bundles=bundle or [], addons=addon or [], like=like, apply=apply))
    rows = [(c.group, c.error or "", c.status) for c in plan.changes]
    rows += [(m, "not set by this tool: assign by hand", "manual") for m in plan.manual]
    title = f"Onboard {plan.upn} ({plan.source}: {', '.join(plan.bundles) or 'no bundle'})"
    r.console.print(_changes_table(title + (" applied" if apply else " preview"), rows))
    if plan.addons:
        r.console.print(f"[dim]add-ons: {escape(', '.join(plan.addons))}[/]")
    n = sum(1 for c in plan.changes if c.status in ("planned", "done", "failed"))
    failed = sum(1 for c in plan.changes if c.status == "failed")
    if apply:
        r.console.print(f"[green]✓ {n - failed} membership(s) added, registry entry written.[/]")
    _preview_note(r.console, apply, n, "membership(s) would be added")
    _groups_finish(r.console, job, json_out, {"mode": "onboard", "applied": apply, "plan": to_jsonable(plan)},
                   EXIT_FAIL if failed else EXIT_OK)


@app.command("offboard")
def offboard(
    upn: Annotated[str, typer.Argument(help="The person leaving.")],
    apply: ApplyOpt = False, tenant: TenantOpt = None, json_out: JsonOpt = False, quiet: QuietOpt = False,
) -> None:
    """Remove a person from every registry group and mark them offboarded (the account is not disabled)."""
    plan, job, r = _groups_execute("offboard", tenant, json_out, quiet,
                                   lambda job: job.offboard(upn, apply=apply))
    rows = [(c.group, c.error or "remove", c.status) for c in plan.changes]
    rows += [(g, "not in the registry: left alone", "kept") for g in plan.kept]
    r.console.print(_changes_table(f"Offboard {plan.upn}" + (" applied" if apply else " preview"), rows))
    failed = sum(1 for c in plan.changes if c.status == "failed")
    if apply:
        r.console.print(f"[green]✓ removed from {len(plan.changes) - failed} group(s); marked offboarded "
                        f"{plan.date}.[/] The account itself is not disabled.")
    _preview_note(r.console, apply, len(plan.changes), "membership(s) would be removed")
    _groups_finish(r.console, job, json_out, {"mode": "offboard", "applied": apply, "plan": to_jsonable(plan)},
                   EXIT_FAIL if failed else EXIT_OK)


# ================================================================================================ tenant


@tenant_app.command("add")
def tenant_add(
    name: Annotated[str, typer.Argument(help="Short name, e.g. 'contoso'.")],
    tenant_id: Annotated[str, typer.Option(prompt="Tenant id (GUID or <name>.onmicrosoft.com)")],
    client_id: Annotated[str, typer.Option(prompt="Application (client) id")],
    sharepoint_host: Annotated[str, typer.Option("--sharepoint-host", "--host",
                                                 prompt="SharePoint host (<name>.sharepoint.com)")],
) -> None:
    """Register a tenant (public-client app registration). See 'm365 tenant setup-help'."""
    console = _console()
    cfg = _load_cfg(console)
    existed = name.strip() in cfg.tenants
    try:
        t = add_tenant(cfg, name, tenant_id=tenant_id, client_id=client_id, sharepoint_host=sharepoint_host)
        path = save_config(cfg)
    except (M365Error, OSError) as exc:
        raise _fail(console, str(exc)) from None
    console.print(f"[green]✓[/] {'Updated' if existed else 'Added'} tenant [{ACCENT}]{escape(t.name)}[/] "
                  f"({escape(t.sharepoint_host)})")
    console.print(f"[dim]  saved to {escape(str(path))}[/]")
    console.print(f"  Next: [{ACCENT}]m365 tenant login {escape(t.name)}[/]")


@tenant_app.command("setup-help")
def tenant_setup_help(
    tenant_id: Annotated[str, typer.Argument(help="Your tenant id or domain.")] = "<yourtenant>.onmicrosoft.com",
) -> None:
    """How to create the Entra app registration this toolkit signs in with."""
    from m365tools.core.auth import tenant_setup_instructions

    _console().print(tenant_setup_instructions(tenant_id), markup=False)


def _signed_in_cell(tc: Any, cache_file_for: Callable[[Any], Path]) -> Text:
    """Cached account name(s) for a tenant, or a dim '-' (never raises; the list is informational)."""
    with contextlib.suppress(Exception):  # informational only
        if cache_file_for(tc).exists():
            from m365tools.core.auth import TokenProvider

            names = [str(a.get("username") or "") for a in TokenProvider(tc).accounts()]
            names = [n for n in names if n]
            if names:
                return Text(", ".join(names), style="green")
    return Text("-", style="dim")


@tenant_app.command("list")
def tenant_list() -> None:
    """List configured tenants."""
    from m365tools.core.auth import cache_file_for

    console = _console()
    cfg = _load_cfg(console)
    if not cfg.tenants:
        console.print("[yellow]![/] No tenants configured. Add one with: m365 tenant add NAME "
                      "(see m365 tenant setup-help)")
        return
    t = Table(title="Tenants", header_style=ACCENT, border_style="dim")
    t.add_column("Name", style="bold", no_wrap=True)
    t.add_column("Tenant id", overflow="fold")
    t.add_column("Host", overflow="fold")
    t.add_column("Signed in as", overflow="fold")
    for name, tc in sorted(cfg.tenants.items()):
        t.add_row(Text(name), Text(tc.tenant_id), Text(tc.sharepoint_host), _signed_in_cell(tc, cache_file_for))
    console.print(t)


@tenant_app.command("login")
def tenant_login(name: Annotated[str, typer.Argument(help="Tenant name.")]) -> None:
    """Sign in (browser or device code) and cache a SharePoint token."""
    console = _console()
    cfg = _load_cfg(console)
    renderer = ConsoleRenderer(console)
    clients = TenantClients(cfg, renderer)

    async def run() -> list[str]:
        try:
            return await clients.login(name)
        finally:
            await clients.aclose()

    try:
        accounts = asyncio.run(run())
    except M365Error as exc:
        raise _fail(console, str(exc)) from None
    console.print(f"[green]✓ Signed in as[/] [bold]{escape(', '.join(accounts) or '?')}[/]")


# ================================================================================================ profile


@profile_app.command("list")
def profile_list() -> None:
    """List stored migration profiles."""
    console = _console()
    cfg = _load_cfg(console)
    if not cfg.profiles:
        console.print(f"No profiles. Define [profiles.NAME] in {escape(str(cfg.path))} or save one from the TUI.")
        return
    t = Table(title="Profiles", expand=True)
    t.add_column("Name", no_wrap=True)
    t.add_column("Source", overflow="fold")
    t.add_column("Destination", overflow="fold")
    for name, p in sorted(cfg.profiles.items()):
        t.add_row(Text(name), Text(p.source.to_endpoint().describe()), Text(p.destination.to_endpoint().describe()))
    console.print(t)


@profile_app.command("show")
def profile_show(name: Annotated[str, typer.Argument(help="Profile name.")]) -> None:
    """Show one profile."""
    console = _console()
    cfg = _load_cfg(console)
    try:
        p = get_profile(cfg, name)
    except ConfigError as exc:
        raise _fail(console, str(exc)) from None
    t = Table(title=f"Profile '{name}'", show_header=False, box=None, padding=(0, 2))
    t.add_column(style="bold")
    t.add_column(overflow="fold")
    t.add_row("Source", f"[{p.source.tenant}] {p.source.to_endpoint().describe()}")
    t.add_row("Destination", f"[{p.destination.tenant}] {p.destination.to_endpoint().describe()}")
    for r in p.routes:
        t.add_row("Route", f"{r.prefix} -> [{r.destination.tenant}] {r.destination.to_endpoint().describe()}")
    t.add_row("Overwrite", p.overwrite)
    t.add_row("Restore metadata", "yes" if p.restore_metadata else "no")
    t.add_row("Excludes", ", ".join(p.excludes) or "-")
    t.add_row("Workers", str(p.workers))
    t.add_row("Lock source after", "yes" if p.lock_source_after else "no")
    t.add_row("User map", p.usermap_path or "-")
    console.print(t, markup=False)


def main() -> None:  # pragma: no cover
    app()


if __name__ == "__main__":  # pragma: no cover
    main()
