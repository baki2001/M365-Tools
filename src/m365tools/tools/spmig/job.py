"""Job orchestration: the only spmig entry point the TUI/CLI call.

    job = Job(gateway_factory_or_tenants, opts, sink, run_dir_factory)
    await job.audit()            -> RunSummary (findings in summary.findings)
    await job.dryrun()           -> (Plan, RunSummary)
    await job.migrate(resume_run_dir=None, force=False) -> RunSummary   (raises BlockingFindingsError)
    await job.verify()           -> (CompareResult, RunSummary)
    await job.lock(apply=False)  -> list[LockChange]
    job.request_stop()           -> graceful stop of a running migrate/dryrun
"""

from __future__ import annotations

import json
from collections.abc import Callable, Mapping
from dataclasses import dataclass, field, replace
from datetime import datetime
from pathlib import Path
from typing import Any

from m365tools.core.errors import ConfigError, NotFoundError, ValidationError
from m365tools.core.events import EventSink, LogLine, PhaseChanged, PhaseClock, null_sink

from .checks import CheckContext, run_checks
from .engine import JournalLike, Migrator, StatusLike
from .gateway import GatewayFactory, RestGateway, SpGateway
from .inventory import build_inventory, inventory_base
from .lock import LockChange, apply_lock, preview_lock
from .models import (
    Action,
    CompareResult,
    Endpoint,
    Finding,
    Inventory,
    JobOptions,
    LibraryRef,
    Plan,
    RunSummary,
)
from .plan import build_plan, plan_to_json
from .usermap import ResolvedUsers, UserMap, build_usermap, collect_users, resolve_users, save_usermap
from .verify import compare

RunDirFactory = Callable[[str], Path]  # mode ("audit", "dryrun", "migrate", "verify", "lock") -> new run dir


class BlockingFindingsError(ValidationError):
    """migrate() refused to start because the audit found blocking problems (pass force=True to override)."""

    def __init__(self, findings: list[Finding], run_dir: Path) -> None:
        blocking = [f for f in findings if f.severity == "blocking"]
        super().__init__(f"{len(blocking)} blocking finding(s); see {run_dir / 'findings.csv'}")
        self.findings = findings
        self.blocking = blocking
        self.run_dir = run_dir


# ------------------------------------------------------------------------------------------- real gateways


class RestGatewayFactory:
    """:class:`GatewayFactory` building :class:`RestGateway` s from ``AppConfig.tenants``.

    One ``TokenProvider`` + ``SpClient`` per tenant, shared by all sites of that tenant. Call :meth:`aclose`
    when done (Job does this at the end of each operation it started)."""

    def __init__(self, tenants: Mapping[str, Any], sink: EventSink = null_sink) -> None:
        self.tenants = tenants
        self.sink = sink
        self._clients: dict[str, Any] = {}

    async def __call__(self, endpoint: Endpoint) -> SpGateway:
        from m365tools.core.auth import TokenProvider, sharepoint_scope
        from m365tools.core.http import SpClient

        tenant = self.tenants.get(endpoint.tenant)
        if tenant is None:
            raise ConfigError(f"Unknown tenant '{endpoint.tenant}'. Known: {', '.join(self.tenants) or '-'}")
        client = self._clients.get(endpoint.tenant)
        if client is None:
            provider = TokenProvider(tenant, prompt=self.sink)
            client = SpClient(provider, sharepoint_scope(tenant), sink=self.sink)
            self._clients[endpoint.tenant] = client
        return RestGateway(client, endpoint.site_url)

    async def aclose(self) -> None:
        for c in self._clients.values():
            await c.aclose()
        self._clients.clear()


def _default_run_dir_factory(opts: JobOptions) -> RunDirFactory:
    def make(mode: str) -> Path:
        from m365tools.core.paths import new_run_dir

        return new_run_dir(opts.profile or "adhoc", mode)

    return make


def _default_journal(run_dir: Path) -> JournalLike:
    from m365tools.core.journal import Journal

    return Journal(run_dir / "journal.jsonl")


def _default_status(run_dir: Path) -> StatusLike:
    from m365tools.core.journal import StatusWriter

    return StatusWriter(run_dir / "status.json")


def _csv(path: Path, columns: list[str], rows: list[list[Any]]) -> None:
    import csv

    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(columns)
        for r in rows:
            w.writerow(["" if c is None else c for c in r])


# ------------------------------------------------------------------------------------------- state


@dataclass
class _Side:
    endpoint: Endpoint
    gw: SpGateway | None = None
    lib: LibraryRef | None = None
    inv: Inventory | None = None
    base: str | None = None
    error: str | None = None


@dataclass
class _State:
    src: _Side
    dst: _Side
    routes: dict[str, _Side] = field(default_factory=dict)

    def route_bases(self) -> dict[str, str]:
        return {k: s.base for k, s in self.routes.items() if s.base}

    def route_invs(self) -> dict[str, Inventory]:
        return {k: s.inv for k, s in self.routes.items() if s.inv is not None}

    def route_targets(self) -> dict[str, tuple[SpGateway, LibraryRef]]:
        return {k: (s.gw, s.lib) for k, s in self.routes.items() if s.gw is not None and s.lib is not None}


class Job:
    def __init__(
        self,
        gateways: GatewayFactory | Mapping[str, Any],
        opts: JobOptions,
        sink: EventSink = null_sink,
        run_dir_factory: RunDirFactory | None = None,
        *,
        journal_factory: Callable[[Path], JournalLike] = _default_journal,
        status_factory: Callable[[Path], StatusLike] = _default_status,
        migrator_kwargs: dict[str, Any] | None = None,
    ) -> None:
        """``gateways`` is a :class:`GatewayFactory` (async ``endpoint -> SpGateway``) or a mapping of tenant
        name -> ``TenantConfig`` (then :class:`RestGatewayFactory` is used)."""
        if isinstance(gateways, Mapping):
            self._factory: GatewayFactory = RestGatewayFactory(gateways, sink)
        else:
            self._factory = gateways
        self.opts = opts
        self._clock = PhaseClock(sink)
        self.sink: EventSink = self._clock  # adds "<phase>: ... (+12.4s)" log lines at every phase start
        self.run_dir_factory = run_dir_factory or _default_run_dir_factory(opts)
        self.journal_factory = journal_factory
        self.status_factory = status_factory
        self.migrator_kwargs = dict(migrator_kwargs or {})
        self._gw_cache: dict[tuple[str, str], SpGateway] = {}
        self._migrator: Migrator | None = None
        # Last results, for the UI to show after an operation.
        self.state: _State | None = None
        self.plan: Plan | None = None
        self.findings: list[Finding] = []
        self.usermap: UserMap | None = None
        self.run_dir: Path | None = None

    # ----------------------------------------------------------------------------------- public API

    def request_stop(self) -> None:
        if self._migrator is not None:
            self._migrator.request_stop()

    async def aclose(self) -> None:
        close = getattr(self._factory, "aclose", None)
        if close is not None:
            await close()

    async def audit(self) -> RunSummary:
        started = datetime.now()
        self._clock.reset()
        run_dir = self._new_run_dir("audit")
        st = await self._load(run_dir)
        findings = await self._audit(st, run_dir)
        summary = RunSummary("audit", started, datetime.now(), run_dir, findings=findings)
        self._report(run_dir, "audit", summary, st=st)
        return summary

    async def dryrun(self) -> tuple[Plan, RunSummary]:
        started = datetime.now()
        self._clock.reset()
        run_dir = self._new_run_dir("dryrun")
        st = await self._load(run_dir)
        plan = self._plan(st, run_dir, phase=True)
        assert st.dst.gw is not None
        self._migrator = Migrator(st.src.gw, st.dst.gw, plan, self.opts, run_dir, self.sink, None, None,
                                  self._load_usermap(), dry_run=True, route_targets=st.route_targets(),
                                  **self.migrator_kwargs)
        try:
            summary = await self._migrator.run()
        finally:
            self._migrator = None
        summary.started = started
        self._report(run_dir, "dryrun", summary, plan=plan)
        return plan, summary

    async def migrate(self, resume_run_dir: Path | None = None, force: bool = False) -> RunSummary:
        started = datetime.now()
        self._clock.reset()
        if resume_run_dir is not None:
            run_dir = Path(resume_run_dir)
            run_dir.mkdir(parents=True, exist_ok=True)
            self.run_dir = run_dir
            self.sink(LogLine("info", f"resuming in {run_dir}"))
        else:
            run_dir = self._new_run_dir("migrate")
        st = await self._load(run_dir)
        findings = await self._audit(st, run_dir)
        blocking = [f for f in findings if f.severity == "blocking"]
        if st.dst.lib is None or st.dst.gw is None:
            raise BlockingFindingsError(findings, run_dir)
        if blocking and not force:
            raise BlockingFindingsError(findings, run_dir)
        if blocking:
            self.sink(LogLine("warn", f"continuing despite {len(blocking)} blocking finding(s) (force)"))
        plan = self.plan or self._plan(st, run_dir)
        assert st.src.gw is not None
        migrator = Migrator(
            st.src.gw, st.dst.gw, plan, self.opts, run_dir, self.sink, self.journal_factory(run_dir),
            self.status_factory(run_dir), self.usermap, route_targets=st.route_targets(), **self.migrator_kwargs,
        )
        self._migrator = migrator
        try:
            summary = await migrator.run()
        finally:
            self._migrator = None
        summary.started = started
        summary.findings = findings
        if self.opts.lock_source_after:
            if summary.stopped_early or summary.failed:
                self.sink(LogLine("warn", "source library NOT locked: the migration did not complete cleanly"))
            else:
                assert st.src.lib is not None
                changes = await preview_lock(st.src.gw, st.src.lib)
                applied = await apply_lock(st.src.gw, st.src.lib, changes, self.sink)
                self._write_lock_csv(run_dir, changes, applied=True)
                self.sink(LogLine("ok", f"source locked: {len(applied)} principal(s) set to Read"))
        summary.finished = datetime.now()
        self._report(run_dir, "migrate", summary, plan=plan, results=migrator.results)
        return summary

    async def verify(self) -> tuple[CompareResult, RunSummary]:
        started = datetime.now()
        self._clock.reset()
        run_dir = self._new_run_dir("verify")
        st = await self._load(run_dir)
        if st.dst.inv is None:
            raise NotFoundError(st.dst.error or "destination library not found")
        plan = self._plan(st, run_dir, write=False)
        main = [e.item for e in plan.entries if e.action != Action.EXCLUDED and e.route is None]
        result = compare(Inventory(st.src.inv.library, st.src.inv.base_url, main), st.dst.inv)
        for prefix, side in st.routes.items():
            if side.inv is None:
                continue
            sub = []
            for e in plan.entries:
                if e.route != prefix or e.action == Action.EXCLUDED:
                    continue
                rel = e.item.rel_path[len(prefix):].lstrip("/")
                if rel:
                    sub.append(replace(e.item, rel_path=rel))
            rres = compare(Inventory(st.src.inv.library, st.src.inv.base_url, sub), side.inv)
            for r in rres.rows:
                r.rel_path = f"{prefix}/{r.rel_path}"
                r.detail = (r.detail + " " if r.detail else "") + f"[route {prefix}]"
            result.rows.extend(rres.rows)
        counts = result.counts()
        problems = sum(counts.get(k, 0) for k in ("missing", "size_differs", "source_newer"))
        _csv(run_dir / "verify.csv",
             ["status", "type", "path", "src_size", "dst_size", "src_modified", "dst_modified", "detail"],
             [[r.status, "folder" if r.is_folder else "file", r.rel_path, r.src_size, r.dst_size,
               r.src_modified.isoformat() if r.src_modified else "",
               r.dst_modified.isoformat() if r.dst_modified else "", r.detail] for r in result.rows])
        summary = RunSummary("verify", started, datetime.now(), run_dir, copied=counts.get("ok", 0),
                             failed=problems)
        self.sink(LogLine("warn" if problems else "ok",
                          "verify: " + ", ".join(f"{k} {v}" for k, v in sorted(counts.items()))))
        self._report(run_dir, "verify", summary, compare_result=result)
        return result, summary

    async def lock(self, apply: bool = False, include_owners: bool = False) -> list[LockChange]:
        self._clock.reset()
        run_dir = self._new_run_dir("lock")
        self.sink(PhaseChanged("lock", "preview" if not apply else "apply"))
        gw = await self._gateway(self.opts.source)
        lib = await self._resolve_library(gw, self.opts.source)
        changes = await preview_lock(gw, lib, include_owners=include_owners)
        if apply:
            await apply_lock(gw, lib, changes, self.sink)
        self._write_lock_csv(run_dir, changes, applied=apply)
        try:
            from m365tools.core.report import lock_report, write_all

            write_all(run_dir, lock_report(self._lock_rows(changes), applied=apply))
        except Exception as exc:  # noqa: BLE001 - reports are best effort
            self.sink(LogLine("warn", f"report not written: {exc}"))
        return changes

    # ----------------------------------------------------------------------------------- internals

    def _new_run_dir(self, mode: str) -> Path:
        run_dir = Path(self.run_dir_factory(mode))
        run_dir.mkdir(parents=True, exist_ok=True)
        self.run_dir = run_dir
        self.sink(LogLine("info", f"run folder: {run_dir}"))
        return run_dir

    async def _gateway(self, ep: Endpoint) -> SpGateway:
        key = (ep.tenant, ep.site_url.rstrip("/").lower())
        gw = self._gw_cache.get(key)
        if gw is None:
            gw = await self._factory(ep)
            self._gw_cache[key] = gw
        return gw

    @staticmethod
    async def _resolve_library(gw: SpGateway, ep: Endpoint) -> LibraryRef:
        if ep.library:
            return await gw.get_library(ep.library)
        return await gw.get_default_library()

    async def _side(self, ep: Endpoint, label: str, required: bool, *, phase: str = "inventory.destination",
                    permissions_phase: str | None = None) -> _Side:
        """Resolve + inventory one endpoint. Only the source gets the unique-permission scan (the checks and
        reports never look at destination permissions, and it costs one request per 5000 items)."""
        side = _Side(ep)
        if phase != "inventory.source":
            self.sink(PhaseChanged(phase, f"{label}: {ep.describe()}"))
        try:
            side.gw = await self._gateway(ep)
            side.lib = await self._resolve_library(side.gw, ep)
        except NotFoundError as exc:
            if required:
                raise
            side.error = str(exc) + (f" (available: {', '.join(exc.available)})" if exc.available else "")
            self.sink(LogLine("error", f"{label}: {side.error}"))
            return side
        side.base = inventory_base(side.lib, ep.folder)
        side.inv = await build_inventory(side.gw, side.lib, ep.folder, self.sink, label=label, phase=phase,
                                         permissions_phase=permissions_phase)
        return side

    async def _load(self, run_dir: Path) -> _State:
        src = await self._side(self.opts.source, "source", required=True, phase="inventory.source",
                               permissions_phase="permissions.source")
        dst = await self._side(self.opts.destination, "destination", required=False)
        st = _State(src, dst)
        for r in self.opts.routes:
            st.routes[r.prefix.strip("/")] = await self._side(r.destination, f"route {r.prefix}", required=False)
        self.state = st
        self.plan = None
        self.usermap = self._load_usermap()
        return st

    def _load_usermap(self) -> UserMap | None:
        p = self.opts.usermap_path
        if p is not None and Path(p).exists():
            return UserMap.load(Path(p))
        return None

    def _plan(self, st: _State, run_dir: Path, write: bool = True, phase: bool = False) -> Plan:
        if st.dst.inv is None or st.dst.base is None:
            raise NotFoundError(st.dst.error or "destination library not found")
        assert st.src.inv is not None
        if phase:
            self.sink(PhaseChanged("plan", f"{len(st.src.inv.items):,} source item(s)"))
        plan = build_plan(st.src.inv, st.dst.inv, self.opts, st.dst.base, st.route_bases(), st.route_invs())
        self.plan = plan
        if write:
            (run_dir / "plan.json").write_text(json.dumps(plan_to_json(plan), ensure_ascii=False, indent=1),
                                               encoding="utf-8")
            c = plan.counts()
            self.sink(LogLine("info", "plan: " + ", ".join(f"{a.value} {n}" for a, n in c.items())))
        return plan

    async def _audit(self, st: _State, run_dir: Path) -> list[Finding]:
        assert st.src.inv is not None and st.src.lib is not None and st.src.gw is not None
        plan = self._plan(st, run_dir) if st.dst.inv is not None else None
        resolved = await self._users(st, run_dir)
        ctx = CheckContext(
            src_inv=st.src.inv, dst_inv=st.dst.inv, src_lib=st.src.lib, dst_lib=st.dst.lib, opts=self.opts,
            plan=plan, src_gw=st.src.gw, dst_gw=st.dst.gw, usermap=self.usermap, dst_error=st.dst.error,
            resolved_users=resolved,
        )
        findings = await run_checks(ctx, self.sink)
        self.findings = findings
        _csv(run_dir / "findings.csv", ["severity", "category", "path", "detail", "fix"],
             [[f.severity.value, f.category, f.path, f.detail, f.fix] for f in findings])
        return findings

    async def _users(self, st: _State, run_dir: Path) -> ResolvedUsers | None:
        """The "users" phase: look every distinct source author/editor up at the destination once (shared by
        the users check and the user-map proposal, which is written here)."""
        assert st.src.inv is not None
        if st.dst.gw is None or not self.opts.restore_metadata:
            why = "destination not available" if st.dst.gw is None else "metadata restore is off"
            self.sink(PhaseChanged("users", f"skipped ({why})"))
            return None
        um = self.usermap
        emails = [e for e in collect_users(st.src.inv) if um is None or not um.resolve(e)]
        try:
            resolved = await resolve_users(emails, st.dst.gw, self.sink)
        except Exception as exc:  # noqa: BLE001 - the users check falls back to its own lookups
            self.sink(LogLine("warn", f"user lookup failed: {exc}"))
            return None
        if um is None:
            try:
                entries = await build_usermap(st.src.inv, st.dst.gw, resolved)
                save_usermap(run_dir / "usermap.csv", entries)
                self.sink(LogLine("info", f"user map proposal written: {run_dir / 'usermap.csv'}"))
            except Exception as exc:  # noqa: BLE001 - proposal only
                self.sink(LogLine("warn", f"user map proposal failed: {exc}"))
        return resolved

    @staticmethod
    def _lock_rows(changes: list[LockChange]) -> list[dict[str, str]]:
        return [{"principal": c.title, "before": ", ".join(c.from_roles),
                 "after": c.to_role if c.will_change else f"unchanged ({c.skipped_reason})"} for c in changes]

    def _write_lock_csv(self, run_dir: Path, changes: list[LockChange], applied: bool) -> None:
        _csv(run_dir / "lock.csv", ["principal_id", "principal", "from_roles", "to_role", "status"],
             [[c.principal_id, c.title, ", ".join(c.from_roles), c.to_role if c.will_change else "",
               ("applied" if applied else "planned") if c.will_change else f"skipped: {c.skipped_reason}"]
              for c in changes])

    def _report(self, run_dir: Path, mode: str, summary: RunSummary, *, st: _State | None = None,
                plan: Plan | None = None, results: list[dict[str, Any]] | None = None,
                compare_result: CompareResult | None = None) -> None:
        """report.html + report.txt via m365tools.core.report when available (best effort)."""
        try:
            from m365tools.core import report as rep
        except ImportError:
            return
        self.sink(PhaseChanged("report", f"{run_dir / 'report.html'}"))
        try:
            if mode == "audit" and st is not None and st.src.inv is not None and st.src.lib is not None:
                opts_desc = {"Source": self.opts.source.describe(), "Destination": self.opts.destination.describe(),
                             "Overwrite": self.opts.overwrite, "Excludes": ", ".join(self.opts.excludes) or "-"}
                r = rep.audit_report(summary, st.src.inv, st.dst.inv, st.src.lib, st.dst.lib, summary.findings,
                                     opts_desc)
            elif mode == "dryrun" and plan is not None:
                r = rep.dryrun_report(plan, summary)
            elif mode == "migrate" and plan is not None:
                rows = [{"path": x["path"], "status": {"FAILED": "failed", "WARN": "warning"}.get(x["result"], ""),
                         "error": x["message"], "metadata_warning": x["message"] if x["result"] == "WARN" else ""}
                        for x in (results or [])]
                r = rep.migrate_report(summary, plan, rows)
            elif mode == "verify" and compare_result is not None:
                r = rep.verify_report(compare_result, summary)
            else:
                return
            rep.write_all(run_dir, r)
        except Exception as exc:  # noqa: BLE001
            self.sink(LogLine("warn", f"report not written: {exc}"))
