"""GroupsJob: the only groups entry point the CLI/TUI call.

    job = GroupsJob(gateway, registry_path, sink, run_dir_factory, tenant="contoso")
    await job.inventory()                         -> Snapshot      (run folder: snapshot.json, matrix.csv, report)
    await job.suggest()                           -> Suggestion
    job.accept(suggestion, {"1": "Servicedesk"}, apply=False) -> list[str]
    await job.drift()                             -> list[Finding]
    await job.explain("evelien@contoso.com")      -> dict
    await job.sync(apply=False)                   -> SyncPlan
    await job.import_(apply=False)                -> list[GroupEntry]
    await job.onboard(upn, bundles=(), addons=(), like=None, apply=False) -> OnboardPlan
    await job.onboard_check(upn)                  -> list[Finding]
    await job.offboard(upn, apply=False)          -> OffboardPlan
    await job.aclose()

Every operation reads a fresh snapshot the first time it needs one and caches it in ``job.snapshot``
(``refresh()`` drops it); after a write the cache is dropped. Writes happen only with ``apply=True``.
"""

from __future__ import annotations

import contextlib
from collections.abc import Callable, Sequence
from datetime import datetime
from pathlib import Path
from typing import Any

from m365tools.core.events import EventSink, LogLine, PhaseChanged, PhaseClock, null_sink
from m365tools.core.report import groups as greport
from m365tools.core.report import write_all, write_csv, write_findings_csv, write_json

from . import drift as drift_mod
from .explain import explain_group, explain_person, resolve_subject
from .gateway import GraphGateway
from .inventory import build_snapshot, write_matrix_csv
from .models import Finding, GroupEntry, Registry, Snapshot
from .onboard import (
    OffboardPlan,
    OnboardPlan,
    apply_offboard,
    apply_onboard,
    plan_offboard,
    plan_onboard,
)
from .registry import load_registry, save_registry
from .suggest import Suggestion
from .suggest import accept as accept_suggestion
from .suggest import suggest as build_suggestion
from .sync import SyncPlan, apply_import, apply_sync, plan_import, plan_sync

RunDirFactory = Callable[[str], Path]
GROUPS_MODES: tuple[str, ...] = ("inventory", "suggest", "drift", "sync", "import", "onboard", "offboard")


def runs_name(tenant: str) -> str:
    """Folder under ``runs/`` holding a tenant's groups runs."""
    return f"groups-{tenant}"


class GroupsJob:
    def __init__(self, gateway: GraphGateway, registry_path: Path, sink: EventSink = null_sink,
                 run_dir_factory: RunDirFactory | None = None, *, tenant: str = "default") -> None:
        self.gateway = gateway
        self.registry_path = Path(registry_path)
        self.clock = PhaseClock(sink)
        self.sink: EventSink = self.clock
        self.tenant = tenant
        self._run_dir_factory = run_dir_factory
        self.snapshot: Snapshot | None = None
        self.registry: Registry = load_registry(self.registry_path)
        self.run_dir: Path | None = None
        self.findings: list[Finding] = []

    # ------------------------------------------------------------------------------------------ plumbing

    def _new_run_dir(self, mode: str) -> Path:
        if self._run_dir_factory is not None:
            d = Path(self._run_dir_factory(mode))
        else:
            from m365tools.core.paths import new_run_dir

            d = new_run_dir(runs_name(self.tenant), mode)
        d.mkdir(parents=True, exist_ok=True)
        with contextlib.suppress(OSError):
            (d / "tmp").rmdir()  # groups runs never need a tmp folder
        self.run_dir = d
        return d

    def reload_registry(self) -> Registry:
        self.registry = load_registry(self.registry_path)
        return self.registry

    def save_registry(self) -> Path:
        path = save_registry(self.registry, self.registry_path)
        self.registry = load_registry(self.registry_path)
        return path

    def refresh(self) -> None:
        self.snapshot = None

    async def _snap(self) -> Snapshot:
        if self.snapshot is None:
            self.clock.reset()
            self.snapshot = await build_snapshot(self.gateway, self.registry.prefix, self.sink)
        return self.snapshot

    def _report(self, run_dir: Path, report: Any) -> None:
        self.sink(PhaseChanged("report"))
        write_all(run_dir, report)
        self.sink(LogLine("debug", f"report written: {run_dir / 'report.html'}"))

    def request_stop(self) -> None:  # nothing long-running to stop; kept for the CLI's Ctrl+C handler
        return None

    async def aclose(self) -> None:
        close = getattr(self.gateway, "aclose", None)
        if close is not None:
            await close()

    # ------------------------------------------------------------------------------------------ read-only

    async def inventory(self) -> Snapshot:
        started = datetime.now()
        self.refresh()
        snap = await self._snap()
        d = self._new_run_dir("inventory")
        write_json(d / "snapshot.json", snap.to_json())
        write_matrix_csv(d / "matrix.csv", snap)
        self.findings = drift_mod.findings(snap, self.registry, str(self.registry_path))
        write_findings_csv(d / "findings.csv", self.findings)
        self._report(d, greport.inventory_report(snap, self.findings, self.tenant, str(self.registry_path),
                                                 started))
        return snap

    async def suggest(self) -> Suggestion:
        started = datetime.now()
        snap = await self._snap()
        sug = build_suggestion(snap, self.registry)
        d = self._new_run_dir("suggest")
        write_csv(d / "candidates.csv", ["Key", "Name", "Existing", "Groups", "People"],
                  [[c.key, c.name, c.existing, ", ".join(c.groups), ", ".join(c.people)] for c in sug.candidates])
        self._report(d, greport.suggest_report(sug, self.tenant, str(self.registry_path), started))
        return sug

    def accept(self, sug: Suggestion, accepted: dict[str, str], *, apply: bool = False,
               include_near: bool = True) -> list[str]:
        """Changes accepting candidates would make to the registry; written only with ``apply``."""
        if not apply:
            scratch = load_registry(self.registry_path)
            return accept_suggestion(scratch, sug, accepted, include_near=include_near)
        changes = accept_suggestion(self.registry, sug, accepted, include_near=include_near)
        self.save_registry()
        return changes

    async def drift(self) -> list[Finding]:
        started = datetime.now()
        snap = await self._snap()
        self.findings = drift_mod.findings(snap, self.registry, str(self.registry_path))
        d = self._new_run_dir("drift")
        write_findings_csv(d / "findings.csv", self.findings)
        self._report(d, greport.drift_report(self.findings, self.tenant, str(self.registry_path), started))
        return self.findings

    async def explain(self, subject: str) -> dict[str, Any]:
        snap = await self._snap()
        kind, key = resolve_subject(snap, self.registry, subject)
        if kind == "group":
            return explain_group(snap, self.registry, key)
        return explain_person(snap, self.registry, key)

    async def onboard_check(self, upn: str) -> list[Finding]:
        snap = await self._snap()
        return drift_mod.person_findings(snap, self.registry, upn)

    # ------------------------------------------------------------------------------------------ writes

    async def sync(self, apply: bool = False) -> SyncPlan:
        started = datetime.now()
        snap = await self._snap()
        plan = plan_sync(snap, self.registry)
        if apply and not plan.empty:
            await apply_sync(self.gateway, self.registry, plan, self.sink)
            self.save_registry()
            self.refresh()
        plan.applied = apply
        d = self._new_run_dir("sync")
        rows = [["create group", e.name, e.purpose, _status(apply, (e.name in dict(plan.created)),
                                                            dict(plan.failed).get(e.name))] for e in plan.create]
        rows += [["fill id", n, gid, "done" if apply else "planned"] for n, gid in plan.fill_ids]
        write_csv(d / "sync.csv", ["Action", "Subject", "Detail", "Status"], rows)
        self._report(d, greport.changes_report("Registry to tenant sync", "sync", rows, apply, self.tenant,
                                               str(self.registry_path), started))
        return plan

    async def import_(self, apply: bool = False) -> list[GroupEntry]:
        started = datetime.now()
        snap = await self._snap()
        entries = plan_import(snap, self.registry)
        if apply and entries:
            apply_import(self.registry, entries)
            self.save_registry()
        d = self._new_run_dir("import")
        rows = [["add to registry", e.name, f"kind {e.kind}, id {e.id}", "done" if apply else "planned"]
                for e in entries]
        write_csv(d / "import.csv", ["Action", "Subject", "Detail", "Status"], rows)
        self._report(d, greport.changes_report("Tenant to registry import", "import", rows, apply, self.tenant,
                                               str(self.registry_path), started))
        return entries

    async def onboard(self, upn: str, *, bundles: Sequence[str] = (), addons: Sequence[str] = (),
                      like: str | None = None, apply: bool = False) -> OnboardPlan:
        started = datetime.now()
        snap = await self._snap()
        plan = plan_onboard(snap, self.registry, upn, bundles=bundles, addons=addons, like=like)
        if apply:
            await apply_onboard(self.gateway, self.registry, plan, self.sink)
            self.save_registry()
            self.refresh()
        d = self._new_run_dir("onboard")
        rows = [["add member", c.group, c.error or plan.upn, c.status] for c in plan.changes]
        rows += [["manual", m, "not set by this tool", "manual"] for m in plan.manual]
        write_csv(d / "onboard.csv", ["Action", "Subject", "Detail", "Status"], rows)
        self._report(d, greport.changes_report(f"Onboard {plan.upn}", "onboard", rows, apply, self.tenant,
                                               str(self.registry_path), started))
        return plan

    async def offboard(self, upn: str, *, apply: bool = False) -> OffboardPlan:
        started = datetime.now()
        snap = await self._snap()
        plan = plan_offboard(snap, self.registry, upn)
        if apply:
            await apply_offboard(self.gateway, self.registry, plan, self.sink)
            self.save_registry()
            self.refresh()
        d = self._new_run_dir("offboard")
        rows = [["remove member", c.group, c.error or plan.upn, c.status] for c in plan.changes]
        rows += [["keep", g, "not in the registry", "kept"] for g in plan.kept]
        write_csv(d / "offboard.csv", ["Action", "Subject", "Detail", "Status"], rows)
        self._report(d, greport.changes_report(f"Offboard {plan.upn}", "offboard", rows, apply, self.tenant,
                                               str(self.registry_path), started))
        return plan


def _status(apply: bool, done: bool, error: str | None) -> str:
    if not apply:
        return "planned"
    return "done" if done else f"failed: {error}" if error else "failed"
