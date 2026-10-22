"""Registry <-> tenant: create missing groups and fill ids (sync), add unregistered groups (import).

Never removes members and never deletes groups.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from m365tools.core.events import EventSink, LogLine, null_sink

from .gateway import GraphGateway
from .models import GroupEntry, Registry, Snapshot, kind_of


@dataclass(slots=True)
class SyncPlan:
    create: list[GroupEntry] = field(default_factory=list)  # registry groups missing in the tenant
    fill_ids: list[tuple[str, str]] = field(default_factory=list)  # (name, tenant id) where the registry id differs
    created: list[tuple[str, str]] = field(default_factory=list)  # filled by apply: (name, new id)
    failed: list[tuple[str, str]] = field(default_factory=list)  # (name, error)
    applied: bool = False

    @property
    def empty(self) -> bool:
        return not self.create and not self.fill_ids


def plan_sync(snap: Snapshot, reg: Registry) -> SyncPlan:
    plan = SyncPlan()
    for e in reg.groups.values():
        g = snap.group(e.name)
        if g is None:
            plan.create.append(e)
        elif e.id != g.id:
            plan.fill_ids.append((e.name, g.id))
    return plan


async def apply_sync(gw: GraphGateway, reg: Registry, plan: SyncPlan, sink: EventSink = null_sink) -> SyncPlan:
    """Create the planned groups (security, mail-disabled, description = purpose) and record ids in ``reg``
    (in memory; the caller saves). A failed create is recorded and the rest continue."""
    for name, gid in plan.fill_ids:
        e = reg.group(name)
        if e is not None:
            e.id = gid
    for e in plan.create:
        try:
            gid = await gw.create_group(e.name, e.purpose)
        except Exception as exc:  # noqa: BLE001 - reported per group, the others still run
            plan.failed.append((e.name, str(exc)))
            sink(LogLine("error", f"create {e.name}: {exc}"))
            continue
        e.id = gid
        plan.created.append((e.name, gid))
        sink(LogLine("ok", f"created {e.name} ({gid})"))
    plan.applied = True
    return plan


def plan_import(snap: Snapshot, reg: Registry) -> list[GroupEntry]:
    """Registry entries for prefix groups that exist in the tenant but not in the registry (purpose left empty)."""
    return [GroupEntry(name=g.name, id=g.id, kind=kind_of(g.name, snap.prefix), purpose="")
            for g in snap.groups if reg.group(g.name) is None]


def apply_import(reg: Registry, entries: list[GroupEntry]) -> list[GroupEntry]:
    added = []
    for e in entries:
        if reg.group(e.name) is None:
            reg.groups[e.name] = e
            added.append(e)
    return added
