"""Build a :class:`Snapshot` from a :class:`GraphGateway`, emitting the inventory phases."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, TypeVar

from m365tools.core.errors import AccessDeniedError
from m365tools.core.events import EventSink, LogLine, PhaseChanged, Stats, null_sink
from m365tools.core.report.csv_ import write_csv

from .gateway import GraphGateway
from .models import AppInfo, GroupInfo, Person, RoleInfo, Snapshot, norm

T = TypeVar("T")

INVENTORY_PHASES: tuple[str, ...] = ("groups", "users", "licences", "apps", "roles", "report")


def _stats(sink: EventSink, done: int, total: int) -> None:
    sink(Stats(done=done, total=total, bytes_done=0, bytes_total=0, mbps=0.0, eta_s=None, failed=0, warnings=0))


async def _optional(sink: EventSink, what: str, coro: Any, fallback: T) -> T:
    """Licences, apps and roles are nice to have: when the signed-in account lacks the permission, warn and
    carry on with the groups and users instead of failing the whole inventory."""
    try:
        return await coro
    except AccessDeniedError as exc:
        sink(LogLine("warn", f"{what} skipped: access denied ({exc}). Grant the Graph permissions listed by "
                             "'m365 tenant setup-help' and sign in again (ctrl+l) to include them."))
        return fallback


async def build_snapshot(gw: GraphGateway, prefix: str, sink: EventSink = null_sink) -> Snapshot:
    """Read groups (with direct user members and owners), enabled member users, licences, enterprise apps that
    require assignment, and directory role members. Phases ``groups -> users -> licences -> apps -> roles``."""
    snap = Snapshot(prefix=prefix, taken=datetime.now())

    sink(PhaseChanged("groups", f"{prefix}*"))
    raw_groups = sorted(await gw.list_groups(prefix), key=lambda g: g.name.lower())
    # startswith() is case-insensitive in Graph; keep that behaviour but drop anything that slipped through.
    raw_groups = [g for g in raw_groups if g.name.lower().startswith(prefix.lower())]
    for i, g in enumerate(raw_groups, 1):
        members = await gw.group_members(g.id)
        owners = await gw.group_owners(g.id)
        snap.groups.append(GroupInfo(
            id=g.id, name=g.name, description=g.description, security_enabled=g.security_enabled,
            mail_enabled=g.mail_enabled, group_types=list(g.group_types),
            members=sorted(norm(m.name) for m in members if m.type == "user"),
            owners=sorted(norm(o.name) if o.type == "user" else o.name for o in owners),
        ))
        _stats(sink, i, len(raw_groups))
    sink(LogLine("info", f"{len(snap.groups)} group(s) starting with {prefix}"))

    sink(PhaseChanged("users"))
    users = [u for u in await gw.list_users() if u.account_enabled and u.user_type.lower() == "member" and u.upn]
    _stats(sink, len(users), len(users))
    sink(LogLine("info", f"{len(users)} enabled member user(s)"))

    sink(PhaseChanged("licences"))
    skus = await _optional(sink, "licences", gw.subscribed_skus(), {})
    for u in sorted(users, key=lambda u: u.upn.lower()):
        snap.people.append(Person(
            id=u.id, upn=norm(u.upn), display_name=u.display_name, job_title=u.job_title,
            department=u.department, account_enabled=u.account_enabled,
            licences=sorted(skus.get(s, s) for s in u.sku_ids),
        ))
    if skus:
        sink(LogLine("info", f"{len(skus)} subscribed SKU(s)"))

    sink(PhaseChanged("apps"))
    upn_by_id = {u.id: norm(u.upn) for u in users}
    apps = [a for a in await _optional(sink, "apps", gw.list_apps(), []) if a.assignment_required]
    for i, a in enumerate(sorted(apps, key=lambda a: a.name.lower()), 1):
        info = AppInfo(id=a.id, name=a.name)
        for asg in await gw.app_assignments(a.id):
            if asg.principal_type.lower() == "user":
                info.users.append(upn_by_id.get(asg.principal_id, norm(asg.principal_name)))
            elif asg.principal_type.lower() == "group":
                info.groups.append(asg.principal_name)
        info.users.sort()
        info.groups.sort(key=str.lower)
        snap.apps.append(info)
        _stats(sink, i, len(apps))
    sink(LogLine("info", f"{len(snap.apps)} enterprise app(s) with assignment required"))

    sink(PhaseChanged("roles"))
    roles = await _optional(sink, "roles", gw.list_roles(), [])
    for i, r in enumerate(sorted(roles, key=lambda r: r.name.lower()), 1):
        members = await gw.role_members(r.id)
        snap.roles.append(RoleInfo(id=r.id, name=r.name,
                                   members=sorted(norm(m.name) for m in members if m.type == "user")))
        _stats(sink, i, len(roles))
    return snap


def matrix_rows(snap: Snapshot) -> tuple[list[str], list[list[str]]]:
    """People x capabilities: columns ``UPN, Name, Job title, Department, <capability labels...>``."""
    per_person = {p.upn: {c.label for c in snap.capabilities(p.upn)} for p in snap.people}
    groups = sorted({lab for caps in per_person.values() for lab in caps if ":" not in lab}
                    | {g.name for g in snap.groups}, key=str.lower)
    others = sorted({lab for caps in per_person.values() for lab in caps if ":" in lab}, key=str.lower)
    cols = groups + others
    rows = [[p.upn, p.display_name, p.job_title, p.department,
             *("x" if c in per_person[p.upn] else "" for c in cols)] for p in snap.people]
    return ["UPN", "Name", "Job title", "Department", *cols], rows


def write_matrix_csv(path: Path, snap: Snapshot) -> None:
    cols, rows = matrix_rows(snap)
    write_csv(path, cols, rows)
