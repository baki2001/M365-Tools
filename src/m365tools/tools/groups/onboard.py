"""Onboarding (add a person directly to every group of their bundles and add-ons) and offboarding."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import date

from m365tools.core.errors import ValidationError
from m365tools.core.events import EventSink, LogLine, null_sink

from .gateway import GraphGateway
from .models import PersonEntry, Registry, Snapshot, norm
from .registry import expectations


@dataclass(slots=True)
class GroupChange:
    group: str
    group_id: str  # "" when the group does not exist in the tenant yet
    status: str = "planned"  # planned | already | not_created | done | failed | kept
    error: str = ""


@dataclass(slots=True)
class OnboardPlan:
    upn: str
    user_id: str
    source: str  # "bundle", "like <upn>", "registry"
    bundles: list[str] = field(default_factory=list)
    addons: list[str] = field(default_factory=list)
    changes: list[GroupChange] = field(default_factory=list)
    manual: list[str] = field(default_factory=list)  # licences/apps/roles the tool cannot set
    applied: bool = False

    @property
    def to_add(self) -> list[GroupChange]:
        return [c for c in self.changes if c.status in ("planned", "done", "failed")]


@dataclass(slots=True)
class OffboardPlan:
    upn: str
    user_id: str
    date: str
    changes: list[GroupChange] = field(default_factory=list)  # registry groups: planned/done/failed
    kept: list[str] = field(default_factory=list)  # prefix groups not in the registry (left alone)
    applied: bool = False


def _user(snap: Snapshot, upn: str) -> str:
    p = snap.person(upn)
    if p is None:
        raise ValidationError(f"{upn}: no enabled member user with this UPN in the tenant")
    return p.id


def _manual(snap: Snapshot, template: list[str], groups: set[str]) -> list[str]:
    """Licences, directly assigned apps and roles that ``template`` people all hold and group membership in
    ``groups`` does not give (the tool does not set these in v1)."""
    if not template:
        return []
    gk = {norm(g) for g in groups}
    has_lic_group = any(snap.group(g) is not None and g.lower().startswith((snap.prefix + "lic-").lower())
                        for g in groups)
    sets = []
    for upn in template:
        p = snap.person(upn)
        caps = {f"Role:{r}" for r in snap.roles_of(upn)}
        caps |= {f"App:{a}" for a, via in snap.apps_of(upn).items() if via == "direct" or norm(via) not in gk}
        if not has_lic_group and p:
            caps |= {f"Lic:{s}" for s in p.licences}
        sets.append(caps)
    return sorted(set.intersection(*sets)) if sets else []


def plan_onboard(snap: Snapshot, reg: Registry, upn: str, *, bundles: Sequence[str] = (),
                 addons: Sequence[str] = (), like: str | None = None) -> OnboardPlan:
    """Plan for ``upn``. With neither ``bundles`` nor ``like`` the person's existing registry entry is used (this
    re-adds missing groups). ``like`` copies that person's registry entry, or their actual groups if unregistered."""
    upn = norm(upn)
    uid = _user(snap, upn)
    template: list[str] = []
    bnames: list[str] = []
    extra: list[str] = []
    if like:
        lk = norm(like)
        le = reg.person(lk)
        if le is not None and not le.offboarded:
            bnames, extra = list(le.bundles), list(le.addons)
        elif snap.person(lk) is not None:
            extra = sorted(snap.groups_of(lk), key=str.lower)
        else:
            raise ValidationError(f"--like {like}: not in the registry or the tenant")
        template, source = [lk], f"like {lk}"
    elif bundles:
        source = "bundle"
    else:
        entry = reg.person(upn)
        if entry is None:
            raise ValidationError(f"{upn} is not in the registry: give --bundle NAME or --like UPN")
        bnames, extra, source = list(entry.bundles), list(entry.addons), "registry"
    for b in bundles:
        bundle = reg.bundle(b)
        if bundle is None:
            raise ValidationError(f"Unknown bundle '{b}' (known: {', '.join(reg.bundles) or 'none'})")
        if bundle.name not in bnames:
            bnames.append(bundle.name)
        if not like:
            template += [p.upn for p in reg.people.values()
                         if not p.offboarded and bundle.name in p.bundles and snap.person(p.upn)]
    for a in addons:
        g = reg.group(a) or snap.group(a)
        if g is None:
            raise ValidationError(f"Unknown group '{a}' (not in the registry or the tenant)")
        if g.name not in extra:
            extra.append(g.name)
    plan = OnboardPlan(upn=upn, user_id=uid, source=source, bundles=bnames, addons=extra)
    wanted: dict[str, str] = {}
    for b in bnames:
        bundle = reg.bundle(b)
        for g in bundle.groups if bundle else []:
            wanted.setdefault(norm(g), g)
    for g in extra:
        wanted.setdefault(norm(g), g)
    have = {norm(g) for g in snap.groups_of(upn)}
    for k, g in sorted(wanted.items()):
        tg = snap.group(g)
        status = "not_created" if tg is None else ("already" if k in have else "planned")
        plan.changes.append(GroupChange(tg.name if tg else g, tg.id if tg else "", status))
    plan.manual = _manual(snap, template, set(wanted.values()))
    return plan


async def apply_onboard(gw: GraphGateway, reg: Registry, plan: OnboardPlan,
                        sink: EventSink = null_sink) -> OnboardPlan:
    """Add the memberships and write the ``[people]`` entry into ``reg`` (in memory; the caller saves)."""
    for c in plan.changes:
        if c.status != "planned":
            continue
        try:
            await gw.add_member(c.group_id, plan.user_id)
            c.status = "done"
            sink(LogLine("ok", f"added {plan.upn} to {c.group}"))
        except Exception as exc:  # noqa: BLE001 - per group, the rest continues
            c.status, c.error = "failed", str(exc)
            sink(LogLine("error", f"add {plan.upn} to {c.group}: {exc}"))
    entry = reg.people.setdefault(plan.upn, PersonEntry(upn=plan.upn))
    for b in plan.bundles:
        if b not in entry.bundles:
            entry.bundles.append(b)
    covered = {norm(g) for g in expectations(reg, plan.upn)} if not entry.offboarded else set()
    for a in plan.addons:
        if norm(a) not in covered and a not in entry.addons:
            entry.addons.append(a)
    entry.offboarded = ""
    plan.applied = True
    return plan


def plan_offboard(snap: Snapshot, reg: Registry, upn: str, today: date | None = None) -> OffboardPlan:
    upn = norm(upn)
    uid = _user(snap, upn)
    plan = OffboardPlan(upn=upn, user_id=uid, date=(today or date.today()).isoformat())
    for name in sorted(snap.groups_of(upn), key=str.lower):
        g = snap.group(name)
        if reg.group(name) is not None and g is not None:
            plan.changes.append(GroupChange(g.name, g.id))
        else:
            plan.kept.append(name)
    return plan


async def apply_offboard(gw: GraphGateway, reg: Registry, plan: OffboardPlan,
                         sink: EventSink = null_sink) -> OffboardPlan:
    """Remove the memberships and mark the ``[people]`` entry offboarded (the account is not disabled)."""
    for c in plan.changes:
        try:
            await gw.remove_member(c.group_id, plan.user_id)
            c.status = "done"
            sink(LogLine("ok", f"removed {plan.upn} from {c.group}"))
        except Exception as exc:  # noqa: BLE001
            c.status, c.error = "failed", str(exc)
            sink(LogLine("error", f"remove {plan.upn} from {c.group}: {exc}"))
    entry = reg.people.setdefault(plan.upn, PersonEntry(upn=plan.upn))
    entry.offboarded = plan.date
    plan.applied = True
    return plan
