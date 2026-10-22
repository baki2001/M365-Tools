"""Structured explanations of one person or one group (plain dicts for CLI, TUI and JSON)."""

from __future__ import annotations

from typing import Any

from m365tools.core.errors import NotFoundError

from .models import Registry, Snapshot, norm
from .registry import bundles_with, expectations


def resolve_subject(snap: Snapshot, reg: Registry, subject: str) -> tuple[str, str]:
    """``("group", name)`` or ``("person", upn)``. Accepts a group name, a UPN, or a unique UPN local part."""
    k = norm(subject)
    g = snap.group(subject) or reg.group(subject)
    if g is not None:
        return "group", g.name
    if snap.person(k) or reg.person(k):
        return "person", k
    hits = sorted({p.upn for p in snap.people if p.upn.split("@")[0] == k}
                  | {u for u in reg.people if u.split("@")[0] == k})
    if len(hits) == 1:
        return "person", hits[0]
    raise NotFoundError(f"No person or group '{subject}'" + (f" (ambiguous: {', '.join(hits)})" if hits else ""))


def _purpose(reg: Registry, name: str) -> str:
    e = reg.group(name)
    return e.purpose if e else ""


def explain_person(snap: Snapshot, reg: Registry, upn: str) -> dict[str, Any]:
    upn = norm(upn)
    p = snap.person(upn)
    entry = reg.person(upn)
    if p is None and entry is None:
        raise NotFoundError(f"No person '{upn}' in the tenant or the registry")
    actual = snap.groups_of(upn)
    expected = expectations(reg, upn)
    a_k, e_k = {norm(g) for g in actual}, {norm(g) for g in expected}
    addon_k = {norm(a) for a in (entry.addons if entry else [])}
    via_bundle: dict[str, list[str]] = {}
    for b in entry.bundles if entry else []:
        bundle = reg.bundle(b)
        for x in bundle.groups if bundle else []:
            via_bundle.setdefault(norm(x), []).append(b)
    groups = []
    for g in sorted(actual | {x for x in expected if norm(x) not in a_k}, key=str.lower):
        why = "addon" if norm(g) in addon_k else ", ".join(via_bundle.get(norm(g), []))
        status = "ok" if norm(g) in a_k and norm(g) in e_k else ("missing" if norm(g) in e_k else "unexplained")
        groups.append({"group": g, "purpose": _purpose(reg, g), "via": why, "status": status})
    return {
        "type": "person", "upn": upn, "name": p.display_name if p else "", "job_title": p.job_title if p else "",
        "department": p.department if p else "", "in_tenant": p is not None, "registered": entry is not None,
        "bundles": list(entry.bundles) if entry else [], "addons": list(entry.addons) if entry else [],
        "note": entry.note if entry else "", "offboarded": entry.offboarded if entry else "",
        "groups": groups,
        "unexplained": [x["group"] for x in groups if x["status"] == "unexplained"],
        "missing": [x["group"] for x in groups if x["status"] == "missing"],
        "licences": list(p.licences) if p else [],
        "apps": [{"app": a, "via": via} for a, via in sorted(snap.apps_of(upn).items())],
        "roles": snap.roles_of(upn),
    }


def explain_group(snap: Snapshot, reg: Registry, name: str) -> dict[str, Any]:
    g = snap.group(name)
    e = reg.group(name)
    if g is None and e is None:
        raise NotFoundError(f"No group '{name}' in the tenant or the registry")
    gname = g.name if g else e.name  # type: ignore[union-attr]
    k = norm(gname)
    members = list(g.members) if g else []
    addon_of = sorted(p.upn for p in reg.people.values() if not p.offboarded and k in {norm(a) for a in p.addons})
    unexplained = [u for u in members if k not in {norm(x) for x in expectations(reg, u)}]
    return {
        "type": "group", "name": gname, "id": g.id if g else e.id,  # type: ignore[union-attr]
        "in_tenant": g is not None, "registered": e is not None,
        "kind": e.kind if e else "", "purpose": e.purpose if e else "", "owner": e.owner if e else "",
        "description": g.description if g else "", "tenant_owners": list(g.owners) if g else [],
        "members": members, "bundles": bundles_with(reg, gname), "addon_of": addon_of,
        "unexplained": unexplained,
        "apps": sorted(a.name for a in snap.apps if k in {norm(x) for x in a.groups}),
    }
