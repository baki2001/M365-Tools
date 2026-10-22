"""Compare the registry expectation with a snapshot and report findings."""

from __future__ import annotations

from .models import Finding, Registry, Severity, Snapshot, norm
from .registry import expectations

W, I = Severity.WARN, Severity.INFO


def _names(xs: set[str]) -> list[str]:
    return sorted(xs, key=str.lower)


def person_findings(snap: Snapshot, reg: Registry, upn: str) -> list[Finding]:
    """Drift for one person (used by ``onboard check`` and as part of :func:`findings`)."""
    upn = norm(upn)
    actual = {norm(g): g for g in snap.groups_of(upn)}
    entry = reg.person(upn)
    if entry is None:
        have = ", ".join(_names(set(actual.values()))) or "no prefix groups"
        return [Finding(I, "unassigned", upn, f"not in [people]; has {have}",
                        "m365 groups suggest, or m365 onboard " + upn + " --bundle NAME --apply")]
    expected = {norm(g): g for g in expectations(reg, upn)}
    out = []
    for k in sorted(set(actual) - set(expected)):
        why = "offboarded on " + entry.offboarded if entry.offboarded else "not in any of its bundles or addons"
        fix = (f"m365 offboard {upn} --apply" if entry.offboarded
               else f"add {actual[k]} to addons of {upn} in the registry, or remove the membership")
        out.append(Finding(W, "unexplained", upn, f"member of {actual[k]}: {why}", fix))
    for k in sorted(set(expected) - set(actual)):
        out.append(Finding(W, "missing", upn, f"expected in {expected[k]} but is not a member",
                           f"m365 onboard {upn} --apply"))
    return out


def findings(snap: Snapshot, reg: Registry, registry_path: str = "the registry") -> list[Finding]:
    out: list[Finding] = []
    for p in snap.people:
        out += person_findings(snap, reg, p.upn)
    in_tenant = {norm(g.name) for g in snap.groups}
    for g in snap.groups:
        if reg.group(g.name) is None:
            out.append(Finding(W, "unregistered", g.name, "prefix group exists in the tenant but not in the registry",
                               "m365 groups import --apply"))
        if not g.owners:
            out.append(Finding(I, "no_owner", g.name, "group has no owner", "assign an owner in Entra"))
        if not g.members:
            out.append(Finding(I, "empty", g.name, "group has no direct user members", ""))
    for e in reg.groups.values():
        if norm(e.name) not in in_tenant:
            out.append(Finding(W, "not_created", e.name, "registry group does not exist in the tenant",
                               "m365 groups sync --apply"))
        if not e.purpose.strip():
            out.append(Finding(I, "no_purpose", e.name, "registry group has no purpose",
                               f"fill in purpose in {registry_path}"))
    return out
