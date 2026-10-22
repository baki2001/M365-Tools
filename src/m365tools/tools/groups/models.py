"""Data model of the groups tool: the tenant snapshot, the registry and findings.

All plain dataclasses. Person and group keys are case-insensitive: UPNs are stored lower-case, group names keep
their tenant spelling and are matched with :func:`norm`.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime
from enum import StrEnum
from typing import Any, Literal

Kind = Literal["doc", "app", "lic", "other"]
KINDS: tuple[str, ...] = ("doc", "app", "lic", "other")
DEFAULT_PREFIX = "SG-"


def norm(name: str) -> str:
    """Comparison key for UPNs and group names."""
    return name.strip().lower()


def kind_of(name: str, prefix: str = DEFAULT_PREFIX) -> str:
    """``SG-Doc-Finance`` -> ``doc``; anything not following ``<prefix><Kind>-<Name>`` -> ``other``."""
    rest = name[len(prefix):] if name.lower().startswith(prefix.lower()) else name
    head = rest.split("-", 1)[0].lower()
    return head if head in KINDS and "-" in rest else "other"


class Severity(StrEnum):
    BLOCKING = "blocking"
    WARN = "warn"
    INFO = "info"


@dataclass(slots=True)
class Finding:
    """Same fields as spmig's Finding. ``path`` is the subject: a UPN or a group name."""

    severity: Severity
    category: str  # unexplained | missing | unregistered | not_created | unassigned | no_owner | empty | no_purpose
    path: str
    detail: str
    fix: str = ""


# ------------------------------------------------------------------------------------------------ snapshot


@dataclass(slots=True)
class Capability:
    """One column of the access matrix: a group, ``App:<name>``, ``Lic:<sku>`` or ``Role:<name>``."""

    kind: Literal["group", "app", "lic", "role"]
    name: str

    @property
    def label(self) -> str:
        return self.name if self.kind == "group" else f"{self.kind.capitalize()}:{self.name}"


@dataclass(slots=True)
class GroupInfo:
    id: str
    name: str
    description: str = ""
    security_enabled: bool = True
    mail_enabled: bool = False
    group_types: list[str] = field(default_factory=list)
    members: list[str] = field(default_factory=list)  # UPNs (direct user members)
    owners: list[str] = field(default_factory=list)  # UPNs / display names


@dataclass(slots=True)
class Person:
    id: str
    upn: str
    display_name: str = ""
    job_title: str = ""
    department: str = ""
    account_enabled: bool = True
    licences: list[str] = field(default_factory=list)  # SKU part numbers


@dataclass(slots=True)
class AppInfo:
    id: str
    name: str
    users: list[str] = field(default_factory=list)  # UPNs assigned directly
    groups: list[str] = field(default_factory=list)  # group display names assigned


@dataclass(slots=True)
class RoleInfo:
    id: str
    name: str
    members: list[str] = field(default_factory=list)  # UPNs


@dataclass(slots=True)
class Snapshot:
    prefix: str
    taken: datetime
    groups: list[GroupInfo] = field(default_factory=list)
    people: list[Person] = field(default_factory=list)
    apps: list[AppInfo] = field(default_factory=list)
    roles: list[RoleInfo] = field(default_factory=list)

    def group(self, name: str) -> GroupInfo | None:
        k = norm(name)
        return next((g for g in self.groups if norm(g.name) == k), None)

    def person(self, upn: str) -> Person | None:
        k = norm(upn)
        return next((p for p in self.people if p.upn == k), None)

    def groups_of(self, upn: str) -> set[str]:
        """Prefix-group names ``upn`` is a direct member of."""
        k = norm(upn)
        return {g.name for g in self.groups if k in g.members}

    def apps_of(self, upn: str) -> dict[str, str]:
        """App name -> how it is assigned ("direct" or the group name)."""
        k = norm(upn)
        mine = {norm(g) for g in self.groups_of(k)}
        out: dict[str, str] = {}
        for a in self.apps:
            if k in a.users:
                out[a.name] = "direct"
            else:
                via = next((g for g in a.groups if norm(g) in mine), None)
                if via:
                    out[a.name] = via
        return out

    def roles_of(self, upn: str) -> list[str]:
        k = norm(upn)
        return sorted(r.name for r in self.roles if k in r.members)

    def capabilities(self, upn: str) -> list[Capability]:
        p = self.person(upn)
        caps = [Capability("group", g) for g in sorted(self.groups_of(upn), key=str.lower)]
        caps += [Capability("app", a) for a in sorted(self.apps_of(upn))]
        caps += [Capability("lic", s) for s in sorted(p.licences if p else [])]
        caps += [Capability("role", r) for r in self.roles_of(upn)]
        return caps

    def to_json(self) -> dict[str, Any]:
        d = asdict(self)
        d["taken"] = self.taken.isoformat()
        return d

    @classmethod
    def from_json(cls, d: dict[str, Any]) -> Snapshot:
        return cls(
            prefix=d.get("prefix", DEFAULT_PREFIX), taken=datetime.fromisoformat(d["taken"]),
            groups=[GroupInfo(**g) for g in d.get("groups", [])],
            people=[Person(**p) for p in d.get("people", [])],
            apps=[AppInfo(**a) for a in d.get("apps", [])],
            roles=[RoleInfo(**r) for r in d.get("roles", [])],
        )


# ------------------------------------------------------------------------------------------------ registry


@dataclass(slots=True)
class GroupEntry:
    name: str
    id: str = ""
    kind: str = "other"
    purpose: str = ""
    owner: str = ""
    extra: dict[str, Any] = field(default_factory=dict)  # unknown keys, preserved on save


@dataclass(slots=True)
class Bundle:
    name: str
    purpose: str = ""
    groups: list[str] = field(default_factory=list)
    extra: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class PersonEntry:
    upn: str
    bundles: list[str] = field(default_factory=list)
    addons: list[str] = field(default_factory=list)
    note: str = ""
    offboarded: str = ""  # ISO date once offboarded
    extra: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class Registry:
    prefix: str = DEFAULT_PREFIX
    groups: dict[str, GroupEntry] = field(default_factory=dict)
    bundles: dict[str, Bundle] = field(default_factory=dict)
    people: dict[str, PersonEntry] = field(default_factory=dict)  # key: lower-case UPN
    extra: dict[str, Any] = field(default_factory=dict)

    def group(self, name: str) -> GroupEntry | None:
        k = norm(name)
        return next((g for g in self.groups.values() if norm(g.name) == k), None)

    def bundle(self, name: str) -> Bundle | None:
        k = norm(name)
        return next((b for b in self.bundles.values() if norm(b.name) == k), None)

    def person(self, upn: str) -> PersonEntry | None:
        return self.people.get(norm(upn))
