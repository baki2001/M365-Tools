"""In-memory GraphGateway for the groups tests, with a small realistic tenant (contoso.com).

12 enabled member users (plus a guest and a disabled account that inventory must skip), 8 ``SG-`` groups (plus a
non-prefix group), 2 enterprise apps that require assignment (plus one that does not), 2 licence SKUs and one
admin role. Every write is recorded in ``calls``.
"""

from __future__ import annotations

import itertools
from dataclasses import dataclass, field
from pathlib import Path

from m365tools.core.errors import M365Error
from m365tools.tools.groups.gateway import DirApp, DirAssignment, DirGroup, DirRef, DirUser
from m365tools.tools.groups.registry import registry_from_dict, save_registry

D = "@contoso.com"
E3, E5 = "sku-e3", "sku-e5"
SKUS = {E3: "ENTERPRISEPACK", E5: "SPE_E5"}

# upn local part -> (display name, job title, department, sku ids)
PEOPLE = {
    "anna": ("Anna de Vries", "Service Desk", "Service", [E3]),
    "bram": ("Bram Jansen", "Service Desk", "Service", [E3]),
    "chris": ("Chris Bakker", "Service Desk", "Service", [E3]),
    "dana": ("Dana Visser", "Accountant", "Finance", [E3]),
    "erik": ("Erik Smit", "Accountant", "Finance", [E3]),
    "evelien": ("Evelien Mulder", "Service Desk", "Service", [E3]),
    "frank": ("Frank de Boer", "Service Desk", "Service", [E3]),
    "gina": ("Gina Bos", "Sales Manager", "Sales", [E3]),
    "hugo": ("Hugo Peters", "Sales Manager", "Sales", [E3]),
    "ivo": ("Ivo Hendriks", "IT Admin", "IT", [E5]),
    "jan": ("Jan Dekker", "Intern", "", []),
    "kim": ("Kim Vos", "Accountant", "Finance", [E3]),
}

SD = ["SG-Doc-Service", "SG-Doc-Hub", "SG-App-Jira", "SG-Lic-E3"]
FIN = ["SG-Doc-Finance", "SG-Doc-Hub", "SG-Lic-E3"]
SALES = ["SG-Doc-Sales", "SG-Doc-Hub", "SG-Lic-E3"]
MEMBERSHIPS = {
    "anna": SD, "bram": SD, "chris": SD, "evelien": [*SD, "SG-Doc-Finance"],
    "frank": ["SG-Doc-Service", "SG-Doc-Hub", "SG-Lic-E3"],
    "dana": FIN, "erik": FIN, "gina": SALES, "hugo": SALES,
    "ivo": ["SG-Doc-Hub", "SG-App-Jira", "SG-Lic-E5"], "jan": [], "kim": ["SG-Doc-Finance"],
}
GROUPS = ["SG-Doc-Service", "SG-Doc-Hub", "SG-Doc-Finance", "SG-Doc-Sales", "SG-App-Jira", "SG-Lic-E3",
          "SG-Lic-E5", "SG-Doc-Archive"]


def uid(local: str) -> str:
    return f"u-{local}"


def gid(name: str) -> str:
    return "g-" + name.lower()


@dataclass
class FakeGroup:
    info: DirGroup
    members: set[str] = field(default_factory=set)  # user ids
    owners: set[str] = field(default_factory=set)


class FakeGraphGateway:
    def __init__(self) -> None:
        self.users: dict[str, DirUser] = {}
        self.groups: dict[str, FakeGroup] = {}
        self.apps: list[tuple[DirApp, list[DirAssignment]]] = []
        self.roles: list[tuple[DirRef, list[str]]] = []
        self.skus = dict(SKUS)
        self.calls: list[tuple] = []
        self.fail_on: set[str] = set()  # group names whose add/remove/create raises
        self.closed = False
        self._ids = itertools.count(1)

    # -- helpers -----------------------------------------------------------------------------------------

    def add_user(self, local: str, name: str, title: str = "", dept: str = "", skus: list[str] | None = None,
                 *, enabled: bool = True, user_type: str = "Member") -> DirUser:
        u = DirUser(uid(local), local + D, name, title, dept, enabled, user_type, list(skus or []))
        self.users[u.id] = u
        return u

    def add_group(self, name: str, owners: list[str] = (), description: str = "") -> FakeGroup:  # type: ignore[assignment]
        g = FakeGroup(DirGroup(gid(name), name, description), owners={uid(o) for o in owners})
        self.groups[g.info.id] = g
        return g

    def by_name(self, name: str) -> FakeGroup:
        return next(g for g in self.groups.values() if g.info.name.lower() == name.lower())

    def members_of(self, name: str) -> set[str]:
        return {self.users[m].upn for m in self.by_name(name).members}

    # -- GraphGateway ------------------------------------------------------------------------------------

    async def list_groups(self, prefix: str) -> list[DirGroup]:
        return [g.info for g in self.groups.values() if g.info.name.lower().startswith(prefix.lower())]

    def _ref(self, i: str) -> DirRef:
        u = self.users.get(i)
        return DirRef(i, u.upn, "user") if u else DirRef(i, i, "servicePrincipal")

    async def group_members(self, group_id: str) -> list[DirRef]:
        return [self._ref(m) for m in sorted(self.groups[group_id].members)]

    async def group_owners(self, group_id: str) -> list[DirRef]:
        return [self._ref(m) for m in sorted(self.groups[group_id].owners)]

    async def list_users(self) -> list[DirUser]:
        return list(self.users.values())

    async def subscribed_skus(self) -> dict[str, str]:
        return dict(self.skus)

    async def list_apps(self) -> list[DirApp]:
        return [a for a, _ in self.apps]

    async def app_assignments(self, app_id: str) -> list[DirAssignment]:
        return next(asg for a, asg in self.apps if a.id == app_id)

    async def list_roles(self) -> list[DirRef]:
        return [r for r, _ in self.roles]

    async def role_members(self, role_id: str) -> list[DirRef]:
        return [self._ref(m) for r, ms in self.roles if r.id == role_id for m in ms]

    async def create_group(self, name: str, description: str) -> str:
        self.calls.append(("create_group", name, description))
        if name in self.fail_on:
            raise M365Error(f"cannot create {name}")
        g = self.add_group(name, description=description)
        return g.info.id

    async def add_member(self, group_id: str, user_id: str) -> None:
        self.calls.append(("add_member", self.groups[group_id].info.name, self.users[user_id].upn))
        if self.groups[group_id].info.name in self.fail_on:
            raise M365Error("add failed")
        self.groups[group_id].members.add(user_id)

    async def remove_member(self, group_id: str, user_id: str) -> None:
        self.calls.append(("remove_member", self.groups[group_id].info.name, self.users[user_id].upn))
        if self.groups[group_id].info.name in self.fail_on:
            raise M365Error("remove failed")
        self.groups[group_id].members.discard(user_id)

    async def aclose(self) -> None:
        self.closed = True

    def writes(self) -> list[tuple]:
        return list(self.calls)


def make_fake() -> FakeGraphGateway:
    gw = FakeGraphGateway()
    for local, (name, title, dept, skus) in PEOPLE.items():
        gw.add_user(local, name, title, dept, skus)
    gw.add_user("guest_ext", "Guest User", user_type="Guest")
    gw.add_user("left", "Left Company", "Service Desk", enabled=False)
    for name in GROUPS:
        gw.add_group(name, owners=[] if name == "SG-Doc-Archive" else ["ivo"], description=f"{name} access")
    gw.add_group("All Staff", owners=["ivo"])
    for local, gs in MEMBERSHIPS.items():
        for g in gs:
            gw.by_name(g).members.add(uid(local))
    gw.by_name("SG-Doc-Service").members.add(uid("left"))  # a disabled account still in a group
    gw.apps = [
        (DirApp("sp-jira", "Jira Cloud", True), [DirAssignment("Group", gid("SG-App-Jira"), "SG-App-Jira"),
                                                 DirAssignment("User", uid("dana"), "Dana Visser")]),
        (DirApp("sp-payroll", "Payroll", True), [DirAssignment("User", uid("dana"), "Dana Visser"),
                                                 DirAssignment("User", uid("erik"), "Erik Smit")]),
        (DirApp("sp-explorer", "Graph Explorer", False), []),
    ]
    gw.roles = [(DirRef("r-ga", "Global Administrator", "role"), [uid("ivo")])]
    return gw


REGISTRY = {
    "prefix": "SG-",
    "custom_top": {"owner_team": "IT"},
    "groups": {
        "SG-Doc-Service": {"kind": "doc", "purpose": "Service site", "owner": "ivo@contoso.com", "x_ticket": 12},
        "SG-Doc-Hub": {"kind": "doc", "purpose": "Intranet hub"},
        "SG-Doc-Finance": {"kind": "doc", "purpose": "Finance site"},
        "SG-Doc-Sales": {"kind": "doc", "purpose": "Sales site"},
        "SG-App-Jira": {"kind": "app", "purpose": "Jira Cloud"},
        "SG-Lic-E3": {"kind": "lic", "purpose": "Microsoft 365 E3"},
        "SG-Lic-E5": {"kind": "lic", "purpose": ""},
        "SG-App-Payroll": {"kind": "app", "purpose": "Payroll app"},
    },
    "bundles": {
        "Servicedesk": {"purpose": "Everyone on the service desk", "groups": SD, "review": "yearly"},
        "Finance": {"purpose": "Finance team", "groups": FIN},
    },
    "people": {
        "anna@contoso.com": {"bundles": ["Servicedesk"]},
        "bram@contoso.com": {"bundles": ["Servicedesk"]},
        "chris@contoso.com": {"bundles": ["Servicedesk"]},
        "Evelien@contoso.com": {"bundles": ["Servicedesk"], "addons": ["SG-Doc-Finance"],
                                "note": "Also does invoicing", "start": "2025-01-06"},
        "frank@contoso.com": {"bundles": ["Servicedesk"]},
        "dana@contoso.com": {"bundles": ["Finance"]},
        "erik@contoso.com": {"bundles": ["Finance"]},
        "gina@contoso.com": {"bundles": [], "addons": ["SG-Doc-Sales"]},
    },
}


def write_registry(path: Path) -> Path:
    return save_registry(registry_from_dict(REGISTRY), path)


def fake_registry():  # -> Registry
    return registry_from_dict(REGISTRY)


async def fake_snapshot(gw: FakeGraphGateway | None = None, prefix: str = "SG-"):  # -> Snapshot
    from m365tools.tools.groups.inventory import build_snapshot

    return await build_snapshot(gw or make_fake(), prefix)
