"""GraphGateway: the only door from the groups tool to Microsoft Graph.

``RestGraphGateway`` adapts :class:`m365tools.core.graph.GraphClient` and normalises Graph JSON into the small
``Dir*`` records below. Tests use ``tests/fake_graph.py`` (``FakeGraphGateway``), which implements the same
Protocol in memory.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Protocol, runtime_checkable


@dataclass(slots=True)
class DirRef:
    """A directory object as seen in a member/owner list."""

    id: str
    name: str  # UPN for users, displayName otherwise
    type: str = "user"  # user | group | servicePrincipal | other


@dataclass(slots=True)
class DirGroup:
    id: str
    name: str
    description: str = ""
    security_enabled: bool = True
    mail_enabled: bool = False
    group_types: list[str] = field(default_factory=list)


@dataclass(slots=True)
class DirUser:
    id: str
    upn: str
    display_name: str = ""
    job_title: str = ""
    department: str = ""
    account_enabled: bool = True
    user_type: str = "Member"
    sku_ids: list[str] = field(default_factory=list)


@dataclass(slots=True)
class DirAssignment:
    principal_type: str  # User | Group | ServicePrincipal
    principal_id: str
    principal_name: str


@dataclass(slots=True)
class DirApp:
    id: str
    name: str
    assignment_required: bool


@runtime_checkable
class GraphGateway(Protocol):
    async def list_groups(self, prefix: str) -> list[DirGroup]: ...
    async def group_members(self, group_id: str) -> list[DirRef]: ...
    async def group_owners(self, group_id: str) -> list[DirRef]: ...
    async def list_users(self) -> list[DirUser]: ...
    async def subscribed_skus(self) -> dict[str, str]: ...  # skuId -> skuPartNumber
    async def list_apps(self) -> list[DirApp]: ...
    async def app_assignments(self, app_id: str) -> list[DirAssignment]: ...
    async def list_roles(self) -> list[DirRef]: ...
    async def role_members(self, role_id: str) -> list[DirRef]: ...
    async def create_group(self, name: str, description: str) -> str: ...  # returns the new id
    async def add_member(self, group_id: str, user_id: str) -> None: ...
    async def remove_member(self, group_id: str, user_id: str) -> None: ...


def _ref(raw: dict[str, Any]) -> DirRef:
    t = str(raw.get("@odata.type") or "#microsoft.graph.user").rsplit(".", 1)[-1]
    t = t if t in ("user", "group", "servicePrincipal") else "other"
    name = raw.get("userPrincipalName") if t == "user" else None
    return DirRef(id=str(raw.get("id", "")), name=str(name or raw.get("displayName") or raw.get("id", "")),
                  type=t)


class RestGraphGateway:
    """GraphGateway over a :class:`GraphClient` (owned: closed by :meth:`aclose`)."""

    def __init__(self, client: Any) -> None:
        self.client = client

    async def aclose(self) -> None:
        await self.client.aclose()

    async def list_groups(self, prefix: str) -> list[DirGroup]:
        rows = await self.client.list_groups_with_prefix(prefix)
        return [DirGroup(id=r["id"], name=r.get("displayName") or "", description=r.get("description") or "",
                         security_enabled=bool(r.get("securityEnabled")), mail_enabled=bool(r.get("mailEnabled")),
                         group_types=list(r.get("groupTypes") or [])) for r in rows]

    async def group_members(self, group_id: str) -> list[DirRef]:
        return [_ref(r) for r in await self.client.list_group_members(group_id)]

    async def group_owners(self, group_id: str) -> list[DirRef]:
        return [_ref(r) for r in await self.client.list_group_owners(group_id)]

    async def list_users(self) -> list[DirUser]:
        out = []
        for r in await self.client.list_users_detailed():
            out.append(DirUser(
                id=r["id"], upn=str(r.get("userPrincipalName") or ""), display_name=r.get("displayName") or "",
                job_title=r.get("jobTitle") or "", department=r.get("department") or "",
                account_enabled=bool(r.get("accountEnabled", True)), user_type=r.get("userType") or "Member",
                sku_ids=[str(lic.get("skuId")) for lic in r.get("assignedLicenses") or [] if lic.get("skuId")],
            ))
        return out

    async def subscribed_skus(self) -> dict[str, str]:
        return {str(r["skuId"]): str(r.get("skuPartNumber") or r["skuId"])
                for r in await self.client.list_subscribed_skus()}

    async def list_apps(self) -> list[DirApp]:
        return [DirApp(id=r["id"], name=r.get("displayName") or r["id"],
                       assignment_required=bool(r.get("appRoleAssignmentRequired")))
                for r in await self.client.list_service_principals()]

    async def app_assignments(self, app_id: str) -> list[DirAssignment]:
        return [DirAssignment(principal_type=str(r.get("principalType") or ""),
                              principal_id=str(r.get("principalId") or ""),
                              principal_name=str(r.get("principalDisplayName") or ""))
                for r in await self.client.list_app_role_assigned_to(app_id)]

    async def list_roles(self) -> list[DirRef]:
        return [DirRef(id=r["id"], name=r.get("displayName") or r["id"], type="role")
                for r in await self.client.list_directory_roles()]

    async def role_members(self, role_id: str) -> list[DirRef]:
        return [_ref(r) for r in await self.client.list_directory_role_members(role_id)]

    async def create_group(self, name: str, description: str) -> str:
        return str((await self.client.create_security_group(name, description)).get("id") or "")

    async def add_member(self, group_id: str, user_id: str) -> None:
        await self.client.add_group_member(group_id, user_id)

    async def remove_member(self, group_id: str, user_id: str) -> None:
        await self.client.remove_group_member(group_id, user_id)
