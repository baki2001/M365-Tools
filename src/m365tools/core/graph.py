"""Microsoft Graph calls the toolkit needs: site search fallback, user lookup for the user map, and the
directory reads/writes of the groups tool (groups, members, users, licences, enterprise apps, roles)."""

from __future__ import annotations

import re
from typing import Any

from m365tools.core.auth import GRAPH_SCOPE
from m365tools.core.events import EventSink, null_sink
from m365tools.core.http import SpClient, TokenSource

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_USER_SELECT = "id,displayName,mail,userPrincipalName,accountEnabled"
_USER_DETAIL_SELECT = "id,displayName,userPrincipalName,jobTitle,department,accountEnabled,userType,assignedLicenses"
_GROUP_SELECT = "id,displayName,description,securityEnabled,mailEnabled,groupTypes"
_MEMBER_SELECT = "id,displayName,userPrincipalName"


def _q(value: str) -> str:
    """Escape a string literal for an OData ``$filter`` (single quotes doubled)."""
    return value.replace("'", "''")


class GraphClient(SpClient):
    """``SpClient`` preconfigured for Microsoft Graph (plain JSON, Graph scope)."""

    default_accept = "application/json"
    json_content_type = "application/json"

    def __init__(
        self,
        token_provider: TokenSource,
        sink: EventSink = null_sink,
        max_attempts: int = 6,
        base_delay: float = 2.0,
    ) -> None:
        super().__init__(token_provider, GRAPH_SCOPE, sink=sink, max_attempts=max_attempts, base_delay=base_delay)

    async def search_sites(self, query: str) -> list[dict[str, str]]:
        """Sites matching ``query``: ``[{"title", "url", "description"}]``."""
        out: list[dict[str, str]] = []
        async for site in self.iter_odata(f"{GRAPH_BASE}/sites", params={"search": query}):
            url = site.get("webUrl") or ""
            if not url:
                continue
            out.append({
                "title": site.get("displayName") or site.get("name") or url,
                "url": url,
                "description": site.get("description") or "",
            })
        return out

    async def _users(self, flt: str) -> list[dict[str, Any]]:
        params = {"$filter": flt, "$select": _USER_SELECT, "$top": "999"}
        return [u async for u in self.iter_odata(f"{GRAPH_BASE}/users", params=params)]

    async def find_user_by_mail(self, mail: str) -> dict[str, Any] | None:
        """User whose ``mail`` (or else ``userPrincipalName``) equals ``mail``; ``None`` if not found."""
        m = _q(mail.strip())
        users = await self._users(f"mail eq '{m}'")
        if not users:
            users = await self._users(f"userPrincipalName eq '{m}'")
        return users[0] if users else None

    async def list_users_by_local_part(self, local: str) -> list[dict[str, Any]]:
        """Users whose mail starts with ``<local>@`` (used to suggest cross-tenant mappings)."""
        return await self._users(f"startswith(mail,'{_q(local.strip())}@')")

    # ------------------------------------------------------------------------------------------- groups tool

    async def _all(self, path: str, select: str, **params: str) -> list[dict[str, Any]]:
        q = {"$select": select, "$top": "999", **params}
        return [x async for x in self.iter_odata(f"{GRAPH_BASE}/{path}", params=q)]

    async def list_groups_with_prefix(self, prefix: str) -> list[dict[str, Any]]:
        """Groups whose displayName starts with ``prefix`` (no ConsistencyLevel needed for startswith)."""
        return await self._all("groups", _GROUP_SELECT, **{"$filter": f"startswith(displayName,'{_q(prefix)}')"})

    async def list_group_members(self, group_id: str) -> list[dict[str, Any]]:
        """Direct members (any object type; ``@odata.type`` tells users apart)."""
        return await self._all(f"groups/{group_id}/members", _MEMBER_SELECT)

    async def list_group_owners(self, group_id: str) -> list[dict[str, Any]]:
        return await self._all(f"groups/{group_id}/owners", _MEMBER_SELECT)

    async def list_users_detailed(self) -> list[dict[str, Any]]:
        """All users with job title, department, type, enabled flag and assigned licences."""
        return await self._all("users", _USER_DETAIL_SELECT)

    async def list_subscribed_skus(self) -> list[dict[str, Any]]:
        """``/subscribedSkus`` rejects ``$top`` ("does not support custom page sizes"), so no paging hint."""
        return [x async for x in self.iter_odata(f"{GRAPH_BASE}/subscribedSkus",
                                                 params={"$select": "skuId,skuPartNumber"})]

    async def list_service_principals(self) -> list[dict[str, Any]]:
        """Service principals with their ``appRoleAssignmentRequired`` flag (filtered by the caller)."""
        return await self._all("servicePrincipals", "id,displayName,appRoleAssignmentRequired")

    async def list_app_role_assigned_to(self, sp_id: str) -> list[dict[str, Any]]:
        """Users/groups assigned to an enterprise app: principalType, principalId, principalDisplayName."""
        return await self._all(f"servicePrincipals/{sp_id}/appRoleAssignedTo",
                               "principalType,principalId,principalDisplayName")

    async def list_directory_roles(self) -> list[dict[str, Any]]:
        """Activated directory roles (no $filter/$top support on this collection)."""
        return [x async for x in self.iter_odata(f"{GRAPH_BASE}/directoryRoles",
                                                 params={"$select": "id,displayName"})]

    async def list_directory_role_members(self, role_id: str) -> list[dict[str, Any]]:
        return [x async for x in self.iter_odata(f"{GRAPH_BASE}/directoryRoles/{role_id}/members",
                                                 params={"$select": _MEMBER_SELECT})]

    async def create_security_group(self, name: str, description: str = "") -> dict[str, Any]:
        """Create a security-enabled, mail-disabled group. Returns the new group (``id`` included)."""
        body: dict[str, Any] = {"displayName": name, "mailEnabled": False, "securityEnabled": True,
                                "mailNickname": mail_nickname(name)}
        if description:
            body["description"] = description
        return await self.post_json(f"{GRAPH_BASE}/groups", json=body)

    async def add_group_member(self, group_id: str, user_id: str) -> None:
        body = {"@odata.id": f"{GRAPH_BASE}/directoryObjects/{user_id}"}
        await self.post_json(f"{GRAPH_BASE}/groups/{group_id}/members/$ref", json=body)

    async def remove_group_member(self, group_id: str, user_id: str) -> None:
        await self.request("DELETE", f"{GRAPH_BASE}/groups/{group_id}/members/{user_id}/$ref")


def mail_nickname(name: str) -> str:
    """A ``mailNickname`` Graph accepts (ASCII letters, digits, '-', '_', '.'; max 64)."""
    nick = re.sub(r"[^A-Za-z0-9._-]", "", name)[:64]
    return nick or "group"
