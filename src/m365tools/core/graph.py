"""Microsoft Graph calls the toolkit needs: site search fallback and user lookup for the user map."""

from __future__ import annotations

from typing import Any

from m365tools.core.auth import GRAPH_SCOPE
from m365tools.core.events import EventSink, null_sink
from m365tools.core.http import SpClient, TokenSource

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_USER_SELECT = "id,displayName,mail,userPrincipalName,accountEnabled"


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
