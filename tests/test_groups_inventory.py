"""Inventory: snapshot content, phases, matrix; Graph calls and RestGraphGateway via respx."""

from __future__ import annotations

import httpx
import respx

from m365tools.core.events import GROUPS_PHASES, CollectingSink, PhaseChanged, Stats
from m365tools.core.graph import GraphClient, mail_nickname
from m365tools.tools.groups.gateway import RestGraphGateway
from m365tools.tools.groups.inventory import build_snapshot, matrix_rows
from m365tools.tools.groups.models import Snapshot

from .fake_graph import fake_snapshot, make_fake

G = "https://graph.microsoft.com/v1.0"


async def test_snapshot_content() -> None:
    sink = CollectingSink()
    snap = await build_snapshot(make_fake(), "SG-", sink)
    assert len(snap.people) == 12  # guest and disabled skipped
    assert snap.person("guest_ext@contoso.com") is None and snap.person("left@contoso.com") is None
    assert len(snap.groups) == 8 and snap.group("All Staff") is None
    assert snap.groups_of("evelien@contoso.com") == {"SG-Doc-Service", "SG-Doc-Hub", "SG-App-Jira", "SG-Lic-E3",
                                                     "SG-Doc-Finance"}
    assert snap.person("ivo@contoso.com").licences == ["SPE_E5"]  # type: ignore[union-attr]
    assert [a.name for a in snap.apps] == ["Jira Cloud", "Payroll"]
    assert snap.apps_of("anna@contoso.com") == {"Jira Cloud": "SG-App-Jira"}
    assert snap.apps_of("dana@contoso.com") == {"Jira Cloud": "direct", "Payroll": "direct"}
    assert snap.roles_of("ivo@contoso.com") == ["Global Administrator"]
    assert snap.group("SG-Doc-Archive").owners == []  # type: ignore[union-attr]
    phases = [e.phase for e in sink.events if isinstance(e, PhaseChanged)]
    assert phases == list(GROUPS_PHASES[:-1])
    assert any(isinstance(e, Stats) and e.done == e.total == 8 for e in sink.events)


async def test_capabilities_and_matrix() -> None:
    snap = await fake_snapshot()
    labels = [c.label for c in snap.capabilities("ivo@contoso.com")]
    assert labels == ["SG-App-Jira", "SG-Doc-Hub", "SG-Lic-E5", "App:Jira Cloud", "Lic:SPE_E5",
                      "Role:Global Administrator"]
    cols, rows = matrix_rows(snap)
    assert cols[:4] == ["UPN", "Name", "Job title", "Department"]
    assert "SG-Doc-Archive" in cols and "App:Payroll" in cols and "Role:Global Administrator" in cols
    row = next(r for r in rows if r[0] == "dana@contoso.com")
    assert row[cols.index("App:Payroll")] == "x" and row[cols.index("SG-Doc-Sales")] == ""


async def test_snapshot_json_round_trip() -> None:
    snap = await fake_snapshot()
    assert Snapshot.from_json(snap.to_json()) == snap


@respx.mock
async def test_graph_calls_through_rest_gateway() -> None:
    class Tokens:
        async def token(self, scope: str, force_refresh: bool = False) -> str:
            return "t"

    groups = respx.get(f"{G}/groups").mock(side_effect=[
        httpx.Response(200, json={"value": [{"id": "g1", "displayName": "SG-Doc-A", "securityEnabled": True,
                                             "mailEnabled": False, "groupTypes": []}],
                                  "@odata.nextLink": f"{G}/groups?page=2"}),
        httpx.Response(200, json={"value": [{"id": "g2", "displayName": "SG-Doc-B"}]}),
    ])
    respx.get(f"{G}/groups/g1/members").mock(return_value=httpx.Response(200, json={"value": [
        {"@odata.type": "#microsoft.graph.user", "id": "u1", "userPrincipalName": "A@contoso.com"},
        {"@odata.type": "#microsoft.graph.group", "id": "g9", "displayName": "Nested"}]}))
    respx.get(f"{G}/users").mock(return_value=httpx.Response(200, json={"value": [
        {"id": "u1", "userPrincipalName": "a@contoso.com", "accountEnabled": True, "userType": "Member",
         "assignedLicenses": [{"skuId": "s1"}]}]}))
    respx.get(f"{G}/subscribedSkus").mock(return_value=httpx.Response(200, json={"value": [
        {"skuId": "s1", "skuPartNumber": "ENTERPRISEPACK"}]}))
    respx.get(f"{G}/servicePrincipals").mock(return_value=httpx.Response(200, json={"value": [
        {"id": "sp1", "displayName": "Jira", "appRoleAssignmentRequired": True},
        {"id": "sp2", "displayName": "Other", "appRoleAssignmentRequired": False}]}))
    respx.get(f"{G}/servicePrincipals/sp1/appRoleAssignedTo").mock(return_value=httpx.Response(200, json={
        "value": [{"principalType": "Group", "principalId": "g1", "principalDisplayName": "SG-Doc-A"}]}))
    roles = respx.get(f"{G}/directoryRoles").mock(return_value=httpx.Response(200, json={"value": [
        {"id": "r1", "displayName": "Global Administrator"}]}))
    respx.get(f"{G}/directoryRoles/r1/members").mock(return_value=httpx.Response(200, json={"value": [
        {"@odata.type": "#microsoft.graph.user", "id": "u1", "userPrincipalName": "a@contoso.com"}]}))
    create = respx.post(f"{G}/groups").mock(return_value=httpx.Response(201, json={"id": "new"}))
    add = respx.post(f"{G}/groups/g1/members/$ref").mock(return_value=httpx.Response(204))
    remove = respx.delete(f"{G}/groups/g1/members/u1/$ref").mock(return_value=httpx.Response(204))

    gw = RestGraphGateway(GraphClient(Tokens(), max_attempts=1))
    try:
        respx.get(f"{G}/groups/g1/owners").mock(return_value=httpx.Response(200, json={"value": []}))
        respx.get(f"{G}/groups/g2/members").mock(return_value=httpx.Response(200, json={"value": []}))
        respx.get(f"{G}/groups/g2/owners").mock(return_value=httpx.Response(200, json={"value": []}))
        snap = await build_snapshot(gw, "SG-")
        assert [g.name for g in snap.groups] == ["SG-Doc-A", "SG-Doc-B"]
        assert snap.groups[0].members == ["a@contoso.com"]
        assert snap.people[0].licences == ["ENTERPRISEPACK"]
        assert [a.name for a in snap.apps] == ["Jira"] and snap.apps[0].groups == ["SG-Doc-A"]
        assert snap.roles[0].members == ["a@contoso.com"]
        assert await gw.create_group("SG-Doc C/x", "purpose") == "new"
        await gw.add_member("g1", "u1")
        await gw.remove_member("g1", "u1")
    finally:
        await gw.aclose()
    q = groups.calls[0].request.url.params
    assert q["$filter"] == "startswith(displayName,'SG-')" and q["$top"] == "999" and "$select" in q
    assert "ConsistencyLevel" not in groups.calls[0].request.headers
    assert "$top" not in roles.calls[0].request.url.params
    import json

    body = json.loads(create.calls[0].request.content)
    assert body == {"displayName": "SG-Doc C/x", "mailEnabled": False, "securityEnabled": True,
                    "mailNickname": "SG-DocCx", "description": "purpose"}
    assert json.loads(add.calls[0].request.content) == {"@odata.id": f"{G}/directoryObjects/u1"}
    assert remove.called


def test_mail_nickname() -> None:
    assert mail_nickname("SG-Doc-Finance") == "SG-Doc-Finance"
    assert mail_nickname("ÄÖ ") == "group"


async def test_inventory_survives_access_denied_on_optional_phases():
    from m365tools.core.errors import AccessDeniedError
    from m365tools.core.events import CollectingSink
    from m365tools.tools.groups.inventory import build_snapshot

    from .fake_graph import make_fake

    gw = make_fake()

    async def denied(*_a, **_k):
        raise AccessDeniedError("Insufficient privileges")

    gw.subscribed_skus = denied  # type: ignore[method-assign]
    gw.list_apps = denied  # type: ignore[method-assign]
    gw.list_roles = denied  # type: ignore[method-assign]
    sink = CollectingSink()
    snap = await build_snapshot(gw, "SG-", sink)
    assert snap.people and snap.groups
    assert snap.apps == [] and snap.roles == []
    warns = [e for e in sink.events if getattr(e, "level", "") == "warn"]
    assert len(warns) == 3


async def test_inventory_leaves_out_external_and_other_domain_accounts():
    from m365tools.tools.groups.gateway import DirUser
    from m365tools.tools.groups.inventory import build_snapshot, is_colleague

    from .fake_graph import make_fake

    assert not is_colleague("x_other.nl#EXT#@contoso.onmicrosoft.com", [])
    assert is_colleague("anna@contoso.com", []) and is_colleague("anna@contoso.com", ["contoso.com"])
    assert not is_colleague("anna@fabrikam.com", ["contoso.com"])
    gw = make_fake()
    gw.users["ext1"] = DirUser("ext1", "x_other.nl#EXT#@contoso.onmicrosoft.com", "X")
    gw.users["oth1"] = DirUser("oth1", "zed@fabrikam.com", "Zed")
    snap = await build_snapshot(gw, "SG-")
    assert all("#ext#" not in p.upn for p in snap.people)
    assert any(p.upn == "zed@fabrikam.com" for p in snap.people)
    snap = await build_snapshot(gw, "SG-", domains=["contoso.com"])
    assert all(p.upn.endswith("@contoso.com") for p in snap.people)
