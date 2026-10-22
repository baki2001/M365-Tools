"""Explain: person and group views, subject resolution."""

from __future__ import annotations

import pytest

from m365tools.core.errors import NotFoundError
from m365tools.tools.groups.explain import explain_group, explain_person, resolve_subject

from .fake_graph import fake_registry, fake_snapshot


async def test_resolve_subject() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    assert resolve_subject(snap, reg, "sg-doc-hub") == ("group", "SG-Doc-Hub")
    assert resolve_subject(snap, reg, "SG-App-Payroll") == ("group", "SG-App-Payroll")  # registry only
    assert resolve_subject(snap, reg, "Evelien") == ("person", "evelien@contoso.com")
    assert resolve_subject(snap, reg, "IVO@contoso.com") == ("person", "ivo@contoso.com")
    with pytest.raises(NotFoundError):
        resolve_subject(snap, reg, "nobody")


async def test_explain_person() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    d = explain_person(snap, reg, "frank@contoso.com")
    assert d["bundles"] == ["Servicedesk"] and d["missing"] == ["SG-App-Jira"] and d["unexplained"] == []
    status = {g["group"]: (g["status"], g["via"]) for g in d["groups"]}
    assert status["SG-App-Jira"] == ("missing", "Servicedesk")
    assert status["SG-Doc-Hub"] == ("ok", "Servicedesk")
    g = explain_person(snap, reg, "gina@contoso.com")
    assert g["unexplained"] == ["SG-Doc-Hub", "SG-Lic-E3"]
    assert {x["group"]: x["via"] for x in g["groups"]}["SG-Doc-Sales"] == "addon"
    dana = explain_person(snap, reg, "dana@contoso.com")
    assert dana["apps"] == [{"app": "Jira Cloud", "via": "direct"}, {"app": "Payroll", "via": "direct"}]
    assert dana["licences"] == ["ENTERPRISEPACK"]
    assert explain_person(snap, reg, "ivo@contoso.com")["roles"] == ["Global Administrator"]
    assert explain_person(snap, reg, "ivo@contoso.com")["registered"] is False


async def test_explain_group() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    d = explain_group(snap, reg, "SG-Doc-Finance")
    assert d["purpose"] == "Finance site" and d["bundles"] == ["Finance"]
    assert d["addon_of"] == ["evelien@contoso.com"]
    assert d["unexplained"] == ["kim@contoso.com"]
    assert d["members"] == ["dana@contoso.com", "erik@contoso.com", "evelien@contoso.com", "kim@contoso.com"]
    j = explain_group(snap, reg, "SG-App-Jira")
    assert j["apps"] == ["Jira Cloud"] and j["tenant_owners"] == ["ivo@contoso.com"]
    p = explain_group(snap, reg, "SG-App-Payroll")
    assert p["in_tenant"] is False and p["members"] == []
    with pytest.raises(NotFoundError):
        explain_group(snap, reg, "SG-Nope")
