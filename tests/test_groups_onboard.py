"""Onboard (--bundle, --like, registry re-apply) and offboard: plans and apply via the fake."""

from __future__ import annotations

from datetime import date

import pytest

from m365tools.core.errors import ValidationError
from m365tools.tools.groups.onboard import apply_offboard, apply_onboard, plan_offboard, plan_onboard

from .fake_graph import fake_registry, fake_snapshot, make_fake


def _status(plan):
    return {c.group: c.status for c in plan.changes}


async def test_bundle_plan_and_manual() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    plan = plan_onboard(snap, reg, "Jan@contoso.com", bundles=["finance"], addons=["SG-Doc-Sales"])
    assert plan.upn == "jan@contoso.com" and plan.source == "bundle" and plan.bundles == ["Finance"]
    assert _status(plan) == {"SG-Doc-Finance": "planned", "SG-Doc-Hub": "planned", "SG-Lic-E3": "planned",
                             "SG-Doc-Sales": "planned"}
    assert plan.manual == ["App:Payroll"]  # every Finance person has Payroll directly; a licence group is planned


async def test_bundle_with_group_not_created() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    reg.bundles["Finance"].groups.append("SG-App-Payroll")
    plan = plan_onboard(snap, reg, "jan@contoso.com", bundles=["Finance"])
    assert _status(plan)["SG-App-Payroll"] == "not_created"
    assert "SG-App-Payroll" not in [c.group for c in plan.to_add]


async def test_like_registered_person() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    plan = plan_onboard(snap, reg, "jan@contoso.com", like="evelien@contoso.com")
    assert plan.source == "like evelien@contoso.com"
    assert plan.bundles == ["Servicedesk"] and plan.addons == ["SG-Doc-Finance"]
    assert len(plan.to_add) == 5 and plan.manual == []


async def test_like_unregistered_person_copies_actual_groups() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    plan = plan_onboard(snap, reg, "jan@contoso.com", like="ivo@contoso.com")
    assert plan.bundles == [] and plan.addons == ["SG-App-Jira", "SG-Doc-Hub", "SG-Lic-E5"]
    assert plan.manual == ["Role:Global Administrator"]  # the role is never copied, only listed


async def test_registry_reapply_and_already() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    plan = plan_onboard(snap, reg, "frank@contoso.com")
    assert plan.source == "registry"
    assert [c.group for c in plan.to_add] == ["SG-App-Jira"]
    assert _status(plan)["SG-Doc-Hub"] == "already"


async def test_plan_errors() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    with pytest.raises(ValidationError, match="not in the registry"):
        plan_onboard(snap, reg, "jan@contoso.com")
    with pytest.raises(ValidationError, match="Unknown bundle"):
        plan_onboard(snap, reg, "jan@contoso.com", bundles=["Nope"])
    with pytest.raises(ValidationError, match="Unknown group"):
        plan_onboard(snap, reg, "jan@contoso.com", bundles=["Finance"], addons=["SG-Nope"])
    with pytest.raises(ValidationError, match="no enabled member"):
        plan_onboard(snap, reg, "ghost@contoso.com", bundles=["Finance"])
    with pytest.raises(ValidationError, match="--like"):
        plan_onboard(snap, reg, "jan@contoso.com", like="ghost@contoso.com")


async def test_apply_onboard() -> None:
    gw = make_fake()
    snap, reg = await fake_snapshot(gw), fake_registry()
    gw.fail_on.add("SG-Doc-Sales")
    plan = plan_onboard(snap, reg, "jan@contoso.com", bundles=["Finance"], addons=["SG-Doc-Sales",
                                                                                 "SG-Doc-Hub"])
    await apply_onboard(gw, reg, plan)
    assert gw.members_of("SG-Doc-Finance") >= {"jan@contoso.com"}
    assert _status(plan)["SG-Doc-Sales"] == "failed"
    assert [c for c in gw.calls if c[0] != "add_member"] == []
    entry = reg.person("jan@contoso.com")
    assert entry is not None and entry.bundles == ["Finance"]
    assert entry.addons == ["SG-Doc-Sales"]  # SG-Doc-Hub is already covered by the bundle


async def test_offboard_plan_and_apply() -> None:
    gw = make_fake()
    snap, reg = await fake_snapshot(gw), fake_registry()
    reg.groups.pop("SG-Lic-E3")  # an unregistered group is kept
    plan = plan_offboard(snap, reg, "anna@contoso.com", today=date(2026, 10, 22))
    assert [c.group for c in plan.changes] == ["SG-App-Jira", "SG-Doc-Hub", "SG-Doc-Service"]
    assert plan.kept == ["SG-Lic-E3"] and plan.date == "2026-10-22"
    await apply_offboard(gw, reg, plan)
    assert all(c.status == "done" for c in plan.changes)
    assert "anna@contoso.com" not in gw.members_of("SG-Doc-Hub")
    assert "anna@contoso.com" in gw.members_of("SG-Lic-E3")
    assert reg.person("anna@contoso.com").offboarded == "2026-10-22"  # type: ignore[union-attr]
    assert {c[0] for c in gw.calls} == {"remove_member"}


async def test_reonboard_clears_offboarded() -> None:
    gw = make_fake()
    snap, reg = await fake_snapshot(gw), fake_registry()
    reg.people["anna@contoso.com"].offboarded = "2026-01-01"
    plan = plan_onboard(snap, reg, "anna@contoso.com", bundles=["Servicedesk"])
    await apply_onboard(gw, reg, plan)
    assert reg.person("anna@contoso.com").offboarded == ""  # type: ignore[union-attr]
