"""Sync (registry -> tenant) and import (tenant -> registry): plans and apply via the fake."""

from __future__ import annotations

from m365tools.tools.groups.sync import apply_import, apply_sync, plan_import, plan_sync

from .fake_graph import fake_registry, fake_snapshot, make_fake


async def test_plan_sync() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    plan = plan_sync(snap, reg)
    assert [e.name for e in plan.create] == ["SG-App-Payroll"]
    assert len(plan.fill_ids) == 7 and ("SG-Doc-Hub", "g-sg-doc-hub") in plan.fill_ids
    for name, gid in plan.fill_ids:
        reg.group(name).id = gid  # type: ignore[union-attr]
    assert plan_sync(snap, reg).fill_ids == []


async def test_apply_sync_creates_and_fills_ids() -> None:
    gw = make_fake()
    snap, reg = await fake_snapshot(gw), fake_registry()
    plan = await apply_sync(gw, reg, plan_sync(snap, reg))
    assert gw.calls == [("create_group", "SG-App-Payroll", "Payroll app")]
    assert reg.group("SG-App-Payroll").id == plan.created[0][1] != ""  # type: ignore[union-attr]
    assert reg.group("SG-Doc-Hub").id == "g-sg-doc-hub"  # type: ignore[union-attr]
    assert plan.applied and not plan.failed
    assert not any(c[0] in ("remove_member",) for c in gw.calls)


async def test_apply_sync_records_failures() -> None:
    gw = make_fake()
    gw.fail_on.add("SG-App-Payroll")
    snap, reg = await fake_snapshot(gw), fake_registry()
    plan = await apply_sync(gw, reg, plan_sync(snap, reg))
    assert plan.failed[0][0] == "SG-App-Payroll" and plan.created == []
    assert reg.group("SG-App-Payroll").id == ""  # type: ignore[union-attr]


async def test_import() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    entries = plan_import(snap, reg)
    assert [(e.name, e.kind, e.purpose, e.id) for e in entries] == [("SG-Doc-Archive", "doc", "", "g-sg-doc-archive")]
    assert apply_import(reg, entries) == entries
    assert reg.group("SG-Doc-Archive") is not None
    assert plan_import(snap, reg) == [] and apply_import(reg, entries) == []
