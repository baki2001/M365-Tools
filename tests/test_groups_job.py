"""GroupsJob: run folders, reports, registry writes only with apply, services helpers."""

from __future__ import annotations

import json
import tomllib
from pathlib import Path

from m365tools.core.config import AppConfig, add_tenant
from m365tools.core.events import GROUPS_PHASES, CollectingSink, PhaseChanged
from m365tools.services import groups_registry_path, groups_tenant, list_group_runs, make_groups_job
from m365tools.tools.groups.job import GroupsJob

from .fake_graph import make_fake, write_registry


def _job(tmp_path: Path, gw=None, sink=None) -> GroupsJob:
    reg = write_registry(tmp_path / "reg.toml")
    n = iter(range(100))
    return GroupsJob(gw or make_fake(), reg, sink or CollectingSink(),
                     lambda mode: tmp_path / "runs" / f"{next(n):02d}-{mode}", tenant="contoso")


async def test_inventory_run_folder_and_phases(tmp_path: Path) -> None:
    sink = CollectingSink()
    job = _job(tmp_path, sink=sink)
    snap = await job.inventory()
    assert job.run_dir is not None
    names = sorted(p.name for p in job.run_dir.iterdir())
    assert names == ["findings.csv", "matrix.csv", "report.html", "report.txt", "snapshot.json"]
    assert json.loads((job.run_dir / "snapshot.json").read_text())["prefix"] == "SG-"
    assert len(snap.people) == 12 and job.findings
    phases = [e.phase for e in sink.events if isinstance(e, PhaseChanged)]
    assert phases == list(GROUPS_PHASES)
    assert "Access groups inventory" in (job.run_dir / "report.html").read_text()
    assert (job.run_dir / "matrix.csv").read_text(encoding="utf-8-sig").startswith("UPN;Name")


async def test_snapshot_cached_between_ops(tmp_path: Path) -> None:
    sink = CollectingSink()
    job = _job(tmp_path, sink=sink)
    await job.drift()
    await job.explain("evelien")
    await job.suggest()
    assert sum(1 for e in sink.events if isinstance(e, PhaseChanged) and e.phase == "groups") == 1
    assert (job.run_dir / "candidates.csv").exists()


async def test_writes_only_with_apply(tmp_path: Path) -> None:
    gw = make_fake()
    job = _job(tmp_path, gw)
    before = job.registry_path.read_text()
    plan = await job.sync()
    await job.import_()
    await job.onboard("jan@contoso.com", bundles=["Finance"])
    await job.offboard("anna@contoso.com")
    sug = await job.suggest()
    job.accept(sug, {"3": "Sales"})
    assert gw.calls == [] and job.registry_path.read_text() == before
    assert plan.create and not plan.applied


async def test_apply_round(tmp_path: Path) -> None:
    gw = make_fake()
    job = _job(tmp_path, gw)
    plan = await job.sync(apply=True)
    assert plan.created and (job.run_dir / "sync.csv").exists()
    await job.import_(apply=True)
    await job.onboard("jan@contoso.com", like="dana@contoso.com", apply=True)
    await job.offboard("frank@contoso.com", apply=True)
    sug = await job.suggest()
    job.accept(sug, {"3": "Sales"}, apply=True)
    raw = tomllib.loads(job.registry_path.read_text())
    assert raw["groups"]["SG-App-Payroll"]["id"].startswith("g-")
    assert raw["groups"]["SG-Doc-Archive"]["purpose"] == ""
    assert raw["people"]["jan@contoso.com"]["bundles"] == ["Finance"]
    assert raw["people"]["frank@contoso.com"]["offboarded"]
    assert raw["bundles"]["Sales"]["groups"] == ["SG-Doc-Hub", "SG-Doc-Sales", "SG-Lic-E3"]
    assert raw["custom_top"] == {"owner_team": "IT"}  # unknown keys survived every write
    assert "jan@contoso.com" in gw.members_of("SG-Doc-Finance")
    findings = await job.drift()  # fresh snapshot after the writes
    cats = {f.category for f in findings}
    assert "not_created" not in cats and "unregistered" not in cats
    assert not [f for f in await job.onboard_check("jan@contoso.com")]
    assert gw.members_of("SG-Doc-Service") == {"anna@contoso.com", "bram@contoso.com", "chris@contoso.com",
                                               "evelien@contoso.com", "left@contoso.com"}


async def test_aclose_closes_gateway(tmp_path: Path) -> None:
    gw = make_fake()
    job = _job(tmp_path, gw)
    await job.aclose()
    assert gw.closed


async def test_services_helpers(tmp_path: Path) -> None:
    cfg = AppConfig()
    add_tenant(cfg, "contoso", tenant_id="contoso.onmicrosoft.com", client_id="c", sharepoint_host="contoso.sharepoint.com")
    assert groups_tenant(cfg) == "contoso"
    path = groups_registry_path("contoso")
    assert path.parts[-2:] == ("groups", "contoso.toml")
    job = make_groups_job(cfg, None, gateway=make_fake())
    assert job.registry_path == path and job.tenant == "contoso"
    await job.drift()
    await job.sync()
    assert [p.name.split("-", 2)[-1] for p in list_group_runs("contoso")] == ["sync", "drift"]
    assert len(list_group_runs(mode="drift")) == 1
    assert list_group_runs("other") == []
