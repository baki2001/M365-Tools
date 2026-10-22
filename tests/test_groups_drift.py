"""Drift: one test per finding kind, plus the per-person view."""

from __future__ import annotations

from m365tools.tools.groups.drift import findings, person_findings
from m365tools.tools.groups.models import Severity

from .fake_graph import fake_registry, fake_snapshot


def _by(fs, cat):
    return [f for f in fs if f.category == cat]


async def test_each_finding_kind() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    fs = findings(snap, reg, "reg.toml")
    assert not [f for f in fs if f.severity == Severity.BLOCKING]
    unexplained = _by(fs, "unexplained")
    assert {(f.path, f.detail.split(":")[0]) for f in unexplained} == {
        ("gina@contoso.com", "member of SG-Doc-Hub"), ("gina@contoso.com", "member of SG-Lic-E3")}
    assert all(f.severity == Severity.WARN for f in unexplained)
    missing = _by(fs, "missing")
    assert [(f.path, f.fix) for f in missing] == [("frank@contoso.com", "m365 onboard frank@contoso.com --apply")]
    assert [f.path for f in _by(fs, "unregistered")] == ["SG-Doc-Archive"]
    assert [f.path for f in _by(fs, "not_created")] == ["SG-App-Payroll"]
    assert _by(fs, "not_created")[0].fix == "m365 groups sync --apply"
    assert sorted(f.path for f in _by(fs, "unassigned")) == ["hugo@contoso.com", "ivo@contoso.com",
                                                             "jan@contoso.com", "kim@contoso.com"]
    assert all(f.severity == Severity.INFO for f in _by(fs, "unassigned"))
    assert [f.path for f in _by(fs, "no_owner")] == ["SG-Doc-Archive"]
    assert [f.path for f in _by(fs, "empty")] == ["SG-Doc-Archive"]
    assert [f.path for f in _by(fs, "no_purpose")] == ["SG-Lic-E5"]
    assert "reg.toml" in _by(fs, "no_purpose")[0].fix
    assert all(f.fix for f in fs if f.category != "empty")


async def test_offboarded_person_still_in_groups() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    reg.people["anna@contoso.com"].offboarded = "2026-10-01"
    fs = person_findings(snap, reg, "Anna@contoso.com")
    assert len(fs) == 4 and all(f.category == "unexplained" for f in fs)
    assert "offboarded on 2026-10-01" in fs[0].detail and fs[0].fix == "m365 offboard anna@contoso.com --apply"


async def test_clean_person_has_no_findings() -> None:
    snap, reg = await fake_snapshot(), fake_registry()
    assert person_findings(snap, reg, "evelien@contoso.com") == []
