"""Suggest: candidate bundles, naming, nearest-candidate Jaccard math, accepting."""

from __future__ import annotations

import pytest

from m365tools.tools.groups.models import Registry
from m365tools.tools.groups.suggest import accept, candidates, jaccard, nearest, suggest

from .fake_graph import fake_registry, fake_snapshot


def test_jaccard() -> None:
    assert jaccard({"a", "b"}, {"a", "b"}) == 1.0
    assert jaccard({"a", "b", "c"}, {"a", "b", "c", "d"}) == 0.75
    assert jaccard(set(), set()) == 0.0
    assert jaccard({"a"}, {"b"}) == 0.0


async def test_candidates_from_shared_sets() -> None:
    snap = await fake_snapshot()
    cands = candidates(snap, Registry())
    assert [(c.key, c.name, len(c.people)) for c in cands] == [
        ("1", "Service Desk", 3), ("2", "Accountant", 2), ("3", "Sales Manager", 2)]
    assert cands[0].groups == ["SG-App-Jira", "SG-Doc-Hub", "SG-Doc-Service", "SG-Lic-E3"]
    assert cands[1].people == ["dana@contoso.com", "erik@contoso.com"]


async def test_candidates_reuse_existing_bundle_names() -> None:
    snap = await fake_snapshot()
    cands = candidates(snap, fake_registry())
    assert [(c.name, c.existing) for c in cands] == [("Servicedesk", "Servicedesk"), ("Finance", "Finance"),
                                                     ("Sales Manager", "")]


async def test_nearest_math() -> None:
    snap = await fake_snapshot()
    near = {m.upn: m for m in nearest(snap, candidates(snap, Registry()))}
    assert set(near) == {"evelien@contoso.com", "frank@contoso.com", "ivo@contoso.com", "kim@contoso.com"}
    ev = near["evelien@contoso.com"]  # SD + Finance: 4/5
    assert (ev.candidate, ev.score, ev.addons, ev.missing) == ("1", 0.8, ["SG-Doc-Finance"], [])
    fr = near["frank@contoso.com"]  # SD minus Jira: 3/4
    assert (fr.candidate, fr.score, fr.addons, fr.missing) == ("1", 0.75, [], ["SG-App-Jira"])
    iv = near["ivo@contoso.com"]  # {Hub, Jira, E5} vs SD: 2/5 = 0.4 beats Finance/Sales 1/5
    assert (iv.candidate, iv.score, iv.addons) == ("1", 0.4, ["SG-Lic-E5"])
    assert iv.missing == ["SG-Doc-Service", "SG-Lic-E3"]
    km = near["kim@contoso.com"]  # {Finance} vs Finance bundle: 1/3
    assert (km.candidate, km.score, km.missing) == ("2", 0.333, ["SG-Doc-Hub", "SG-Lic-E3"])
    assert "jan@contoso.com" not in near  # no groups at all


async def test_accept_writes_bundles_and_people() -> None:
    snap = await fake_snapshot()
    reg = Registry()
    sug = suggest(snap, reg)
    changes = accept(reg, sug, {"1": "Servicedesk", "Sales Manager": ""})
    assert reg.bundles["Servicedesk"].groups == sug.candidates[0].groups
    assert "Sales Manager" in reg.bundles
    assert reg.people["anna@contoso.com"].bundles == ["Servicedesk"]
    assert reg.people["evelien@contoso.com"].addons == ["SG-Doc-Finance"]  # near match, score 0.8
    assert reg.people["frank@contoso.com"].bundles == ["Servicedesk"]
    assert "ivo@contoso.com" not in reg.people  # 0.4 < threshold
    assert reg.people["hugo@contoso.com"].bundles == ["Sales Manager"]
    assert any(line.startswith("bundle Servicedesk") for line in changes)


async def test_accept_extends_existing_entries_only() -> None:
    snap = await fake_snapshot()
    reg = fake_registry()
    sug = suggest(snap, reg)
    accept(reg, sug, {"3": "Sales"}, include_near=False)
    assert reg.people["gina@contoso.com"].bundles == ["Sales"]
    assert reg.people["gina@contoso.com"].addons == ["SG-Doc-Sales"]  # kept
    with pytest.raises(KeyError):
        accept(reg, sug, {"99": "x"})
