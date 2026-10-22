"""Registry TOML: round trip, unknown keys preserved, never delete, expectations."""

from __future__ import annotations

import tomllib
from pathlib import Path

from m365tools.tools.groups.models import Bundle, GroupEntry, PersonEntry, Registry, kind_of
from m365tools.tools.groups.registry import (
    bundles_with,
    expectations,
    load_registry,
    registry_from_dict,
    save_registry,
)

from .fake_graph import REGISTRY, fake_registry, write_registry


def test_round_trip_preserves_unknown_keys(tmp_path: Path) -> None:
    path = write_registry(tmp_path / "contoso.toml")
    reg = load_registry(path)
    assert reg.extra == {"custom_top": {"owner_team": "IT"}}
    assert reg.groups["SG-Doc-Service"].extra == {"x_ticket": 12}
    assert reg.bundles["Servicedesk"].extra == {"review": "yearly"}
    ev = reg.person("EVELIEN@contoso.com")
    assert ev is not None and ev.extra == {"start": "2025-01-06"} and ev.note == "Also does invoicing"
    save_registry(reg, path)
    raw = tomllib.loads(path.read_text(encoding="utf-8"))
    assert raw["custom_top"] == {"owner_team": "IT"}
    assert raw["groups"]["SG-Doc-Service"]["x_ticket"] == 12
    assert raw["bundles"]["Servicedesk"]["review"] == "yearly"
    assert raw["people"]["evelien@contoso.com"]["start"] == "2025-01-06"
    assert load_registry(path) == reg


def test_missing_file_is_empty_registry(tmp_path: Path) -> None:
    reg = load_registry(tmp_path / "nope.toml")
    assert reg.prefix == "SG-" and not reg.groups and not reg.people


def test_save_never_deletes_entries(tmp_path: Path) -> None:
    path = write_registry(tmp_path / "r.toml")
    partial = Registry(prefix="SG-")
    partial.groups["SG-Doc-New"] = GroupEntry("SG-Doc-New", kind="doc", purpose="new")
    partial.people["zoe@contoso.com"] = PersonEntry("zoe@contoso.com", bundles=["Finance"])
    save_registry(partial, path)
    reg = load_registry(path)
    assert set(REGISTRY["groups"]) <= set(reg.groups)
    assert "SG-Doc-New" in reg.groups and "Servicedesk" in reg.bundles
    assert reg.person("anna@contoso.com") is not None and reg.person("zoe@contoso.com") is not None
    assert reg.extra["custom_top"] == {"owner_team": "IT"}


def test_save_merges_person_case_insensitively(tmp_path: Path) -> None:
    path = write_registry(tmp_path / "r.toml")  # key written lower-case already
    reg = load_registry(path)
    reg.people["evelien@contoso.com"].addons.append("SG-Doc-Sales")
    save_registry(reg, path)
    raw = tomllib.loads(path.read_text(encoding="utf-8"))
    assert [k for k in raw["people"] if k.lower() == "evelien@contoso.com"] == ["evelien@contoso.com"]
    assert raw["people"]["evelien@contoso.com"]["addons"] == ["SG-Doc-Finance", "SG-Doc-Sales"]


def test_expectations() -> None:
    reg = fake_registry()
    assert expectations(reg, "evelien@contoso.com") == {"SG-Doc-Service", "SG-Doc-Hub", "SG-App-Jira",
                                                        "SG-Lic-E3", "SG-Doc-Finance"}
    assert expectations(reg, "gina@contoso.com") == {"SG-Doc-Sales"}
    assert expectations(reg, "hugo@contoso.com") == set()
    reg.people["anna@contoso.com"].offboarded = "2026-10-01"
    assert expectations(reg, "anna@contoso.com") == set()
    reg.people["x@contoso.com"] = PersonEntry("x@contoso.com", bundles=["Unknown"], addons=["SG-A"])
    assert expectations(reg, "x@contoso.com") == {"SG-A"}


def test_kind_and_bundles_with() -> None:
    assert kind_of("SG-Doc-Finance") == "doc"
    assert kind_of("SG-App-Jira") == "app"
    assert kind_of("sg-lic-E3") == "lic"
    assert kind_of("SG-Team") == "other"
    assert kind_of("ACL-Doc-X", "ACL-") == "doc"
    reg = fake_registry()
    assert bundles_with(reg, "sg-doc-hub") == ["Servicedesk", "Finance"]
    reg.bundles["Empty"] = Bundle("Empty")
    assert bundles_with(reg, "SG-Doc-Sales") == []


def test_kind_inferred_when_missing() -> None:
    reg = registry_from_dict({"groups": {"SG-App-Teams": {"purpose": "x"}}})
    assert reg.groups["SG-App-Teams"].kind == "app"
