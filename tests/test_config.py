from __future__ import annotations

from pathlib import Path

import pytest

from m365tools.core.config import (
    AppConfig,
    add_tenant,
    get_profile,
    get_tenant,
    load_config,
    profile_to_job_options,
    save_config,
    tenant_for_url,
)
from m365tools.core.errors import ConfigError

SAMPLE = """
[tenants.contoso]
tenant_id = "contoso.onmicrosoft.com"
client_id = "11111111-2222-3333-4444-555555555555"
sharepoint_host = "contoso.sharepoint.com"

[profiles.servicemanagement]
source      = { tenant="contoso", site_url="https://contoso.sharepoint.com/sites/OldDocuments", library="Servicemanagement" }
destination = { tenant="contoso", site_url="https://contoso.sharepoint.com/sites/Service" }
excludes    = ["0. Archief/**"]
routes      = [{ prefix="0. Archief", destination={ tenant="contoso", site_url="https://contoso.sharepoint.com/sites/Service", library="Archief" } }]
"""


@pytest.fixture
def cfg_path(tmp_path: Path) -> Path:
    p = tmp_path / "config.toml"
    p.write_text(SAMPLE, encoding="utf-8")
    return p


def test_missing_file_is_empty(tmp_path: Path) -> None:
    cfg = load_config(tmp_path / "nope.toml")
    assert cfg.tenants == {} and cfg.profiles == {}


def test_load_sample(cfg_path: Path) -> None:
    cfg = load_config(cfg_path)
    t = get_tenant(cfg, "contoso")
    assert t.name == "contoso"
    assert t.sharepoint_host == "contoso.sharepoint.com"
    prof = get_profile(cfg, "servicemanagement")
    assert prof.name == "servicemanagement"
    assert prof.destination.library is None
    assert prof.routes[0].destination.library == "Archief"


def test_roundtrip(cfg_path: Path, tmp_path: Path) -> None:
    cfg = load_config(cfg_path)
    add_tenant(cfg, "other", tenant_id="other.onmicrosoft.com", client_id="abc",
               sharepoint_host="https://Other.sharepoint.com/")
    out = tmp_path / "sub" / "out.toml"
    save_config(cfg, out)
    text = out.read_text(encoding="utf-8")
    assert "name" not in text
    assert "None" not in text
    again = load_config(out)
    assert again.model_dump() == cfg.model_dump()
    assert again.tenants["other"].sharepoint_host == "other.sharepoint.com"


def test_save_default_path(tmp_path: Path) -> None:
    cfg = AppConfig()
    add_tenant(cfg, "t", tenant_id="t.onmicrosoft.com", client_id="c", sharepoint_host="t.sharepoint.com")
    p = save_config(cfg)
    assert p == tmp_path / "config" / "config.toml"
    assert load_config().tenants["t"].client_id == "c"


def test_profile_to_job_options(cfg_path: Path) -> None:
    opts = profile_to_job_options(load_config(cfg_path), "servicemanagement")
    assert opts.profile == "servicemanagement"
    assert opts.source.library == "Servicemanagement"
    assert opts.source.site_url.endswith("/OldDocuments")
    assert opts.destination.library is None and opts.destination.folder == ""
    assert opts.excludes == ["0. Archief/**"]
    assert opts.routes[0].prefix == "0. Archief"
    assert opts.routes[0].destination.library == "Archief"
    assert opts.overwrite == "skip_if_same_or_newer"
    assert opts.workers == 2 and opts.restore_metadata and not opts.lock_source_after
    assert opts.usermap_path is None


def test_missing_tenant_error(cfg_path: Path) -> None:
    cfg = load_config(cfg_path)
    cfg.profiles["servicemanagement"].destination.tenant = "ghost"
    with pytest.raises(ConfigError, match="unknown tenant 'ghost'.*contoso"):
        profile_to_job_options(cfg, "servicemanagement")
    with pytest.raises(ConfigError, match="Unknown tenant 'ghost'"):
        get_tenant(cfg, "ghost")


def test_unknown_profile(cfg_path: Path) -> None:
    with pytest.raises(ConfigError, match="servicemanagement"):
        get_profile(load_config(cfg_path), "nope")


def test_invalid_toml(tmp_path: Path) -> None:
    p = tmp_path / "bad.toml"
    p.write_text("[tenants.x\n", encoding="utf-8")
    with pytest.raises(ConfigError, match="invalid TOML"):
        load_config(p)
    p.write_text('[tenants.x]\ntenant_id="a"\n', encoding="utf-8")
    with pytest.raises(ConfigError, match="invalid configuration"):
        load_config(p)


def test_tenant_for_url(cfg_path: Path) -> None:
    cfg = load_config(cfg_path)
    assert tenant_for_url(cfg, "https://contoso.sharepoint.com/sites/x").name == "contoso"
    assert tenant_for_url(cfg, "https://CONTOSO-my.sharepoint.com/personal/x").name == "contoso"
    assert tenant_for_url(cfg, "https://fabrikam.sharepoint.com/sites/x") is None
