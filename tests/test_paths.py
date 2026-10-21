from __future__ import annotations

import re
from pathlib import Path

import pytest

from m365tools.core import paths


def test_env_overrides(tmp_path: Path) -> None:
    assert paths.config_dir() == tmp_path / "config"
    assert paths.data_dir() == tmp_path / "data"
    assert paths.cache_dir() == tmp_path / "cache"
    assert paths.runs_dir() == tmp_path / "data" / "runs"
    assert paths.config_path() == tmp_path / "config" / "config.toml"


def test_defaults_without_override(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    for name in ("CONFIG", "DATA", "CACHE"):
        monkeypatch.delenv(f"M365TOOLS_{name}_DIR", raising=False)
    for name in ("XDG_CONFIG_HOME", "XDG_DATA_HOME", "XDG_CACHE_HOME"):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv("HOME", str(tmp_path))
    assert paths.config_dir() == tmp_path / ".config" / "m365tools"
    assert paths.cache_dir() == tmp_path / ".cache" / "m365tools"
    assert paths.data_dir() == tmp_path / ".local" / "share" / "m365tools"


def test_new_run_dir(tmp_path: Path) -> None:
    run = paths.new_run_dir("servicemanagement", "dryrun")
    assert run.parent == tmp_path / "data" / "runs" / "servicemanagement"
    assert re.fullmatch(r"\d{8}-\d{6}-dryrun", run.name)
    assert (run / "tmp").is_dir()
    second = paths.new_run_dir("servicemanagement", "dryrun")
    assert second != run and second.is_dir()
