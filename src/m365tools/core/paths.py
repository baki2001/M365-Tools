"""XDG-style directories for config, cache, data and per-run folders.

Every directory can be overridden through an environment variable (used by tests and by people who want the
toolkit to live elsewhere): ``M365TOOLS_CONFIG_DIR``, ``M365TOOLS_CACHE_DIR``, ``M365TOOLS_DATA_DIR``.
Otherwise the standard ``XDG_*_HOME`` variables are honoured, falling back to ``~/.config``, ``~/.cache`` and
``~/.local/share``.
"""

from __future__ import annotations

import os
import re
from datetime import datetime
from pathlib import Path

APP_NAME = "m365tools"


def _resolve(override_env: str, xdg_env: str, default: str) -> Path:
    override = os.environ.get(override_env)
    if override:
        return Path(override).expanduser()
    base = os.environ.get(xdg_env)
    root = Path(base).expanduser() if base else Path.home() / default
    return root / APP_NAME


def config_dir() -> Path:
    """Directory holding ``config.toml`` (default ``~/.config/m365tools``)."""
    return _resolve("M365TOOLS_CONFIG_DIR", "XDG_CONFIG_HOME", ".config")


def cache_dir() -> Path:
    """Directory holding MSAL token caches (default ``~/.cache/m365tools``)."""
    return _resolve("M365TOOLS_CACHE_DIR", "XDG_CACHE_HOME", ".cache")


def data_dir() -> Path:
    """Directory holding run folders and other data (default ``~/.local/share/m365tools``)."""
    return _resolve("M365TOOLS_DATA_DIR", "XDG_DATA_HOME", ".local/share")


def config_path() -> Path:
    """Default location of the TOML config file."""
    return config_dir() / "config.toml"


def runs_dir() -> Path:
    """Root of all run folders: ``<data_dir>/runs``."""
    return data_dir() / "runs"


_UNSAFE = re.compile(r"[^A-Za-z0-9._-]+")


def _safe(name: str) -> str:
    cleaned = _UNSAFE.sub("_", name.strip()).strip("._")
    return cleaned or "adhoc"


def new_run_dir(profile_or_adhoc: str, mode: str) -> Path:
    """Create ``<runs>/<name>/<YYYYmmdd-HHMMSS>-<mode>/`` (with a ``tmp/`` subdir) and return it.

    If a folder with the same timestamp already exists (two runs within one second) a ``-2``, ``-3`` ... suffix is
    appended so runs never share a folder.
    """
    parent = runs_dir() / _safe(profile_or_adhoc)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    base = f"{stamp}-{_safe(mode)}"
    candidate = parent / base
    n = 2
    while True:
        try:
            candidate.mkdir(parents=True, exist_ok=False)
            break
        except FileExistsError:
            candidate = parent / f"{base}-{n}"
            n += 1
    (candidate / "tmp").mkdir(exist_ok=True)
    return candidate
