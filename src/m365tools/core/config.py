"""Configuration models and TOML persistence (``~/.config/m365tools/config.toml``).

The TOML layout is documented in docs/ARCHITECTURE.md. Tenant and profile names are the TOML table keys; the
``name`` attribute on the models is filled from the key on load and is not written back to the file.
"""

from __future__ import annotations

import os
import tempfile
import tomllib
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import tomli_w
from pydantic import BaseModel, ConfigDict, Field, field_validator
from pydantic import ValidationError as PydanticValidationError

from m365tools.core.errors import ConfigError
from m365tools.core.paths import config_path
from m365tools.tools.spmig.models import Endpoint, JobOptions, OverwritePolicy, Route


def _normalise_host(value: str) -> str:
    v = value.strip()
    if "://" in v:
        v = urlparse(v).netloc
    return v.strip("/").lower()


class _Model(BaseModel):
    model_config = ConfigDict(extra="forbid", validate_assignment=True)


class TenantConfig(_Model):
    """One Entra tenant plus the public-client app registration used to sign in."""

    tenant_id: str
    client_id: str
    sharepoint_host: str
    name: str = ""

    @field_validator("sharepoint_host")
    @classmethod
    def _host(cls, v: str) -> str:
        host = _normalise_host(v)
        if not host:
            raise ValueError("sharepoint_host must be a host name like 'contoso.sharepoint.com'")
        return host

    @property
    def root_url(self) -> str:
        """``https://<sharepoint_host>`` (root site collection)."""
        return f"https://{self.sharepoint_host}"


class EndpointConfig(_Model):
    tenant: str
    site_url: str
    library: str | None = None
    folder: str = ""

    @field_validator("site_url")
    @classmethod
    def _site(cls, v: str) -> str:
        return v.strip().rstrip("/")

    def to_endpoint(self) -> Endpoint:
        return Endpoint(tenant=self.tenant, site_url=self.site_url, library=self.library, folder=self.folder)


class RouteConfig(_Model):
    prefix: str
    destination: EndpointConfig


class ProfileConfig(_Model):
    source: EndpointConfig
    destination: EndpointConfig
    overwrite: OverwritePolicy = "skip_if_same_or_newer"
    restore_metadata: bool = True
    excludes: list[str] = Field(default_factory=list)
    routes: list[RouteConfig] = Field(default_factory=list)
    lock_source_after: bool = False
    workers: int = Field(default=2, ge=1, le=16)
    usermap_path: str | None = None
    name: str = ""


class AppConfig(_Model):
    tenants: dict[str, TenantConfig] = Field(default_factory=dict)
    profiles: dict[str, ProfileConfig] = Field(default_factory=dict)
    path: Path | None = Field(default=None, exclude=True)

    def model_post_init(self, context: Any, /) -> None:
        for key, t in self.tenants.items():
            if not t.name:
                t.name = key
        for key, p in self.profiles.items():
            if not p.name:
                p.name = key


# ---------------------------------------------------------------------------------------------------- load/save


def load_config(path: Path | None = None) -> AppConfig:
    """Load the config file. A missing file yields an empty :class:`AppConfig`."""
    p = path or config_path()
    if not p.exists():
        return AppConfig(path=p)
    try:
        with p.open("rb") as fh:
            raw = tomllib.load(fh)
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"{p}: invalid TOML: {exc}") from exc
    except OSError as exc:
        raise ConfigError(f"{p}: cannot read config: {exc}") from exc
    try:
        cfg = AppConfig.model_validate(raw)
    except PydanticValidationError as exc:
        raise ConfigError(f"{p}: invalid configuration:\n{exc}") from exc
    cfg.path = p
    return cfg


def _dump(cfg: AppConfig) -> dict[str, Any]:
    data = cfg.model_dump(mode="json", exclude_none=True)
    for t in data.get("tenants", {}).values():
        t.pop("name", None)
    for prof in data.get("profiles", {}).values():
        prof.pop("name", None)
    return data


def save_config(cfg: AppConfig, path: Path | None = None) -> Path:
    """Write the config atomically (tmp file + ``os.replace``). Returns the path written."""
    p = path or cfg.path or config_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    payload = tomli_w.dumps(_dump(cfg))
    fd, tmp = tempfile.mkstemp(prefix=".config-", suffix=".toml", dir=p.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(payload)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, p)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise
    cfg.path = p
    return p


# ---------------------------------------------------------------------------------------------------- helpers


def add_tenant(
    cfg: AppConfig,
    name: str,
    *,
    tenant_id: str,
    client_id: str,
    sharepoint_host: str,
    replace: bool = True,
) -> TenantConfig:
    """Add (or replace) a tenant entry in ``cfg`` (not saved; call :func:`save_config`)."""
    name = name.strip()
    if not name:
        raise ConfigError("Tenant name must not be empty.")
    if name in cfg.tenants and not replace:
        raise ConfigError(f"Tenant '{name}' already exists. Use a different name or replace it.")
    try:
        tenant = TenantConfig(
            name=name, tenant_id=tenant_id.strip(), client_id=client_id.strip(), sharepoint_host=sharepoint_host
        )
    except PydanticValidationError as exc:
        raise ConfigError(f"Invalid tenant '{name}':\n{exc}") from exc
    cfg.tenants[name] = tenant
    return tenant


def get_tenant(cfg: AppConfig, name: str) -> TenantConfig:
    try:
        return cfg.tenants[name]
    except KeyError:
        known = ", ".join(sorted(cfg.tenants)) or "none configured"
        raise ConfigError(
            f"Unknown tenant '{name}' (known: {known}). Add it with: m365 tenant add {name}"
        ) from None


def get_profile(cfg: AppConfig, name: str) -> ProfileConfig:
    try:
        return cfg.profiles[name]
    except KeyError:
        known = ", ".join(sorted(cfg.profiles)) or "none configured"
        where = cfg.path or config_path()
        raise ConfigError(
            f"Unknown profile '{name}' (known: {known}). Define [profiles.{name}] in {where}."
        ) from None


def tenant_for_url(cfg: AppConfig, url: str) -> TenantConfig | None:
    """Return the tenant whose ``sharepoint_host`` matches the host of ``url`` (also ``-my``/``-admin`` hosts)."""
    host = _normalise_host(url)
    if not host:
        return None
    for t in cfg.tenants.values():
        if t.sharepoint_host == host:
            return t
    for t in cfg.tenants.values():
        stem, _, rest = t.sharepoint_host.partition(".")
        for suffix in ("-my", "-admin"):
            if host == f"{stem}{suffix}.{rest}":
                return t
    return None


def profile_to_job_options(cfg: AppConfig, name: str) -> JobOptions:
    """Convert a stored profile into :class:`JobOptions`, validating that every referenced tenant exists."""
    prof = get_profile(cfg, name)
    endpoints = [("source", prof.source), ("destination", prof.destination)]
    endpoints += [(f"route '{r.prefix}'", r.destination) for r in prof.routes]
    for label, ep in endpoints:
        if ep.tenant not in cfg.tenants:
            known = ", ".join(sorted(cfg.tenants)) or "none configured"
            raise ConfigError(
                f"Profile '{name}': {label} refers to unknown tenant '{ep.tenant}' (known: {known})."
            )
    usermap = Path(prof.usermap_path).expanduser() if prof.usermap_path else None
    return JobOptions(
        source=prof.source.to_endpoint(),
        destination=prof.destination.to_endpoint(),
        overwrite=prof.overwrite,
        restore_metadata=prof.restore_metadata,
        excludes=list(prof.excludes),
        routes=[Route(prefix=r.prefix, destination=r.destination.to_endpoint()) for r in prof.routes],
        lock_source_after=prof.lock_source_after,
        workers=prof.workers,
        usermap_path=usermap,
        profile=name,
    )
