"""Glue shared by the CLI and the TUI: building jobs, run folders, tenant sessions, site/library lookups.

Nothing here prints; callers render. Kept separate from ``core`` because it wires core + tools together for a UI.
"""

from __future__ import annotations

import contextlib
import json
from collections.abc import Callable, Mapping
from dataclasses import asdict
from pathlib import Path
from typing import Any

from m365tools.core.config import (
    AppConfig,
    EndpointConfig,
    ProfileConfig,
    RouteConfig,
    get_tenant,
    tenant_for_url,
)
from m365tools.core.errors import ConfigError, NotFoundError
from m365tools.core.events import EventSink, null_sink
from m365tools.tools.spmig.job import RestGatewayFactory, RunDirFactory
from m365tools.tools.spmig.lock import LockChange
from m365tools.tools.spmig.models import Endpoint, Finding, JobOptions, Route, RunSummary, Severity

OPTIONS_FILE = "options.json"
DEFAULT_LIBRARY_LABEL = "(default library)"
# A name no real library has: asking get_library() for it yields NotFoundError(available=[...]).
_LIST_PROBE = "__m365tools_list_libraries__"
SEVERITY_ORDER = (Severity.BLOCKING, Severity.WARN, Severity.INFO)


# ------------------------------------------------------------------------------------------------ jobs


def make_gateway_factory(cfg: AppConfig, sink: EventSink = null_sink) -> RestGatewayFactory:
    """Production gateway factory (one authenticated SharePoint client per tenant)."""
    return RestGatewayFactory(cfg.tenants, sink)


def options_to_json(opts: JobOptions) -> dict[str, Any]:
    d = asdict(opts)
    d["usermap_path"] = str(opts.usermap_path) if opts.usermap_path else None
    return d


def options_from_json(d: Mapping[str, Any]) -> JobOptions:
    def ep(x: Mapping[str, Any]) -> Endpoint:
        return Endpoint(tenant=x["tenant"], site_url=x["site_url"], library=x.get("library"),
                        folder=x.get("folder") or "")

    return JobOptions(
        source=ep(d["source"]), destination=ep(d["destination"]),
        overwrite=d.get("overwrite", "skip_if_same_or_newer"),
        restore_metadata=bool(d.get("restore_metadata", True)),
        excludes=list(d.get("excludes") or []),
        routes=[Route(r["prefix"], ep(r["destination"])) for r in d.get("routes") or []],
        lock_source_after=bool(d.get("lock_source_after", False)),
        workers=int(d.get("workers", 2)),
        usermap_path=Path(d["usermap_path"]) if d.get("usermap_path") else None,
        profile=d.get("profile"),
    )


def save_run_options(run_dir: Path, opts: JobOptions) -> None:
    (Path(run_dir) / OPTIONS_FILE).write_text(json.dumps(options_to_json(opts), indent=1), encoding="utf-8")


def load_run_options(run_dir: Path, cfg: AppConfig | None = None) -> JobOptions:
    """Options a run was started with: ``options.json`` in the run folder, else the profile the folder is under."""
    p = Path(run_dir) / OPTIONS_FILE
    if p.exists():
        return options_from_json(json.loads(p.read_text(encoding="utf-8")))
    profile = Path(run_dir).parent.name
    if cfg is not None and profile in cfg.profiles:
        from m365tools.core.config import profile_to_job_options

        return profile_to_job_options(cfg, profile)
    raise ConfigError(f"{run_dir}: no {OPTIONS_FILE} and no matching profile; pass the options explicitly")


def make_run_dir_factory(opts: JobOptions) -> RunDirFactory:
    """New run folder under ``runs/<profile|adhoc>/`` that also records the job options (for resume)."""

    def make(mode: str) -> Path:
        from m365tools.core.paths import new_run_dir

        run_dir = new_run_dir(opts.profile or "adhoc", mode)
        with contextlib.suppress(OSError):
            save_run_options(run_dir, opts)
        return run_dir

    return make


def list_runs(mode: str = "migrate") -> list[Path]:
    """Run folders of ``mode`` (newest first)."""
    from m365tools.core.paths import runs_dir

    root = runs_dir()
    if not root.exists():
        return []
    runs = [p for p in root.glob(f"*/*-{mode}*") if p.is_dir()]
    return sorted(runs, key=lambda p: p.name, reverse=True)


def default_tenant(cfg: AppConfig, site_url: str | None, fallback: str | None = None) -> str:
    """Tenant name for ``site_url``: by host, else ``fallback``, else the only configured tenant."""
    if site_url:
        t = tenant_for_url(cfg, site_url)
        if t is not None:
            return t.name
    if fallback:
        return fallback
    if len(cfg.tenants) == 1:
        return next(iter(cfg.tenants))
    known = ", ".join(sorted(cfg.tenants)) or "none configured"
    raise ConfigError(f"Cannot tell which tenant {site_url or 'the site'} belongs to (known: {known}); "
                      "pass the tenant explicitly")


def opts_to_profile(opts: JobOptions) -> ProfileConfig:
    def ec(ep: Endpoint) -> EndpointConfig:
        return EndpointConfig(tenant=ep.tenant, site_url=ep.site_url, library=ep.library, folder=ep.folder)

    return ProfileConfig(
        source=ec(opts.source), destination=ec(opts.destination), overwrite=opts.overwrite,
        restore_metadata=opts.restore_metadata, excludes=list(opts.excludes),
        routes=[RouteConfig(prefix=r.prefix, destination=ec(r.destination)) for r in opts.routes],
        lock_source_after=opts.lock_source_after, workers=opts.workers,
        usermap_path=str(opts.usermap_path) if opts.usermap_path else None,
    )


# ------------------------------------------------------------------------------------------------ output


def finding_to_json(f: Finding) -> dict[str, str]:
    return {"severity": str(f.severity), "category": f.category, "path": f.path, "detail": f.detail, "fix": f.fix}


def summary_to_json(s: RunSummary, **extra: Any) -> dict[str, Any]:
    out: dict[str, Any] = {
        "mode": s.mode, "started": s.started.isoformat(), "finished": s.finished.isoformat(),
        "duration_s": round(s.duration_s, 3), "run_dir": str(s.run_dir),
        "report": str(Path(s.run_dir) / "report.html"), "copied": s.copied, "skipped": s.skipped,
        "failed": s.failed, "metadata_warnings": s.metadata_warnings, "bytes_copied": s.bytes_copied,
        "stopped_early": s.stopped_early,
        "findings": [finding_to_json(f) for f in s.findings],
        "blocking": len(s.blocking),
    }
    out.update(extra)
    return out


def lock_change_to_json(c: LockChange) -> dict[str, Any]:
    return {"principal_id": c.principal_id, "principal": c.title, "from_roles": list(c.from_roles),
            "to_role": c.to_role if c.will_change else None, "will_change": c.will_change,
            "skipped_reason": c.skipped_reason}


def sorted_findings(findings: list[Finding]) -> list[Finding]:
    order = {s: i for i, s in enumerate(SEVERITY_ORDER)}
    return sorted(findings, key=lambda f: (order.get(f.severity, 9), f.category, f.path))


# ------------------------------------------------------------------------------------------------ sessions


class TenantClients:
    """Lazily created, shared per-tenant clients for interactive lookups (site search, libraries, login)."""

    def __init__(self, cfg: AppConfig, sink: EventSink | Callable[[Any], None] = null_sink) -> None:
        self.cfg = cfg
        self.sink = sink
        self._providers: dict[str, Any] = {}
        self._sp: dict[str, Any] = {}
        self._graph: dict[str, Any] = {}
        self._closing: list[Any] = []

    def provider(self, tenant: str) -> Any:
        from m365tools.core.auth import TokenProvider

        if tenant not in self._providers:
            self._providers[tenant] = TokenProvider(get_tenant(self.cfg, tenant), prompt=self._prompt)
        return self._providers[tenant]

    def forget(self, tenant: str) -> None:
        """Drop cached clients for ``tenant`` (after its configuration changed)."""
        self._providers.pop(tenant, None)
        for d in (self._sp, self._graph):
            c = d.pop(tenant, None)
            if c is not None:
                self._closing.append(c)

    def _prompt(self, ev: Any) -> None:
        self.sink(ev)

    def sp(self, tenant: str) -> Any:
        from m365tools.core.auth import sharepoint_scope
        from m365tools.core.http import SpClient

        if tenant not in self._sp:
            self._sp[tenant] = SpClient(self.provider(tenant), sharepoint_scope(get_tenant(self.cfg, tenant)),
                                        sink=self._prompt)
        return self._sp[tenant]

    def graph(self, tenant: str) -> Any:
        from m365tools.core.graph import GraphClient

        if tenant not in self._graph:
            self._graph[tenant] = GraphClient(self.provider(tenant), sink=self._prompt)
        return self._graph[tenant]

    async def login(self, tenant: str) -> list[str]:
        """Acquire a SharePoint token (interactive if needed) and return the cached account names."""
        from m365tools.core.auth import sharepoint_scope

        prov = self.provider(tenant)
        await prov.token(sharepoint_scope(get_tenant(self.cfg, tenant)))
        return [str(a.get("username") or "?") for a in prov.accounts()]

    async def search_sites(self, tenant: str, query: str) -> list[Any]:
        from m365tools.core import sharepoint

        t = get_tenant(self.cfg, tenant)
        return await sharepoint.search_sites(self.sp(tenant), query, root_url=t.root_url,
                                             graph_client=self.graph(tenant))

    async def list_libraries(self, tenant: str, site_url: str) -> list[str]:
        """Titles of the visible document libraries of ``site_url``."""
        from m365tools.core import sharepoint

        try:
            lib = await sharepoint.get_library(self.sp(tenant), site_url.rstrip("/"), _LIST_PROBE)
        except NotFoundError as exc:
            return list(exc.available)
        return [lib.title]

    async def aclose(self) -> None:
        for c in [*self._sp.values(), *self._graph.values(), *self._closing]:
            with contextlib.suppress(Exception):
                await c.aclose()
        self._sp.clear()
        self._graph.clear()
        self._closing.clear()


# ------------------------------------------------------------------------------------------------ groups


def groups_tenant(cfg: AppConfig, name: str | None = None) -> str:
    """The tenant a groups command works on: ``name`` if given (must exist), else the only/first tenant."""
    if name:
        get_tenant(cfg, name)
        return name
    if not cfg.tenants:
        raise ConfigError("No tenants configured. Add one with: m365 tenant add NAME")
    return next(iter(cfg.tenants))


def groups_registry_path(tenant: str) -> Path:
    """``<config_dir>/groups/<tenant>.toml``."""
    from m365tools.core.paths import config_dir

    return config_dir() / "groups" / f"{tenant}.toml"


def make_groups_gateway(cfg: AppConfig, tenant: str, sink: EventSink = null_sink) -> Any:
    """Production GraphGateway for ``tenant`` (one TokenProvider + GraphClient; closed by the job)."""
    from m365tools.core.auth import TokenProvider
    from m365tools.core.graph import GraphClient
    from m365tools.tools.groups.gateway import RestGraphGateway

    provider = TokenProvider(get_tenant(cfg, tenant), prompt=sink)
    return RestGraphGateway(GraphClient(provider, sink=sink))


def make_groups_job(cfg: AppConfig, tenant: str | None = None, sink: EventSink = null_sink,
                    gateway: Any = None) -> Any:
    """GroupsJob for ``tenant`` with its registry file and run folders under ``runs/groups-<tenant>/``."""
    from m365tools.tools.groups.job import GroupsJob

    name = groups_tenant(cfg, tenant)
    gw = gateway if gateway is not None else make_groups_gateway(cfg, name, sink)
    return GroupsJob(gw, groups_registry_path(name), sink, tenant=name)


def list_group_runs(tenant: str | None = None, mode: str | None = None) -> list[Path]:
    """groups run folders (newest first), optionally of one tenant and/or one mode."""
    from m365tools.core.paths import runs_dir

    root = runs_dir()
    if not root.exists():
        return []
    parent = f"groups-{tenant}" if tenant else "groups-*"
    runs = [p for p in root.glob(f"{parent}/*-{mode}*" if mode else f"{parent}/*") if p.is_dir()]
    return sorted(runs, key=lambda p: p.name, reverse=True)


def to_jsonable(obj: Any) -> Any:
    """Dataclasses (plans, snapshots, findings) as plain JSON-able values."""
    import dataclasses

    if dataclasses.is_dataclass(obj) and not isinstance(obj, type):
        return {k: to_jsonable(v) for k, v in asdict(obj).items()}
    if isinstance(obj, list | tuple):
        return [to_jsonable(x) for x in obj]
    if isinstance(obj, dict):
        return {str(k): to_jsonable(v) for k, v in obj.items()}
    return obj if obj is None or isinstance(obj, str | int | float | bool) else str(obj)
