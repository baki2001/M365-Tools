"""CLI tests: typer CliRunner with a fake Job (canned results) and, end-to-end, the in-memory gateway."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, ClassVar

import pytest
from typer.testing import CliRunner

from m365tools import cli
from m365tools.core.config import AppConfig, ProfileConfig, add_tenant, load_config, save_config
from m365tools.core.events import LogLine, Stats
from m365tools.core.sharepoint import SiteHit
from m365tools.tools.spmig.job import BlockingFindingsError
from m365tools.tools.spmig.lock import LockChange
from m365tools.tools.spmig.models import Action, Finding, Plan, RunSummary, Severity

from .fake_gateway import factory_for, make_inv, make_item, make_lib

SRC = "https://contoso.sharepoint.com/sites/src"
DST = "https://contoso.sharepoint.com/sites/dst"
ENV = {"COLUMNS": "200"}

runner = CliRunner()


def write_cfg() -> AppConfig:
    cfg = AppConfig()
    add_tenant(cfg, "contoso", tenant_id="contoso.onmicrosoft.com", client_id="cid",
               sharepoint_host="contoso.sharepoint.com")
    cfg.profiles["svc"] = ProfileConfig.model_validate({
        "source": {"tenant": "contoso", "site_url": SRC, "library": "Servicemanagement"},
        "destination": {"tenant": "contoso", "site_url": DST},
        "excludes": ["0. Archief/**"],
        "routes": [{"prefix": "0. Archief", "destination": {"tenant": "contoso", "site_url": DST,
                                                             "library": "Archief"}}],
    })
    save_config(cfg)
    return cfg


def summary(mode: str, run_dir: Path, **kw: Any) -> RunSummary:
    now = datetime.now()
    return RunSummary(mode, now, now, run_dir, **kw)  # type: ignore[arg-type]


class FakeJob:
    """Stands in for spmig.job.Job. Configure through class attributes before invoking the CLI."""

    findings: ClassVar[list[Finding]] = []
    raise_blocking = False
    migrate_kw: ClassVar[dict[str, Any]] = {}
    last: FakeJob | None = None

    def __init__(self, gateways: Any, opts: Any, sink: Any, run_dir_factory: Any = None) -> None:
        self.opts, self.sink, self.run_dir_factory = opts, sink, run_dir_factory
        self.gateways = gateways
        self.run_dir: Path | None = None
        self.calls: list[tuple[str, dict[str, Any]]] = []
        FakeJob.last = self

    def _dir(self, mode: str) -> Path:
        self.run_dir = self.run_dir_factory(mode)
        return self.run_dir

    async def audit(self) -> RunSummary:
        self.calls.append(("audit", {}))
        run_dir = self._dir("audit")
        if FakeJob.raise_blocking:
            raise BlockingFindingsError(list(FakeJob.findings), run_dir)
        self.sink(LogLine("info", "inventory done"))
        return summary("audit", run_dir, findings=list(FakeJob.findings))

    async def dryrun(self) -> tuple[Plan, RunSummary]:
        self.calls.append(("dryrun", {}))
        lib = make_lib()
        inv = make_inv([make_item("a.txt", size=5)], lib)
        from m365tools.tools.spmig.models import PlanEntry

        plan = Plan([PlanEntry(Action.COPY, inv.items[0], "/sites/b/Docs/a.txt")], lib, lib, "/a", "/b")
        self.sink(Stats(1, 1, 0, 0, 0.0, None, 0, 0))
        return plan, summary("dryrun", self._dir("dryrun"), copied=1)

    async def migrate(self, resume_run_dir: Path | None = None, force: bool = False) -> RunSummary:
        self.calls.append(("migrate", {"resume_run_dir": resume_run_dir, "force": force}))
        run_dir = Path(resume_run_dir) if resume_run_dir else self._dir("migrate")
        if FakeJob.raise_blocking and not force:
            raise BlockingFindingsError(list(FakeJob.findings), run_dir)
        self.sink(Stats(2, 2, 30, 30, 1.5, 0.0, 0, 0))
        return summary("migrate", run_dir, **{"copied": 2, "bytes_copied": 30, **FakeJob.migrate_kw})

    async def lock(self, apply: bool = False, include_owners: bool = False) -> list[LockChange]:
        self.calls.append(("lock", {"apply": apply, "include_owners": include_owners}))
        self._dir("lock")
        return [LockChange(4, "Members", ["Edit"]),
                LockChange(5, "Visitors", ["Read"], skipped_reason="cannot write (left untouched)")]

    def request_stop(self) -> None:
        self.calls.append(("stop", {}))

    async def aclose(self) -> None:
        pass


@pytest.fixture
def fake_job(monkeypatch: pytest.MonkeyPatch) -> type[FakeJob]:
    FakeJob.findings, FakeJob.raise_blocking, FakeJob.migrate_kw, FakeJob.last = [], False, {}, None
    monkeypatch.setattr(cli, "Job", FakeJob)
    monkeypatch.setattr(cli, "make_gateway_factory", lambda cfg, sink: "fake-factory")
    write_cfg()
    return FakeJob


WARN = Finding(Severity.WARN, "checked_out", "Klanten/a.docx", "checked out to Jan", "check in")
BLOCK = Finding(Severity.BLOCKING, "write_access", "Gedeelde documenten", "403 Forbidden", "grant edit")


# ------------------------------------------------------------------------------------------------ tenant/profile


def test_tenant_add_writes_config_and_lists():
    res = runner.invoke(cli.app, ["tenant", "add", "contoso", "--tenant-id", "contoso.onmicrosoft.com",
                                  "--client-id", "abcd1234", "--sharepoint-host", "https://Contoso.sharepoint.com/"],
                        env=ENV)
    assert res.exit_code == 0, res.output
    cfg = load_config()
    assert cfg.tenants["contoso"].sharepoint_host == "contoso.sharepoint.com"
    assert cfg.tenants["contoso"].client_id == "abcd1234"
    res = runner.invoke(cli.app, ["tenant", "list"], env=ENV)
    assert res.exit_code == 0 and "contoso.onmicrosoft.com" in res.output
    res = runner.invoke(cli.app, ["tenant", "add", "fabrikam", "--tenant-id", "f.onmicrosoft.com", "--client-id",
                                  "x", "--host", "fabrikam.sharepoint.com"], env=ENV)
    assert res.exit_code == 0 and set(load_config().tenants) == {"contoso", "fabrikam"}


def test_tenant_setup_help():
    res = runner.invoke(cli.app, ["tenant", "setup-help", "contoso.onmicrosoft.com"], env=ENV)
    assert res.exit_code == 0 and "AllSites.FullControl" in res.output


def test_profile_list_and_show():
    write_cfg()
    res = runner.invoke(cli.app, ["profile", "list"], env=ENV)
    assert res.exit_code == 0 and "svc" in res.output
    res = runner.invoke(cli.app, ["profile", "show", "svc"], env=ENV)
    assert res.exit_code == 0 and "0. Archief/**" in res.output and "Archief" in res.output
    res = runner.invoke(cli.app, ["profile", "show", "nope"], env=ENV)
    assert res.exit_code == 1 and "Unknown profile" in res.output


# ------------------------------------------------------------------------------------------------ spmig (fake Job)


def test_audit_exit_codes(fake_job):
    res = runner.invoke(cli.app, ["spmig", "audit", "--profile", "svc"], env=ENV)
    assert res.exit_code == 0, res.output
    assert "No findings" in res.output and "report.html" in res.output

    fake_job.findings = [WARN]
    res = runner.invoke(cli.app, ["spmig", "audit", "--profile", "svc"], env=ENV)
    assert res.exit_code == 0, res.output
    assert "checked_out" in res.output and "WARN" in res.output

    fake_job.findings = [WARN, BLOCK]
    res = runner.invoke(cli.app, ["spmig", "audit", "--profile", "svc"], env=ENV)
    assert res.exit_code == 1, res.output
    out = res.output
    assert "BLOCKING" in out and out.index("BLOCKING") < out.index("WARN")  # grouped, blocking first

    fake_job.raise_blocking = True
    res = runner.invoke(cli.app, ["spmig", "audit", "--profile", "svc"], env=ENV)
    assert res.exit_code == 2, res.output


def test_migrate_blocking_exit_2_and_force(fake_job):
    fake_job.findings, fake_job.raise_blocking = [BLOCK], True
    res = runner.invoke(cli.app, ["spmig", "migrate", "--profile", "svc"], env=ENV)
    assert res.exit_code == 2, res.output
    assert "write_access" in res.output and "--force" in res.output
    res = runner.invoke(cli.app, ["spmig", "migrate", "--profile", "svc", "--force"], env=ENV)
    assert res.exit_code == 0, res.output
    assert fake_job.last.calls[-1] == ("migrate", {"resume_run_dir": None, "force": True})


def test_migrate_json(fake_job):
    res = runner.invoke(cli.app, ["spmig", "migrate", "--profile", "svc", "--json", "--quiet"], env=ENV)
    assert res.exit_code == 0, res.output
    data = json.loads(res.stdout)
    assert data["mode"] == "migrate" and data["copied"] == 2 and data["bytes_copied"] == 30
    assert data["report"].endswith("report.html") and Path(data["run_dir"]).is_dir()
    # the run folder records the options, so --resume can find them later
    assert json.loads((Path(data["run_dir"]) / "options.json").read_text())["profile"] == "svc"


def test_migrate_failures_and_stopped_early(fake_job):
    fake_job.migrate_kw = {"failed": 1}
    assert runner.invoke(cli.app, ["spmig", "migrate", "-p", "svc"], env=ENV).exit_code == 1
    fake_job.migrate_kw = {"stopped_early": True}
    res = runner.invoke(cli.app, ["spmig", "migrate", "-p", "svc"], env=ENV)
    assert res.exit_code == 130 and "--resume" in res.output


def test_migrate_resume_loads_options_from_run_dir(fake_job):
    res = runner.invoke(cli.app, ["spmig", "migrate", "-p", "svc", "--json"], env=ENV)
    run_dir = json.loads(res.stdout)["run_dir"]
    res = runner.invoke(cli.app, ["spmig", "migrate", "--resume", run_dir], env=ENV)
    assert res.exit_code == 0, res.output
    job = fake_job.last
    assert job.calls[-1] == ("migrate", {"resume_run_dir": Path(run_dir), "force": False})
    assert job.opts.profile == "svc" and job.opts.excludes == ["0. Archief/**"]


def test_explicit_endpoints_and_flags(fake_job):
    res = runner.invoke(cli.app, [
        "spmig", "dryrun", "--src-site", SRC + "/", "--src-library", "Docs", "--src-folder", "Klanten",
        "--dst-site", DST, "--exclude", "*.tmp", "-x", "~$*", "--overwrite", "always", "--no-metadata",
        "--workers", "4", "--json",
    ], env=ENV)
    assert res.exit_code == 0, res.output
    o = fake_job.last.opts
    assert (o.source.tenant, o.source.site_url, o.source.library, o.source.folder) == ("contoso", SRC, "Docs",
                                                                                       "Klanten")
    assert (o.destination.tenant, o.destination.site_url, o.destination.library) == ("contoso", DST, None)
    assert o.excludes == ["*.tmp", "~$*"] and o.overwrite == "always" and not o.restore_metadata
    assert o.workers == 4
    assert json.loads(res.stdout)["plan"]["copy"] == 1


def test_missing_endpoints_is_an_error(fake_job):
    res = runner.invoke(cli.app, ["spmig", "audit"], env=ENV)
    assert res.exit_code == 1 and "--profile" in res.output


def test_lock_previews_without_apply(fake_job):
    res = runner.invoke(cli.app, ["spmig", "lock", "--profile", "svc"], env=ENV)
    assert res.exit_code == 0, res.output
    assert fake_job.last.calls == [("lock", {"apply": False, "include_owners": False})]
    assert "Preview only" in res.output and "Members" in res.output
    res = runner.invoke(cli.app, ["spmig", "lock", "--src-site", SRC, "--apply", "--include-owners", "--json"],
                        env=ENV)
    assert res.exit_code == 0, res.output
    assert fake_job.last.calls == [("lock", {"apply": True, "include_owners": True})]
    data = json.loads(res.stdout)
    assert data["applied"] is True and [c["will_change"] for c in data["changes"]] == [True, False]


def test_sites_search(monkeypatch):
    write_cfg()

    async def fake_search(self, tenant, query):
        return [SiteHit("Service", f"https://contoso.sharepoint.com/sites/{query}")]

    monkeypatch.setattr(cli.TenantClients, "search_sites", fake_search)
    res = runner.invoke(cli.app, ["spmig", "sites", "contoso", "svc", "--json"], env=ENV)
    assert res.exit_code == 0, res.output
    assert json.loads(res.stdout)["sites"][0]["url"].endswith("/sites/svc")


# ------------------------------------------------------------------------------------------------ real Job


def test_audit_and_migrate_end_to_end_with_fake_gateways(monkeypatch):
    from .test_job import DST_URL, SRC_URL, setup

    src, dst = setup()
    monkeypatch.setattr(cli, "make_gateway_factory", lambda cfg, sink: factory_for({SRC_URL: src, DST_URL: dst}))
    cfg = write_cfg()
    assert cfg.profiles["svc"].source.site_url == SRC_URL

    res = runner.invoke(cli.app, ["spmig", "audit", "-p", "svc"], env=ENV)
    assert res.exit_code == 0, res.output
    assert "gone.user@contoso.com" in res.output  # unresolvable user -> warning row

    res = runner.invoke(cli.app, ["spmig", "migrate", "-p", "svc", "--json"], env=ENV)
    assert res.exit_code == 0, res.output
    data = json.loads(res.stdout)
    assert (data["copied"], data["failed"], data["metadata_warnings"]) == (3, 0, 1)
    assert Path(data["report"]).exists()
    assert dst.file("Klanten/a.pdf").content == b"aaa"
    assert dst.file("2019/old.pdf", "Archief").content == b"old"

    res = runner.invoke(cli.app, ["spmig", "verify", "-p", "svc", "--json"], env=ENV)
    assert res.exit_code == 0, res.output
    assert json.loads(res.stdout)["counts"].get("missing", 0) == 0

    dst.write_denied = True  # blocking finding -> migrate refuses with exit 2
    res = runner.invoke(cli.app, ["spmig", "migrate", "-p", "svc"], env=ENV)
    assert res.exit_code == 2, res.output


def test_ctrl_c_before_transfer_cancels_with_130(fake_job, monkeypatch):
    import asyncio
    import os
    import signal

    async def slow_audit(self):
        os.kill(os.getpid(), signal.SIGINT)  # handled by the CLI's handler (main thread), not KeyboardInterrupt
        await asyncio.sleep(5)

    monkeypatch.setattr(FakeJob, "audit", slow_audit)
    res = runner.invoke(cli.app, ["spmig", "audit", "-p", "svc"], env=ENV)
    assert res.exit_code == 130, res.output
    assert ("stop", {}) in fake_job.last.calls and "interrupted" in res.output
    assert signal.getsignal(signal.SIGINT) is signal.default_int_handler  # handler restored


def test_resume_of_missing_run_dir_is_refused(fake_job, tmp_path):
    res = runner.invoke(cli.app, ["spmig", "migrate", "--resume", str(tmp_path / "nope")], env=ENV)
    assert res.exit_code == 1 and "does not exist" in res.output
    assert fake_job.last is None


def test_site_override_infers_tenant_from_host(fake_job):
    cfg = load_config()
    add_tenant(cfg, "fabrikam", tenant_id="fabrikam.onmicrosoft.com", client_id="cid2",
               sharepoint_host="fabrikam.sharepoint.com")
    save_config(cfg)
    other = "https://fabrikam.sharepoint.com/sites/new"
    res = runner.invoke(cli.app, ["spmig", "dryrun", "-p", "svc", "--dst-site", other], env=ENV)
    assert res.exit_code == 0, res.output
    o = fake_job.last.opts
    assert (o.destination.tenant, o.destination.site_url) == ("fabrikam", other)
    assert o.source.tenant == "contoso"
