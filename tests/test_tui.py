"""Headless TUI tests (Textual pilot) with a fake Job and fake site search."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import ClassVar

from textual.widgets import DataTable, Input, ListView

from m365tools.core.config import AppConfig, ProfileConfig, TenantConfig
from m365tools.core.events import LogLine
from m365tools.core.sharepoint import SiteHit
from m365tools.tools.spmig.models import Finding, RunSummary, Severity
from m365tools.tui.app import M365App
from m365tools.tui.screens.audit import AuditScreen
from m365tools.tui.screens.endpoint import EndpointScreen, SourceScreen
from m365tools.tui.screens.home import HomeScreen
from m365tools.tui.screens.options import OptionsScreen

SRC = "https://contoso.sharepoint.com/sites/src"
DST = "https://contoso.sharepoint.com/sites/dst"


def make_cfg() -> AppConfig:
    cfg = AppConfig(
        tenants={"contoso": TenantConfig(tenant_id="contoso.onmicrosoft.com", client_id="cid",
                                         sharepoint_host="contoso.sharepoint.com")},
        profiles={"svc": ProfileConfig.model_validate({
            "source": {"tenant": "contoso", "site_url": SRC, "library": "Docs"},
            "destination": {"tenant": "contoso", "site_url": DST},
            "excludes": ["0. Archief/**"],
        })},
    )
    return cfg


FINDINGS = [
    Finding(Severity.WARN, "checked_out", "a/b.docx", "checked out to Jan"),
    Finding(Severity.BLOCKING, "path_length", "x" * 50, "path too long", "shorten"),
    Finding(Severity.INFO, "versions", "", "only latest version copied"),
]


class FakeJob:
    instances: ClassVar[list[FakeJob]] = []

    def __init__(self, gateways, opts, sink, run_dir_factory=None) -> None:
        self.opts, self.sink, self.run_dir = opts, sink, None
        self.calls: list[str] = []
        FakeJob.instances.append(self)

    async def audit(self) -> RunSummary:
        self.calls.append("audit")
        self.run_dir = Path("/tmp/fake-run")
        self.sink(LogLine("info", "auditing"))
        now = datetime.now()
        return RunSummary("audit", now, now, self.run_dir, findings=list(FINDINGS))

    def request_stop(self) -> None:
        self.calls.append("stop")

    async def aclose(self) -> None:
        pass


class FakeClients:
    def __init__(self) -> None:
        self.queries: list[tuple[str, str]] = []

    async def search_sites(self, tenant: str, query: str) -> list[SiteHit]:
        self.queries.append((tenant, query))
        return [SiteHit("Service", f"https://contoso.sharepoint.com/sites/{query}-service"),
                SiteHit("Archive", f"https://contoso.sharepoint.com/sites/{query}-archive")]

    async def list_libraries(self, tenant: str, site_url: str) -> list[str]:
        return ["Documents", "Archief"]

    async def aclose(self) -> None:
        pass


def make_app(clients: FakeClients | None = None) -> M365App:
    return M365App(make_cfg(), job_cls=FakeJob, gateway_factory=lambda cfg, sink: object(),
                   clients=clients or FakeClients())  # type: ignore[arg-type]


async def test_boots_to_home_and_lists_profiles():
    app = make_app()
    async with app.run_test(size=(120, 40)) as pilot:
        await pilot.pause()
        assert isinstance(app.screen, HomeScreen)
        assert len(app.screen.query_one("#profiles", ListView).children) == 1


async def test_source_search_shows_hits(monkeypatch):
    monkeypatch.setattr(EndpointScreen, "DEBOUNCE_S", 0.01)
    clients = FakeClients()
    app = make_app(clients)
    async with app.run_test(size=(120, 40)) as pilot:
        await pilot.press("n")
        await pilot.pause()
        assert isinstance(app.screen, SourceScreen)
        app.screen.query_one("#search", Input).focus()
        await pilot.press(*"serv")
        await pilot.pause(0.2)
        await app.workers.wait_for_complete()
        await pilot.pause()
        hits = app.screen.query_one("#hits", ListView)
        assert len(hits.children) == 2
        assert clients.queries[-1] == ("contoso", "serv")
        # pick the first hit -> site URL filled + libraries listed
        hits.focus()
        hits.index = 0
        await pilot.press("enter")
        await pilot.pause()
        await app.workers.wait_for_complete()
        await pilot.pause()
        assert app.screen.query_one("#site", Input).value.endswith("/serv-service")
        assert app.screen.query_one("#libraries", DataTable).row_count == 3  # default + 2


async def test_profile_to_audit_shows_findings():
    FakeJob.instances.clear()
    app = make_app()
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        lv = app.screen.query_one("#profiles", ListView)
        lv.focus()
        lv.index = 0
        await pilot.press("enter")
        await pilot.pause()
        assert isinstance(app.screen, OptionsScreen)
        app.screen.action_audit()
        await pilot.pause()
        assert isinstance(app.screen, AuditScreen)
        await app.workers.wait_for_complete()
        await pilot.pause()
        table = app.screen.query_one("#findings", DataTable)
        assert table.row_count == 3
        first = table.get_row_at(0)
        assert str(first[0]) == "BLOCKING"
        # blocking findings gate the Migrate button until Force is ticked
        assert app.screen.query_one("#migrate").disabled
        await pilot.press("f")
        await pilot.pause()
        assert not app.screen.query_one("#migrate").disabled
        job = FakeJob.instances[-1]
        assert job.calls == ["audit"] and job.opts.excludes == ["0. Archief/**"]


async def test_full_flow_with_real_job_and_fake_gateways():
    """audit -> dry run -> migrate -> done -> verify -> lock preview, using the in-memory gateway."""
    from m365tools.tools.spmig.job import Job
    from m365tools.tui.screens.done import ReportDoneScreen
    from m365tools.tui.screens.lock import LockScreen
    from m365tools.tui.screens.migrate import MigrateScreen
    from m365tools.tui.screens.plan import PlanScreen
    from m365tools.tui.screens.verify import VerifyScreen

    from .fake_gateway import FakeGateway, factory_for
    from .test_job import setup

    src, dst = setup()
    src.add_file("Klanten/c.txt", b"cc")
    dst.add_user("gone.user@contoso.com")

    def job_cls(gateways, opts, sink, run_dir_factory):
        return Job(factory_for({SRC: src, DST: dst}), opts, sink, run_dir_factory,
                   migrator_kwargs={"retry_pause": 0.0})

    cfg = make_cfg()
    cfg.profiles["svc"].source.library = "Servicemanagement"
    app = M365App(cfg, job_cls=job_cls, gateway_factory=lambda c, s: None, clients=FakeClients())  # type: ignore
    assert isinstance(src, FakeGateway)

    async def settle() -> None:
        await pilot.pause()
        await app.workers.wait_for_complete()
        await pilot.pause()

    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        lv = app.screen.query_one("#profiles", ListView)
        lv.focus()
        lv.index = 0
        await pilot.press("enter")
        await pilot.pause()
        app.screen.action_audit()
        await settle()
        assert isinstance(app.screen, AuditScreen)
        assert app.state.audit is not None and not app.state.audit.blocking
        await pilot.press("d")
        await settle()
        assert isinstance(app.screen, PlanScreen)
        assert app.screen.query_one("#entries", DataTable).row_count > 0
        assert not dst.write_calls()
        await pilot.press("m")
        await settle()
        await settle()
        assert isinstance(app.screen, ReportDoneScreen)
        assert not any(isinstance(s, MigrateScreen) for s in app.screen_stack)  # switched, not stacked
        assert app.screen.summary.copied == 3 and app.screen.summary.failed == 0
        assert dst.file("Klanten/c.txt").content == b"cc"
        await pilot.press("v")
        await settle()
        assert isinstance(app.screen, VerifyScreen)
        assert app.screen.result is not None
        await pilot.press("escape")
        await pilot.pause()
        await pilot.press("l")
        await settle()
        assert isinstance(app.screen, LockScreen)
        assert app.screen.query_one("#changes", DataTable).row_count == 1
        assert src.reader_set == []  # preview only


class PausableJob(FakeJob):
    """migrate() runs 'files' until request_stop(), then reports stopped_early; the resumed run completes."""

    def __init__(self, *a, **kw) -> None:
        super().__init__(*a, **kw)
        import asyncio

        self.stop = asyncio.Event()

    async def migrate(self, resume_run_dir=None, force=False) -> RunSummary:
        from m365tools.core.events import PhaseChanged, Stats

        self.calls.append(f"migrate:{resume_run_dir}")
        run_dir = Path(resume_run_dir or "/tmp/fake-migrate")
        now = datetime.now()
        self.sink(PhaseChanged("files", "3 file(s)"))
        if resume_run_dir is None:
            self.sink(Stats(1, 3, 10, 30, 1.0, 2.0, 0, 0, copied=1))
            await self.stop.wait()
            return RunSummary("migrate", now, now, run_dir, copied=1, stopped_early=True, bytes_copied=10)
        self.sink(Stats(3, 3, 30, 30, 1.0, 0.0, 0, 0, copied=2, skipped=1))
        return RunSummary("migrate", now, now, run_dir, copied=2, skipped=1, bytes_copied=20)

    def request_stop(self) -> None:
        self.calls.append("stop")
        self.stop.set()


async def test_migrate_pause_and_resume():
    from m365tools.tui.screens.done import ReportDoneScreen
    from m365tools.tui.screens.migrate import MigrateScreen

    app = M365App(make_cfg(), job_cls=PausableJob, gateway_factory=lambda c, s: None,
                  clients=FakeClients())  # type: ignore[arg-type]
    async with app.run_test(size=(120, 40)) as pilot:
        await pilot.pause()
        job = app.new_job()
        app.push_screen(MigrateScreen())
        await pilot.pause(0.05)
        screen = app.screen
        assert isinstance(screen, MigrateScreen) and screen.running
        assert "1/3" in str(screen.query_one("#stats").render())
        await pilot.press("p")  # pause -> request_stop -> stopped_early
        await app.workers.wait_for_complete()
        await pilot.pause()
        assert screen.paused and job.calls == ["migrate:None", "stop"]
        await pilot.press("p")  # resume in the same run folder
        await app.workers.wait_for_complete()
        await pilot.pause()
        assert job.calls[-1] == "migrate:/tmp/fake-migrate"
        assert isinstance(app.screen, ReportDoneScreen)
        assert (app.screen.summary.copied, app.screen.summary.skipped) == (3, 0)  # legs added up
