"""Headless TUI tests (Textual pilot). The real Job runs against the in-memory FakeGateway; site search and
library listing go through a fake TenantClients."""

from __future__ import annotations

import asyncio
from datetime import datetime
from pathlib import Path
from typing import Any, ClassVar

from textual.command import CommandPalette
from textual.widgets import ContentSwitcher, DataTable, OptionList, ProgressBar

from m365tools.core.config import AppConfig, ProfileConfig, TenantConfig
from m365tools.core.events import PhaseChanged, Stats
from m365tools.core.sharepoint import SiteHit
from m365tools.tools.spmig.job import Job
from m365tools.tools.spmig.models import RunSummary
from m365tools.tui.app import M365App
from m365tools.tui.modals import ConfirmModal, PickerModal, PromptModal, TextModal, TypedConfirmModal
from m365tools.tui.widgets import Sidebar

from .fake_gateway import factory_for
from .test_job import DST_URL, SRC_URL, setup

SITES = [SiteHit("Service", "https://contoso.sharepoint.com/sites/Service"),
         SiteHit("Servicedesk", "https://contoso.sharepoint.com/sites/Servicedesk"),
         SiteHit("Finance", "https://contoso.sharepoint.com/sites/Finance")]


def make_cfg() -> AppConfig:
    return AppConfig(
        tenants={"t": TenantConfig(tenant_id="contoso.onmicrosoft.com", client_id="cid",
                                   sharepoint_host="contoso.sharepoint.com")},
        profiles={"svc": ProfileConfig.model_validate({
            "source": {"tenant": "t", "site_url": SRC_URL, "library": "Servicemanagement"},
            "destination": {"tenant": "t", "site_url": DST_URL},
            "excludes": ["0. Archief/**"],
            "routes": [{"prefix": "0. Archief",
                        "destination": {"tenant": "t", "site_url": DST_URL, "library": "Archief"}}],
        })},
    )


class FakeClients:
    def __init__(self) -> None:
        self.queries: list[tuple[str, str]] = []

    async def search_sites(self, tenant: str, query: str) -> list[SiteHit]:
        self.queries.append((tenant, query))
        return [h for h in SITES if query.lower() in h.title.lower()]

    async def list_libraries(self, tenant: str, site_url: str) -> list[str]:
        return ["Documents", "Archief"]

    async def login(self, tenant: str) -> list[str]:
        return ["admin@contoso.com"]

    async def aclose(self) -> None:
        pass


def make_app(src: Any = None, dst: Any = None, clients: FakeClients | None = None,
             job_cls: Any = None) -> M365App:
    if src is None:
        src, dst = setup()
    gateways = factory_for({SRC_URL: src, DST_URL: dst})

    def real_job(_gateways: Any, opts: Any, sink: Any, run_dir_factory: Any) -> Job:
        return Job(gateways, opts, sink, run_dir_factory, migrator_kwargs={"retry_pause": 0.0})

    return M365App(make_cfg(), job_cls=job_cls or real_job, gateway_factory=lambda cfg, sink: None,
                   clients=clients or FakeClients(), start="spmig")  # type: ignore[arg-type]


async def settle(app: M365App, pilot: Any) -> None:
    await pilot.pause()
    await app.workers.wait_for_complete()
    await pilot.pause()


def rendered(renderable: Any) -> str:
    from rich.console import Console

    console = Console(width=140, record=True, file=open("/dev/null", "w"))  # noqa: SIM115
    console.print(renderable)
    return console.export_text()


def sidebar_text(app: M365App, index: int) -> str:
    return str(app.screen.query_one(Sidebar).get_option_at_index(index).prompt)


# ------------------------------------------------------------------------------------------------ layout


async def test_boot_renders_seven_sections_and_number_keys_switch():
    app = make_app()
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        sidebar = app.screen.query_one(Sidebar)
        assert sidebar.option_count == 7
        titles = [sidebar_text(app, i) for i in range(7)]
        for title in ("Job", "Checks", "Plan", "Run", "Verify", "Lock", "Runs"):
            assert any(title in t for t in titles)
        assert all("○" in t for t in titles[1:6])
        switcher = app.screen.query_one(ContentSwitcher)
        assert switcher.current == "job"
        for key, section in (("2", "checks"), ("4", "run"), ("7", "runs"), ("1", "job")):
            await pilot.press(key)
            await pilot.pause()
            assert switcher.current == section
        # sidebar j/k moves the selection, tab enters the main pane, esc goes back
        await pilot.press("j")
        await pilot.pause()
        assert switcher.current == "checks"
        await pilot.press("tab")
        assert app.screen.focused is not None and app.screen.focused.id == "findings"
        await pilot.press("escape")
        assert isinstance(app.screen.focused, Sidebar)


async def test_help_and_palette():
    app = make_app()
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.press("question_mark")
        await pilot.pause()
        assert isinstance(app.screen, TextModal)
        await pilot.press("escape")
        await pilot.pause()
        await pilot.press("colon")
        await pilot.pause()
        assert isinstance(app.screen, CommandPalette)
        await pilot.press(*"load profile svc")
        await pilot.pause(0.3)
        await pilot.press("enter")
        await settle(app, pilot)
        assert app.opts.profile == "svc" and app.opts.source.library == "Servicemanagement"


# ------------------------------------------------------------------------------------------------ job


async def test_edit_source_site_search_filters_and_sets_endpoint(monkeypatch):
    monkeypatch.setattr(PickerModal, "DEBOUNCE_S", 0.01)
    clients = FakeClients()
    app = make_app(clients=clients)
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.press("tab")  # into the Job cards (source focused)
        await pilot.press("e")  # one tenant configured: straight to the site search
        await pilot.pause()
        assert isinstance(app.screen, PickerModal)
        await pilot.press(*"serv")
        await pilot.pause(0.1)
        await pilot.pause(0.2)
        options = app.screen.query_one(OptionList)
        assert options.option_count == 2  # Service, Servicedesk; not Finance
        assert clients.queries[-1] == ("t", "serv")
        await pilot.press("down", "enter")  # pick Servicedesk
        await pilot.pause(0.2)
        assert isinstance(app.screen, PickerModal)  # library picker, default library first
        assert "(default library)" in str(app.screen.query_one(OptionList).get_option_at_index(0).prompt)
        assert app.screen.query_one(OptionList).option_count == 3
        await pilot.press("enter")
        await pilot.pause(0.2)
        assert isinstance(app.screen, PromptModal)  # folder
        await pilot.press(*"Klanten", "enter")
        await pilot.pause(0.2)
        src = app.opts.source
        assert (src.tenant, src.site_url, src.library, src.folder) == (
            "t", "https://contoso.sharepoint.com/sites/Servicedesk", None, "Klanten")


# ------------------------------------------------------------------------------------------------ audit / migrate


async def test_audit_shows_findings_and_blocking_count_then_migrate_demands_force():
    src, dst = setup()
    dst.write_denied = True  # -> BLOCKING write_access finding
    app = make_app(src, dst)
    async with app.run_test(size=(140, 45)) as pilot:
        app.load_profile("svc")
        await pilot.press("a")
        await settle(app, pilot)
        checks = app.screen.query_one("#checks")
        table = app.screen.query_one("#findings", DataTable)
        assert table.row_count == len(checks.findings) > 0
        assert str(table.get_row_at(0)[0]) == "● blocking"
        assert "1 blocking" in sidebar_text(app, 1) and "✗" in sidebar_text(app, 1)
        # enter opens the detail panel, esc closes it
        assert app.screen.focused is table  # a moved focus to the results
        await pilot.press("enter")
        assert app.screen.query_one("#finding-detail").display
        await pilot.press("escape")
        assert not app.screen.query_one("#finding-detail").display

        await pilot.press("m")
        await pilot.pause()
        assert isinstance(app.screen, TypedConfirmModal) and app.screen.word == "force"
        await pilot.press(*"forc", "enter")
        await pilot.pause()
        assert isinstance(app.screen, TypedConfirmModal)  # wrong word: still asking
        await pilot.press("escape")
        await settle(app, pilot)
        assert not dst.write_calls()


async def test_full_flow_audit_dryrun_migrate_verify_lock():
    src, dst = setup()
    app = make_app(src, dst)
    async with app.run_test(size=(140, 45)) as pilot:
        app.load_profile("svc")
        await pilot.press("a")
        await settle(app, pilot)
        assert app.last_audit is not None and not app.last_audit.blocking
        assert "files" in sidebar_text(app, 0)  # inventory summary on the Job entry
        await pilot.press("d")
        await settle(app, pilot)
        assert app.screen.query_one("#plan-table", DataTable).row_count > 0
        assert "copy" in sidebar_text(app, 2)
        await pilot.press("f")  # action filter: create_folder
        assert app.screen.query_one("#plan").kind_filter == "create_folder"
        assert not dst.write_calls()

        await pilot.press("m")
        await pilot.pause()
        assert isinstance(app.screen, ConfirmModal)
        await pilot.press("y")
        await settle(app, pilot)
        await settle(app, pilot)
        run = app.screen.query_one("#run")
        assert run.result is not None and run.result.copied == 3 and run.result.failed == 0
        assert dst.file("Klanten/a.pdf").content == b"aaa"
        assert "●" in sidebar_text(app, 3)

        await pilot.press("v")
        await settle(app, pilot)
        assert app.screen.query_one("#verify").result is not None

        await pilot.press("6", "l")
        await settle(app, pilot)
        assert app.screen.query_one("#lock-table", DataTable).row_count == 1
        assert src.reader_set == []  # preview only
        await pilot.press("A")
        await pilot.pause()
        assert isinstance(app.screen, TypedConfirmModal) and app.screen.word == "lock"
        await pilot.press(*"lock", "enter")
        await settle(app, pilot)
        assert src.reader_set == [4]

        await pilot.press("7")
        await pilot.pause()
        assert app.screen.query_one("#runs-table", DataTable).row_count >= 5


async def test_worker_error_becomes_toast_and_log_line():
    src, dst = setup()
    src.libraries.clear()  # source library missing -> NotFoundError from the Job
    app = make_app(src, dst)
    async with app.run_test(size=(140, 45)) as pilot:
        app.load_profile("svc")
        await pilot.press("a")
        await settle(app, pilot)
        assert app.busy is None
        assert any(n.severity == "error" for n in app._notifications)
        assert "✗" in sidebar_text(app, 1)


# ------------------------------------------------------------------------------------------------ dashboard


class PausableJob:
    """migrate() runs 'files' until request_stop(), then reports stopped_early; the resumed leg completes."""

    instances: ClassVar[list[PausableJob]] = []

    def __init__(self, _gateways: Any, opts: Any, sink: Any, _run_dir_factory: Any = None) -> None:
        self.opts, self.sink, self.run_dir = opts, sink, None
        self.calls: list[str] = []
        self.stop = asyncio.Event()
        PausableJob.instances.append(self)

    async def migrate(self, resume_run_dir: Path | None = None, force: bool = False) -> RunSummary:
        self.calls.append(f"migrate:{resume_run_dir}")
        run_dir = Path(resume_run_dir or "/tmp/fake-migrate")
        now = datetime.now()
        self.sink(PhaseChanged("files", "3 file(s)"))
        if resume_run_dir is None:
            self.sink(Stats(1, 3, 10, 30, 1.5, 20.0, 0, 1, copied=1))
            await self.stop.wait()
            return RunSummary("migrate", now, now, run_dir, copied=1, stopped_early=True, bytes_copied=10)
        self.sink(Stats(3, 3, 30, 30, 1.0, 0.0, 0, 0, copied=2, skipped=1))
        return RunSummary("migrate", now, now, run_dir, copied=2, skipped=1, bytes_copied=20)

    def request_stop(self) -> None:
        self.calls.append("stop")
        self.stop.set()

    async def aclose(self) -> None:
        pass


async def test_run_dashboard_updates_from_stats_and_pause_resume():
    app = make_app(job_cls=PausableJob)
    async with app.run_test(size=(140, 45)) as pilot:
        app.load_profile("svc")
        await pilot.press("m")
        await pilot.pause()
        await pilot.press("y")
        await pilot.pause(0.1)
        run = app.screen.query_one("#run")
        assert run.running
        bar = app.screen.query_one("#files-bar", ProgressBar)
        assert (bar.progress, bar.total) == (1, 3)
        stats = rendered(app.screen.query_one("#run-stats").content)
        assert "1/3" in stats and "1.50 MB/s" in stats
        assert "33%" in sidebar_text(app, 3)
        job = PausableJob.instances[-1]
        await pilot.press("p")  # pause -> request_stop -> stopped_early
        await settle(app, pilot)
        assert run.paused and job.calls == ["migrate:None", "stop"]
        await pilot.press("p")  # resume in the same run folder
        await settle(app, pilot)
        assert job.calls[-1] == "migrate:/tmp/fake-migrate"
        assert (run.result.copied, run.result.skipped) == (3, 0)  # legs added up
        assert app.screen.query_one("#files-bar", ProgressBar).progress == 3


class GatedAuditJob:
    """audit() walks the first two phases, then waits until the test releases it."""

    instances: ClassVar[list[GatedAuditJob]] = []

    def __init__(self, _gateways: Any, opts: Any, sink: Any, _run_dir_factory: Any = None) -> None:
        self.opts, self.sink, self.run_dir = opts, sink, None
        self.gate = asyncio.Event()
        GatedAuditJob.instances.append(self)

    async def audit(self) -> RunSummary:
        self.sink(PhaseChanged("inventory.source", "source: /sites/src/Docs"))
        self.sink(Stats(102_149, 102_149, 2_400 * 1024**2, 0, 0.0, None, 0, 0))
        self.sink(PhaseChanged("permissions.source", "102,149 items"))
        self.sink(Stats(45_000, 102_149, 0, 0, 0.0, None, 0, 0))
        await self.gate.wait()
        for phase in ("inventory.destination", "users", "checks", "report"):
            self.sink(PhaseChanged(phase))
        now = datetime.now()
        return RunSummary("audit", now, now, Path("/tmp/fake-audit"), findings=[])

    async def aclose(self) -> None:
        pass


async def test_checks_step_list_shows_running_step_then_collapses():
    from m365tools.tui.widgets import StepList

    app = make_app(job_cls=GatedAuditJob)
    async with app.run_test(size=(140, 45)) as pilot:
        app.load_profile("svc")
        await pilot.press("a")
        await pilot.pause(0.1)
        steps = app.screen.query_one("#checks").query_one(StepList)
        assert steps.display and steps.running
        lines = rendered(steps.render()).splitlines()
        assert lines[0].startswith("●") and "inventory source" in lines[0] and "102,149 items" in lines[0]
        assert "· 2.34 GB" in lines[0]
        assert lines[1][0] in "◐◓◑◒" and "permissions scan" in lines[1] and "45,000/102,149" in lines[1]
        assert "━" in lines[1]  # progress bar on the running step
        assert lines[2].startswith("○") and "inventory destination" in lines[2]
        assert [ln.split()[1] for ln in lines[:6]] == ["inventory", "permissions", "inventory", "users", "checks",
                                                      "report"]
        GatedAuditJob.instances[-1].gate.set()
        await settle(app, pilot)
        assert steps.display and not steps.running and steps.collapsed
        summary = rendered(steps.render()).strip()
        assert summary.startswith("● audit took ") and "slowest" in summary
        assert all(st.state == "done" for st in steps.steps)
