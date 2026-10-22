"""Pilot tests for the toolbox home and the groups screen. The real GroupsJob runs against the in-memory
FakeGraphGateway (tests/fake_graph.py) with the sample registry."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from textual.command import CommandPalette
from textual.widgets import ContentSwitcher, DataTable, Static

from m365tools.tools.groups.job import GroupsJob
from m365tools.tui.app import M365App
from m365tools.tui.groups import GroupList, GroupsScreen
from m365tools.tui.home import HomeScreen, ToolList
from m365tools.tui.modals import PickerModal, PromptModal, TextModal, TypedConfirmModal
from m365tools.tui.workspace import Workspace

from .fake_graph import FakeGraphGateway, make_fake, write_registry
from .test_tui import FakeClients, make_cfg, rendered, settle


class Env:
    def __init__(self, tmp: Path) -> None:
        self.gw: FakeGraphGateway = make_fake()
        self.registry = write_registry(tmp / "groups" / "t.toml")
        self.jobs: list[GroupsJob] = []

    def factory(self, _cfg: Any, tenant: str | None, sink: Any) -> GroupsJob:
        job = GroupsJob(self.gw, self.registry, sink, tenant=tenant or "t")
        self.jobs.append(job)
        return job


def make_app(env: Env, start: str = "home") -> M365App:
    return M365App(make_cfg(), gateway_factory=lambda cfg, sink: None, clients=FakeClients(),  # type: ignore[arg-type]
                   groups_job_factory=env.factory, start=start)


def screen_text(widget: Any) -> str:
    return rendered(widget.render())


def detail_text(app: M365App) -> str:
    return rendered(app.screen.query_one("#detail-body", Static).content)


async def inventory(app: M365App, pilot: Any) -> None:
    await pilot.press("i")
    await settle(app, pilot)
    await settle(app, pilot)


# ------------------------------------------------------------------------------------------------ home


async def test_home_lists_tools_and_opens_both(tmp_path):
    env = Env(tmp_path)
    app = make_app(env)
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        assert isinstance(app.screen, HomeScreen)
        tools = app.screen.query_one(ToolList)
        assert tools.option_count == 2
        both = rendered(tools.get_option("spmig").prompt) + rendered(tools.get_option("groups").prompt)
        assert "SharePoint library migration" in both and "access groups, drift and onboarding" in both
        assert "no runs yet" in both
        assert "m365 toolbox" in screen_text(app.screen.query_one("#header"))

        await pilot.press("1")
        await pilot.pause()
        assert isinstance(app.screen, Workspace)
        assert "▸ spmig" in screen_text(app.screen.query_one("#header"))
        await pilot.press("g")
        await pilot.pause()
        assert isinstance(app.screen, HomeScreen)

        await pilot.press("j", "enter")  # highlight groups, open it
        await settle(app, pilot)
        assert isinstance(app.screen, GroupsScreen)
        assert "▸ groups" in screen_text(app.screen.query_one("#header"))
        await pilot.press("g")
        await pilot.pause()
        assert isinstance(app.screen, HomeScreen)
        await pilot.press("2")
        await pilot.pause()
        assert isinstance(app.screen, GroupsScreen)


async def test_g_refuses_to_leave_while_busy(tmp_path):
    app = make_app(Env(tmp_path), start="groups")
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        app.busy = "inventory"
        await pilot.press("g")
        await pilot.pause()
        assert isinstance(app.screen, GroupsScreen)
        assert any("inventory is running" in n.message for n in app._notifications)
        app.busy = None
        await pilot.press("g")
        await pilot.pause()
        assert isinstance(app.screen, HomeScreen)


# ------------------------------------------------------------------------------------------------ explorer


async def test_empty_state_then_inventory_fills_people(tmp_path):
    env = Env(tmp_path)
    app = make_app(env, start="groups")
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        screen = app.screen
        assert "No snapshot yet. Press i to run an inventory." in rendered(screen.query_one("#gstate").content)
        # registry people are listed before any inventory
        assert screen.query_one(GroupList).option_count == 8
        await inventory(app, pilot)
        assert screen.snap is not None and len(screen.snap.people) == 12
        assert screen.query_one(GroupList).option_count == 12
        state = rendered(screen.query_one("#gstate").content)
        assert "12 people" in state and "unexplained" in state and "\n" in state and "snapshot just now" in state
        assert not env.gw.writes()


async def test_filter_and_enter_shows_person(tmp_path):
    env = Env(tmp_path)
    app = make_app(env, start="groups")
    async with app.run_test(size=(140, 45)) as pilot:
        await inventory(app, pilot)
        await pilot.press("slash")
        await pilot.pause()
        assert isinstance(app.screen, PromptModal)
        await pilot.press(*"evelien")
        await pilot.pause()
        await pilot.press("enter")
        await pilot.pause()
        screen = app.screen
        assert isinstance(screen, GroupsScreen)
        assert screen.query_one(GroupList).option_count == 1
        await pilot.press("enter")
        await pilot.pause()
        assert screen.focused is screen.query_one("#detail")
        text = detail_text(app)
        for word in ("Evelien Mulder", "Servicedesk", "SG-Doc-Finance", "add-on", "Intranet hub", "ENTERPRISEPACK",
                     "Jira Cloud"):
            assert word in text, word
        await pilot.press("escape")
        assert screen.focused is screen.query_one(GroupList)
        await pilot.press("escape")  # clears the filter
        await pilot.pause()
        assert screen.query_one(GroupList).option_count == 12


async def test_groups_tab_and_edit_purpose_writes_registry(tmp_path):
    env = Env(tmp_path)
    app = make_app(env, start="groups")
    async with app.run_test(size=(140, 45)) as pilot:
        await inventory(app, pilot)
        await pilot.press("2")
        await pilot.pause()
        screen = app.screen
        assert screen.mode == "groups"
        await pilot.press("slash", *"lic-e5", "enter")
        await pilot.pause()
        assert screen.current_row.key == "SG-Lic-E5"
        assert "not set: press e" in detail_text(app)
        await pilot.press("e")
        await pilot.pause()
        await pilot.press(*"Microsoft 365 E5", "enter")
        await pilot.pause()
        await pilot.press("ctrl+u", *"ivo@contoso.com", "enter")
        await settle(app, pilot)
        assert "Microsoft 365 E5" in env.registry.read_text()
        assert "Microsoft 365 E5" in detail_text(app)


async def test_drift_lists_findings_and_enter_jumps(tmp_path):
    env = Env(tmp_path)
    app = make_app(env, start="groups")
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.press("d")  # no snapshot yet: drift reads one first
        await settle(app, pilot)
        await settle(app, pilot)
        screen = app.screen
        assert screen.query_one(ContentSwitcher).current == "drift"
        table = screen.query_one("#drift-table", DataTable)
        assert table.row_count == len(screen.findings) > 0
        assert screen.focused is table
        idx = next(i for i, f in enumerate(screen.findings) if f.path == "kim@contoso.com")
        table.move_cursor(row=idx)
        await pilot.pause()
        await pilot.press("enter")
        await pilot.pause()
        assert screen.mode == "people" and screen.view == "detail"
        assert screen.current_row.key == "kim@contoso.com"
        assert "Kim Vos" in detail_text(app) and "back to drift" in detail_text(app)
        await pilot.press("escape")
        await pilot.pause()
        assert screen.view == "drift" and screen.focused is table
        # a group finding jumps to the Groups tab
        idx = next(i for i, f in enumerate(screen.findings) if f.category == "no_owner")
        table.move_cursor(row=idx)
        await pilot.press("enter")
        await pilot.pause()
        assert screen.mode == "groups" and screen.current_row.key == "SG-Doc-Archive"


async def test_onboard_preview_then_typed_apply(tmp_path):
    env = Env(tmp_path)
    app = make_app(env, start="groups")
    async with app.run_test(size=(140, 45)) as pilot:
        await inventory(app, pilot)
        await pilot.press("o")
        await pilot.pause()
        assert isinstance(app.screen, PromptModal)
        await pilot.press(*"jan@contoso.com", "enter")
        await pilot.pause()
        assert isinstance(app.screen, PickerModal)
        await pilot.press(*"like dana", "enter")
        await pilot.pause()
        assert isinstance(app.screen, PickerModal)  # add-ons: first entry is "done"
        await pilot.press("enter")
        await settle(app, pilot)
        screen = app.screen
        assert isinstance(screen, GroupsScreen) and screen.view == "changes"
        table = screen.query_one("#changes-table", DataTable)
        assert table.row_count == 3  # Dana's Finance bundle
        notes = rendered(screen.query_one("#changes-notes").content)
        assert "like dana@contoso.com" in notes and "manual" in notes and "Payroll" in notes
        assert not env.gw.writes()

        await pilot.press("A")
        await pilot.pause()
        assert isinstance(app.screen, TypedConfirmModal) and app.screen.word == "apply"
        await pilot.press(*"aply", "enter")
        await pilot.pause()
        assert isinstance(app.screen, TypedConfirmModal)  # wrong word: still asking
        await pilot.press("escape")
        await settle(app, pilot)
        assert not env.gw.writes()

        await pilot.press("A")
        await pilot.pause()
        await pilot.press(*"apply", "enter")
        await settle(app, pilot)
        assert {c[1] for c in env.gw.writes() if c[0] == "add_member"} == {
            "SG-Doc-Finance", "SG-Doc-Hub", "SG-Lic-E3"}
        assert "jan@contoso.com" in env.registry.read_text()
        assert "done" in rendered(screen.query_one("#changes-head").content)
        assert screen.snap_stale


async def test_offboard_needs_the_word_offboard(tmp_path):
    env = Env(tmp_path)
    app = make_app(env, start="groups")
    async with app.run_test(size=(140, 45)) as pilot:
        await inventory(app, pilot)
        await pilot.press("slash", *"anna", "enter")
        await pilot.pause()
        await pilot.press("x")
        await settle(app, pilot)
        screen = app.screen
        assert screen.view == "changes" and screen.query_one("#changes-table", DataTable).row_count == 4
        await pilot.press("A")
        await pilot.pause()
        assert isinstance(app.screen, TypedConfirmModal) and app.screen.word == "offboard"
        await pilot.press(*"offboard", "enter")
        await settle(app, pilot)
        assert sum(1 for c in env.gw.writes() if c[0] == "remove_member") == 4


async def test_suggest_accept_writes_bundle_after_apply(tmp_path):
    env = Env(tmp_path)
    app = make_app(env, start="groups")
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.press("s")
        await settle(app, pilot)
        await settle(app, pilot)
        screen = app.screen
        assert screen.view == "suggest"
        table = screen.query_one("#suggest-table", DataTable)
        assert table.row_count == len(screen.suggestion.candidates) > 0
        sales = next(i for i, c in enumerate(screen.suggestion.candidates) if "gina@contoso.com" in c.people)
        table.move_cursor(row=sales)
        await pilot.press("a")
        await pilot.pause()
        assert isinstance(app.screen, PromptModal)
        await pilot.press("ctrl+u", *"Sales", "enter")
        await settle(app, pilot)
        assert screen.view == "changes" and screen.pending is not None
        assert "[bundles.Sales]" not in env.registry.read_text()
        await pilot.press("A", *"apply", "enter")
        await settle(app, pilot)
        assert "Sales" in env.jobs[-1].registry.bundles
        await pilot.press("escape")
        await pilot.pause()
        assert screen.view == "suggest"


async def test_sync_and_import_previews(tmp_path):
    env = Env(tmp_path)
    app = make_app(env, start="groups")
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.press("y")
        await settle(app, pilot)
        await settle(app, pilot)
        screen = app.screen
        assert screen.view == "changes"
        assert "SG-App-Payroll" in rendered(screen.query_one("#changes-table", DataTable).get_row_at(0)[1])
        await pilot.press("escape")
        await pilot.press("m")
        await settle(app, pilot)
        assert screen.view == "changes"
        assert screen.query_one("#changes-table", DataTable).row_count == 1  # SG-Doc-Archive is unregistered
        assert not env.gw.writes()


async def test_groups_help_and_palette(tmp_path):
    app = make_app(Env(tmp_path), start="groups")
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.press("question_mark")
        await pilot.pause()
        assert isinstance(app.screen, TextModal)
        await pilot.press("escape")
        await pilot.press("colon")
        await pilot.pause()
        assert isinstance(app.screen, CommandPalette)
        await pilot.press(*"show bundles")
        await pilot.pause(0.3)
        await pilot.press("enter")
        await settle(app, pilot)
        assert isinstance(app.screen, GroupsScreen) and app.screen.mode == "bundles"
        assert app.screen.query_one(GroupList).option_count == 2
        await pilot.press("enter")
        await pilot.pause()
        assert "Everyone on the service desk" in detail_text(app) or "Finance team" in detail_text(app)


def test_key_bar_keeps_essentials_when_narrow():
    from m365tools.tui.widgets import fit_keys

    local = [("enter", "open"), ("j/k", "move")]
    glob = [("g", "home"), ("/", "filter"), ("i", "inventory"), ("d", "drift"), ("s", "suggest"),
            ("o", "onboard"), (":", "commands"), ("?", "help"), ("q", "quit")]
    loc, glo = fit_keys(local, glob, 70)
    keys = [k for k, _d in loc + glo]
    assert {":", "?", "q"} <= set(keys) and "j/k" not in keys
