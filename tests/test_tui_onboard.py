"""Pilot tests for the onboard screen (third tool): no-snapshot path, onboard, offboard and check."""

from __future__ import annotations

from textual.widgets import DataTable

from m365tools.tui.home import HomeScreen
from m365tools.tui.modals import PickerModal, PromptModal, TypedConfirmModal
from m365tools.tui.onboard import OnboardScreen

from .test_tui import rendered, settle
from .test_tui_groups import Env, make_app


def text(app, sel: str) -> str:
    return rendered(app.screen.query_one(sel).content)


async def test_home_opens_onboard_and_g_returns(tmp_path):
    app = make_app(Env(tmp_path))
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        await pilot.press("3")
        await pilot.pause()
        assert isinstance(app.screen, OnboardScreen)
        assert "▸ onboard" in rendered(app.screen.query_one("#header").render())
        await pilot.press("g")
        await pilot.pause()
        assert isinstance(app.screen, HomeScreen)


async def test_no_snapshot_offers_inventory(tmp_path):
    env = Env(tmp_path)
    app = make_app(env, start="onboard")
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        screen = app.screen
        assert "No snapshot yet. Press i to run an inventory." in text(app, "#gstate")
        await pilot.press("o")
        await settle(app, pilot)
        assert not isinstance(app.screen, PromptModal)
        assert any("press i" in n.message for n in app._notifications)
        await pilot.press("i")
        await settle(app, pilot)
        await settle(app, pilot)
        assert screen.snap is not None and len(screen.snap.people) == 12
        assert "Bring people in and out" in rendered(screen.query_one("#ostart").content)
        assert not env.gw.writes()


async def onboarded(app, pilot) -> None:
    await pilot.press("i")
    await settle(app, pilot)
    await settle(app, pilot)


async def test_onboard_preview_then_typed_apply(tmp_path):
    env = Env(tmp_path)
    app = make_app(env, start="onboard")
    async with app.run_test(size=(140, 45)) as pilot:
        await onboarded(app, pilot)
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
        assert isinstance(screen, OnboardScreen) and screen.view == "preview"
        assert screen.query_one("#ochanges", DataTable).row_count == 3  # Dana's Finance bundle
        assert "Payroll" in text(app, "#omanual")
        assert 'bundles = ["Finance"]' in text(app, "#oentry")
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

        await pilot.press("A", *"apply", "enter")
        await settle(app, pilot)
        assert {c[1] for c in env.gw.writes() if c[0] == "add_member"} == {
            "SG-Doc-Finance", "SG-Doc-Hub", "SG-Lic-E3"}
        assert "jan@contoso.com" in env.registry.read_text()
        assert "applied" in text(app, "#ohead")
        assert screen.query_one("#oruns").option_count >= 2  # preview and apply runs


async def test_offboard_needs_the_word_offboard(tmp_path):
    env = Env(tmp_path)
    app = make_app(env, start="onboard")
    async with app.run_test(size=(140, 45)) as pilot:
        await onboarded(app, pilot)
        await pilot.press("x")
        await pilot.pause()
        assert isinstance(app.screen, PickerModal)
        await pilot.press(*"anna", "enter")
        await settle(app, pilot)
        screen = app.screen
        assert screen.view == "preview" and screen.query_one("#ochanges", DataTable).row_count == 4
        assert "offboarded" in text(app, "#oentry")
        await pilot.press("A")
        await pilot.pause()
        assert isinstance(app.screen, TypedConfirmModal) and app.screen.word == "offboard"
        await pilot.press(*"offboard", "enter")
        await settle(app, pilot)
        assert sum(1 for c in env.gw.writes() if c[0] == "remove_member") == 4


async def test_check_shows_findings_for_one_person(tmp_path):
    app = make_app(Env(tmp_path), start="onboard")
    async with app.run_test(size=(140, 45)) as pilot:
        await onboarded(app, pilot)
        await pilot.press("c")
        await pilot.pause()
        await pilot.press(*"frank@contoso.com", "enter")
        await settle(app, pilot)
        screen = app.screen
        assert screen.view == "check"
        assert screen.query_one("#ocheck-table", DataTable).row_count == 1
        assert "missing" in rendered(screen.query_one("#ocheck-table", DataTable).get_row_at(0)[1])
        await pilot.press("escape")
        assert screen.view == "start"


# ------------------------------------------------------------------------------------------------ esc, quit, switch


async def esc_until_home(app, pilot, limit: int = 8) -> int:
    for n in range(1, limit + 1):
        await pilot.press("escape")
        await settle(app, pilot)
        if isinstance(app.screen, HomeScreen):
            return n
    raise AssertionError(f"esc did not reach home, stuck on {app.screen!r}")


async def test_esc_walks_every_screen_back_to_home(tmp_path):
    from m365tools.tui.groups import GroupsScreen
    from m365tools.tui.workspace import Workspace

    env = Env(tmp_path)
    app = make_app(env)
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        # spmig: a section's main pane, then a modal on top
        await pilot.press("1", "2", "tab", "question_mark")
        await pilot.pause()
        from m365tools.tui.modals import TextModal

        assert isinstance(app.screen, TextModal)
        await esc_until_home(app, pilot)
        # groups: inventory, filter, detail, then drift; and a jump from drift into a detail
        await pilot.press("2", "i")
        await settle(app, pilot)
        await settle(app, pilot)
        assert isinstance(app.screen, GroupsScreen)
        await pilot.press("slash", *"anna", "enter", "enter")
        await pilot.pause()
        await esc_until_home(app, pilot)
        await pilot.press("2", "d")
        await settle(app, pilot)
        await pilot.press("enter")  # jump to the first finding
        await pilot.pause()
        assert await esc_until_home(app, pilot) >= 3  # detail -> drift -> list -> home
        # groups: a preview and an open picker / prompt
        await pilot.press("2", "y")
        await settle(app, pilot)
        await pilot.press("e")  # no group selected on People: just a toast
        await esc_until_home(app, pilot)
        # onboard: a picker in the middle of the flow, then a preview
        await pilot.press("3", "o")
        await pilot.pause()
        await pilot.press(*"jan@contoso.com", "enter")
        await pilot.pause()
        assert isinstance(app.screen, PickerModal)
        await esc_until_home(app, pilot)
        await pilot.press("3", "x")
        await pilot.pause()
        await pilot.press(*"anna", "enter")
        await settle(app, pilot)
        assert app.screen.view == "preview"
        await esc_until_home(app, pilot)
        # ctrl+number jumps between tools
        await pilot.press("2")
        await pilot.pause()
        await pilot.press("ctrl+1")
        await pilot.pause()
        assert isinstance(app.screen, Workspace)
        await pilot.press("ctrl+3")
        await pilot.pause()
        assert isinstance(app.screen, OnboardScreen)
        await pilot.press("2")  # digits switch tools on the onboard screen
        await pilot.pause()
        assert isinstance(app.screen, GroupsScreen)


async def test_quit_always_asks(tmp_path):
    from m365tools.tui.modals import ConfirmModal

    app = make_app(Env(tmp_path), start="groups")
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.press("q")
        await pilot.pause()
        assert isinstance(app.screen, ConfirmModal) and "Quit m365 toolbox?" in app.screen.dialog_title
        await pilot.press("escape")
        await pilot.pause()
        assert not isinstance(app.screen, ConfirmModal) and app.is_running
        app.busy = "drift"
        await pilot.press("q")
        await pilot.pause()
        assert "drift is running" in str(app.screen.body)
        await pilot.press("n")
        await pilot.pause()
        app.busy = None
        await pilot.press("q")
        await pilot.pause()
        await pilot.press("y")
        await pilot.pause()
        assert app.return_code is not None or not app.is_running
