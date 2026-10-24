"""The onboard screen: a small guided flow to bring people in (onboard), take them out (offboard) and check one
person against the registry. Uses the same GroupsJob and snapshot as the groups screen.

    o  new person     x  offboard     c  check one person     i  inventory (when there is no snapshot)
    A  apply the preview (asks a typed word)     esc  back to the start     L  log     enter  report of a run
"""

from __future__ import annotations

import csv
from collections.abc import Callable
from pathlib import Path
from typing import Any, ClassVar

from rich.table import Table
from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.widgets import ContentSwitcher, OptionList, Static

from m365tools.core.events import GROUPS_PHASES
from m365tools.tools.groups import drift as drift_mod

from . import theme
from .groups import SEV_ORDER, GroupsOpsScreen, _member_rows
from .groups_view import FINDING_LABEL, change_status, group_rows, people_rows, state_ribbon
from .home import ago
from .modals import Choice, PickerModal, PromptModal, TextModal, TypedConfirmModal
from .sections.base import Empty
from .theme import ACCENT, BLUE, DIM, FG, GREEN, ON_FILL, ORANGE, RED, SEVERITY_STYLE, SUB, YELLOW
from .widgets import Activity, BarList, BarOption, HeaderBar, KeyBar, StepList, VimLog, VimTable
from .workspace import BACK_HELP

HELP = {
    "start": ("Onboarding gives a new colleague the same access groups as someone in the same role, or as a "
              "bundle from the registry. o starts it: type their sign-in name, pick 'like <colleague>' or a "
              "bundle, add extra groups if needed, then read the preview. Nothing is written until you press A "
              "and type apply. x offboards a leaver the same way; c checks one person against the registry. "
              "Licences, apps and roles are listed as manual work: do those in the admin center."),
    "preview": ("This is a preview: nothing has been written yet. 'planned' groups will get the person, 'already' "
                "means they are a member, 'not created' means the group is only in the registry (create it from "
                "the groups tool with y). The right box shows the registry entry that will be saved. A applies "
                "after you type the word; esc drops the preview."),
    "check": ("The check compares one person with the registry. 'missing' groups are expected by their bundle or "
              "add-ons but not given yet (onboard them again to fix it); 'unexplained' groups have no bundle or "
              "add-on that says why. Nothing is changed by looking."),
}


class RunList(BarList):
    BINDINGS: ClassVar[list[BindingType]] = [Binding("enter", "select", "report")]


def run_summary(run: Path) -> tuple[str, str, str]:
    """(mode, upn, state) of an onboard / offboard run folder, read from its CSV."""
    mode = "offboard" if "-offboard" in run.name else "onboard"
    upn, state = "?", "preview"
    try:
        with (run / f"{mode}.csv").open(encoding="utf-8-sig", newline="") as fh:
            for row in csv.DictReader(fh, delimiter=";"):
                detail, status = row.get("Detail", ""), row.get("Status", "")
                if "@" in detail and upn == "?":
                    upn = detail
                if status in ("done", "failed"):
                    state = "applied"
    except OSError:
        pass
    return mode, upn, state


class OnboardScreen(GroupsOpsScreen):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("tab,shift+tab", "toggle_pane", "switch pane", show=False),
        Binding("escape", "back", "back", show=False),
        Binding("o", "onboard", "new person"),
        Binding("x", "offboard", "offboard"),
        Binding("c", "check", "check"),
        Binding("i", "inventory", "inventory"),
        Binding("A", "apply", "apply"),
        Binding("L", "toggle_log", "log", show=False),
        Binding("colon", "app.command_palette", "commands", key_display=":"),
        Binding("question_mark", "help", "help", key_display="?"),
        Binding("ctrl+l", "app.login", "sign in", show=False),
        Binding("g", "app.go_home", "home", priority=True),
        Binding("ctrl+1", "app.tool(1)", "spmig", show=False, priority=True),
        Binding("ctrl+2", "app.tool(2)", "groups", show=False, priority=True),
        Binding("ctrl+3", "app.tool(3)", "onboard", show=False, priority=True),
        Binding("1", "app.tool(1)", "spmig", show=False),
        Binding("2", "app.tool(2)", "groups", show=False),
        Binding("3", "app.tool(3)", "onboard", show=False),
        Binding("q", "app.quit_app", "quit"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.view = "start"
        self.error = ""
        self.pending: tuple[str, str, Callable[[Any], Any], Callable[[Any], None]] | None = None
        self.findings: list[Any] = []
        self.runs: list[Path] = []
        self.log_open = False

    # -------------------------------------------------------------------------------------------- layout

    def compose(self) -> ComposeResult:
        yield HeaderBar("onboard", id="header")
        with Horizontal(id="body"):
            with Vertical(id="onav"):
                yield Static("", id="oactions")
                yield Static(Text("RECENT", style=f"bold {DIM}"), classes="onav-label")
                yield RunList(id="oruns")
                yield Static("", id="oruns-empty")
            with Vertical(id="omain"):
                yield Static("", id="gstate")
                with ContentSwitcher(id="oswitch", initial="ostart"):
                    yield Empty("", id="ostart")
                    with Vertical(id="opreview"):
                        yield Static("", id="ohead", classes="counts")
                        yield VimTable(id="ochanges")
                        with Horizontal(id="oboxes"):
                            manual = Static("", id="omanual", classes="obox")
                            manual.border_title = "by hand (admin center)"
                            yield manual
                            entry = Static("", id="oentry", classes="obox")
                            entry.border_title = "registry entry"
                            yield entry
                        yield Static("", id="onotes", classes="gnote")
                    with Vertical(id="ocheck"):
                        yield Static("", id="ocheck-head", classes="counts")
                        yield VimTable(id="ocheck-table")
                with Vertical(id="ops"):
                    yield StepList(GROUPS_PHASES, id="gsteps")
                    yield Activity(id="gactivity")
                    log = VimLog(id="glog")
                    log.border_title = "log"
                    yield log
        yield KeyBar(id="keybar")

    def on_mount(self) -> None:
        self.query_one("#ochanges", VimTable).add_fit_columns(("action", 13), ("group", 18), ("purpose", None),
                                                              ("status", 15))
        self.query_one("#ocheck-table", VimTable).add_fit_columns(("severity", 8), ("what", 12), ("detail", None),
                                                                  ("fix", None))
        self.query_one("#glog").display = False
        self.load()
        self.query_one(RunList).focus()

    def load(self) -> None:
        self.error, self.pending, self.view = "", None, "start"
        self.snap, self.snap_origin, self.snap_stale = None, "", False
        try:
            job = self.m365.groups_job()
        except Exception as exc:  # noqa: BLE001 - shown in the pane
            self.error = str(exc)
        else:
            self.pick_snapshot(job)
        self.after_op()

    def tenant_changed(self) -> None:
        self.load()

    def after_op(self) -> None:
        self._render_actions()
        self._render_runs()
        self._render_state()
        self._render_view()
        self.refresh_bindings()

    # -------------------------------------------------------------------------------------------- rendering

    def _render_actions(self) -> None:
        g = Table.grid(padding=(0, 1))
        g.add_column(no_wrap=True)
        g.add_column(style=SUB, no_wrap=True)
        for key, what in (("o", "new person"), ("x", "offboard a leaver"), ("c", "check one person"),
                          ("i", "inventory")):
            g.add_row(theme.chip(key), what)
        self.query_one("#oactions", Static).update(g)

    def _render_runs(self) -> None:
        from m365tools.services import list_group_runs

        job = self.job
        runs: list[Path] = []
        if job is not None:
            runs = [r for r in list_group_runs(job.tenant) if "-onboard" in r.name or "-offboard" in r.name][:20]
        self.runs = runs
        lst = self.query_one(RunList)
        lst.set_options([BarOption(_run_maker(r), id=str(i)) for i, r in enumerate(runs)])
        if runs:
            lst.highlighted = 0
        empty = self.query_one("#oruns-empty", Static)
        empty.display = not runs
        empty.update(Text("No onboard or offboard runs yet.", style=DIM))

    def _render_state(self) -> None:
        job = self.job
        if self.error:
            self.query_one("#gstate", Static).update(Text(self.error, style=ORANGE))
            return
        findings = drift_mod.findings(self.snap, self.registry) if self.snap is not None else []
        self.query_one("#gstate", Static).update(state_ribbon(
            self.snap, self.registry, findings, origin=self.snap_origin, stale=self.snap_stale,
            registry_name=job.registry_path.name if job is not None else "none"))

    def _render_view(self) -> None:
        switch = self.query_one(ContentSwitcher)
        if self.view == "start":
            switch.current = "ostart"
            self.query_one("#ostart", Empty).update(self._start_text())
        else:
            switch.current = "opreview" if self.view == "preview" else "ocheck"

    def _start_text(self) -> Text:
        if self.error:
            return theme.empty_text("onboard cannot start", f"{self.error}\n\nAdd a tenant with: m365 tenant add NAME")
        if self.snap is None:
            return theme.empty_text(
                "No snapshot yet. Press i to run an inventory.",
                "Onboarding copies access from a colleague or a bundle, so it needs to know who has what.\n"
                "The inventory reads the tenant and writes nothing.", [("i", "inventory"), ("?", "help")])
        return theme.empty_text(
            "Bring people in and out",
            "o gives a new colleague the groups of a bundle or of someone in the same role.\n"
            "x takes a leaver out of every registry group. Every change is previewed first.",
            [("o", "new person"), ("x", "offboard"), ("c", "check one person")])

    def show_view(self, view: str) -> None:
        self.view = view
        self._render_view()
        self.refresh_bindings()
        target = {"preview": "#ochanges", "check": "#ocheck-table"}.get(view)
        if target:
            self.query_one(target).focus()

    def check_action(self, action: str, parameters: tuple[object, ...]) -> bool | None:
        if action == "apply":
            return self.view == "preview" and self.pending is not None
        return True

    # -------------------------------------------------------------------------------------------- navigation

    def action_toggle_pane(self) -> None:
        f = self.focused
        if f is not None and self.query_one("#omain") in f.ancestors:
            self.query_one(RunList).focus()
        elif self.view != "start":
            self.show_view(self.view)

    def action_back(self) -> None:
        if self.view != "start":
            self.pending = None
            self.show_view("start")
            self.query_one(RunList).focus()
        else:
            self.m365.action_go_home()

    def on_option_list_option_selected(self, event: OptionList.OptionSelected) -> None:
        if event.option_list.id == "oruns" and event.option.id is not None:
            self.m365.show_report_path(self.runs[int(event.option.id)])

    def action_toggle_log(self) -> None:
        self.log_open = not self.log_open
        self.query_one("#glog").display = self.log_open

    # -------------------------------------------------------------------------------------------- actions

    def _need_snapshot(self) -> bool:
        if self.m365.busy:
            self.notify(f"Busy: {self.m365.busy} is running.", severity="warning")
            return False
        if self.error:
            self.notify(self.error, severity="error")
            return False
        if self.snap is None:
            self.notify("No snapshot yet: press i to run an inventory first.", severity="warning")
            return False
        return True

    def _people_choices(self, exclude: str = "") -> list[Choice]:
        return [Choice(f"like  {r.name}", ("like", r.key), f"{r.key} · {r.hint}")
                for r in people_rows(self.snap, self.registry)
                if r.key != exclude and r.state not in ("offboarded", "gone")]

    async def _pick_person(self, title: str, note: str) -> str | None:
        choices = [Choice(r.name, r.key, f"{r.key} · {r.hint}") for r in people_rows(self.snap, self.registry)]
        upn = await self.app.push_screen_wait(PickerModal(
            title, choices, placeholder="type a name or a UPN", note=note,
            accept_text=lambda t: t.lower() if "@" in t else None))
        return upn.strip().lower() if isinstance(upn, str) and upn.strip() else None

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_inventory(self) -> None:
        ok, snap = await self.run_op("inventory", lambda job: job.inventory())
        if ok and snap is not None:
            self.notify(f"Inventory done: {len(snap.people)} people, {len(snap.groups)} groups.")

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_onboard(self) -> None:
        if not self._need_snapshot():
            return
        reg = self.registry
        upn = await self.app.push_screen_wait(PromptModal(
            "Onboard · who", placeholder="new.colleague@contoso.com",
            note="The person's UPN (sign-in name). They must already exist in Entra ID."))
        if not upn or not upn.strip():
            return
        upn = upn.strip().lower()
        choices = [Choice(f"bundle  {b.name}", ("bundle", b.name), f"{len(b.groups)} groups · {b.purpose}")
                   for b in sorted(reg.bundles.values(), key=lambda b: b.name.lower())]
        choices += self._people_choices(exclude=upn)
        if reg.person(upn) is not None:
            choices.insert(0, Choice("their registry entry", ("registry", ""), "re-add what the registry says"))
        how = await self.app.push_screen_wait(PickerModal(
            f"Onboard {upn} · give access like", choices, placeholder="type to filter: a bundle or a colleague",
            note="A bundle adds its groups. 'like' copies a colleague's bundles and add-ons."))
        if how is None:
            return
        addons: list[str] = []
        groups = sorted({r.name for r in group_rows(self.snap, reg)}, key=str.lower)
        while True:
            left = [Choice("done", None, "no more add-ons" if addons else "no add-ons")]
            left += [Choice(g, g, (reg.group(g).purpose if reg.group(g) else "")) for g in groups
                     if g not in addons]
            pick = await self.app.push_screen_wait(PickerModal(
                f"Onboard {upn} · add-ons" + (f" ({', '.join(addons)})" if addons else ""), left,
                placeholder="type a group, or enter for done",
                note="Extra groups on top of the bundle, for this person only. Pick 'done' to continue."))
            if pick is None or pick == "done" or not isinstance(pick, str):
                break
            addons.append(pick)
        kind, value = how
        kwargs: dict[str, Any] = {"addons": tuple(addons)}
        if kind == "bundle":
            kwargs["bundles"] = (value,)
        elif kind == "like":
            kwargs["like"] = value
        await self.preview_onboard(upn, kwargs)

    async def preview_onboard(self, upn: str, kwargs: dict[str, Any]) -> None:
        ok, plan = await self.run_op("onboard preview", lambda job: job.onboard(upn, apply=False, **kwargs))
        if not ok or plan is None:
            return
        todo = [c for c in plan.changes if c.status == "planned"]
        pending = None
        if todo:
            pending = ("apply", f"Add {upn} to {theme.plural(len(todo), 'group')}?",
                       lambda job: job.onboard(upn, apply=True, **kwargs),
                       lambda p: self._show_onboard(p, done=True))
        self._show_onboard(plan, pending=pending)

    def _show_onboard(self, plan: Any, *, pending: Any = None, done: bool = False) -> None:
        rows = _member_rows(plan, "add member", self.registry)
        manual = Text()
        for m in plan.manual:
            manual.append("○ ", style=ORANGE).append(f"{m}\n", style=FG)
        if not plan.manual:
            manual.append("nothing: groups cover it all", style=DIM)
        entry = _entry_text(plan.upn, [("bundles", plan.bundles), ("addons", plan.addons)])
        notes = Text.assemble(("from ", SUB), (plan.source, FG))
        self._show_preview(f"Onboard {plan.upn}" + (": done" if done else ""), rows, manual, entry, notes, pending)

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_offboard(self) -> None:
        if not self._need_snapshot():
            return
        upn = await self._pick_person("Offboard · who", "Pick the leaver, or type their UPN. Their registry "
                                                          "groups are previewed before anything changes.")
        if upn is None:
            return
        ok, plan = await self.run_op("offboard preview", lambda job: job.offboard(upn, apply=False))
        if not ok or plan is None:
            return
        todo = [c for c in plan.changes if c.status == "planned"]
        pending = None
        if todo:
            pending = ("offboard", f"Remove {upn} from {theme.plural(len(todo), 'group')} and mark them offboarded?",
                       lambda job: job.offboard(upn, apply=True), lambda p: self._show_offboard(p, done=True))
        self._show_offboard(plan, pending=pending)

    def _show_offboard(self, plan: Any, *, pending: Any = None, done: bool = False) -> None:
        rows = _member_rows(plan, "remove member", self.registry)
        rows += [("keep", g, "not in the registry: left alone", "kept") for g in plan.kept]
        kept = Text("Licences, apps and roles are not touched.\n", style=SUB)
        if plan.kept:
            kept.append("Groups outside the registry are kept: remove those by hand if needed.", style=SUB)
        entry = _entry_text(plan.upn, [("offboarded", plan.date)])
        self._show_preview(f"Offboard {plan.upn}" + (": done" if done else ""), rows, kept, entry, Text(), pending,
                           manual_title="not changed")

    def _show_preview(self, title: str, rows: list[tuple[str, str, str, str]], manual: Text, entry: Text,
                      notes: Text, pending: Any, manual_title: str = "by hand (admin center)") -> None:
        self.pending = pending
        n = sum(1 for r in rows if r[3] == "planned")
        head = Text.assemble((title, f"bold {FG}"), ("   ", ""))
        if pending is not None:
            head.append_text(theme.pill(f"preview · {n} to do", YELLOW)).append("  not written yet", style=SUB)
        elif any(r[3] in ("done", "failed") for r in rows):
            head.append_text(theme.pill("applied", GREEN))
        else:
            head.append("nothing to do", style=GREEN)
        self.query_one("#ohead", Static).update(head)
        table = self.query_one("#ochanges", VimTable)
        table.clear()
        for i, (action, subject, detail, status) in enumerate(rows):
            table.add_row(Text(action, style=SUB), Text(subject, style=FG, no_wrap=True, overflow="ellipsis"),
                          Text(detail, style=SUB, no_wrap=True, overflow="ellipsis"), change_status(status),
                          key=str(i))
        box = self.query_one("#omanual", Static)
        box.border_title = manual_title
        box.update(manual)
        self.query_one("#oentry", Static).update(entry)
        out = Text()
        out.append_text(notes)
        if pending is not None:
            if out:
                out.append("\n")
            out.append("A ", style=f"bold {BLUE}").append("applies (you type ", style=SUB)
            out.append(pending[0], style=f"bold {RED}").append(")   ", style=SUB)
            out.append("esc ", style=f"bold {BLUE}").append("drops the preview", style=SUB)
        self.query_one("#onotes", Static).update(out)
        self.show_view("preview")

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_apply(self) -> None:
        if self.pending is None:
            return
        word, question, apply, done = self.pending
        body = Text.assemble((question, FG), ("\n\nThis changes ", SUB),
                             ("group memberships in the tenant and the registry", f"bold {FG}"), (".", SUB))
        if not await self.app.push_screen_wait(TypedConfirmModal(question, body, word)):
            return
        self.pending = None
        ok, result = await self.run_op(word, apply)
        if ok:
            self.snap_stale = self.snap is not None
            done(result)
            self._render_state()
            self.notify("Applied. The registry was saved." + (" Press i to refresh the snapshot."
                                                              if self.snap_stale else ""))

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_check(self) -> None:
        if not self._need_snapshot():
            return
        upn = await self._pick_person("Check · who", "Compare one person with the registry. Nothing is changed.")
        if upn is None:
            return
        ok, findings = await self.run_op("check", lambda job: job.onboard_check(upn))
        if not ok:
            return
        self.findings = sorted(findings or [], key=lambda f: (SEV_ORDER.get(str(f.severity), 9), f.category))
        head = Text.assemble((f"Check {upn}", f"bold {FG}"), ("   ", ""))
        if self.findings:
            head.append_text(theme.pill(theme.plural(len(self.findings), "finding"), YELLOW))
        else:
            head.append("● matches the registry", style=GREEN)
        self.query_one("#ocheck-head", Static).update(head)
        table = self.query_one("#ocheck-table", VimTable)
        table.clear()
        for i, f in enumerate(self.findings):
            color = SEVERITY_STYLE.get(str(f.severity), SUB)
            table.add_row(theme.glyph_word(str(f.severity), color), Text(FINDING_LABEL.get(f.category, f.category),
                                                                          style=f"bold {FG}"),
                          Text(f.detail, style=SUB, no_wrap=True, overflow="ellipsis"),
                          Text(f.fix, style=GREEN, no_wrap=True, overflow="ellipsis"), key=str(i))
        self.show_view("check")

    # -------------------------------------------------------------------------------------------- help, palette

    def action_help(self) -> None:
        from rich.console import Group

        from .workspace import help_group

        keys = [("o", "new person: like a colleague or a bundle"), ("x", "offboard a leaver"),
                ("c", "check one person against the registry"), ("i", "inventory: read the tenant, writes nothing"),
                ("A", "apply the preview (asks a typed word)"), ("esc", "drop the preview, back to the start"),
                ("enter", "report path of a recent run"), ("L", "show / hide the log")]
        glob = [(":", "command palette"), ("ctrl+l", "sign in"), ("?", "this help"), ("g", "toolbox home"),
                ("q", "quit")]
        cols = Table.grid(expand=True, padding=(0, 4))
        cols.add_column(ratio=1)
        cols.add_column(ratio=1)
        cols.add_row(help_group("onboard", keys), Group(help_group("anywhere", glob), Text(""),
                                                        help_group("back and switch", BACK_HELP)))
        body = Group(Text(HELP[self.view], style=SUB), Text(""), cols)
        self.app.push_screen(TextModal(f"Keys · onboard · {self.view}", body, fit=True))

    def palette_commands(self) -> list[tuple[str, Any, str]]:
        return [
            ("Onboard a new person", self.action_onboard, "give them access like a colleague or a bundle  (o)"),
            ("Offboard a person", self.action_offboard, "remove a leaver from every registry group  (x)"),
            ("Check one person", self.action_check, "compare one person with the registry  (c)"),
            ("Inventory", self.action_inventory, "read the tenant so onboarding knows who has what  (i)"),
            ("Toggle log", self.action_toggle_log, "(L)"),
            ("Help", self.action_help, "keys and what this view means  (?)"),
        ]


def _entry_text(upn: str, fields: list[tuple[str, Any]]) -> Text:
    t = Text(no_wrap=True, overflow="ellipsis")
    t.append("[people.", style=DIM).append(f'"{upn}"', style=ACCENT).append("]\n", style=DIM)
    for k, v in fields:
        value = "[" + ", ".join(f'"{x}"' for x in v) + "]" if isinstance(v, list) else f'"{v}"'
        t.append(f"{k} = ", style=SUB).append(f"{value}\n", style=FG)
    return t


def _run_maker(run: Path) -> Callable[[bool], Text]:
    from .home import _parse

    mode, upn, state = run_summary(run)
    when = ago(_parse(run).started)

    def make(selected: bool) -> Text:
        t = Text(no_wrap=True, overflow="ellipsis")
        if selected:
            return t.append(f"{mode:<9}{upn}\n  {state} · {when}", style=ON_FILL)
        t.append(f"{mode:<9}", style=GREEN if mode == "onboard" else ORANGE).append(upn, style=FG)
        return t.append(f"\n  {state} · {when}", style=SUB if state == "applied" else DIM)

    return make


