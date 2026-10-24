"""The groups screen: an explorer over people, groups and bundles, with drift and registry upkeep as actions.
Onboarding and offboarding have their own screen (onboard.py).

Left: a filterable list in three tabs (People, Groups, Bundles). Right: the state of the registry and the
snapshot, then the detail of what is highlighted (or the result of the last action: drift findings, bundle
suggestions, a preview of changes). An operation's step list and log sit at the bottom of the right pane.

    1 2 3  t     People / Groups / Bundles      /       filter            enter   open (right pane)
    tab          switch list <-> right pane     esc     back              L       show / hide the log
    i  inventory    d  drift    s  suggest bundles    y  sync    m  import
    e  edit purpose and owner    r  reload registry    A  apply the previewed changes (asks a typed word)

Browsing is local: details are computed from the last snapshot (kept from the last inventory) and the registry,
so moving through the list never calls Microsoft Graph. Every action that writes shows a preview first.
"""

from __future__ import annotations

import contextlib
import json
from collections.abc import Awaitable, Callable
from datetime import datetime
from typing import TYPE_CHECKING, Any, ClassVar

from rich.console import Group
from rich.table import Table
from rich.text import Text
from textual import work
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical
from textual.screen import Screen
from textual.widgets import ContentSwitcher, DataTable, OptionList, Static

from m365tools.core.events import GROUPS_PHASES, LogLine, ProgressEvent
from m365tools.tools.groups import drift as drift_mod
from m365tools.tools.groups.explain import explain_group, explain_person
from m365tools.tools.groups.models import Finding, GroupEntry, Registry, Snapshot, kind_of

from . import theme
from .groups_view import (
    FINDING_LABEL,
    Row,
    bundle_detail,
    bundle_rows,
    change_status,
    filtered,
    group_detail,
    group_rows,
    people_rows,
    person_detail,
    snapshot_age,
    state_ribbon,
)
from .modals import ConfirmModal, PromptModal, TextModal, TypedConfirmModal
from .sections.base import Empty
from .theme import ACCENT, DIM, FG, GREEN, ORANGE, SEVERITY_STYLE, SUB, YELLOW
from .widgets import Activity, BarList, BarOption, HeaderBar, KeyBar, StepList, VimLog, VimScroll, VimTable

if TYPE_CHECKING:
    from .app import M365App

TABS = (("groups", "Groups"), ("people", "People"), ("bundles", "Bundles"))
SEV_ORDER = {"blocking": 0, "warn": 1, "info": 2}

HELP = {
    "people": ("People lists everyone in the tenant snapshot plus everyone in the registry. The glyph says how "
               "they compare: ● matches the registry, ▲ has drift (a missing or unexplained group), ○ is not in "
               "the registry yet, ◌ is in the registry but no longer in the tenant. Highlight someone to see "
               "their bundles, add-ons and every group with why they have it. Nothing here writes anything."),
    "groups": ("Groups lists every prefix group in the tenant and every group in the registry. ● is in both, "
               "▲ exists in the tenant but nobody registered it (m imports it), ○ is in the registry but not "
               "created yet (y creates it). e sets the purpose and owner: write them so a colleague knows what "
               "the group gives access to and who approves new members."),
    "bundles": ("Bundles are role packages kept in the registry, like 'Servicedesk': a list of groups someone in "
                "that role needs. Onboarding a person with a bundle adds them to each of its groups. s suggests "
                "bundles from who has what today."),
    "drift": ("Drift compares the registry with the tenant. 'missing' means a bundle or add-on says the person "
              "should be in a group but is not; 'unexplained' means they are in a group and nothing says why. "
              "enter jumps to the person or group, esc comes back here. Nothing is changed by looking."),
    "suggest": ("Suggestions group people by the exact set of groups they have. Each candidate could become a "
                "bundle: a accepts the highlighted one under a name you choose and shows what the registry would "
                "get; you confirm by typing apply. Only the registry file changes, never the tenant."),
    "changes": ("A preview of changes. Nothing has been written yet: read the list, then A applies it after you "
                "type the confirmation word. 'manual' items (licences, apps, roles) are not set by this tool; "
                "do them in the admin center."),
}


class GroupList(BarList):
    BINDINGS: ClassVar[list[BindingType]] = [Binding("enter", "select", "open")]


class GroupsOpsScreen(Screen[None]):
    """What the groups and onboard screens share: the GroupsJob of the current tenant, its snapshot, and running
    an operation with the step list (``#gsteps``), activity line (``#gactivity``) and log (``#glog``)."""

    snap: Snapshot | None = None
    snap_origin = ""
    snap_stale = False

    @property
    def m365(self) -> M365App:
        return self.app  # type: ignore[return-value]

    @property
    def job(self) -> Any:
        return self.m365.groups_current

    @property
    def registry(self) -> Registry:
        job = self.job
        return job.registry if job is not None else Registry()

    def pick_snapshot(self, job: Any) -> None:
        """The job's snapshot of this session, else the newest saved inventory."""
        if job.snapshot is not None:
            self.snap, self.snap_origin = job.snapshot, "this session"
        else:
            self.snap = self._saved_snapshot(job)
            self.snap_origin = "last inventory" if self.snap else ""

    @staticmethod
    def _saved_snapshot(job: Any) -> Snapshot | None:
        from m365tools.services import list_group_runs

        with contextlib.suppress(Exception):
            for run in list_group_runs(job.tenant, "inventory"):
                path = run / "snapshot.json"
                if path.exists():
                    return Snapshot.from_json(json.loads(path.read_text(encoding="utf-8")))
        return None

    def write_log(self, ev: LogLine) -> None:
        self.query_one("#glog", VimLog).write(theme.log_text(ev))

    def on_job_event(self, ev: ProgressEvent) -> None:
        steps = self.query_one("#gsteps", StepList)
        if steps.running:
            steps.feed(ev)
        act = self.query_one("#gactivity", Activity)
        if act.active:
            act.feed(ev)

    async def run_op(self, name: str, op: Callable[[Any], Awaitable[Any]]) -> tuple[bool, Any]:
        """Run ``op(job)`` as operation ``name`` (one at a time across the toolbox). Never raises."""
        app = self.m365
        job = app.begin(name, self, tool="groups")
        if job is None:
            return False, None
        steps, act = self.query_one("#gsteps", StepList), self.query_one("#gactivity", Activity)
        use_steps = name == "inventory" or job.snapshot is None
        if use_steps:
            steps.start(name)
        else:
            steps.display = False
            act.start(name)
        ok, result = False, None
        try:
            result = await op(job)
            ok = True
        except Exception as exc:  # noqa: BLE001 - every failure becomes a toast + log line
            app.report_error(exc, name)
        finally:
            if use_steps:
                app.call_later(steps.finish, ok)
            act.stop()
            app.end()
        if job.snapshot is not None and job.snapshot is not self.snap:
            self.snap, self.snap_origin, self.snap_stale = job.snapshot, "this session", False
        self.after_op()
        return ok, result

    def after_op(self) -> None:
        """Re-render after an operation (override)."""


    def on_screen_resume(self) -> None:
        """Another screen may have taken a newer snapshot (or written to the registry) meanwhile."""
        job = self.job
        if job is not None and job.snapshot is not None and job.snapshot is not self.snap:
            self.pick_snapshot(job)
            self.snap_stale = False
        if job is not None:
            self.after_op()

    def _names(self, upn: str) -> str:
        p = self.snap.person(upn) if self.snap else None
        return p.display_name if p and p.display_name else upn.split("@")[0]


class GroupsScreen(GroupsOpsScreen):
    BINDINGS: ClassVar[list[BindingType]] = [
        Binding("tab,shift+tab", "toggle_pane", "switch pane", show=False),
        Binding("escape", "back", "back", show=False),
        Binding("1", "tab('groups')", "tabs", key_display="1/2/3"),
        Binding("2", "tab('people')", "people", show=False),
        Binding("3", "tab('bundles')", "bundles", show=False),
        Binding("0", "start", "get started", show=False),
        Binding("t", "next_tab", "tab", show=False),
        Binding("slash", "filter", "filter", key_display="/"),
        Binding("i", "inventory", "inventory"),
        Binding("m", "import_groups", "import"),
        Binding("e", "edit", "edit"),
        Binding("s", "suggest", "suggest"),
        Binding("d", "drift", "drift"),
        Binding("y", "sync", "sync"),
        Binding("a", "accept", "accept"),
        Binding("A", "apply", "apply"),
        Binding("D", "domains", "domains", show=False),
        Binding("r", "reload", "reload", show=False),
        Binding("L", "toggle_log", "log", show=False),
        Binding("colon", "app.command_palette", "commands", key_display=":"),
        Binding("question_mark", "help", "help", key_display="?"),
        Binding("ctrl+l", "app.login", "sign in", show=False),
        Binding("g", "app.go_home", "home", priority=True),
        Binding("ctrl+1", "app.tool(1)", "spmig", show=False, priority=True),
        Binding("ctrl+2", "app.tool(2)", "groups", show=False, priority=True),
        Binding("ctrl+3", "app.tool(3)", "onboard", show=False, priority=True),
        Binding("q", "app.quit_app", "quit"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.mode = "groups"
        self.filter_text = ""
        self.rows: list[Row] = []
        self.snap: Snapshot | None = None
        self.snap_origin = ""
        self.snap_stale = False
        self.local_findings: list[Finding] = []
        self.findings: list[Finding] = []
        self.suggestion: Any = None
        self.view = "start"
        self.subject: str | None = None  # Row id shown in the detail
        self.back_view: str | None = None  # where esc returns to from a detail opened by a jump
        self.pending: tuple[str, str, Callable[[Any], Awaitable[Any]], Callable[[Any], None]] | None = None
        self.changes_from: str | None = None
        self.error = ""
        self.log_open = False
        self._last_hl: str | None = None  # highlight set by code: its event is not a user move

    # -------------------------------------------------------------------------------------------- layout

    def compose(self) -> ComposeResult:
        yield HeaderBar("groups", id="header")
        with Horizontal(id="body"):
            with Vertical(id="gnav"):
                yield Static("", id="gtabs")
                yield Static("", id="gfilter")
                yield GroupList(id="glist")
                yield Static("", id="glist-empty")
            with Vertical(id="gmain"):
                yield Static("", id="gstate")
                with ContentSwitcher(id="gswitch", initial="start"):
                    start = VimScroll(id="start", classes="gpanel")
                    with start:
                        yield Static("", id="start-body")
                    detail = VimScroll(id="detail", classes="gpanel")
                    with detail:
                        yield Static("", id="detail-body")
                    with Vertical(id="drift"):
                        yield Static("", id="drift-counts", classes="counts")
                        yield VimTable(id="drift-table")
                        yield Static("", id="drift-fix", classes="gnote")
                    with Vertical(id="suggest"):
                        yield Static("", id="suggest-head", classes="counts")
                        yield VimTable(id="suggest-table")
                        yield Static("", id="suggest-detail", classes="gnote")
                    with Vertical(id="changes"):
                        yield Static("", id="changes-head", classes="counts")
                        yield VimTable(id="changes-table")
                        yield Static("", id="changes-notes", classes="gnote")
                    yield Empty("", id="gempty")
                with Vertical(id="ops"):
                    yield StepList(GROUPS_PHASES, id="gsteps")
                    yield Activity(id="gactivity")
                    log = VimLog(id="glog")
                    log.border_title = "log"
                    yield log
        yield KeyBar(id="keybar")

    def on_mount(self) -> None:
        self.query_one("#drift-table", VimTable).add_fit_columns(("severity", 8), ("who", 24), ("what", None),
                                                                  ("fix", None))
        self.query_one("#suggest-table", VimTable).add_fit_columns(("#", 3), ("bundle name", 22), ("people", 7),
                                                                    ("groups", None), ("exists as", 14))
        self.query_one("#changes-table", VimTable).add_fit_columns(("action", 15), ("subject", 24), ("detail", None),
                                                                    ("status", 18))
        self.query_one("#glog").display = False
        self.load()
        self.query_one(GroupList).focus()

    # -------------------------------------------------------------------------------------------- data

    def load(self) -> None:
        """(Re)create the job for the current tenant, read its registry and the newest saved snapshot."""
        self.error = ""
        self.snap, self.snap_origin, self.snap_stale = None, "", False
        self.findings, self.suggestion, self.pending = [], None, None
        try:
            job = self.m365.groups_job()
        except Exception as exc:  # noqa: BLE001 - shown in the pane
            self.error = str(exc)
            self.refresh_all()
            return
        self.pick_snapshot(job)
        self.view = "start"
        self.refresh_all()

    def tenant_changed(self) -> None:
        self.load()

    def after_op(self) -> None:
        self.refresh_all()

    # -------------------------------------------------------------------------------------------- rendering

    def refresh_all(self) -> None:
        reg = self.registry
        self.local_findings = drift_mod.findings(self.snap, reg) if self.snap is not None else []
        self._rebuild_list()
        self._render_state()
        self._render_view()
        self.refresh_bindings()

    def _all_rows(self) -> list[Row]:
        reg = self.registry
        if self.mode == "people":
            return people_rows(self.snap, reg)
        if self.mode == "groups":
            return group_rows(self.snap, reg)
        return bundle_rows(reg)

    def _rebuild_list(self, keep: str | None = None) -> None:
        lst = self.query_one(GroupList)
        if keep is None:
            keep = self.current_row.id if self.current_row else self.subject
        self.rows = filtered(self._all_rows(), self.filter_text)
        lst.set_options([BarOption(_maker(r), id=r.id) for r in self.rows])
        if self.rows:
            ids = [r.id for r in self.rows]
            idx = ids.index(keep) if keep in ids else 0
            self._last_hl = ids[idx]
            lst.highlighted = idx
        self._render_tabs()
        empty = self.query_one("#glist-empty", Static)
        empty.display = not self.rows
        if not self.rows:
            if self.filter_text:
                msg = Text.assemble(("Nothing matches ", SUB), (self.filter_text, f"bold {ACCENT}"), (".\n", SUB),
                                    ("/ changes the filter, esc clears it.", DIM))
            elif self.mode == "bundles":
                msg = Text.assemble(("No bundles yet.\n", f"bold {FG}"),
                                    ("Press s to suggest bundles from who has what today.", SUB))
            else:
                msg = Text.assemble((f"No {self.mode} yet.\n", f"bold {FG}"),
                                    ("Press i to run an inventory.", SUB))
            empty.update(msg)

    def _render_tabs(self) -> None:
        reg = self.registry
        counts = {"people": len(people_rows(self.snap, reg)), "groups": len(group_rows(self.snap, reg)),
                  "bundles": len(reg.bundles)}
        t = Text(no_wrap=True)
        for i, (key, title) in enumerate(TABS, 1):
            if key == self.mode:
                t.append_text(theme.pill(f"{i} {title} {counts[key]}", ACCENT))
            else:
                t.append_text(theme.chip(str(i))).append(f" {title}", style=FG)
            t.append(" ")
        self.query_one("#gtabs", Static).update(t)
        f = self.query_one("#gfilter", Static)
        f.display = bool(self.filter_text)
        if self.filter_text:
            f.update(Text.assemble(("/ ", DIM), (self.filter_text, f"bold {ACCENT}"),
                                   (f"  {len(self.rows)} shown", SUB)))

    def _render_state(self) -> None:
        job = self.job
        name = job.registry_path.name if job is not None else "none"
        if self.error:
            self.query_one("#gstate", Static).update(Text(self.error, style=ORANGE))
            return
        self.query_one("#gstate", Static).update(state_ribbon(
            self.snap, self.registry, self.local_findings, origin=self.snap_origin, stale=self.snap_stale,
            registry_name=name))

    @property
    def current_row(self) -> Row | None:
        lst = self.query_one(GroupList)
        idx = lst.highlighted
        if idx is None or not 0 <= idx < len(self.rows):
            return None
        return self.rows[idx]

    def show_view(self, view: str) -> None:
        self.view = view
        self._render_view()
        self.refresh_bindings()

    def _render_view(self) -> None:
        switch = self.query_one(ContentSwitcher)
        if self.error:
            switch.current = "gempty"
            self.query_one("#gempty", Empty).update(theme.empty_text(
                "groups cannot start", f"{self.error}\n\nAdd a tenant with: m365 tenant add NAME", []))
            return
        if self.view == "detail" and self.current_row is None and self.subject is None:
            self.view = "start"
        if self.view == "start":
            switch.current = "start"
            self.query_one("#start-body", Static).update(self._start_card())
            return
        if self.view == "detail":
            row = self.current_row
            switch.current = "detail"
            self._render_detail(self.subject if self.subject and self.back_view else (row.id if row else
                                                                                      self.subject))
            return
        switch.current = self.view

    def _start_card(self) -> Any:
        """The "get started" checklist: what to do next, derived from the snapshot and the registry."""
        from rich.console import Group

        reg, snap = self.registry, self.snap
        cats = [f.category for f in self.local_findings]
        warn = sum(1 for f in self.local_findings if str(f.severity) == "warn")
        info = sum(1 for f in self.local_findings if str(f.severity) == "info")
        unreg = cats.count("unregistered")
        lacking = [g for g in reg.groups.values() if not g.purpose.strip() or not g.owner.strip()]
        n_people = len(snap.people) if snap else 0
        steps: list[tuple[str, str, str, str]] = []  # (state, key, title, detail)
        steps.append(("done" if snap else "todo", "i", "Run an inventory",
                      f"snapshot {snapshot_age(snap.taken)}" if snap else "reads the tenant, writes nothing"))
        if reg.domains:
            steps.append(("done", "D", "Limit people to your own domain",
                          f"{n_people} people at {', '.join(reg.domains)}" if snap else ", ".join(reg.domains)))
        else:
            steps.append(("doing" if snap else "todo", "D", "Limit people to your own domain",
                          f"{n_people} people in the snapshot, external accounts included" if snap
                          else "only colleagues, no partners or test accounts"))
        steps.append(("todo" if snap is None else "doing" if unreg else "done", "m",
                      "Import your groups into the registry",
                      f"{theme.plural(unreg, 'group')} not in the registry yet" if unreg else
                      (f"{len(reg.groups)} groups registered" if snap else "after the inventory")))
        steps.append(("todo" if not reg.groups else "doing" if lacking else "done", "e",
                      "Give every group a purpose and owner",
                      f"{theme.plural(len(lacking), 'group')} still lack one: 1, a group, then e" if lacking else
                      ("all groups explained" if reg.groups else "after the import")))
        steps.append(("done" if reg.bundles else "todo", "s", "Suggest bundles",
                      f"{theme.plural(len(reg.bundles), 'bundle')}: " + ", ".join(sorted(reg.bundles)) if reg.bundles
                      else "role packages from who has what today"))
        steps.append(("todo" if snap is None else "doing" if warn else "done", "d", "Check drift weekly",
                      f"{warn} warn · {info} info" if snap else "after the inventory"))
        glyph = {"done": ("●", GREEN), "doing": ("◐", YELLOW), "todo": ("○", DIM)}
        nxt = next((i for i, st in enumerate(steps) if st[0] != "done"), None)
        g = Text(no_wrap=True, overflow="ellipsis")
        for i, (state, key, title, detail) in enumerate(steps):
            gl, color = glyph[state]
            g.append(f"{gl} ", style=color).append(f"{i + 1}  ", style=DIM).append_text(theme.chip(key))
            g.append(f" {title}\n", style=f"bold {FG}" if i == nxt else FG)
            g.append(f"         {detail}\n", style=SUB)
        head = Text.assemble(("Get started", f"bold {ACCENT}"), ("   ", ""),
                             (f"{sum(1 for st in steps if st[0] == 'done')} of {len(steps)} done", SUB))
        parts: list[Any] = [head, Text(""), g]
        if nxt is not None:
            parts.append(Text.assemble(("next  ", DIM), (steps[nxt][2], f"bold {FG}"), ("  press ", SUB),
                                       (steps[nxt][1], f"bold {theme.BLUE}")))
        parts += [Text(""), Text.assemble(
            ("1", f"bold {theme.BLUE}"), (" groups  ", SUB), ("2", f"bold {theme.BLUE}"), (" people  ", SUB),
            ("3", f"bold {theme.BLUE}"), (" bundles  ", SUB), ("enter", f"bold {theme.BLUE}"), (" open  ", SUB),
            ("0/esc", f"bold {theme.BLUE}"), (" back here", SUB))]
        return Group(*parts)

    def _render_detail(self, row_id: str | None) -> None:
        body = self.query_one("#detail-body", Static)
        if row_id is None:
            body.update("")
            return
        self.subject = row_id
        kind, key = row_id[0], row_id[2:]
        snap = self.snap or Snapshot(self.registry.prefix, datetime.now())
        try:
            if kind == "p":
                content: Any = person_detail(explain_person(snap, self.registry, key))
            elif kind == "g":
                content = group_detail(explain_group(snap, self.registry, key), self._names)
            else:
                content = bundle_detail(self.registry, self.snap, key, self._names)
        except Exception as exc:  # noqa: BLE001 - shown in the pane
            content = Text(f"{type(exc).__name__}: {exc}", style=ORANGE)
        if self.snap is None and kind != "b":
            content = Group(content, Text(""), Text("No snapshot yet: only the registry side is shown. "
                                                    "Press i to run an inventory.", style=YELLOW))
        if self.back_view:
            content = Group(Text.assemble(("esc ", f"bold {theme.BLUE}"), (f"back to {self.back_view}", SUB)),
                            Text(""), content)
        body.update(content)
        self.query_one("#detail", VimScroll).scroll_home(animate=False)

    # -------------------------------------------------------------------------------------------- list events

    def on_option_list_option_highlighted(self, event: OptionList.OptionHighlighted) -> None:
        if event.option_list.id != "glist" or not event.option.id or event.option.id == self._last_hl:
            return
        self._last_hl = event.option.id
        if self.focused is event.option_list or self.view == "detail":
            self.back_view = None
            self.subject = event.option.id
            if self.view != "detail":
                self.view = "detail"
            self._render_view()
            self.refresh_bindings()

    def on_option_list_option_selected(self, event: OptionList.OptionSelected) -> None:
        if event.option_list.id == "glist":
            self.back_view = None
            self.subject = event.option.id
            self.show_view("detail")
            self.query_one("#detail").focus()

    def on_data_table_row_highlighted(self, event: DataTable.RowHighlighted) -> None:
        if event.data_table.id == "drift-table":
            self._render_drift_fix()
        elif event.data_table.id == "suggest-table":
            self._render_candidate()

    def on_data_table_row_selected(self, event: DataTable.RowSelected) -> None:
        event.stop()
        if event.data_table.id == "drift-table":
            self.jump_to_finding()

    # -------------------------------------------------------------------------------------------- navigation

    def _right_focused(self) -> bool:
        f = self.focused
        return f is not None and self.query_one("#gmain") in f.ancestors

    def _focus_right(self) -> None:
        target = {"start": "#start", "detail": "#detail", "drift": "#drift-table", "suggest": "#suggest-table",
                  "changes": "#changes-table"}.get(self.view)
        if target and self.query_one(ContentSwitcher).current == self.view:
            self.query_one(target).focus()

    def action_toggle_pane(self) -> None:
        if self._right_focused():
            self.query_one(GroupList).focus()
        else:
            self._focus_right()

    def action_start(self) -> None:
        self.back_view, self.pending = None, None
        self.show_view("start")
        self.query_one(GroupList).focus()

    def action_back(self) -> None:
        if self.view == "detail" and self.back_view:
            back, self.back_view = self.back_view, None
            self.show_view(back)
            self._focus_right()
            return
        if self.view == "changes" and self.changes_from:
            back, self.changes_from = self.changes_from, None
            self.pending = None
            self.show_view(back)
            self._focus_right()
            return
        if self.view != "start":
            self.action_start()
            return
        if self.filter_text:
            self.filter_text = ""
            self._rebuild_list()
            self._render_view()
        else:
            self.m365.action_go_home()

    def action_tab(self, mode: str) -> None:
        if mode == self.mode:
            return
        self.mode = mode
        self.filter_text = ""
        self.back_view = None
        self.subject = None
        self.view = "start"
        self._rebuild_list()
        self._render_view()
        self.query_one(GroupList).focus()
        self.refresh_bindings()

    def action_next_tab(self) -> None:
        keys = [k for k, _t in TABS]
        self.action_tab(keys[(keys.index(self.mode) + 1) % len(keys)])

    def action_filter(self) -> None:
        before = self.filter_text

        def live(text: str) -> None:
            self.filter_text = text.strip()
            self._rebuild_list()
            if self.view == "detail":
                self._render_view()

        def done(value: str | None) -> None:
            self.filter_text = before if value is None else value.strip()
            self._rebuild_list()
            self._render_view()
            self.query_one(GroupList).focus()

        self.app.push_screen(PromptModal(f"Filter {self.mode}", self.filter_text,
                                         placeholder="name, upn, bundle, title ...", on_change=live, bottom=True),
                             done)

    def select_row(self, row_id: str) -> bool:
        ids = [r.id for r in self.rows]
        if row_id not in ids:
            return False
        self._last_hl = row_id
        self.query_one(GroupList).highlighted = ids.index(row_id)
        return True

    def jump_to_finding(self) -> None:
        table = self.query_one("#drift-table", VimTable)
        row = table.cursor_row
        if not 0 <= row < len(self.findings):
            return
        path = self.findings[row].path
        reg = self.registry
        is_group = reg.group(path) is not None or (self.snap is not None and self.snap.group(path) is not None)
        mode = "groups" if is_group else "people"
        row_id = ("g:" + (reg.group(path) or self.snap.group(path)).name  # type: ignore[union-attr]
                  if is_group else "p:" + path.lower())
        self.mode, self.filter_text = mode, ""
        self.view = "detail"
        self.subject = row_id
        self._rebuild_list(keep=row_id)
        self.back_view = "drift"
        self.subject = row_id
        self._render_view()
        self.refresh_bindings()
        self.query_one("#detail").focus()

    # -------------------------------------------------------------------------------------------- bindings state

    def check_action(self, action: str, parameters: tuple[object, ...]) -> bool | None:
        if action == "accept":
            return self.view == "suggest" and self.suggestion is not None
        if action == "apply":
            return self.view == "changes" and self.pending is not None
        return True

    # -------------------------------------------------------------------------------------------- operations

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_inventory(self) -> None:
        ok, snap = await self.run_op("inventory", lambda job: job.inventory())
        if ok and snap is not None:
            self.notify(f"Inventory done: {len(snap.people)} people, {len(snap.groups)} groups.")

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_drift(self) -> None:
        ok, findings = await self.run_op("drift", lambda job: job.drift())
        if not ok:
            return
        self.findings = sorted(findings or [], key=lambda f: (SEV_ORDER.get(str(f.severity), 9), f.category,
                                                              f.path.lower()))
        self._render_drift()
        self.show_view("drift")
        self._focus_right()

    def _render_drift(self) -> None:
        from collections import Counter

        c = Counter(f.category for f in self.findings)
        color = {f.category: SEVERITY_STYLE.get(str(f.severity), SUB) for f in self.findings}
        head = Text()
        for i, (cat, n) in enumerate(sorted(c.items(), key=lambda kv: (-kv[1], kv[0]))):
            head.append("   " if i else "").append("● ", style=color[cat]).append(f"{n} ", style=f"bold {FG}")
            head.append(FINDING_LABEL.get(cat, cat), style=SUB)
        if not self.findings:
            head = Text("● No drift: the tenant matches the registry.", style=theme.GREEN)
        self.query_one("#drift-counts", Static).update(head)
        table = self.query_one("#drift-table", VimTable)
        table.clear()
        for i, f in enumerate(self.findings):
            color = SEVERITY_STYLE.get(str(f.severity), SUB)
            what = Text.assemble((FINDING_LABEL.get(f.category, f.category), f"bold {FG}"), (f"  {f.detail}", SUB))
            what.no_wrap, what.overflow = True, "ellipsis"
            table.add_row(theme.glyph_word(str(f.severity), color), Text(f.path, style=FG, no_wrap=True,
                                                                          overflow="ellipsis"),
                          what, Text(f.fix or "", style=theme.GREEN, no_wrap=True, overflow="ellipsis"),
                          key=str(i))
        self._render_drift_fix()

    def _render_drift_fix(self) -> None:
        table = self.query_one("#drift-table", VimTable)
        row = table.cursor_row
        note = self.query_one("#drift-fix", Static)
        if not self.findings or not 0 <= row < len(self.findings):
            note.update("")
            return
        f = self.findings[row]
        note.update(Text.assemble((f.path, f"bold {FG}"), ("  ", ""), (f.detail, SUB), ("\nfix  ", DIM),
                                  (f.fix or "no suggestion", theme.GREEN if f.fix else DIM),
                                  ("\nenter ", f"bold {theme.BLUE}"), ("opens this person or group", SUB)))

    # -- suggest -----------------------------------------------------------------------------------

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_suggest(self) -> None:
        ok, sug = await self.run_op("suggest", lambda job: job.suggest())
        if not ok or sug is None:
            return
        self.suggestion = sug
        self._render_suggest()
        self.show_view("suggest")
        self._focus_right()

    def _render_suggest(self) -> None:
        sug = self.suggestion
        n_near = sum(1 for m in sug.near if m.candidate)
        self.query_one("#suggest-head", Static).update(Text.assemble(
            theme.pill(f"{len(sug.candidates)} candidates", ACCENT), ("  ", ""),
            (f"{n_near} people close to one", SUB), ("   a ", f"bold {theme.BLUE}"),
            ("accepts the highlighted candidate", SUB)))
        table = self.query_one("#suggest-table", VimTable)
        table.clear()
        for c in sug.candidates:
            table.add_row(Text(c.key, style=DIM), Text(c.name, style=f"bold {FG}"), Text(str(len(c.people)),
                                                                                           style=FG),
                          Text(", ".join(c.groups), style=SUB, no_wrap=True, overflow="ellipsis"),
                          Text(c.existing or "new", style=theme.GREEN if c.existing else ORANGE), key=c.key)
        self._render_candidate()

    def _current_candidate(self) -> Any:
        sug = self.suggestion
        if sug is None or not sug.candidates:
            return None
        row = self.query_one("#suggest-table", VimTable).cursor_row
        return sug.candidates[row] if 0 <= row < len(sug.candidates) else None

    def _render_candidate(self) -> None:
        c = self._current_candidate()
        note = self.query_one("#suggest-detail", Static)
        if c is None:
            note.update(Text("No candidates: fewer than two people share the same set of groups.", style=SUB))
            return
        near = [m for m in self.suggestion.near if m.candidate == c.key]
        t = Table.grid(padding=(0, 2))
        t.add_column(style=SUB, no_wrap=True, width=8)
        t.add_column(style=FG)
        t.add_row("groups", ", ".join(c.groups))
        t.add_row("people", ", ".join(self._names(u) for u in c.people))
        if near:
            close = []
            for m in near:
                extra = " ".join([*(f"+{a}" for a in m.addons), *(f"-{x}" for x in m.missing)])
                close.append(f"{self._names(m.upn)} ({extra})" if extra else self._names(m.upn))
            t.add_row("close", Text("; ".join(close), style=SUB))
        note.update(t)

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_accept(self) -> None:
        c = self._current_candidate()
        if c is None:
            return
        name = await self.app.push_screen_wait(PromptModal(
            "Accept as bundle", c.existing or c.name, placeholder="bundle name, e.g. Servicedesk",
            note=f"{len(c.people)} people share these {len(c.groups)} groups. Name the bundle; people close to it "
                 "get it too, with their extra groups as add-ons."))
        if not name or not name.strip():
            return
        name = name.strip()
        sug, key = self.suggestion, c.key
        try:
            changes = self.job.accept(sug, {key: name}, apply=False)
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.m365.report_error(exc, "accept")
            return

        async def apply(job: Any) -> list[str]:
            return job.accept(sug, {key: name}, apply=True)

        def done(result: Any) -> None:
            self._show_changes(f"Bundle {name}: registry updated", [("registry", ch, "", "done") for ch in result],
                               origin="suggest")

        self._show_changes(f"Accept '{name}': changes to the registry",
                           [("registry", ch, "", "planned") for ch in changes],
                           pending=("apply", f"Write bundle '{name}' to the registry?", apply, done),
                           origin="suggest")

    # -- previews and apply -------------------------------------------------------------------------

    def _show_changes(self, title: str, rows: list[tuple[str, str, str, str]], *, notes: Text | None = None,
                      pending: Any = None, origin: str | None = None) -> None:
        self.pending = pending
        self.changes_from = origin
        n = sum(1 for r in rows if r[3] == "planned")
        head = Text.assemble((title, f"bold {FG}"), ("   ", ""))
        if pending is not None:
            head.append_text(theme.pill(f"preview · {n} to do", YELLOW)).append("  nothing written yet", style=SUB)
        elif any(r[3] == "planned" for r in rows):
            head.append("nothing to apply", style=SUB)
        self.query_one("#changes-head", Static).update(head)
        table = self.query_one("#changes-table", VimTable)
        table.clear()
        for i, (action, subject, detail, status) in enumerate(rows):
            table.add_row(Text(action, style=SUB), Text(subject, style=FG, no_wrap=True, overflow="ellipsis"),
                          Text(detail, style=SUB, no_wrap=True, overflow="ellipsis"), change_status(status),
                          key=str(i))
        out = Text()
        if notes is not None:
            out.append_text(notes)
        if pending is not None:
            if out:
                out.append("\n\n")
            out.append("A ", style=f"bold {theme.BLUE}").append("applies these changes (you type ", style=SUB)
            out.append(pending[0], style=f"bold {theme.RED}").append(")   ", style=SUB)
            out.append("esc ", style=f"bold {theme.BLUE}").append("drops the preview", style=SUB)
        elif not rows:
            out.append("Nothing to do.", style=theme.GREEN)
        self.query_one("#changes-notes", Static).update(out)
        self.show_view("changes")
        self._focus_right()

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_apply(self) -> None:
        if self.pending is None:
            return
        word, question, apply, done = self.pending
        body = Text.assemble((question, FG), ("\n\nThis changes ", SUB),
                             ("the registry file" if word == "apply" and self.changes_from == "suggest"
                              else "the tenant and the registry", f"bold {FG}"), (".", SUB))
        ok = await self.app.push_screen_wait(TypedConfirmModal(question, body, word))
        if not ok:
            return
        self.pending = None
        ok, result = await self.run_op(f"{word}", apply)
        if ok:
            self.snap_stale = self.snap is not None
            done(result)
            self._render_state()
            self.notify("Applied. The registry was saved." + (" Press i to refresh the snapshot."
                                                              if self.snap_stale else ""))

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_sync(self) -> None:
        ok, plan = await self.run_op("sync preview", lambda job: job.sync(apply=False))
        if ok and plan is not None:
            self._show_changes("Sync: registry → tenant", _sync_rows(plan, False),
                               notes=Text("Creates registry groups that do not exist in the tenant yet, and fills "
                                          "in group ids in the registry.", style=SUB),
                               pending=None if plan.empty else (
                                   "apply", f"Create {theme.plural(len(plan.create), 'group')} in the tenant?",
                                   lambda job: job.sync(apply=True),
                                   lambda p: self._show_changes("Sync: done", _sync_rows(p, True)))
                               )

    @work(exclusive=True, group="gop", exit_on_error=False)
    async def action_import_groups(self) -> None:
        ok, entries = await self.run_op("import preview", lambda job: job.import_(apply=False))
        if ok and entries is not None:
            rows = [("add to registry", e.name, f"kind {e.kind}", "planned") for e in entries]
            self._show_changes("Import: tenant → registry", rows,
                               notes=Text("Adds prefix groups that exist in the tenant but not in the registry. "
                                          "Fill in purpose and owner afterwards with e.", style=SUB),
                               pending=None if not entries else (
                                   "apply", f"Add {theme.plural(len(entries), 'group')} to the registry?",
                                   lambda job: job.import_(apply=True),
                                   lambda es: self._show_changes("Import: done", [
                                       ("add to registry", e.name, f"kind {e.kind}", "done") for e in es])))

    # -- registry edits ----------------------------------------------------------------------------

    def _selected_group(self) -> str | None:
        if self.subject and self.subject.startswith("g:") and self.view == "detail":
            return self.subject[2:]
        row = self.current_row
        return row.key if row is not None and row.kind == "group" else None

    @work(exclusive=True, group="gedit", exit_on_error=False)
    async def action_edit(self) -> None:
        if self.m365.busy:
            self.notify(f"Busy: {self.m365.busy} is running.", severity="warning")
            return
        name = self._selected_group()
        if name is None or self.job is None:
            self.notify("Select a group first (2 Groups), then e edits its purpose and owner.", severity="warning")
            return
        reg = self.registry
        entry = reg.group(name)
        purpose = await self.app.push_screen_wait(PromptModal(
            f"{name} · purpose", entry.purpose if entry else "", placeholder="what this group gives access to",
            note="One line a colleague understands, e.g. 'Edit access to the Finance site'."))
        if purpose is None:
            return
        owner = await self.app.push_screen_wait(PromptModal(
            f"{name} · owner", entry.owner if entry else "", placeholder="owner@contoso.com",
            note="Who approves new members (a person or a team)."))
        if owner is None:
            return
        if entry is None:
            g = self.snap.group(name) if self.snap else None
            entry = GroupEntry(name=name, id=g.id if g else "", kind=kind_of(name, reg.prefix))
            reg.groups[name] = entry
        entry.purpose, entry.owner = purpose.strip(), owner.strip()
        try:
            path = self.job.save_registry()
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.m365.report_error(exc, "save registry")
            return
        self.refresh_all()
        self.notify(f"Saved {name} to {path.name}.")

    def action_reload(self) -> None:
        if self.m365.busy:
            self.notify(f"Busy: {self.m365.busy} is running.", severity="warning")
            return
        if self.job is None:
            self.load()
            return
        try:
            reg = self.job.reload_registry()
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.m365.report_error(exc, "reload registry")
            return
        self.refresh_all()
        self.notify(f"Registry reloaded: {len(reg.groups)} groups, {len(reg.bundles)} bundles, "
                    f"{len(reg.people)} people.")

    @work(exclusive=True, group="gedit", exit_on_error=False)
    async def action_domains(self) -> None:
        if self.m365.busy:
            self.notify(f"Busy: {self.m365.busy} is running.", severity="warning")
            return
        if self.job is None:
            self.notify(self.error or "No tenant.", severity="error")
            return
        reg = self.registry
        account = self.m365.account.split(",")[0].strip()
        guess = ", ".join(reg.domains) or (account.rsplit("@", 1)[1] if "@" in account else "")
        value = await self.app.push_screen_wait(PromptModal(
            "Your own domains", guess, placeholder="contoso.com, contoso.nl",
            note="Only people whose sign-in name ends in one of these count as colleagues. Partners, guests and "
                 "test accounts on other domains are left out. Empty = everyone."))
        if value is None:
            return
        reg.domains = [d.strip().lower().lstrip("@") for d in value.split(",") if d.strip()]
        try:
            self.job.save_registry()
        except Exception as exc:  # noqa: BLE001 - shown to the user
            self.m365.report_error(exc, "save registry")
            return
        self.refresh_all()
        again = await self.app.push_screen_wait(ConfirmModal(
            "Run the inventory again?", "The people list changes only after a new inventory. Run it now?"))
        if again:
            self.action_inventory()

    def action_toggle_log(self) -> None:
        self.log_open = not self.log_open
        self.query_one("#glog").display = self.log_open

    def action_report(self) -> None:
        self.m365.show_report_path(self.job.run_dir if self.job is not None else None)

    # -------------------------------------------------------------------------------------------- help, palette

    def action_help(self) -> None:
        from .workspace import BACK_HELP, help_group

        topic = self.mode if self.view in ("detail", "start") else self.view
        browse = [("1 2 3", "Groups / People / Bundles"), ("0", "get started checklist"), ("t", "next tab"), ("j / k", "move"),
                  ("/", "filter the list"), ("enter", "open in the right pane"), ("tab", "switch list / right pane"),
                  ("esc", "back"), ("L", "show / hide the log")]
        actions = [("i", "inventory: read the tenant, writes nothing"), ("d", "drift: registry vs tenant"),
                   ("s", "suggest bundles (a accepts one)"), ("y", "sync: create registry groups in the tenant"),
                   ("m", "import tenant groups into the registry"), ("e", "edit purpose / owner of a group"),
                   ("r", "reload the registry file"), ("D", "your own domains (who counts as a colleague)"), ("A", "apply a preview (asks a typed word)")]
        glob = [(":", "command palette"), ("ctrl+l", "sign in"), ("?", "this help"), ("g", "toolbox home"),
                ("q", "quit")]
        cols = Table.grid(expand=True, padding=(0, 4))
        cols.add_column(ratio=1)
        cols.add_column(ratio=1)
        legend = [("●", "matches the registry"), ("▲", "drift, or not in the registry (groups)"),
                  ("○", "not in the registry yet / group not created"), ("◌", "registered, gone from the tenant"),
                  ("⊘", "offboarded")]
        cols.add_row(Group(help_group("browse", browse), Text(""), help_group("anywhere", glob), Text(""),
                           help_group("glyphs in the list", legend)),
                     Group(help_group("actions", actions), Text(""), help_group("back and switch", BACK_HELP)))
        body = Group(Text(HELP[topic], style=SUB), Text(""), cols)
        self.app.push_screen(TextModal(f"Keys · groups · {topic}", body, fit=True))

    def palette_commands(self) -> list[tuple[str, Any, str]]:
        return [
            ("Inventory", self.action_inventory, "read groups, users, licences, apps and roles; writes nothing  (i)"),
            ("Drift", self.action_drift, "compare the registry with the tenant  (d)"),
            ("Suggest bundles", self.action_suggest, "candidate bundles from who has what  (s)"),
            ("Sync registry to tenant", self.action_sync, "create missing registry groups, preview first  (y)"),
            ("Import tenant groups", self.action_import_groups, "add unregistered groups to the registry  (m)"),
            ("Edit group purpose and owner", self.action_edit, "for the selected group  (e)"),
            ("Reload registry", self.action_reload, "re-read the registry file  (r)"),
            ("Show people", lambda: self.action_tab("people"), "(1)"),
            ("Show groups", lambda: self.action_tab("groups"), "(2)"),
            ("Show bundles", lambda: self.action_tab("bundles"), "(3)"),
            ("Filter", self.action_filter, "(/)"),
            ("Toggle log", self.action_toggle_log, "(L)"),
            ("Report path", self.action_report, "the last run's report.html"),
            ("Help", self.action_help, "keys and what this view means  (?)"),
        ]


def _maker(row: Row) -> Callable[[bool], Text]:
    def make(selected: bool) -> Text:
        return row.prompt(selected)

    return make


def _member_rows(plan: Any, action: str, reg: Registry | None = None) -> list[tuple[str, str, str, str]]:
    def purpose(name: str) -> str:
        e = reg.group(name) if reg is not None else None
        return e.purpose if e else ""

    return [(action, c.group, c.error or purpose(c.group), c.status) for c in plan.changes]


def _sync_rows(plan: Any, applied: bool) -> list[tuple[str, str, str, str]]:
    created, failed = dict(plan.created), dict(plan.failed)
    rows = []
    for e in plan.create:
        status = "planned" if not applied else ("done" if e.name in created else "failed")
        rows.append(("create group", e.name, failed.get(e.name) or e.purpose, status))
    rows += [("fill id", n, gid, "done" if applied else "planned") for n, gid in plan.fill_ids]
    return rows


