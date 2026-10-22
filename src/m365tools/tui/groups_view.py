"""Rich renderers for the groups screen: list rows, the state ribbon and the person / group / bundle details.

Pure functions over the engine's data (Snapshot, Registry, the dicts of ``explain``); no widgets here."""

from __future__ import annotations

from collections import Counter
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from rich.console import Group, RenderableType
from rich.table import Table
from rich.text import Text

from m365tools.tools.groups.drift import person_findings
from m365tools.tools.groups.models import Finding, Registry, Snapshot, norm
from m365tools.tools.groups.registry import bundles_with

from . import theme
from .theme import BLUE, CRUST, DIM, FG, GREEN, MAUVE, PEACH, RED, SUB, TEAL, YELLOW

LIST_TEXT = 34  # list width minus padding and scrollbar

# ------------------------------------------------------------------------------------------------ states

# person / group states in the list: glyph, colour, word
STATE = {
    "ok": ("●", GREEN, "ok"),
    "drift": ("▲", YELLOW, "drift"),
    "unassigned": ("○", SUB, "not in the registry"),
    "gone": ("◌", PEACH, "not in the tenant"),
    "offboarded": ("⊘", DIM, "offboarded"),
    "unknown": ("·", DIM, "no snapshot"),
    "unregistered": ("▲", YELLOW, "not in the registry"),
    "not_created": ("○", PEACH, "not created in the tenant"),
}
GROUP_STATUS = {"ok": ("●", GREEN), "missing": ("○", PEACH), "unexplained": ("▲", YELLOW)}
CHANGE_STYLE = {"planned": BLUE, "already": DIM, "done": GREEN, "failed": RED, "not_created": PEACH,
                "kept": DIM, "manual": PEACH}
CHANGE_LABEL = {"already": "already member", "not_created": "group not created"}
FINDING_LABEL = {"unexplained": "unexplained", "missing": "missing", "unregistered": "unregistered",
                 "not_created": "not created", "unassigned": "unassigned", "no_owner": "no owner",
                 "empty": "empty group", "no_purpose": "no purpose"}


def person_state(snap: Snapshot | None, reg: Registry, upn: str) -> str:
    entry = reg.person(upn)
    if entry is not None and entry.offboarded:
        return "offboarded"
    if snap is None:
        return "unknown" if entry is not None else "unassigned"
    if snap.person(upn) is None:
        return "gone"
    if entry is None:
        return "unassigned"
    return "drift" if person_findings(snap, reg, upn) else "ok"


def group_state(snap: Snapshot | None, reg: Registry, name: str) -> str:
    registered = reg.group(name) is not None
    if snap is None:
        return "unknown"
    in_tenant = snap.group(name) is not None
    if in_tenant and registered:
        return "ok"
    return "unregistered" if in_tenant else "not_created"


# ------------------------------------------------------------------------------------------------ list rows


@dataclass(slots=True)
class Row:
    """One entry of the left list: ``id`` is ``p:<upn>``, ``g:<name>`` or ``b:<name>``."""

    id: str
    name: str
    hint: str
    state: str
    haystack: str

    @property
    def kind(self) -> str:
        return {"p": "person", "g": "group", "b": "bundle"}[self.id[0]]

    @property
    def key(self) -> str:
        return self.id[2:]

    def prompt(self, selected: bool, width: int = LIST_TEXT) -> Text:
        glyph, color, _w = STATE.get(self.state, ("·", DIM, ""))
        hint = theme_fit(self.hint, max(width - 2 - min(len(self.name), 18) - 2, 6))
        name = theme_fit(self.name, width - 2 - len(hint) - 1)
        pad = max(width - 2 - len(name) - len(hint), 1)
        t = Text(no_wrap=True, overflow="ellipsis")
        if selected:
            t.append(f"{glyph} ", style=f"bold {CRUST}").append(name, style=f"bold {CRUST}")
            return t.append(" " * pad).append(hint, style=CRUST)
        t.append(f"{glyph} ", style=color).append(name, style=FG)
        return t.append(" " * pad).append(hint, style=DIM)


def theme_fit(text: str, width: int) -> str:
    return text if len(text) <= width else text[: max(width - 1, 0)] + "…"


def people_rows(snap: Snapshot | None, reg: Registry) -> list[Row]:
    upns = {p.upn for p in snap.people} if snap else set()
    upns |= set(reg.people)
    rows = []
    for upn in upns:
        p = snap.person(upn) if snap else None
        e = reg.person(upn)
        name = (p.display_name if p and p.display_name else upn.split("@")[0])
        state = person_state(snap, reg, upn)
        if e is not None and e.offboarded:
            hint = "offboarded"
        elif e is not None:
            hint = ", ".join(e.bundles) or ("add-ons only" if e.addons else "no bundle")
        else:
            hint = "unassigned"
        hay = " ".join([upn, name, p.job_title if p else "", p.department if p else "", hint]).lower()
        rows.append(Row(f"p:{upn}", name, hint, state, hay))
    return sorted(rows, key=lambda r: r.name.lower())


def group_rows(snap: Snapshot | None, reg: Registry) -> list[Row]:
    names: dict[str, str] = {norm(g.name): g.name for g in snap.groups} if snap else {}
    for e in reg.groups.values():
        names.setdefault(norm(e.name), e.name)
    rows = []
    for name in names.values():
        g = snap.group(name) if snap else None
        e = reg.group(name)
        hint = f"{len(g.members)}" if g else ("not created" if snap else "")
        hay = " ".join([name, e.purpose if e else "", e.owner if e else "", e.kind if e else ""]).lower()
        rows.append(Row(f"g:{name}", name, hint, group_state(snap, reg, name), hay))
    return sorted(rows, key=lambda r: r.name.lower())


def bundle_rows(reg: Registry) -> list[Row]:
    rows = []
    for b in reg.bundles.values():
        people = sum(1 for p in reg.people.values() if b.name in p.bundles and not p.offboarded)
        hint = f"{len(b.groups)} groups · {people}"
        rows.append(Row(f"b:{b.name}", b.name, hint, "ok", f"{b.name} {b.purpose} {' '.join(b.groups)}".lower()))
    return sorted(rows, key=lambda r: r.name.lower())


def filtered(rows: Iterable[Row], text: str) -> list[Row]:
    words = text.lower().split()
    return [r for r in rows if all(w in r.haystack for w in words)]


# ------------------------------------------------------------------------------------------------ ribbon


def snapshot_age(taken: datetime) -> str:
    secs = (datetime.now() - taken).total_seconds()
    if secs < 90:
        return "just now"
    if secs < 3600:
        return f"{secs / 60:.0f} min old"
    if secs < 48 * 3600:
        return f"{secs / 3600:.0f} h old"
    return f"{secs / 86400:.0f} days old"


def state_ribbon(snap: Snapshot | None, reg: Registry, findings: list[Finding], *, origin: str, stale: bool,
                 registry_name: str) -> Text:
    """Line 1: the snapshot; line 2: the registry; line 3 (with a snapshot): compact coloured counts.
    Lines never wrap: they are cut with an ellipsis."""
    t = Text(no_wrap=True, overflow="ellipsis")
    if snap is None:
        t.append("○ ", style=DIM).append("No snapshot yet. ", style=f"bold {FG}")
        t.append("Press ", style=SUB).append("i", style=f"bold {BLUE}").append(" to run an inventory.", style=SUB)
    else:
        t.append("● ", style=YELLOW if stale else GREEN).append("snapshot ", style=SUB)
        t.append(snapshot_age(snap.taken), style=FG)
        if origin:
            t.append(f" ({origin})", style=DIM)
        if stale:
            t.append(" · changed since: press i", style=YELLOW)
    t.append("\n").append("  registry ", style=SUB).append(registry_name, style=FG)
    t.append(f" · {len(reg.groups)} groups · {len(reg.bundles)} bundles · {len(reg.people)} people", style=SUB)
    if snap is None:
        return t
    c = Counter(f.category for f in findings)
    counts = [(len(snap.people), "people", BLUE), (len(snap.groups), "groups", BLUE),
              (c["unregistered"], "unregistered", YELLOW), (c["unexplained"], "unexplained", YELLOW),
              (c["missing"], "missing", PEACH)]
    t.append("\n")
    for i, (n, label, color) in enumerate(counts):
        if i:
            t.append(" · ", style=DIM)
        t.append("● " if n else "○ ", style=color if n else DIM).append(f"{n} ", style=f"bold {FG}" if n else DIM)
        t.append(label, style=SUB if n else DIM)
    return t


# ------------------------------------------------------------------------------------------------ details


def heading(text: str) -> Text:
    return Text(text, style=f"bold {MAUVE}")


def kv(rows: list[tuple[str, RenderableType]]) -> Table:
    t = Table.grid(padding=(0, 2))
    t.add_column(style=SUB, no_wrap=True, width=11)
    t.add_column(style=FG)
    for k, v in rows:
        t.add_row(k, v)
    return t


def flags(*pairs: tuple[bool, str, str]) -> Text:
    """``● in tenant  ○ not registered``."""
    t = Text()
    for i, (on, yes, no) in enumerate(pairs):
        if i:
            t.append("   ")
        t.append("● " if on else "○ ", style=GREEN if on else PEACH).append(yes if on else no, style=SUB)
    return t


def none(text: str = "none") -> Text:
    return Text(text, style=DIM)


def joined(items: Iterable[str], style: str = FG) -> Text:
    items = list(items)
    return Text(", ".join(items), style=style) if items else none()


def title_line(title: str, right: Text) -> Table:
    g = Table.grid(expand=True)
    g.add_column(no_wrap=True, overflow="ellipsis")
    g.add_column(justify="right", no_wrap=True)
    g.add_row(Text(title, style=f"bold {FG}"), right)
    return g


def person_detail(d: dict[str, Any]) -> Group:
    sub = " · ".join(x for x in (d["upn"], d["job_title"], d["department"]) if x)
    parts: list[RenderableType] = [
        title_line(d["name"] or d["upn"], flags((d["in_tenant"], "in tenant", "not in tenant"),
                                               (d["registered"], "registered", "not registered"))),
        Text(sub, style=SUB),
    ]
    if d["offboarded"]:
        parts.append(Text(f"offboarded on {d['offboarded']}", style=PEACH))
    if d["note"]:
        parts.append(Text(f"note: {d['note']}", style=SUB))
    parts += [Text(""), kv([("bundles", joined(d["bundles"], f"bold {FG}")), ("add-ons", joined(d["addons"]))])]
    parts += [Text(""), heading(f"GROUPS  {len(d['groups'])}")]
    if d["groups"]:
        g = Table.grid(padding=(0, 2))
        g.add_column(no_wrap=True, width=13)
        g.add_column(no_wrap=True)
        g.add_column(style=SUB, no_wrap=True, overflow="ellipsis", ratio=1)
        g.add_column(style=DIM, no_wrap=True, overflow="ellipsis")
        for x in d["groups"]:
            glyph, color = GROUP_STATUS[x["status"]]
            via = ("add-on" if x["via"] == "addon" else f"via {x['via']}") if x["via"] else ""
            g.add_row(theme.glyph_word(x["status"], color, glyph), Text(x["group"], style=FG),
                      Text(x["purpose"] or "no purpose", style=SUB if x["purpose"] else DIM), via)
        parts.append(g)
    else:
        parts.append(none("no prefix groups"))
    if d["unexplained"] or d["missing"]:
        parts.append(Text(""))
        if d["missing"]:
            parts.append(Text.assemble(("○ missing  ", PEACH), ("expected by a bundle or add-on but not a member: ",
                                                                SUB), ("o", f"bold {BLUE}"),
                                       (" onboard re-adds them", SUB)))
        if d["unexplained"]:
            parts.append(Text.assemble(("▲ unexplained  ", YELLOW),
                                       ("member, but no bundle or add-on says why: add an add-on or remove it", SUB)))
    apps = [Text.assemble((a["app"], FG), (f" ({'direct' if a['via'] == 'direct' else 'via ' + a['via']})", DIM))
            for a in d["apps"]]
    app_text = Text(", ").join(apps) if apps else none()
    parts += [Text(""), heading("ALSO HAS"),
              kv([("licences", joined(d["licences"])), ("apps", app_text),
                  ("roles", joined(d["roles"], f"bold {PEACH}"))])]
    return Group(*parts)


def group_detail(d: dict[str, Any], names: Callable[[str], str]) -> Group:
    kind = d["kind"] or "unregistered"
    parts: list[RenderableType] = [
        title_line(d["name"], flags((d["in_tenant"], "in tenant", "not in tenant"),
                                    (d["registered"], "registered", "not registered"))),
        Text(f"{kind} group" + (f" · {d['description']}" if d["description"] else ""), style=SUB),
        Text(""),
        kv([("purpose", Text(d["purpose"], style=FG) if d["purpose"] else none("not set: press e")),
            ("owner", Text(d["owner"], style=FG) if d["owner"] else none("not set: press e")),
            ("Entra owner", joined(d["tenant_owners"]) if d["tenant_owners"] else none("none in Entra")),
            ("bundles", joined(d["bundles"], f"bold {FG}")),
            ("apps", joined(d["apps"])),
            ("id", Text(d["id"] or "", style=DIM))]),
    ]
    addon, unexplained = set(d["addon_of"]), set(d["unexplained"])
    parts += [Text(""), heading(f"MEMBERS  {len(d['members'])}")]
    if d["members"]:
        g = Table.grid(padding=(0, 2))
        g.add_column(no_wrap=True, width=13)
        g.add_column(no_wrap=True)
        g.add_column(style=SUB, no_wrap=True)
        for upn in sorted(d["members"]):
            if upn in unexplained:
                why = theme.glyph_word("unexplained", YELLOW, "▲")
            elif upn in addon:
                why = theme.glyph_word("add-on", TEAL)
            else:
                why = theme.glyph_word("bundle", GREEN)
            g.add_row(why, Text(names(upn), style=FG), Text(upn, style=DIM))
        parts.append(g)
    else:
        parts.append(none("no direct members"))
    holders = sorted(addon - set(d["members"]))
    if holders:
        parts += [Text(""), Text.assemble(("○ add-on but not a member: ", PEACH), (", ".join(holders), FG))]
    return Group(*parts)


def bundle_detail(reg: Registry, snap: Snapshot | None, name: str, names: Callable[[str], str]) -> Group:
    b = reg.bundle(name)
    if b is None:
        return Group(Text(f"No bundle '{name}'.", style=RED))
    people = sorted(p.upn for p in reg.people.values() if b.name in p.bundles and not p.offboarded)
    parts: list[RenderableType] = [
        title_line(b.name, Text(f"{len(b.groups)} groups · {len(people)} people", style=SUB)),
        Text(b.purpose or "no purpose", style=SUB if b.purpose else DIM), Text(""),
        heading(f"GROUPS  {len(b.groups)}"),
    ]
    g = Table.grid(padding=(0, 2))
    g.add_column(no_wrap=True, width=2)
    g.add_column(no_wrap=True)
    g.add_column(style=SUB)
    for gname in b.groups:
        e = reg.group(gname)
        exists = snap is None or snap.group(gname) is not None
        g.add_row(Text("●" if exists else "○", style=GREEN if exists else PEACH), Text(gname, style=FG),
                  Text(e.purpose if e and e.purpose else "no purpose", style=SUB if e and e.purpose else DIM))
    parts += [g, Text(""), heading(f"PEOPLE  {len(people)}")]
    if people:
        p = Table.grid(padding=(0, 2))
        p.add_column(no_wrap=True, width=2)
        p.add_column(no_wrap=True)
        p.add_column(style=DIM, no_wrap=True)
        for upn in people:
            glyph, color, _w = STATE[person_state(snap, reg, upn)]
            p.add_row(Text(glyph, style=color), Text(names(upn), style=FG), upn)
        parts.append(p)
    else:
        parts.append(none("nobody has this bundle yet"))
    others = [x for x in bundles_with_any(reg, b.groups) if x != b.name]
    if others:
        parts += [Text(""), Text.assemble(("shares groups with ", SUB), (", ".join(others), FG))]
    return Group(*parts)


def bundles_with_any(reg: Registry, groups: list[str]) -> list[str]:
    out: list[str] = []
    for gname in groups:
        for b in bundles_with(reg, gname):
            if b not in out:
                out.append(b)
    return out


def change_status(status: str) -> Text:
    return theme.glyph_word(CHANGE_LABEL.get(status, status), CHANGE_STYLE.get(status.split(":")[0], SUB))
