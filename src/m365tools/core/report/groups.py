"""Reports of the groups tool (inventory, drift, suggest, sync/import, onboard/offboard)."""

from __future__ import annotations

from collections import Counter
from datetime import datetime
from typing import Any

from .model import Report, Section

_SEV = {"blocking": 0, "warn": 1, "info": 2}


def _report(title: str, mode: str, tenant: str, registry: str, started: datetime, cards: list[tuple[str, str]],
            sections: list[Section]) -> Report:
    return Report(title=title, mode=mode, source=f"tenant {tenant}", destination=f"registry {registry}",
                  started=started, finished=datetime.now(), summary=cards, sections=sections)


def findings_section(findings: list[Any]) -> Section:
    ordered = sorted(findings, key=lambda f: (_SEV.get(str(f.severity), 9), f.category, f.path.lower()))
    return Section("Findings", "table", ["Severity", "Category", "Subject", "Detail", "Fix"],
                   [[str(f.severity), f.category, f.path, f.detail, f.fix] for f in ordered],
                   note=f"{len(findings)} finding(s), grouped by severity.", severity_col=0)


def inventory_report(snap: Any, findings: list[Any], tenant: str, registry: str, started: datetime) -> Report:
    sev = Counter(str(f.severity) for f in findings)
    cards = [("Groups", str(len(snap.groups))), ("People", str(len(snap.people))), ("Apps", str(len(snap.apps))),
             ("Roles", str(len(snap.roles))), ("Warnings", str(sev["warn"])), ("Info", str(sev["info"]))]
    groups = [[g.name, str(len(g.members)), ", ".join(g.owners) or "-", g.description] for g in snap.groups]
    people = [[p.upn, p.display_name, p.job_title, p.department,
               ", ".join(sorted(snap.groups_of(p.upn), key=str.lower)), ", ".join(p.licences)] for p in snap.people]
    apps = [[a.name, ", ".join(a.users) or "-", ", ".join(a.groups) or "-"] for a in snap.apps]
    roles = [[r.name, ", ".join(r.members) or "-"] for r in snap.roles]
    return _report("Access groups inventory", "inventory", tenant, registry, started, cards, [
        Section("Groups", "table", ["Group", "Members", "Owners", "Description"], groups),
        Section("People", "table", ["UPN", "Name", "Job title", "Department", "Groups", "Licences"], people),
        Section("Enterprise apps (assignment required)", "table", ["App", "Users", "Groups"], apps),
        Section("Directory roles", "table", ["Role", "Members"], roles),
        findings_section(findings),
    ])


def drift_report(findings: list[Any], tenant: str, registry: str, started: datetime, title: str = "Access drift"
                 ) -> Report:
    sev = Counter(str(f.severity) for f in findings)
    cat = Counter(f.category for f in findings)
    cards = [("Warnings", str(sev["warn"])), ("Info", str(sev["info"]))]
    return _report(title, "drift", tenant, registry, started, cards, [
        Section("By category", "table", ["Category", "Count"], [[k, str(v)] for k, v in sorted(cat.items())]),
        findings_section(findings),
    ])


def suggest_report(sug: Any, tenant: str, registry: str, started: datetime) -> Report:
    names = {c.key: c.name for c in sug.candidates}
    cards = [("Candidates", str(len(sug.candidates))), ("Others", str(len(sug.near)))]
    return _report("Bundle suggestions", "suggest", tenant, registry, started, cards, [
        Section("Candidate bundles", "table", ["#", "Proposed name", "Existing", "Groups", "People"],
                [[c.key, c.name, c.existing or "-", ", ".join(c.groups), ", ".join(c.people)]
                 for c in sug.candidates]),
        Section("Nearest candidate for everyone else", "table", ["Person", "Candidate", "Score", "+ addons",
                                                                 "- missing"],
                [[m.upn, names.get(m.candidate, "-"), f"{m.score:.2f}", ", ".join(m.addons),
                  ", ".join(m.missing)] for m in sug.near]),
    ])


def changes_report(title: str, mode: str, rows: list[list[str]], applied: bool, tenant: str, registry: str,
                   started: datetime, extra: list[Section] | None = None) -> Report:
    cards = [("Changes", str(len(rows))), ("Applied", "yes" if applied else "no (preview)")]
    return _report(title, mode, tenant, registry, started, cards,
                   [Section("Changes", "table", ["Action", "Subject", "Detail", "Status"], rows), *(extra or [])])
