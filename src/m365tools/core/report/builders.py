"""Turn spmig objects into Reports."""

from __future__ import annotations

import dataclasses
from collections import Counter
from collections.abc import Iterable
from typing import Any

from m365tools.tools.spmig.models import (
    Action,
    CompareResult,
    Finding,
    Inventory,
    LibraryRef,
    Plan,
    RunSummary,
    Severity,
)

from .format import fmt_duration, fmt_size
from .model import Report, Section

_SEV_ORDER = {Severity.BLOCKING: 0, Severity.WARN: 1, Severity.INFO: 2}
NOT_MIGRATED = [
    "Version history (only the latest version of each file is copied)",
    "Unique (item-level) permissions and sharing groups",
    "Custom column values (only standard metadata: author, editor, dates)",
    "Sharing links (anonymous and specific-people links must be re-created)",
]


def _new(summary: RunSummary, title: str, src: str, dst: str, cards: list[tuple[str, str]],
         sections: list[Section], not_migrated: list[str] | None = None) -> Report:
    return Report(title=title, mode=summary.mode, source=src, destination=dst, started=summary.started,
                  finished=summary.finished, summary=cards, sections=sections,
                  not_migrated=not_migrated or [])


def _lib_desc(lib: LibraryRef | None) -> str:
    return f"{lib.site_url} / {lib.title}" if lib else "-"


def _yn(b: bool) -> str:
    return "yes" if b else "no"


def audit_report(summary: RunSummary, src_inv: Inventory, dst_inv: Inventory | None, src_lib: LibraryRef,
                 dst_lib: LibraryRef | None, findings: list[Finding], opts_desc: dict) -> Report:
    files = src_inv.files()
    folders = src_inv.folders()
    sev = Counter(f.severity for f in findings)
    cards = [
        ("Files", f"{len(files):,}"), ("Folders", f"{len(folders):,}"),
        ("Total size", fmt_size(src_inv.total_bytes())),
        ("Blocking", str(sev[Severity.BLOCKING])), ("Warnings", str(sev[Severity.WARN])),
        ("Info", str(sev[Severity.INFO])), ("Duration", fmt_duration(summary.duration_s)),
    ]
    kv = [["Library", src_lib.title], ["Site", src_lib.site_url], ["Items (reported)", f"{src_lib.item_count:,}"],
          ["Files", f"{len(files):,}"], ["Folders", f"{len(folders):,}"],
          ["Total size", fmt_size(src_inv.total_bytes())],
          ["Unique permissions", str(sum(1 for i in src_inv.items if i.unique_permissions))]]
    kv += [[str(k), str(v)] for k, v in opts_desc.items()]
    if dst_inv is not None:
        kv.append(["Destination files", f"{len(dst_inv.files()):,}"])

    def sett(lib: LibraryRef | None, name: str, conv=_yn) -> str:
        return conv(getattr(lib, name)) if lib else "-"

    rows = []
    for label, attr, conv in [("Versioning", "enable_versioning", _yn),
                              ("Major version limit", "major_version_limit", str),
                              ("Minor versions", "enable_minor_versions", _yn),
                              ("Force checkout", "force_checkout", _yn),
                              ("Content approval", "enable_moderation", _yn)]:
        s, d = sett(src_lib, attr, conv), sett(dst_lib, attr, conv)
        rows.append([label, s, d, "differs" if dst_lib and s != d else ""])
    ordered = sorted(findings, key=lambda f: _SEV_ORDER[f.severity])
    ext = Counter((i.ext or "(none)") for i in files)
    ext_size: Counter[str] = Counter()
    for i in files:
        ext_size[i.ext or "(none)"] += i.size
    big = sorted(files, key=lambda i: i.size, reverse=True)[:10]
    top: dict[str, list[int]] = {}
    for i in files:
        t = top.setdefault(i.top_folder, [0, 0])
        t[0] += 1
        t[1] += i.size
    top_rows = sorted(top.items(), key=lambda kv_: kv_[1][1], reverse=True)
    sections = [
        Section("Source overview", "kv", rows=kv),
        Section("Library settings", "table", ["Setting", "Source", "Destination", ""], rows),
        Section("Findings", "table", ["Severity", "Category", "Path", "Detail", "Fix"],
                [[f.severity.value, f.category, f.path, f.detail, f.fix] for f in ordered],
                note=f"{len(findings)} finding(s), grouped by severity.", severity_col=0),
        Section("File types (top 15)", "table", ["Extension", "Files", "Size"],
                [[e, f"{n:,}", fmt_size(ext_size[e])] for e, n in ext.most_common(15)]),
        Section("Largest files (top 10)", "table", ["Path", "Size", "Modified"],
                [[i.rel_path, fmt_size(i.size), f"{i.modified:%Y-%m-%d}" if i.modified else ""] for i in big]),
        Section("Top folders", "table", ["Folder", "Files", "Size"],
                [[k, f"{v[0]:,}", fmt_size(v[1])] for k, v in top_rows]),
    ]
    return _new(summary, "SharePoint migration audit", _lib_desc(src_lib), _lib_desc(dst_lib), cards,
                sections, NOT_MIGRATED)


def dryrun_report(plan: Plan, summary: RunSummary) -> Report:
    counts = plan.counts()
    cards = [("To copy", f"{counts[Action.COPY]:,}"), ("Size to copy", fmt_size(plan.bytes_to_copy())),
             ("Folders to create", f"{counts[Action.CREATE_FOLDER]:,}"),
             ("Skipped", f"{counts[Action.SKIP]:,}"), ("Conflicts", f"{counts[Action.CONFLICT]:,}"),
             ("Excluded", f"{counts[Action.EXCLUDED]:,}")]
    conflicts = [e for e in plan.entries if e.action == Action.CONFLICT]
    excluded = [e for e in plan.entries if e.action == Action.EXCLUDED]
    routed = [e for e in plan.entries if e.route]
    sections = [
        Section("Counts", "table", ["Action", "Items"], [[a.value, f"{n:,}"] for a, n in counts.items()]),
        Section("Conflicts", "table", ["Path", "Reason"], [[e.item.rel_path, e.reason] for e in conflicts]),
        Section("Excluded", "table", ["Path", "Reason"], [[e.item.rel_path, e.reason] for e in excluded],
                collapsed=True),
        Section("Routed to other destinations", "table", ["Path", "Route", "Target"],
                [[e.item.rel_path, e.route or "", e.target_server_relative_url] for e in routed], collapsed=True),
    ]
    return _new(summary, "SharePoint migration dry run", _lib_desc(plan.src), _lib_desc(plan.dst), cards,
                sections, NOT_MIGRATED)


def migrate_report(summary: RunSummary, plan: Plan, results_rows: list[dict]) -> Report:
    mbps = summary.bytes_copied / 1024 / 1024 / summary.duration_s if summary.duration_s > 0 else 0.0
    cards = [("Copied", f"{summary.copied:,}"), ("Skipped", f"{summary.skipped:,}"),
             ("Failed", f"{summary.failed:,}"), ("Warnings", f"{summary.metadata_warnings:,}"),
             ("Data copied", fmt_size(summary.bytes_copied)), ("Duration", fmt_duration(summary.duration_s)),
             ("Speed", f"{mbps:.2f} MB/s")]
    failed = [r for r in results_rows if str(r.get("status", "")).lower() in ("failed", "error")]
    warns = [r for r in results_rows if r.get("metadata_warning") or str(r.get("status", "")).lower() == "warning"]
    sections = [
        Section("Failed files", "table", ["Path", "Error"],
                [[str(r.get("path", r.get("rel_path", ""))), str(r.get("error", r.get("detail", "")))]
                 for r in failed]),
        Section("Metadata warnings", "table", ["Path", "Warning"],
                [[str(r.get("path", r.get("rel_path", ""))),
                  str(r.get("metadata_warning") or r.get("detail", ""))] for r in warns]),
    ]
    if summary.stopped_early:
        sections.insert(0, Section("Run status", "text", text="The run was stopped early; resume to continue."))
    return _new(summary, "SharePoint migration", _lib_desc(plan.src), _lib_desc(plan.dst), cards, sections,
                NOT_MIGRATED)


def verify_report(result: CompareResult, summary: RunSummary) -> Report:
    c = result.counts()
    cards = [(k.replace("_", " ").capitalize(), f"{c.get(k, 0):,}")
             for k in ("ok", "missing", "size_differs", "source_newer", "only_in_destination")]

    def sz(n: int | None) -> str:
        return fmt_size(n) if n is not None else "-"

    def date(d) -> str:
        return f"{d:%Y-%m-%d %H:%M}" if d else ""

    def tbl(title: str, status: str, cols: list[str], fn, collapsed=False) -> Section:
        return Section(title, "table", cols, [fn(r) for r in result.rows if r.status == status],
                       collapsed=collapsed)

    sections = [
        Section("Counts", "table", ["Status", "Items"], [[k, f"{v:,}"] for k, v in sorted(c.items())]),
        tbl("Missing in destination", "missing", ["Path", "Source size"], lambda r: [r.rel_path, sz(r.src_size)]),
        tbl("Size differs", "size_differs", ["Path", "Source", "Destination", "Detail"],
            lambda r: [r.rel_path, sz(r.src_size), sz(r.dst_size), r.detail]),
        tbl("Source newer", "source_newer", ["Path", "Source modified", "Destination modified"],
            lambda r: [r.rel_path, date(r.src_modified), date(r.dst_modified)]),
        tbl("Only in destination", "only_in_destination", ["Path", "Size"],
            lambda r: [r.rel_path, sz(r.dst_size)], collapsed=True),
    ]
    return _new(summary, "SharePoint migration verification", "source", "destination", cards, sections)


def _get(obj: Any, key: str, default: str = "") -> Any:
    if isinstance(obj, dict):
        return obj.get(key, default)
    return getattr(obj, key, default)


def lock_report(changes: Iterable[Any], applied: bool) -> Report:
    changes = list(changes)
    now = __import__("datetime").datetime.now()
    rows = []
    for ch in changes:
        if dataclasses.is_dataclass(ch) and not isinstance(ch, type):
            ch = dataclasses.asdict(ch)
        rows.append([str(_get(ch, "principal", _get(ch, "user", ""))), str(_get(ch, "before", _get(ch, "old", ""))),
                     str(_get(ch, "after", _get(ch, "new", "")))])
    status = "Applied" if applied else "Preview only (nothing changed)"
    return Report(
        title="Source library lock", mode="lock", source="", destination="", started=now, finished=now,
        summary=[("Status", status), ("Principals", str(len(rows)))],
        sections=[Section("Permission changes", "table", ["Principal", "Before", "After"], rows,
                          note="Only principals with Edit/Add/Delete/ManageLists are downgraded.")])
