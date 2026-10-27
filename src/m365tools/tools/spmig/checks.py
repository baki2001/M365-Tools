"""Pre-flight checks. Each rule is an async ``(CheckContext) -> list[Finding]`` registered in :data:`CHECKS`.

Severities follow the table in docs/ARCHITECTURE.md. Rules never write anything except two probes at the
destination: the write-access probe creates and deletes one ``__m365_probe_<uuid>`` folder, and the metadata
probe uploads one ``__m365_probe_<uuid>.txt``, sets its dates, reads them back and deletes it.
"""

from __future__ import annotations

import re
from collections.abc import Awaitable, Callable
from dataclasses import dataclass

from m365tools.core.errors import AccessDeniedError, M365Error, NotFoundError
from m365tools.core.events import EventSink, LogLine, PhaseChanged, Stats, null_sink

from .gateway import SpGateway
from .inventory import inventory_base
from .models import (
    MAX_PATH_LEN,
    PATH_WARN_LEN,
    Action,
    Finding,
    Inventory,
    Item,
    JobOptions,
    LibraryRef,
    Plan,
    Severity,
    sp_date_ok,
)
from .plan import ROUTE_UNAVAILABLE, glob_match, is_excluded
from .usermap import ResolvedUsers, UserMap, collect_users

MB = 1024 * 1024
ESTIMATE_MBPS = 0.7
QUOTA_HEADROOM = 1.1
LARGE_FILE_BYTES = 100 * MB
ILLEGAL_CHARS = re.compile(r'["*:<>?/\\|]')
# Names SharePoint Online and OneDrive refuse outright (Microsoft's "Restrictions and limitations in OneDrive and
# SharePoint"). A *leading* dot is allowed there (".gitignore", draw.io's ".$name.drawio.bkp" backups).
RESERVED_NAMES: frozenset[str] = frozenset({
    ".lock", "con", "prn", "aux", "nul", "desktop.ini",
    *(f"com{i}" for i in range(10)), *(f"lpt{i}" for i in range(10)),
})
DRAWIO_BACKUP = "**/.$*.bkp"

# Columns SharePoint (or Microsoft 365 services) add by themselves; never reported as "missing at destination".
SYSTEM_FIELDS: frozenset[str] = frozenset({
    "_ExtendedDescription", "MediaServiceImageTags", "MediaServiceOCR", "MediaServiceAutoTags",
    "MediaServiceLocation", "MediaServiceGenerationTime", "MediaServiceEventHashCode", "MediaServiceDateTaken",
    "MediaServiceKeyPoints", "MediaServiceAutoKeyPoints", "MediaServiceMetadata", "MediaServiceFastMetadata",
    "MediaServiceObjectDetectorVersions", "MediaServiceSearchProperties", "MediaServiceBillingMetadata",
    "MediaLengthInSeconds", "lcf76f155ced4ddcb4097134ff3c332f", "TaxCatchAll", "TaxCatchAllLabel",
    "_dlc_DocId", "_dlc_DocIdUrl", "_dlc_DocIdPersistId", "SharedWithUsers", "SharedWithDetails",
    "_ip_UnifiedCompliancePolicyProperties", "_ip_UnifiedCompliancePolicyUIAction", "ComplianceAssetId",
    "_ComplianceFlags", "_ComplianceTag", "_ComplianceTagWrittenTime", "_ComplianceTagUserId", "_IsRecord",
    "_DisplayName", "_CopySource", "_ModerationComments", "_ModerationStatus", "_CommentCount", "_LikeCount",
    "_CommentFlags", "_Status", "_ColorTag", "_activity", "_VirusStatus", "_VirusVendorID", "_VirusInfo",
    "_RansomwareAnomalyMetaInfo", "_DraftOwnerId", "_MainLinkSettings", "_AuthorByline", "_TopicHeader",
    "_SPSitePageFlags", "_SPCallToAction", "_OriginalSourceUrl", "_OriginalSourceSiteId",
    "_OriginalSourceWebId", "_OriginalSourceListId", "_OriginalSourceItemId", "Title", "FileLeafRef",
    "ContentType", "Created", "Modified", "Author", "Editor", "CheckoutUser", "DocIcon", "LinkFilename",
    "LinkFilenameNoMenu", "FileSizeDisplay", "ItemChildCount", "FolderChildCount", "AppAuthor", "AppEditor",
    "ParentVersionString", "ParentLeafName", "Edit", "SelectTitle", "_UIVersionString", "TemplateUrl",
    "xd_ProgID", "xd_Signature", "Combine", "RepairDocument", "ThumbnailExists", "PreviewExists",
})
_SYSTEM_LOWER = frozenset(f.lower() for f in SYSTEM_FIELDS)


@dataclass
class CheckContext:
    src_inv: Inventory
    dst_inv: Inventory | None
    src_lib: LibraryRef
    dst_lib: LibraryRef | None
    opts: JobOptions
    plan: Plan | None
    src_gw: SpGateway
    dst_gw: SpGateway | None
    usermap: UserMap | None = None
    dst_error: str | None = None  # why the destination library could not be resolved
    resolved_users: ResolvedUsers | None = None  # pre-resolved destination users (Job's "users" phase)

    @property
    def dst_base(self) -> str | None:
        if self.plan is not None:
            return self.plan.dst_base
        if self.dst_lib is not None:
            return inventory_base(self.dst_lib, self.opts.destination.folder)
        return None

    def migrating_files(self) -> list[Item]:
        """Files that will (or may) be copied: plan COPY/CONFLICT, else all non-excluded source files."""
        if self.plan is not None:
            return [e.item for e in self.plan.entries
                    if not e.item.is_folder and e.action in (Action.COPY, Action.CONFLICT)]
        return [i for i in self.src_inv.files() if is_excluded(i.rel_path, self.opts.excludes) is None]

    def targets(self) -> list[tuple[Item, str]]:
        """(item, target server-relative URL) for every non-excluded source item."""
        if self.plan is not None:
            return [(e.item, e.target_server_relative_url) for e in self.plan.entries
                    if e.action != Action.EXCLUDED]
        base = self.dst_base or self.src_inv.base_url
        return [(i, f"{base}/{i.rel_path}") for i in self.src_inv.items
                if is_excluded(i.rel_path, self.opts.excludes) is None]


Check = Callable[[CheckContext], Awaitable[list[Finding]]]


def _fmt_duration(seconds: float) -> str:
    seconds = int(seconds)
    h, rem = divmod(seconds, 3600)
    m = rem // 60
    return f"{h} h {m:02d} min" if h else f"{max(m, 1)} min"


def _gb(n: float) -> str:
    return f"{n / 1024 / 1024 / 1024:.2f} GB"


# ------------------------------------------------------------------------------------------------- rules


async def check_destination_library(ctx: CheckContext) -> list[Finding]:
    if ctx.dst_lib is not None:
        return _unavailable_routes(ctx)
    ep = ctx.opts.destination
    return [Finding(
        Severity.BLOCKING, "destination", ep.describe(),
        ctx.dst_error or "Destination library could not be resolved.",
        "Check the destination site URL and library name (leave the library empty to use the site's default "
        "document library) and make sure your account can open the site.",
    )]


def _unavailable_routes(ctx: CheckContext) -> list[Finding]:
    """Routes whose destination could not be resolved: their content is left out of the plan."""
    if ctx.plan is None:
        return []
    counts: dict[str, int] = {}
    for e in ctx.plan.entries:
        if e.action == Action.EXCLUDED and e.route is not None and e.reason.startswith(ROUTE_UNAVAILABLE):
            counts[e.route] = counts.get(e.route, 0) + 1
    by_prefix = {r.prefix.strip("/"): r for r in ctx.opts.routes}
    out: list[Finding] = []
    for prefix, n in counts.items():
        r = by_prefix.get(prefix)
        out.append(Finding(
            Severity.BLOCKING, "destination", prefix,
            f"Route destination {r.destination.describe() if r else '?'} could not be resolved; "
            f"{n} item(s) under '{prefix}' would not be migrated.",
            "Check the route's site URL and library name, or remove the route.",
        ))
    return out


async def check_write_access(ctx: CheckContext) -> list[Finding]:
    if ctx.dst_lib is None or ctx.dst_gw is None:
        return []
    folder = ctx.dst_base or ctx.dst_lib.root_server_relative_url
    try:
        try:
            await ctx.dst_gw.probe_write_access(folder)
        except NotFoundError:
            folder = ctx.dst_lib.root_server_relative_url
            await ctx.dst_gw.probe_write_access(folder)
    except AccessDeniedError as exc:
        return [Finding(
            Severity.BLOCKING, "write_access", folder, f"Cannot create items at the destination: {exc}",
            "Ask a site owner to grant your account Edit (or Full Control) on the destination library, then "
            "sign in again.",
        )]
    return []


async def check_metadata(ctx: CheckContext) -> list[Finding]:
    """Set dates on a test file at the destination and read them back: catches a date format or time zone the
    site does not accept before hours of copying (every file would otherwise end up without its dates)."""
    if ctx.dst_lib is None or ctx.dst_gw is None or not ctx.opts.restore_metadata:
        return []
    folder = ctx.dst_base or ctx.dst_lib.root_server_relative_url
    try:
        try:
            problem = await ctx.dst_gw.probe_metadata(ctx.dst_lib, folder)
        except NotFoundError:
            folder = ctx.dst_lib.root_server_relative_url
            problem = await ctx.dst_gw.probe_metadata(ctx.dst_lib, folder)
    except AccessDeniedError:
        return []  # reported by check_write_access
    except M365Error as exc:
        return [Finding(
            Severity.BLOCKING, "metadata", folder,
            f"Created/Modified could not be set on a test file at the destination: {exc}",
            "Copied files would keep the upload date and the migrating account. Report this, or turn off "
            "'restore metadata' in the job to copy without the original dates and authors.",
        )]
    if problem:
        return [Finding(
            Severity.WARN, "metadata", folder,
            f"Dates on a test file at the destination came back different: {problem}.",
            "Check the destination site's time zone (Site settings > Regional settings).",
        )]
    return []


async def check_dates(ctx: CheckContext) -> list[Finding]:
    out: list[Finding] = []
    for i in ctx.migrating_files():
        bad = [f"{n} {d:%Y-%m-%d}" for n, d in (("Created", i.created), ("Modified", i.modified))
               if d is not None and not sp_date_ok(d)]
        if bad:
            out.append(Finding(
                Severity.WARN, "dates", i.rel_path,
                f"{', '.join(bad)} is outside what SharePoint can store (1900-8900).",
                "No action needed: the other date (or the upload time) is used instead.",
            ))
    return out


async def check_quota(ctx: CheckContext) -> list[Finding]:
    if ctx.dst_lib is None or ctx.dst_gw is None:
        return []
    total = ctx.src_inv.total_bytes()
    try:
        used, quota = await ctx.dst_gw.get_site_usage()
    except Exception as exc:  # noqa: BLE001 - usage is optional information
        return [Finding(
            Severity.INFO, "quota", ctx.dst_gw.site_url,
            f"Storage usage of the destination site is not available ({exc}); quota was not checked.",
            "Check free storage in the SharePoint admin center (Active sites > Storage) before migrating "
            f"{_gb(total)}.",
        )]
    if quota <= 0:
        return [Finding(
            Severity.INFO, "quota", ctx.dst_gw.site_url,
            "The destination site reports no storage quota; free space was not checked.",
            "Nothing to do unless the tenant itself is low on storage.",
        )]
    free = quota - used
    need = total * QUOTA_HEADROOM
    if free >= need:
        return []
    return [Finding(
        Severity.BLOCKING, "quota", ctx.dst_gw.site_url,
        f"Destination has {_gb(free)} free; source needs {_gb(total)} (+10% headroom = {_gb(need)}).",
        "Raise the destination site's storage limit in the SharePoint admin center, or exclude folders.",
    )]


def _name_problem(name: str) -> str | None:
    bad = sorted(set(ILLEGAL_CHARS.findall(name)))
    if bad:
        return "contains illegal character(s) " + " ".join(bad)
    if name != name.strip(" "):
        return "starts or ends with a space"
    if name.endswith("."):
        return "ends with a dot"
    if name.startswith("~$"):
        return "is an Office lock/temp file (~$)"
    if name.lower() in RESERVED_NAMES:
        return "is a reserved name"
    if "_vti_" in name.lower():
        return "contains '_vti_'"
    return None


async def check_paths(ctx: CheckContext) -> list[Finding]:
    out: list[Finding] = []
    for item, target in ctx.targets():
        n = len(target)
        if n > MAX_PATH_LEN:
            out.append(Finding(
                Severity.BLOCKING, "path_length", item.rel_path,
                f"Target path is {n} characters (SharePoint limit {MAX_PATH_LEN}).",
                "Shorten folder or file names, or migrate into a shallower destination folder.",
            ))
        elif n > PATH_WARN_LEN:
            out.append(Finding(
                Severity.INFO, "path_length", item.rel_path,
                f"Target path is {n} characters (> {PATH_WARN_LEN}); desktop Office and OneDrive sync may "
                "struggle with it.",
                "Optional: shorten the folder structure for users who sync this library.",
            ))
        problem = _name_problem(item.name)
        if problem:
            out.append(Finding(
                Severity.BLOCKING, "illegal_name", item.rel_path, f"Name '{item.name}' {problem}.",
                "Rename the item in the source library, or add it to the excludes "
                "(e.g. '**/~$*' for Office lock files).",
            ))
    return out


async def check_temp_files(ctx: CheckContext) -> list[Finding]:
    """draw.io keeps a ``.$<name>.drawio.bkp`` backup next to every diagram. Valid in SharePoint, but clutter."""
    files = [i for i in ctx.migrating_files() if glob_match(DRAWIO_BACKUP, i.rel_path)]
    if not files:
        return []
    examples = ", ".join(i.rel_path for i in files[:3])
    return [Finding(
        Severity.INFO, "temp_files", ctx.src_inv.base_url,
        f"{len(files)} draw.io backup file(s) (.$*.bkp) will be copied (e.g. {examples}).",
        f"Optional: add the exclude '{DRAWIO_BACKUP}' to leave them behind; the real .drawio files are unaffected.",
    )]


async def check_unique_permissions(ctx: CheckContext) -> list[Finding]:
    items = [i for i in ctx.src_inv.items if i.unique_permissions]
    if not items:
        return []
    examples = ", ".join(i.rel_path for i in items[:5])
    return [Finding(
        Severity.WARN, "unique_permissions", ctx.src_inv.base_url,
        f"{len(items)} item(s) have unique permissions; these are not migrated (e.g. {examples}).",
        "Re-apply the needed item-level sharing at the destination after the migration, or accept that "
        "items inherit the destination library's permissions.",
    )]


async def check_checked_out(ctx: CheckContext) -> list[Finding]:
    return [
        Finding(
            Severity.WARN, "checked_out", i.rel_path, f"Checked out to {i.checked_out_to}.",
            "Ask the user to check the file in (or discard the check-out) first; otherwise only the last "
            "checked-in version is copied.",
        )
        for i in ctx.migrating_files() if i.checked_out_to
    ]


def _major(version: str | None) -> int:
    try:
        return int(float(version or "1"))
    except ValueError:
        return 1


async def check_versions(ctx: CheckContext) -> list[Finding]:
    files = [i for i in ctx.migrating_files() if _major(i.version) > 1]
    if not files:
        return []
    return [Finding(
        Severity.INFO, "versions", ctx.src_inv.base_url,
        f"{len(files)} file(s) have version history (up to v{max(_major(i.version) for i in files)}); "
        "only the latest version is copied.",
        "Keep the source library read-only for a while if older versions may still be needed.",
    )]


async def check_columns(ctx: CheckContext) -> list[Finding]:
    if ctx.dst_lib is None or ctx.dst_gw is None:
        return []
    src_fields = await ctx.src_gw.get_fields(ctx.src_lib)
    dst_fields = await ctx.dst_gw.get_fields(ctx.dst_lib)
    dst_names = {f.internal_name.lower() for f in dst_fields} | {f.title.lower() for f in dst_fields if f.title}
    out: list[Finding] = []
    for f in src_fields:
        if f.internal_name.lower() in _SYSTEM_LOWER:
            continue
        if f.internal_name.lower() in dst_names or (f.title and f.title.lower() in dst_names):
            continue
        label = f.title or f.internal_name
        out.append(Finding(
            Severity.WARN, "columns", label,
            f"Column '{label}' ({f.internal_name}{', ' + f.type_name if f.type_name else ''}) exists in the "
            "source library but not at the destination; its values will not be migrated.",
            "Create the column at the destination (same name and type) before migrating, or accept losing it.",
        ))
    return out


async def check_users(ctx: CheckContext) -> list[Finding]:
    if ctx.dst_lib is None or ctx.dst_gw is None or not ctx.opts.restore_metadata:
        return []
    out: list[Finding] = []
    files = ctx.migrating_files()
    inv = Inventory(library=ctx.src_inv.library, base_url=ctx.src_inv.base_url, items=files)
    for email, (name, count) in sorted(collect_users(inv).items()):
        if ctx.usermap is not None and ctx.usermap.resolve(email):
            continue
        if ctx.resolved_users is not None and email in ctx.resolved_users:
            if ctx.resolved_users[email] is not None:
                continue
        elif await ctx.dst_gw.find_site_user(email) is not None:
            continue
        who = f"{name} <{email}>" if name else email
        out.append(Finding(
            Severity.WARN, "users", email,
            f"{who} (author/editor of {count} file(s)) is not known at the destination; those files keep "
            "their dates but get the migrating account as Created By/Modified By.",
            "Map the user to an existing account in the user map CSV (source_email;target_email), or accept "
            "the fallback (typical for people who left).",
        ))
    return out


async def check_conflicts(ctx: CheckContext) -> list[Finding]:
    out: list[Finding] = []
    if ctx.plan is not None:
        for e in ctx.plan.entries:
            if e.action == Action.CONFLICT:
                out.append(Finding(
                    Severity.WARN, "conflict", e.item.rel_path,
                    f"Exists at the destination and differs: {e.reason}. It will be overwritten.",
                    "Check that nobody edited the destination copy; otherwise set overwrite to 'never' or "
                    "exclude the file.",
                ))
        return out
    if ctx.dst_inv is None:
        return out
    idx = ctx.dst_inv.index()
    for i in ctx.src_inv.files():
        d = idx.get(i.rel_path.lower())
        if d is None or d.is_folder:
            continue
        size_differs = d.size != i.size and not i.is_office
        if size_differs or (d.modified and i.modified and d.modified != i.modified):
            out.append(Finding(
                Severity.WARN, "conflict", i.rel_path, "Exists at the destination with a different size/date.",
                "Check which copy is current before migrating.",
            ))
    return out


async def check_onenote(ctx: CheckContext) -> list[Finding]:
    return [
        Finding(
            Severity.WARN, "onenote", i.rel_path,
            "OneNote notebook: copied as a folder of .one files, which may lose sync state and history.",
            "Prefer moving the notebook with OneNote itself, or exclude it and recreate it at the destination.",
        )
        for i in ctx.src_inv.items
        if (i.prog_id or "").lower() == "onenote.notebook" and is_excluded(i.rel_path, ctx.opts.excludes) is None
    ]


async def check_large_files(ctx: CheckContext) -> list[Finding]:
    return [
        Finding(
            Severity.INFO, "large_file", i.rel_path, f"Large file ({i.size / MB:.0f} MB).",
            "No action needed; large files are uploaded in chunks and may take a while.",
        )
        for i in ctx.migrating_files() if i.size >= LARGE_FILE_BYTES
    ]


_SETTINGS = (
    ("enable_versioning", "Versioning"),
    ("major_version_limit", "Major version limit"),
    ("enable_minor_versions", "Minor versions"),
    ("force_checkout", "Require check-out"),
    ("enable_moderation", "Content approval"),
)


async def check_library_settings(ctx: CheckContext) -> list[Finding]:
    if ctx.dst_lib is None:
        return []
    out: list[Finding] = []
    for attr, label in _SETTINGS:
        a, b = getattr(ctx.src_lib, attr), getattr(ctx.dst_lib, attr)
        if a != b:
            out.append(Finding(
                Severity.INFO, "library_settings", label,
                f"{label}: source = {a}, destination = {b}.",
                "Adjust the destination library settings (Library settings > Versioning settings) if the "
                "source behaviour should be kept.",
            ))
    return out


async def check_duration(ctx: CheckContext) -> list[Finding]:
    files = ctx.migrating_files()
    total = sum(i.size for i in files)
    secs = total / (ESTIMATE_MBPS * MB)
    return [Finding(
        Severity.INFO, "duration", ctx.src_inv.base_url,
        f"{len(files)} file(s), {total / MB:.1f} MB: estimated {_fmt_duration(secs)} at {ESTIMATE_MBPS} MB/s.",
        "Plan the migration outside office hours if it takes long; it can be stopped and resumed.",
    )]


CHECKS: list[Check] = [
    check_destination_library,
    check_write_access,
    check_metadata,
    check_quota,
    check_paths,
    check_temp_files,
    check_unique_permissions,
    check_checked_out,
    check_versions,
    check_columns,
    check_users,
    check_conflicts,
    check_onenote,
    check_large_files,
    check_dates,
    check_library_settings,
    check_duration,
]

_ORDER = {Severity.BLOCKING: 0, Severity.WARN: 1, Severity.INFO: 2}


async def run_checks(ctx: CheckContext, sink: EventSink = null_sink) -> list[Finding]:
    """Run every rule; a rule that crashes becomes a WARN finding instead of aborting the audit."""
    findings: list[Finding] = []
    n = len(CHECKS)
    sink(PhaseChanged("checks", f"{n} checks"))
    for i, check in enumerate(CHECKS):
        name = check.__name__.removeprefix("check_")
        sink(PhaseChanged("checks", name.replace("_", " ")))
        sink(Stats(done=i, total=n, bytes_done=0, bytes_total=0, mbps=0.0, eta_s=None, failed=0, warnings=0))
        try:
            findings.extend(await check(ctx))
        except Exception as exc:  # noqa: BLE001
            sink(LogLine("warn", f"check '{name}' failed: {exc}"))
            findings.append(Finding(
                Severity.WARN, "check_failed", name, f"The '{name}' check could not run: {exc}",
                "Re-run the audit; if it keeps failing, check this point manually.",
            ))
    sink(Stats(done=n, total=n, bytes_done=0, bytes_total=0, mbps=0.0, eta_s=None, failed=0, warnings=0))
    findings.sort(key=lambda f: _ORDER[f.severity])
    counts = {s: sum(1 for f in findings if f.severity == s) for s in Severity}
    sink(LogLine("warn" if counts[Severity.BLOCKING] else "info",
                 f"checks: {counts[Severity.BLOCKING]} blocking, {counts[Severity.WARN]} warnings, "
                 f"{counts[Severity.INFO]} info"))
    return findings
