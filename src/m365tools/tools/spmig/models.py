"""Domain models for the SharePoint library migrator. Field names are the contract (see docs/ARCHITECTURE.md)."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import StrEnum
from pathlib import Path
from typing import Literal

OFFICE_EXTENSIONS: frozenset[str] = frozenset(
    {".docx", ".docm", ".dotx", ".xlsx", ".xlsm", ".xltx", ".pptx", ".pptm", ".potx",
     ".doc", ".xls", ".ppt", ".vsdx", ".one"}
)
MAX_PATH_LEN = 400
PATH_WARN_LEN = 260


class Severity(StrEnum):
    BLOCKING = "blocking"
    WARN = "warn"
    INFO = "info"


class Action(StrEnum):
    CREATE_FOLDER = "create_folder"
    COPY = "copy"
    SKIP = "skip"
    CONFLICT = "conflict"
    EXCLUDED = "excluded"


OverwritePolicy = Literal["skip_if_same_or_newer", "always", "never"]


@dataclass(slots=True)
class Item:
    id: int
    is_folder: bool
    rel_path: str
    name: str
    ext: str
    size: int
    created: datetime | None
    modified: datetime | None
    author_email: str | None
    author_name: str | None
    editor_email: str | None
    editor_name: str | None
    version: str | None
    checked_out_to: str | None
    prog_id: str | None
    server_relative_url: str
    unique_permissions: bool = False
    etag: str | None = None

    @property
    def is_office(self) -> bool:
        return self.ext in OFFICE_EXTENSIONS

    @property
    def top_folder(self) -> str:
        return self.rel_path.split("/", 1)[0] if "/" in self.rel_path else "(root)"


@dataclass(slots=True)
class LibraryRef:
    site_url: str
    web_server_relative_url: str
    list_id: str
    title: str
    root_server_relative_url: str
    item_count: int
    enable_versioning: bool = False
    major_version_limit: int = 0
    enable_minor_versions: bool = False
    force_checkout: bool = False
    enable_moderation: bool = False

    @property
    def root_name(self) -> str:
        return self.root_server_relative_url.rstrip("/").rsplit("/", 1)[-1]


@dataclass(slots=True)
class Inventory:
    library: LibraryRef
    base_url: str  # server-relative root of the inventory (library root or sub-folder), no trailing slash
    items: list[Item] = field(default_factory=list)

    def files(self) -> list[Item]:
        return [i for i in self.items if not i.is_folder]

    def folders(self) -> list[Item]:
        return [i for i in self.items if i.is_folder]

    def total_bytes(self) -> int:
        return sum(i.size for i in self.items if not i.is_folder)

    def index(self) -> dict[str, Item]:
        return {i.rel_path.lower(): i for i in self.items}


@dataclass(slots=True)
class Endpoint:
    tenant: str
    site_url: str
    library: str | None = None  # None => site's default document library
    folder: str = ""

    def describe(self) -> str:
        lib = self.library or "(default library)"
        return f"{self.site_url} / {lib}" + (f" / {self.folder}" if self.folder else "")


@dataclass(slots=True)
class Route:
    prefix: str  # rel_path prefix in the source, e.g. "0. Archief"
    destination: Endpoint


@dataclass(slots=True)
class JobOptions:
    source: Endpoint
    destination: Endpoint
    overwrite: OverwritePolicy = "skip_if_same_or_newer"
    restore_metadata: bool = True
    excludes: list[str] = field(default_factory=list)
    routes: list[Route] = field(default_factory=list)
    lock_source_after: bool = False
    workers: int = 2
    usermap_path: Path | None = None
    profile: str | None = None


@dataclass(slots=True)
class PlanEntry:
    action: Action
    item: Item
    target_server_relative_url: str
    reason: str = ""
    route: str | None = None


@dataclass(slots=True)
class Plan:
    entries: list[PlanEntry]
    src: LibraryRef
    dst: LibraryRef
    src_base: str
    dst_base: str

    def counts(self) -> dict[Action, int]:
        out: dict[Action, int] = {a: 0 for a in Action}
        for e in self.entries:
            out[e.action] += 1
        return out

    def bytes_to_copy(self) -> int:
        return sum(e.item.size for e in self.entries if e.action == Action.COPY)

    def to_copy(self) -> list[PlanEntry]:
        return [e for e in self.entries if e.action == Action.COPY]

    def folders_to_create(self) -> list[PlanEntry]:
        return [e for e in self.entries if e.action == Action.CREATE_FOLDER]


@dataclass(slots=True)
class Finding:
    severity: Severity
    category: str
    path: str
    detail: str
    fix: str = ""


@dataclass(slots=True)
class UserMapEntry:
    source_email: str
    target_email: str | None
    display_name: str
    status: Literal["mapped", "unmapped", "auto"] = "unmapped"


@dataclass(slots=True)
class RunSummary:
    mode: Literal["audit", "dryrun", "migrate", "verify", "lock"]
    started: datetime
    finished: datetime
    run_dir: Path
    copied: int = 0
    skipped: int = 0
    failed: int = 0
    metadata_warnings: int = 0
    bytes_copied: int = 0
    stopped_early: bool = False
    findings: list[Finding] = field(default_factory=list)

    @property
    def duration_s(self) -> float:
        return (self.finished - self.started).total_seconds()

    @property
    def blocking(self) -> list[Finding]:
        return [f for f in self.findings if f.severity == Severity.BLOCKING]


@dataclass(slots=True)
class CompareRow:
    status: Literal["ok", "missing", "size_differs", "source_newer", "only_in_destination"]
    is_folder: bool
    rel_path: str
    src_size: int | None
    dst_size: int | None
    src_modified: datetime | None
    dst_modified: datetime | None
    detail: str = ""


@dataclass(slots=True)
class CompareResult:
    rows: list[CompareRow]

    def counts(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for r in self.rows:
            out[r.status] = out.get(r.status, 0) + 1
        return out
