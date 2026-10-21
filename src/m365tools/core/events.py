"""Progress events emitted by core/tools and rendered by TUI/CLI. Core never prints."""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Literal, Protocol

Step = Literal["folder", "download", "upload", "metadata", "verify", "permissions"]
Level = Literal["debug", "info", "ok", "warn", "error"]

# Stable phase ids of the audit / dry-run pre-phase (emitted in this order) and their human labels. "inventory"
# and "permissions" are prefixes: UIs that only care about "scanning vs. transferring" match on them.
PHASE_LABELS: dict[str, str] = {
    "inventory.source": "inventory source",
    "permissions.source": "permissions scan",
    "inventory.destination": "inventory destination",
    "users": "users",
    "plan": "plan",
    "checks": "checks",
    "report": "report",
}
AUDIT_PHASES: tuple[str, ...] = (
    "inventory.source", "permissions.source", "inventory.destination", "users", "checks", "report",
)
DRYRUN_PHASES: tuple[str, ...] = ("inventory.source", "permissions.source", "inventory.destination", "plan")
SCAN_PHASES: frozenset[str] = frozenset(PHASE_LABELS) | {"inventory", "permissions"}


def phase_label(phase: str) -> str:
    return PHASE_LABELS.get(phase, phase)


@dataclass(slots=True)
class PhaseChanged:
    phase: str
    detail: str = ""


@dataclass(slots=True)
class ItemStarted:
    rel_path: str
    size: int
    index: int
    total: int


@dataclass(slots=True)
class ItemStep:
    rel_path: str
    step: Step


@dataclass(slots=True)
class ItemDone:
    rel_path: str
    ok: bool
    seconds: float
    message: str = ""
    skipped: bool = False


@dataclass(slots=True)
class Throttled:
    seconds: float
    attempt: int
    what: str


@dataclass(slots=True)
class LogLine:
    level: Level
    text: str
    at: datetime = field(default_factory=datetime.now)
    phase: str = ""  # set on the "<phase> started (+12.4s)" lines from PhaseClock (renderers may skip them)


@dataclass(slots=True)
class Stats:
    done: int
    total: int
    bytes_done: int
    bytes_total: int
    mbps: float
    eta_s: float | None
    failed: int
    warnings: int
    copied: int = 0
    skipped: int = 0


@dataclass(slots=True)
class AuthPrompt:
    """Device-code or browser instructions that must be shown to the user."""

    tenant: str
    message: str
    url: str | None = None
    code: str | None = None


ProgressEvent = PhaseChanged | ItemStarted | ItemStep | ItemDone | Throttled | LogLine | Stats | AuthPrompt


class EventSink(Protocol):
    def __call__(self, ev: ProgressEvent) -> None: ...


def null_sink(ev: ProgressEvent) -> None:
    return None


class CollectingSink:
    """Test helper and simple buffer: keeps every event."""

    def __init__(self) -> None:
        self.events: list[ProgressEvent] = []

    def __call__(self, ev: ProgressEvent) -> None:
        self.events.append(ev)


class PhaseClock:
    """Sink wrapper: forwards every event and, on the first ``PhaseChanged`` of each new phase id, adds an info
    ``LogLine`` "<label>: <detail> (+12.4s)" with the seconds since :meth:`reset` (start of the operation)."""

    def __init__(self, inner: EventSink) -> None:
        self.inner = inner
        self.t0 = time.monotonic()
        self.last: str | None = None

    def reset(self) -> None:
        self.t0 = time.monotonic()
        self.last = None

    def elapsed(self) -> float:
        return time.monotonic() - self.t0

    def __call__(self, ev: ProgressEvent) -> None:
        self.inner(ev)
        if isinstance(ev, PhaseChanged) and ev.phase != self.last:
            self.last = ev.phase
            detail = f": {ev.detail}" if ev.detail else ""
            self.inner(LogLine("info", f"{phase_label(ev.phase)}{detail} (+{self.elapsed():.1f}s)",
                               phase=ev.phase))
