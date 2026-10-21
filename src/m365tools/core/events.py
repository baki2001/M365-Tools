"""Progress events emitted by core/tools and rendered by TUI/CLI. Core never prints."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Literal, Protocol

Step = Literal["folder", "download", "upload", "metadata", "verify", "permissions"]
Level = Literal["debug", "info", "ok", "warn", "error"]


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
