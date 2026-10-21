"""Report data model."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Literal


@dataclass
class Section:
    title: str
    kind: Literal["table", "kv", "text"]
    columns: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    text: str = ""
    note: str = ""
    severity_col: int | None = None
    collapsed: bool = False


@dataclass
class Report:
    title: str
    mode: str
    source: str
    destination: str
    started: datetime
    finished: datetime
    summary: list[tuple[str, str]]
    sections: list[Section]
    not_migrated: list[str] = field(default_factory=list)
    footer: str = ""
