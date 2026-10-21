"""JSON writer with a tolerant default handler."""

from __future__ import annotations

import dataclasses
import json
from datetime import date, datetime
from enum import Enum
from pathlib import Path
from typing import Any


def _default(o: Any) -> Any:
    if isinstance(o, datetime | date):
        return o.isoformat()
    if isinstance(o, Path):
        return str(o)
    if isinstance(o, Enum):
        return o.value
    if dataclasses.is_dataclass(o) and not isinstance(o, type):
        return dataclasses.asdict(o)
    if isinstance(o, set | frozenset):
        return sorted(o, key=str)
    raise TypeError(f"Cannot serialise {type(o).__name__}")


def write_json(path: Path, obj: Any) -> None:
    Path(path).write_text(
        json.dumps(obj, default=_default, indent=2, ensure_ascii=False), encoding="utf-8"
    )
