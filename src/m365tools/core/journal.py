"""Resumability primitives: an append-only JSONL journal and a throttled atomic ``status.json`` writer."""

from __future__ import annotations

import json
import os
import tempfile
import time
from pathlib import Path
from typing import Any


def _default(obj: Any) -> Any:
    if isinstance(obj, Path):
        return str(obj)
    if hasattr(obj, "isoformat"):
        return obj.isoformat()
    if isinstance(obj, set | frozenset):
        return sorted(obj)
    return str(obj)


def atomic_write_text(path: Path, text: str) -> None:
    """Write ``text`` to ``path`` via a temp file in the same directory and ``os.replace``."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=f".{path.name}-", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(text)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, path)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise


class Journal:
    """Append-only ``journal.jsonl``. Each entry is flushed and fsync'ed so a crash loses nothing written."""

    def __init__(self, path: Path) -> None:
        self.path = path

    def append(self, entry: dict[str, Any]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        line = json.dumps(entry, default=_default, ensure_ascii=False)
        with self.path.open("a", encoding="utf-8") as fh:
            fh.write(line + "\n")
            fh.flush()
            os.fsync(fh.fileno())

    def load(self) -> list[dict[str, Any]]:
        """All entries; a truncated/corrupt line (e.g. from a crash mid-write) is skipped."""
        if not self.path.exists():
            return []
        out: list[dict[str, Any]] = []
        with self.path.open(encoding="utf-8") as fh:
            for raw in fh:
                raw = raw.strip()
                if not raw:
                    continue
                try:
                    obj = json.loads(raw)
                except json.JSONDecodeError:
                    continue
                if isinstance(obj, dict):
                    out.append(obj)
        return out

    def done_paths(self) -> set[str]:
        """``rel_path`` (or ``path``) values of entries whose ``ok`` is true."""
        done: set[str] = set()
        for e in self.load():
            if e.get("ok") is True:
                p = e.get("rel_path", e.get("path"))
                if isinstance(p, str):
                    done.add(p)
        return done

    def failed_paths(self) -> set[str]:
        """``rel_path`` values that have a failed entry (``ok`` false) and no successful one."""
        failed: set[str] = set()
        ok: set[str] = set()
        for e in self.load():
            p = e.get("rel_path", e.get("path"))
            if not isinstance(p, str):
                continue
            (ok if e.get("ok") is True else failed).add(p)
        return failed - ok

    def metadata_pending(self) -> set[str]:
        """``rel_path`` values whose latest entry is a copy whose Created/Modified/Author/Editor could not be set
        (``metadata_failed`` true): the file is at the destination but its metadata should be tried again."""
        latest: dict[str, dict[str, Any]] = {}
        for e in self.load():
            p = e.get("rel_path", e.get("path"))
            if isinstance(p, str):
                latest[p] = e
        return {p for p, e in latest.items() if e.get("ok") is True and e.get("metadata_failed") is True}


class StatusWriter:
    """Merges status fields and writes ``status.json`` atomically, at most once per ``interval_s``."""

    def __init__(self, path: Path, interval_s: float = 3.0) -> None:
        self.path = path
        self.interval_s = interval_s
        self.fields: dict[str, Any] = {}
        self._last_write: float | None = None
        self._dirty = False

    def update(self, force: bool = False, **fields: Any) -> bool:
        """Merge ``fields``; write if forced or the interval elapsed. Returns True when the file was written."""
        self.fields.update(fields)
        self._dirty = True
        now = time.monotonic()
        if force or self._last_write is None or now - self._last_write >= self.interval_s:
            self._write(now)
            return True
        return False

    def _write(self, now: float) -> None:
        payload = dict(self.fields)
        payload["updated_at"] = time.strftime("%Y-%m-%dT%H:%M:%S%z")
        atomic_write_text(self.path, json.dumps(payload, default=_default, ensure_ascii=False, indent=2))
        self._last_write = now
        self._dirty = False

    def close(self) -> None:
        """Force a final write."""
        self._write(time.monotonic())
