from __future__ import annotations

import json
from pathlib import Path

import pytest

from m365tools.core import journal as journal_mod
from m365tools.core.journal import Journal, StatusWriter


def test_journal_append_load_done(tmp_path: Path) -> None:
    j = Journal(tmp_path / "run" / "journal.jsonl")
    assert j.load() == [] and j.done_paths() == set()
    j.append({"rel_path": "a/b.docx", "ok": True, "size": 10})
    j.append({"rel_path": "a/c.pdf", "ok": False, "error": "boom"})
    j.append({"path": "d.txt", "ok": True, "at": tmp_path})
    with j.path.open("a", encoding="utf-8") as fh:
        fh.write('{"rel_path": "trunc')  # crash mid-write
    entries = j.load()
    assert len(entries) == 3
    assert entries[1]["error"] == "boom"
    assert j.done_paths() == {"a/b.docx", "d.txt"}


class FakeClock:
    def __init__(self) -> None:
        self.t = 1000.0

    def __call__(self) -> float:
        return self.t


def test_status_writer_throttle_and_force(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    clock = FakeClock()
    monkeypatch.setattr(journal_mod.time, "monotonic", clock)
    path = tmp_path / "status.json"
    sw = StatusWriter(path, interval_s=3.0)

    assert sw.update(done=1, total=10) is True  # first update always writes
    assert json.loads(path.read_text())["done"] == 1

    clock.t += 1.0
    assert sw.update(done=2) is False  # throttled
    assert json.loads(path.read_text())["done"] == 1

    assert sw.update(force=True, current="x.docx") is True
    data = json.loads(path.read_text())
    assert data["done"] == 2 and data["current"] == "x.docx" and data["total"] == 10

    clock.t += 1.0
    assert sw.update(done=3) is False
    clock.t += 3.0
    assert sw.update(done=4) is True
    assert json.loads(path.read_text())["done"] == 4

    clock.t += 0.1
    sw.update(done=5)
    sw.close()
    assert json.loads(path.read_text())["done"] == 5
    assert not list(tmp_path.glob("*.tmp"))


def test_journal_failed_paths(tmp_path: Path) -> None:
    j = Journal(tmp_path / "journal.jsonl")
    j.append({"rel_path": "a.pdf", "ok": False})
    j.append({"rel_path": "b.pdf", "ok": False})
    j.append({"rel_path": "b.pdf", "ok": True})
    assert j.failed_paths() == {"a.pdf"}
