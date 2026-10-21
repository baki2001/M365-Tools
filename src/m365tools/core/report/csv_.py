"""CSV writers (Excel NL friendly: ';' delimiter, UTF-8 with BOM)."""

from __future__ import annotations

import csv
from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import Any


def write_csv(path: Path, columns: Sequence[str], rows: Iterable[Sequence[Any]]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";", quoting=csv.QUOTE_MINIMAL)
        w.writerow(columns)
        for r in rows:
            w.writerow(["" if c is None else c for c in r])


def write_findings_csv(path: Path, findings: Iterable[Any]) -> None:
    write_csv(
        path,
        ["Severity", "Category", "Path", "Detail", "Fix"],
        [[str(f.severity.value if hasattr(f.severity, "value") else f.severity),
          f.category, f.path, f.detail, f.fix] for f in findings],
    )
