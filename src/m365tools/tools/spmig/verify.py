"""Compare a source and a destination inventory after migration."""

from __future__ import annotations

from datetime import UTC, datetime, timedelta

from .models import CompareResult, CompareRow, Inventory

MTIME_TOLERANCE = timedelta(seconds=2)


def _aware(dt: datetime) -> datetime:
    """Naive datetimes are UTC (as SharePoint REST returns them); keeps naive/aware comparisons from raising."""
    return dt.replace(tzinfo=UTC) if dt.tzinfo is None else dt


def compare(src: Inventory, dst: Inventory) -> CompareResult:
    """Per source item: ``missing`` / ``size_differs`` (never for Office files: SharePoint rewrites them on
    upload) / ``source_newer`` (beyond 2 s) / ``ok``; plus ``only_in_destination`` rows."""
    dst_idx = dst.index()
    src_keys: set[str] = set()
    rows: list[CompareRow] = []
    for s in src.items:
        key = s.rel_path.lower()
        src_keys.add(key)
        d = dst_idx.get(key)
        if d is None:
            rows.append(CompareRow("missing", s.is_folder, s.rel_path, None if s.is_folder else s.size, None,
                                   s.modified, None, "not found at destination"))
            continue
        row = CompareRow("ok", s.is_folder, s.rel_path, None if s.is_folder else s.size,
                         None if d.is_folder else d.size, s.modified, d.modified)
        if s.is_folder != d.is_folder:
            row.status, row.detail = "missing", "exists at destination as a " + ("folder" if d.is_folder else "file")
        elif not s.is_folder:
            if s.size != d.size and not s.is_office:
                row.status, row.detail = "size_differs", f"{s.size} vs {d.size} bytes"
            elif s.modified and d.modified and _aware(s.modified) > _aware(d.modified) + MTIME_TOLERANCE:
                row.status, row.detail = "source_newer", "source modified after the destination copy"
            elif s.size != d.size:
                row.detail = "size differs (Office file, expected)"
        rows.append(row)
    for d in dst.items:
        if d.rel_path.lower() not in src_keys:
            rows.append(CompareRow("only_in_destination", d.is_folder, d.rel_path, None,
                                   None if d.is_folder else d.size, None, d.modified))
    return CompareResult(rows=rows)
