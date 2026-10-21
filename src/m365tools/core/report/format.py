"""Formatting helpers."""

from __future__ import annotations


def fmt_size(n: float) -> str:
    """Bytes to '2.40 GB' (1024 based); bytes below 1 KB are shown as '512 B'."""
    if n < 1024:
        return f"{int(n)} B"
    v = float(n)
    for unit in ("KB", "MB", "GB", "TB"):
        v /= 1024
        if v < 1024 or unit == "TB":
            return f"{v:.2f} {unit}"
    return f"{v:.2f} TB"  # pragma: no cover


def fmt_duration(seconds: float) -> str:
    s = round(seconds)
    d, s = divmod(s, 86400)
    h, s = divmod(s, 3600)
    m, s = divmod(s, 60)
    base = f"{h:02d}:{m:02d}:{s:02d}"
    return f"{d}d {base}" if d else base
