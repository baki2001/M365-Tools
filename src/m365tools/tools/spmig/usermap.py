"""Source -> destination user mapping for Author/Editor metadata.

CSV format (``;`` delimited, UTF-8 with BOM so Excel opens it cleanly):
``source_email;target_email;display_name;status`` where status is ``auto`` (found automatically),
``mapped`` (set by a human) or ``unmapped``.
"""

from __future__ import annotations

import csv
from pathlib import Path

from m365tools.core.events import EventSink, PhaseChanged, Stats, null_sink

from .gateway import SiteUser, SpGateway
from .models import Inventory, UserMapEntry

ResolvedUsers = dict[str, SiteUser | None]  # lower-cased source email -> destination site user (None = unknown)

CSV_COLUMNS = ["source_email", "target_email", "display_name", "status"]
_VALID_STATUS = {"auto", "mapped", "unmapped"}


def collect_users(inv: Inventory) -> dict[str, tuple[str, int]]:
    """Unique author/editor emails (lower-cased) of the inventory's files -> (display name, file count)."""
    out: dict[str, tuple[str, int]] = {}
    for item in inv.files():
        seen: set[str] = set()
        for email, name in ((item.author_email, item.author_name), (item.editor_email, item.editor_name)):
            if not email:
                continue
            key = email.strip().lower()
            if not key or key in seen:
                continue
            seen.add(key)
            prev_name, count = out.get(key, ("", 0))
            out[key] = (prev_name or (name or ""), count + 1)
    return out


async def resolve_users(
    emails: list[str], dst_gw: SpGateway, sink: EventSink = null_sink, *, phase: str = "users"
) -> ResolvedUsers:
    """Look every (distinct) email up at the destination once, emitting ``PhaseChanged(phase)`` and
    ``Stats(resolved, distinct)`` progress. The result feeds :func:`build_usermap` and the users check."""
    distinct = sorted({e.strip().lower() for e in emails if e and e.strip()})
    n = len(distinct)
    sink(PhaseChanged(phase, f"{n:,} distinct author(s)/editor(s)"))
    out: ResolvedUsers = {}
    step = max(1, n // 100)
    for i, email in enumerate(distinct, 1):
        out[email] = await dst_gw.find_site_user(email)
        if i % step == 0 or i == n:
            sink(Stats(done=i, total=n, bytes_done=0, bytes_total=0, mbps=0.0, eta_s=None, failed=0, warnings=0))
    if n == 0:
        sink(Stats(done=0, total=0, bytes_done=0, bytes_total=0, mbps=0.0, eta_s=None, failed=0, warnings=0))
    return out


async def build_usermap(
    src_inv: Inventory, dst_gw: SpGateway, resolved: ResolvedUsers | None = None
) -> list[UserMapEntry]:
    """Propose a mapping: same email known at destination => ``auto``; else a unique match on the local part
    (if the gateway offers ``find_users_by_local_part``) => ``auto``; else ``unmapped``. ``resolved`` (from
    :func:`resolve_users`) avoids looking the same email up again."""
    by_local = getattr(dst_gw, "find_users_by_local_part", None)
    entries: list[UserMapEntry] = []
    for email, (name, _count) in sorted(collect_users(src_inv).items()):
        hit = resolved[email] if resolved is not None and email in resolved else await dst_gw.find_site_user(email)
        if hit is not None:
            entries.append(UserMapEntry(email, hit.email or email, name or hit.title, "auto"))
            continue
        target: str | None = None
        if by_local is not None:
            candidates = await by_local(email.split("@", 1)[0])
            if len(candidates) == 1:
                target = candidates[0].email
        if target:
            entries.append(UserMapEntry(email, target, name, "auto"))
        else:
            entries.append(UserMapEntry(email, None, name, "unmapped"))
    return entries


def save_usermap(path: Path, entries: list[UserMapEntry]) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(CSV_COLUMNS)
        for e in entries:
            w.writerow([e.source_email, e.target_email or "", e.display_name, e.status])


def load_usermap(path: Path) -> list[UserMapEntry]:
    out: list[UserMapEntry] = []
    with Path(path).open("r", encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh, delimiter=";"):
            src = (row.get("source_email") or "").strip()
            if not src:
                continue
            target = (row.get("target_email") or "").strip() or None
            status = (row.get("status") or "").strip().lower()
            if status not in _VALID_STATUS:
                status = "mapped" if target else "unmapped"
            if target is None:
                status = "unmapped"
            out.append(UserMapEntry(src, target, (row.get("display_name") or "").strip(), status))  # type: ignore[arg-type]
    return out


class UserMap:
    """Lookup helper: ``resolve(email)`` -> target email for ``auto``/``mapped`` entries, else None."""

    def __init__(self, entries: list[UserMapEntry] | None = None) -> None:
        self.entries = list(entries or [])
        self._map: dict[str, UserMapEntry] = {e.source_email.strip().lower(): e for e in self.entries}

    @classmethod
    def load(cls, path: Path) -> UserMap:
        return cls(load_usermap(path))

    def save(self, path: Path) -> None:
        save_usermap(path, self.entries)

    def __contains__(self, email: object) -> bool:
        return isinstance(email, str) and email.strip().lower() in self._map

    def __len__(self) -> int:
        return len(self._map)

    def entry(self, email: str) -> UserMapEntry | None:
        return self._map.get(email.strip().lower())

    def resolve(self, email: str | None) -> str | None:
        if not email:
            return None
        e = self._map.get(email.strip().lower())
        if e is None or e.status == "unmapped" or not e.target_email:
            return None
        return e.target_email
