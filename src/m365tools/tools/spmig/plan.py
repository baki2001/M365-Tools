"""Turn a source and destination inventory into a migration :class:`Plan`.

Order of entries: folders (shallowest first), then files (by path). Decision per item:

1. route   - rel_path equals/starts with a route prefix -> target under that route's base, ``entry.route=prefix``
             (exclude patterns that only address the route prefix itself, e.g. ``"0. Archief/**"``, do not apply
             to routed items: the route is what handles them)
2. exclude - any glob in ``opts.excludes`` matches                    -> EXCLUDED
3. folder  - exists at destination -> SKIP, else CREATE_FOLDER
4. file    - missing at destination -> COPY; exists ->
             ``never`` -> SKIP; ``always`` -> COPY;
             ``skip_if_same_or_newer`` -> SKIP when dst.modified >= src.modified - 60 s, else CONFLICT
             (the engine copies CONFLICT entries; they are reported separately).
"""

from __future__ import annotations

from dataclasses import asdict
from datetime import UTC, datetime, timedelta
from fnmatch import fnmatchcase
from typing import Any

from .models import Action, Inventory, Item, JobOptions, LibraryRef, Plan, PlanEntry

MTIME_TOLERANCE = timedelta(seconds=60)  # dates may be stored to the minute
ROUTE_UNAVAILABLE = "route destination unavailable"


# --------------------------------------------------------------------------------------------------- globs


def glob_match(pattern: str, rel_path: str) -> bool:
    """Case-insensitive glob on a ``/``-separated rel_path.

    ``*`` and ``**`` both cross ``/`` (fnmatch semantics); a leading ``**/`` also matches at the root and a
    trailing ``/**`` also matches the folder itself (``"X/**"`` matches ``"X"``).
    """
    pat = pattern.strip().replace("\\", "/").strip("/").lower()
    path = rel_path.strip("/").lower()
    if not pat:
        return False
    if fnmatchcase(path, pat):
        return True
    if pat.endswith("/**") and fnmatchcase(path, pat[:-3]):
        return True
    return pat.startswith("**/") and glob_match(pat[3:], path)


def is_excluded(rel_path: str, patterns: list[str]) -> str | None:
    """Return the first matching pattern, or None."""
    for p in patterns:
        if glob_match(p, rel_path):
            return p
    return None


def _route_for(rel_path: str, prefixes: list[str]) -> str | None:
    low = rel_path.lower()
    best: str | None = None
    for p in prefixes:
        pl = p.strip("/").lower()
        if pl and (low == pl or low.startswith(pl + "/")) and (best is None or len(pl) > len(best)):
            best = p.strip("/")
    return best


def _pattern_targets_route(pattern: str, prefix: str) -> bool:
    """True for a pattern that addresses only the route prefix as a whole (``"0. Archief"``, ``"0. Archief/**"``,
    ``"0. Archief/*"``). Patterns for something *inside* the route (``"0. Archief/Concepten/**"``) are not."""
    pat = pattern.strip().replace("\\", "/").strip("/").lower()
    pre = prefix.strip("/").lower()
    for suffix in ("/**", "/*"):
        if pat.endswith(suffix):
            pat = pat[: -len(suffix)]
            break
    return pat == pre


def _aware(dt: datetime) -> datetime:
    """Naive datetimes are UTC (SharePoint REST returns UTC); makes naive/aware values comparable."""
    return dt.replace(tzinfo=UTC) if dt.tzinfo is None else dt


# --------------------------------------------------------------------------------------------------- plan


def _decide_file(src: Item, dst: Item | None, policy: str) -> tuple[Action, str]:
    if dst is None:
        return Action.COPY, "new"
    if dst.is_folder:
        return Action.CONFLICT, "a folder with this name exists at the destination"
    if policy == "never":
        return Action.SKIP, "exists at destination (overwrite=never)"
    if policy == "always":
        return Action.COPY, "exists at destination (overwrite=always)"
    if dst.size == 0 and src.size > 0:
        # Typically the empty placeholder of an interrupted chunked upload: its Modified is the upload time, so
        # the date rule below would wrongly keep it.
        return Action.CONFLICT, "destination is empty (incomplete earlier upload?)"
    if src.modified is None or dst.modified is None:
        if src.size == dst.size:
            return Action.SKIP, "exists, same size (no dates to compare)"
        return Action.CONFLICT, "exists with a different size (no dates to compare)"
    if _aware(dst.modified) >= _aware(src.modified) - MTIME_TOLERANCE:
        return Action.SKIP, "destination is same or newer"
    return Action.CONFLICT, "source is newer than destination"


def build_plan(
    src: Inventory,
    dst: Inventory,
    opts: JobOptions,
    dst_base: str,
    route_bases: dict[str, str] | None = None,
    route_invs: dict[str, Inventory] | None = None,
) -> Plan:
    """Build the plan. ``route_bases`` maps route prefix -> server-relative base URL of the route destination;
    ``route_invs`` optionally maps prefix -> inventory of that destination (for skip/conflict decisions)."""
    route_bases = {k.strip("/"): v.rstrip("/") for k, v in (route_bases or {}).items()}
    route_invs = {k.strip("/"): v for k, v in (route_invs or {}).items()}
    dst_base = dst_base.rstrip("/")
    dst_index = dst.index()
    route_index = {k: inv.index() for k, inv in route_invs.items()}
    # Every configured route prefix, also those whose destination could not be resolved: their items must not
    # silently fall through to the main destination.
    prefixes = list(dict.fromkeys([*route_bases, *(r.prefix.strip("/") for r in opts.routes)]))

    folders: list[PlanEntry] = []
    files: list[PlanEntry] = []
    for item in src.items:
        rel = item.rel_path.strip("/")
        route = _route_for(rel, prefixes)
        sub = ""
        if route is not None and route not in route_bases:
            entry = PlanEntry(Action.EXCLUDED, item, "", f"{ROUTE_UNAVAILABLE} ({route})", route)
            (folders if item.is_folder else files).append(entry)
            continue
        if route is not None:
            sub = rel[len(route):].lstrip("/")
            base = route_bases[route]
            target = f"{base}/{sub}" if sub else base
            index = route_index.get(route, {})
            existing = index.get(sub.lower()) if sub else None
            patterns = [p for p in opts.excludes if not _pattern_targets_route(p, route)]
        else:
            target = f"{dst_base}/{rel}"
            existing = dst_index.get(rel.lower())
            patterns = opts.excludes

        pat = is_excluded(rel, patterns)
        if pat is not None:
            entry = PlanEntry(Action.EXCLUDED, item, target, f"excluded by '{pat}'", route)
        elif item.is_folder:
            if route is not None and not sub:
                entry = PlanEntry(Action.SKIP, item, target, "route root (content goes to route destination)",
                                  route)
            elif existing is not None:
                entry = PlanEntry(Action.SKIP, item, target, "folder exists", route)
            else:
                entry = PlanEntry(Action.CREATE_FOLDER, item, target, "", route)
        else:
            action, reason = _decide_file(item, existing, opts.overwrite)
            entry = PlanEntry(action, item, target, reason, route)
        (folders if item.is_folder else files).append(entry)

    folders.sort(key=lambda e: (e.item.rel_path.count("/"), e.item.rel_path.lower()))
    files.sort(key=lambda e: e.item.rel_path.lower())
    return Plan(entries=folders + files, src=src.library, dst=dst.library, src_base=src.base_url,
                dst_base=dst_base)


# --------------------------------------------------------------------------------------------------- json


def _dt(v: datetime | None) -> str | None:
    return v.isoformat() if v else None


def _parse_dt(v: str | None) -> datetime | None:
    return datetime.fromisoformat(v) if v else None


def _item_to_json(item: Item) -> dict[str, Any]:
    d = asdict(item)
    d["created"] = _dt(item.created)
    d["modified"] = _dt(item.modified)
    return d


def _item_from_json(d: dict[str, Any]) -> Item:
    d = dict(d)
    d["created"] = _parse_dt(d.get("created"))
    d["modified"] = _parse_dt(d.get("modified"))
    return Item(**d)


def plan_to_json(plan: Plan) -> dict[str, Any]:
    counts = {str(k): v for k, v in plan.counts().items()}
    return {
        "version": 1,
        "src": asdict(plan.src),
        "dst": asdict(plan.dst),
        "src_base": plan.src_base,
        "dst_base": plan.dst_base,
        "counts": counts,
        "bytes_to_copy": plan.bytes_to_copy(),
        "entries": [
            {
                "action": str(e.action),
                "target": e.target_server_relative_url,
                "reason": e.reason,
                "route": e.route,
                "item": _item_to_json(e.item),
            }
            for e in plan.entries
        ],
    }


def plan_from_json(data: dict[str, Any]) -> Plan:
    return Plan(
        entries=[
            PlanEntry(
                action=Action(e["action"]),
                item=_item_from_json(e["item"]),
                target_server_relative_url=e["target"],
                reason=e.get("reason", ""),
                route=e.get("route"),
            )
            for e in data["entries"]
        ],
        src=LibraryRef(**data["src"]),
        dst=LibraryRef(**data["dst"]),
        src_base=data["src_base"],
        dst_base=data["dst_base"],
    )
