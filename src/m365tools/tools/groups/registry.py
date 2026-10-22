"""Registry TOML file: load, save (merge, never delete) and expectations.

The file is edited by humans and by the tool. Unknown keys (top level, per group, bundle and person) are kept in
``extra`` and written back. :func:`save_registry` merges with what is on disk at save time, so an entry that was
added by hand while the tool was running, or that the in-memory registry lacks, is never dropped. TOML comments are
not preserved (``tomli_w`` cannot write them).
"""

from __future__ import annotations

import os
import tomllib
from pathlib import Path
from typing import Any

import tomli_w

from m365tools.core.errors import ConfigError

from .models import DEFAULT_PREFIX, Bundle, GroupEntry, PersonEntry, Registry, kind_of, norm

_GROUP_KEYS = ("id", "kind", "purpose", "owner")
_BUNDLE_KEYS = ("purpose", "groups")
_PERSON_KEYS = ("bundles", "addons", "note", "offboarded")
_TOP_KEYS = ("prefix", "domains", "groups", "bundles", "people")


def _split(d: dict[str, Any], known: tuple[str, ...]) -> dict[str, Any]:
    return {k: v for k, v in d.items() if k not in known}


def _strs(v: Any) -> list[str]:
    return [str(x) for x in v] if isinstance(v, list) else []


def registry_from_dict(raw: dict[str, Any]) -> Registry:
    prefix = str(raw.get("prefix") or DEFAULT_PREFIX)
    reg = Registry(prefix=prefix, domains=[d.lower().lstrip("@") for d in _strs(raw.get("domains"))],
                   extra=_split(raw, _TOP_KEYS))
    for name, g in (raw.get("groups") or {}).items():
        g = g if isinstance(g, dict) else {}
        reg.groups[name] = GroupEntry(name=name, id=str(g.get("id", "")), kind=str(g.get("kind") or
                                      kind_of(name, prefix)), purpose=str(g.get("purpose", "")),
                                      owner=str(g.get("owner", "")), extra=_split(g, _GROUP_KEYS))
    for name, b in (raw.get("bundles") or {}).items():
        b = b if isinstance(b, dict) else {}
        reg.bundles[name] = Bundle(name=name, purpose=str(b.get("purpose", "")), groups=_strs(b.get("groups")),
                                   extra=_split(b, _BUNDLE_KEYS))
    for upn, p in (raw.get("people") or {}).items():
        p = p if isinstance(p, dict) else {}
        reg.people[norm(upn)] = PersonEntry(upn=norm(upn), bundles=_strs(p.get("bundles")),
                                            addons=_strs(p.get("addons")), note=str(p.get("note", "")),
                                            offboarded=str(p.get("offboarded", "")),
                                            extra=_split(p, _PERSON_KEYS))
    return reg


def _drop_empty(d: dict[str, Any], keep: tuple[str, ...] = ()) -> dict[str, Any]:
    return {k: v for k, v in d.items() if k in keep or v not in ("", [], None)}


def registry_to_dict(reg: Registry) -> dict[str, Any]:
    out: dict[str, Any] = {"prefix": reg.prefix, "domains": list(reg.domains), **reg.extra}
    out["groups"] = {g.name: {**_drop_empty({"id": g.id, "kind": g.kind, "owner": g.owner}),
                              "purpose": g.purpose, **g.extra} for g in reg.groups.values()}
    out["bundles"] = {b.name: {"purpose": b.purpose, "groups": list(b.groups), **b.extra}
                      for b in reg.bundles.values()}
    out["people"] = {p.upn: {**_drop_empty({"bundles": list(p.bundles), "addons": list(p.addons),
                                            "note": p.note, "offboarded": p.offboarded}, keep=("bundles",)),
                             **p.extra} for p in reg.people.values()}
    return out


def load_registry(path: Path) -> Registry:
    """Registry at ``path``; an empty one (default prefix) when the file does not exist yet."""
    path = Path(path)
    if not path.exists():
        return Registry()
    try:
        raw = tomllib.loads(path.read_text(encoding="utf-8"))
    except (tomllib.TOMLDecodeError, OSError) as exc:
        raise ConfigError(f"{path}: cannot read registry: {exc}") from exc
    return registry_from_dict(raw)


def _merge(disk: dict[str, Any], mine: dict[str, Any]) -> dict[str, Any]:
    """``mine`` wins per key; tables on disk that ``mine`` lacks are kept (recursively for the entry tables)."""
    out = dict(disk)
    for k, v in mine.items():
        if k in ("groups", "bundles", "people") and isinstance(disk.get(k), dict):
            table = dict(disk[k])
            lower = {norm(n): n for n in table} if k == "people" else {}
            for name, entry in v.items():
                if k == "people":
                    table.pop(lower.get(norm(name), name), None)
                table[name] = entry
            out[k] = table
        else:
            out[k] = v
    return out


def save_registry(reg: Registry, path: Path) -> Path:
    """Write ``reg`` to ``path`` atomically, merged with the file on disk (entries are never deleted)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = registry_to_dict(reg)
    if path.exists():
        try:
            disk = tomllib.loads(path.read_text(encoding="utf-8"))
        except (tomllib.TOMLDecodeError, OSError) as exc:
            raise ConfigError(f"{path}: refusing to overwrite an unreadable registry: {exc}") from exc
        data = _merge(disk, data)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(tomli_w.dumps(data), encoding="utf-8")
    os.replace(tmp, path)
    return path


def expectations(reg: Registry, upn: str) -> set[str]:
    """Expected groups of ``upn``: union of its bundles' groups plus its add-ons (empty when unregistered or
    offboarded). Names are returned as written in the registry."""
    p = reg.person(upn)
    if p is None or p.offboarded:
        return set()
    out: set[str] = set(p.addons)
    for b in p.bundles:
        bundle = reg.bundle(b)
        if bundle is not None:
            out.update(bundle.groups)
    return out


def bundles_with(reg: Registry, group: str) -> list[str]:
    k = norm(group)
    return [b.name for b in reg.bundles.values() if k in {norm(g) for g in b.groups}]
