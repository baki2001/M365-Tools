"""Bundle suggestions from actual memberships (no ML).

People are grouped by their exact set of prefix groups. Every set shared by at least two people is a candidate
bundle, named after the most common job title (else department) of its people. Everyone else gets the nearest
candidate by Jaccard similarity, expressed as ``candidate + addons - missing``.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field

from .models import Bundle, PersonEntry, Registry, Snapshot, norm


@dataclass(slots=True)
class Candidate:
    key: str  # stable id within one suggestion: "1", "2", ...
    name: str  # proposed bundle name (editable)
    groups: list[str]
    people: list[str]
    existing: str = ""  # name of a registry bundle with exactly these groups, if any


@dataclass(slots=True)
class NearMatch:
    upn: str
    candidate: str  # Candidate.key; "" when nothing overlaps
    score: float
    addons: list[str] = field(default_factory=list)  # groups the person has beyond the candidate
    missing: list[str] = field(default_factory=list)  # candidate groups the person lacks


@dataclass(slots=True)
class Suggestion:
    candidates: list[Candidate]
    near: list[NearMatch]

    def candidate(self, key_or_name: str) -> Candidate | None:
        k = norm(key_or_name)
        return next((c for c in self.candidates if c.key == k or norm(c.name) == k), None)


def jaccard(a: set[str], b: set[str]) -> float:
    if not a and not b:
        return 0.0
    return len(a & b) / len(a | b)


def _name_for(snap: Snapshot, people: list[str], taken: set[str], n: int) -> str:
    persons = [p for p in (snap.person(u) for u in people) if p]
    for attr in ("job_title", "department"):
        counts = Counter(getattr(p, attr).strip() for p in persons if getattr(p, attr).strip())
        if counts:
            base = counts.most_common(1)[0][0]
            break
    else:
        base = f"Bundle-{n}"
    name, i = base, 2
    while norm(name) in taken:
        name, i = f"{base}-{i}", i + 1
    taken.add(norm(name))
    return name


def candidates(snap: Snapshot, reg: Registry, min_people: int = 2) -> list[Candidate]:
    sets: dict[frozenset[str], list[str]] = {}
    for p in snap.people:
        gs = frozenset(snap.groups_of(p.upn))
        if gs:
            sets.setdefault(gs, []).append(p.upn)
    existing = {frozenset(norm(g) for g in b.groups): b.name for b in reg.bundles.values()}
    taken = {norm(b) for b in reg.bundles}
    shared = sorted(((gs, ups) for gs, ups in sets.items() if len(ups) >= min_people),
                    key=lambda x: (-len(x[1]), -len(x[0]), sorted(x[0])))
    out = []
    for n, (gs, ups) in enumerate(shared, 1):
        ex = existing.get(frozenset(norm(g) for g in gs), "")
        name = ex or _name_for(snap, ups, taken, n)
        out.append(Candidate(key=str(n), name=name, groups=sorted(gs, key=str.lower), people=sorted(ups),
                             existing=ex))
    return out


def nearest(snap: Snapshot, cands: list[Candidate]) -> list[NearMatch]:
    """Nearest candidate for everyone with prefix groups who is not exactly in one (ties: bigger candidate,
    then lower key)."""
    inside = {u for c in cands for u in c.people}
    out = []
    for p in snap.people:
        mine = snap.groups_of(p.upn)
        if not mine or p.upn in inside:
            continue
        best: tuple[float, int, int, Candidate] | None = None
        for c in cands:
            s = jaccard({norm(g) for g in mine}, {norm(g) for g in c.groups})
            rank = (s, len(c.people), -int(c.key), c)
            if s > 0 and (best is None or rank[:3] > best[:3]):
                best = rank
        if best is None:
            out.append(NearMatch(p.upn, "", 0.0, addons=sorted(mine, key=str.lower)))
            continue
        c = best[3]
        cg = {norm(g) for g in c.groups}
        mg = {norm(g) for g in mine}
        out.append(NearMatch(p.upn, c.key, round(best[0], 3),
                             addons=sorted((g for g in mine if norm(g) not in cg), key=str.lower),
                             missing=[g for g in c.groups if norm(g) not in mg]))
    return out


def suggest(snap: Snapshot, reg: Registry) -> Suggestion:
    cands = candidates(snap, reg)
    return Suggestion(cands, nearest(snap, cands))


NEAR_MIN_SCORE = 0.5


def accept(reg: Registry, sug: Suggestion, accepted: dict[str, str], *, include_near: bool = True,
           min_score: float = NEAR_MIN_SCORE) -> list[str]:
    """Write accepted candidates into ``reg`` (in memory): ``accepted`` maps candidate key/name -> bundle name.

    The bundle gets the candidate's groups (an existing bundle keeps its groups). Its exact people get the bundle;
    with ``include_near`` people whose nearest candidate it is (Jaccard >= ``min_score``) get the bundle plus
    their extra groups as add-ons.
    Existing person entries are only extended, never overwritten. Returns human-readable change lines."""
    changes: list[str] = []
    for key, name in accepted.items():
        c = sug.candidate(key)
        if c is None:
            raise KeyError(key)
        name = name or c.name
        if reg.bundle(name) is None:
            reg.bundles[name] = Bundle(name=name, purpose="", groups=list(c.groups))
            changes.append(f"bundle {name}: {', '.join(c.groups)}")
        bname = reg.bundle(name).name  # type: ignore[union-attr]
        members = [(u, []) for u in c.people]
        if include_near:
            members += [(m.upn, m.addons) for m in sug.near if m.candidate == c.key and m.score >= min_score]
        for upn, addons in members:
            entry = reg.people.setdefault(upn, PersonEntry(upn=upn))
            added = []
            if bname not in entry.bundles:
                entry.bundles.append(bname)
                added.append(f"bundle {bname}")
            for a in addons:
                if a not in entry.addons:
                    entry.addons.append(a)
                    added.append(f"addon {a}")
            if added:
                changes.append(f"person {upn}: + " + ", ".join(added))
    return changes
