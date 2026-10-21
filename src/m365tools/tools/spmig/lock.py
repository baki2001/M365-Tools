"""Make a (migrated) source library read-only. Never widens access: only principals that can write are
downgraded to Read; everyone else (Read, Limited Access, ...) is left untouched."""

from __future__ import annotations

from dataclasses import dataclass, field

from m365tools.core.events import EventSink, ItemDone, LogLine, PhaseChanged, null_sink

from .gateway import SpGateway
from .models import LibraryRef

READ_ROLE = "Read"


@dataclass(slots=True)
class LockChange:
    principal_id: int
    title: str
    from_roles: list[str] = field(default_factory=list)
    to_role: str = READ_ROLE
    skipped_reason: str | None = None

    @property
    def will_change(self) -> bool:
        return self.skipped_reason is None


async def preview_lock(gw: SpGateway, lib: LibraryRef, include_owners: bool = False) -> list[LockChange]:
    """List what a lock would do. Read-only."""
    owner_id = await gw.get_associated_owner_group_id()
    out: list[LockChange] = []
    for ra in await gw.get_role_assignments(lib):
        change = LockChange(ra.principal_id, ra.principal_title, list(ra.role_names))
        if not ra.can_write:
            change.skipped_reason = "cannot write (left untouched)"
        elif owner_id is not None and ra.principal_id == owner_id and not include_owners:
            change.skipped_reason = "site owners group (kept)"
        out.append(change)
    return out


async def apply_lock(
    gw: SpGateway, lib: LibraryRef, changes: list[LockChange], sink: EventSink = null_sink
) -> list[LockChange]:
    """Apply the non-skipped changes from :func:`preview_lock`. Breaks inheritance (copying current
    assignments) once if the library still inherits. Returns the changes applied."""
    todo = [c for c in changes if c.will_change]
    sink(PhaseChanged("lock", f"{len(todo)} principal(s) to downgrade"))
    if not todo:
        return []
    if not await gw.has_unique_role_assignments(lib):
        sink(LogLine("info", f"breaking permission inheritance on '{lib.title}' (copying current roles)"))
        await gw.break_inheritance(lib, copy=True)
    done: list[LockChange] = []
    for c in todo:
        await gw.set_reader_only(lib, c.principal_id)
        done.append(c)
        sink(ItemDone(c.title, True, 0.0, f"{', '.join(c.from_roles)} -> {c.to_role}"))
        sink(LogLine("ok", f"{c.title}: {', '.join(c.from_roles)} -> {c.to_role}"))
    return done
