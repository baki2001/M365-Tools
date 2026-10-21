"""Build an :class:`Inventory` of a library (or a sub-folder of it) through an :class:`SpGateway`."""

from __future__ import annotations

from m365tools.core.events import EventSink, LogLine, PhaseChanged, Stats, null_sink

from .gateway import SpGateway
from .models import Inventory, Item, LibraryRef

PAGE_SIZE = 2000  # matches the GetItems RowLimit used by core.sharepoint.iter_items


def inventory_base(lib: LibraryRef, folder: str) -> str:
    """Server-relative base URL of an inventory: library root plus optional sub-folder, no trailing slash."""
    root = lib.root_server_relative_url.rstrip("/")
    folder = folder.strip("/")
    return f"{root}/{folder}" if folder else root


def _stats(done: int, total: int, bytes_done: int = 0) -> Stats:
    return Stats(done=done, total=total, bytes_done=bytes_done, bytes_total=0, mbps=0.0, eta_s=None, failed=0,
                 warnings=0)


async def build_inventory(
    gw: SpGateway,
    library: LibraryRef,
    folder: str = "",
    sink: EventSink = null_sink,
    *,
    page_size: int = PAGE_SIZE,
    label: str = "",
    phase: str = "inventory.source",
    permissions_phase: str | None = "permissions.source",
) -> Inventory:
    """Enumerate every item below ``folder`` (recursive). Progress is emitted once per page of ``page_size``.

    Emits ``PhaseChanged(phase)`` for the listing and, unless ``permissions_phase`` is None (then the scan is
    skipped), ``PhaseChanged(permissions_phase)`` plus page-wise :class:`Stats` for the unique-permission scan.
    """
    base = inventory_base(library, folder)
    prefix = base.lower() + "/"
    what = label or library.title
    sink(PhaseChanged(phase, f"{what}: {base}"))

    items: list[Item] = []
    seen = 0
    bytes_seen = 0
    total = max(library.item_count, 0)

    def progress() -> None:
        sink(_stats(seen, max(total, seen), bytes_seen))

    async for raw in gw.iter_items(library):
        seen += 1
        item = gw.build_item(raw, library, base)
        if item is not None and item.server_relative_url.lower().startswith(prefix) and item.rel_path:
            items.append(item)
            if not item.is_folder:
                bytes_seen += item.size
        if seen % page_size == 0:
            progress()
    progress()

    if items and permissions_phase is not None:
        scan_total = max(total, seen)
        sink(PhaseChanged(permissions_phase, f"{scan_total:,} items"))
        scanned = 0

        def on_page(n: int) -> None:
            nonlocal scanned
            scanned += n
            sink(_stats(min(scanned, scan_total), scan_total))

        sink(_stats(0, scan_total))
        by_id = {i.id: i for i in items}
        unique = 0
        async for item_id in gw.iter_unique_permission_ids(library, on_page=on_page):
            it = by_id.get(item_id)
            if it is not None:
                it.unique_permissions = True
                unique += 1
        sink(_stats(scan_total, scan_total))
        if unique:
            sink(LogLine("info", f"{what}: {unique} item(s) with unique permissions"))

    items.sort(key=lambda i: (not i.is_folder, i.rel_path.lower()))
    inv = Inventory(library=library, base_url=base, items=items)
    sink(LogLine("info", f"{what}: {len(inv.files())} files, {len(inv.folders())} folders, "
                         f"{inv.total_bytes() / 1024 / 1024:.1f} MB"))
    return inv
