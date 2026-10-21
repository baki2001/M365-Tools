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


async def build_inventory(
    gw: SpGateway,
    library: LibraryRef,
    folder: str = "",
    sink: EventSink = null_sink,
    *,
    page_size: int = PAGE_SIZE,
    label: str = "",
) -> Inventory:
    """Enumerate every item below ``folder`` (recursive). Progress is emitted once per page of ``page_size``."""
    base = inventory_base(library, folder)
    prefix = base.lower() + "/"
    what = label or library.title
    sink(PhaseChanged("inventory", f"{what}: {base}"))

    items: list[Item] = []
    seen = 0
    bytes_seen = 0
    total = max(library.item_count, 0)

    def progress() -> None:
        sink(Stats(done=seen, total=max(total, seen), bytes_done=bytes_seen, bytes_total=0, mbps=0.0,
                   eta_s=None, failed=0, warnings=0))

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

    if items:
        by_id = {i.id: i for i in items}
        unique = 0
        async for item_id in gw.iter_unique_permission_ids(library):
            it = by_id.get(item_id)
            if it is not None:
                it.unique_permissions = True
                unique += 1
        if unique:
            sink(LogLine("info", f"{what}: {unique} item(s) with unique permissions"))

    items.sort(key=lambda i: (not i.is_folder, i.rel_path.lower()))
    inv = Inventory(library=library, base_url=base, items=items)
    sink(LogLine("info", f"{what}: {len(inv.files())} files, {len(inv.folders())} folders, "
                         f"{inv.total_bytes() / 1024 / 1024:.1f} MB"))
    return inv
