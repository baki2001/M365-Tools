import asyncio
from datetime import datetime

from m365tools.core.errors import TransferError
from m365tools.core.events import CollectingSink, ItemDone, LogLine, Stats
from m365tools.tools.spmig.engine import Migrator
from m365tools.tools.spmig.inventory import build_inventory
from m365tools.tools.spmig.models import Endpoint, JobOptions
from m365tools.tools.spmig.plan import build_plan
from m365tools.tools.spmig.usermap import UserMap, UserMapEntry

from .fake_gateway import FakeGateway

SRC = Endpoint("t", "https://contoso.sharepoint.com/sites/src", "Docs")
DST = Endpoint("t", "https://contoso.sharepoint.com/sites/dst")
T = datetime(2022, 3, 4, 5, 6, 7)


class MemJournal:
    def __init__(self, done=()):
        self.entries = [{"rel_path": p, "ok": True} for p in done]

    def done_paths(self):
        return {e["rel_path"] for e in self.entries if e["ok"]}

    def append(self, record):
        self.entries.append(record)


class MemStatus:
    def __init__(self):
        self.updates = []

    def update(self, force=False, **fields):
        self.updates.append((force, fields))


def gateways(n_files=3):
    src = FakeGateway("https://contoso.sharepoint.com/sites/src")
    src.add_library("Docs")
    dst = FakeGateway("https://contoso.sharepoint.com/sites/dst")
    dst.add_library("Documenten", "Gedeelde documenten")
    dst.add_user("jan@contoso.com")
    for i in range(n_files):
        src.add_file(f"Klanten/Sub/f{i}.pdf", f"content-{i}".encode(), modified=T, created=T,
                     author_email="jan@contoso.com", editor_email="jan@contoso.com")
    return src, dst


async def make(src, dst, tmp_path, *, journal=None, usermap=None, sink=None, **kw):
    opts = JobOptions(source=SRC, destination=DST, workers=kw.pop("workers", 2))
    src_inv = await build_inventory(src, src.lib().ref)
    dst_inv = await build_inventory(dst, dst.lib().ref)
    plan = build_plan(src_inv, dst_inv, opts, dst_inv.base_url, {})
    kw.setdefault("retry_pause", 0.0)
    return Migrator(src, dst, plan, opts, tmp_path / "run", sink or CollectingSink(),
                    journal if journal is not None else MemJournal(), MemStatus(), usermap, **kw)


async def test_full_run_copies_files_and_sets_metadata(tmp_path):
    src, dst = gateways(3)
    sink = CollectingSink()
    m = await make(src, dst, tmp_path, sink=sink)
    summary = await m.run()
    assert (summary.copied, summary.failed, summary.skipped, summary.metadata_warnings) == (3, 0, 0, 0)
    assert summary.bytes_copied == sum(len(f"content-{i}") for i in range(3))
    assert not summary.stopped_early
    for i in range(3):
        e = dst.file(f"Klanten/Sub/f{i}.pdf")
        assert e is not None and e.content == f"content-{i}".encode()
        assert e.modified == T and e.created == T and e.author_email == "jan@contoso.com"
    assert dst.file("Klanten") is not None and dst.file("Klanten").is_folder
    assert len(m.journal.entries) == 3 and all(x["ok"] for x in m.journal.entries)
    assert (tmp_path / "run" / "results.csv").read_text(encoding="utf-8-sig").startswith(
        "time;result;path;size_mb;seconds;message")
    assert not (tmp_path / "run" / "tmp").exists()
    assert m.status.updates[-1][0] is True  # final forced status write
    stats = [e for e in sink.events if isinstance(e, Stats)]
    assert stats[-1].done == 3 and stats[-1].bytes_done == stats[-1].bytes_total


async def test_resume_skips_journaled(tmp_path):
    src, dst = gateways(3)
    m = await make(src, dst, tmp_path, journal=MemJournal(["Klanten/Sub/f0.pdf"]))
    summary = await m.run()
    assert summary.copied == 2 and summary.skipped == 1
    downloads = [c[1] for c in src.calls if c[0] == "download_file"]
    assert not any(d.endswith("f0.pdf") for d in downloads)
    assert dst.file("Klanten/Sub/f0.pdf") is None


async def test_user_not_found_falls_back_to_dates_only(tmp_path):
    src, dst = gateways(1)
    src.add_file("gone.docx", b"x", modified=T, author_email="gone.user@contoso.com", editor_email="gone.user@contoso.com")
    sink = CollectingSink()
    summary = await (await make(src, dst, tmp_path, sink=sink)).run()
    assert summary.failed == 0 and summary.copied == 2 and summary.metadata_warnings == 1
    e = dst.file("gone.docx")
    assert e.modified == T and e.author_email == "migrator@contoso.com"
    assert any(isinstance(ev, LogLine) and ev.level == "warn" and "gone.docx" in ev.text for ev in sink.events)


async def test_usermap_maps_author(tmp_path):
    src, dst = gateways(0)
    src.add_file("a.txt", b"x", author_email="old@legacy.nl", editor_email="old@legacy.nl")
    um = UserMap([UserMapEntry("old@legacy.nl", "jan@contoso.com", "Jan", "mapped")])
    summary = await (await make(src, dst, tmp_path, usermap=um)).run()
    assert summary.metadata_warnings == 0
    assert dst.file("a.txt").author_email == "jan@contoso.com"


async def test_request_stop_stops_early(tmp_path):
    src, dst = gateways(6)
    src.download_delay = 0.01
    m = None

    def sink(ev):
        if isinstance(ev, ItemDone) and m is not None:
            m.request_stop()

    m = await make(src, dst, tmp_path, sink=sink, workers=1)
    summary = await m.run()
    assert summary.stopped_early
    assert 1 <= summary.copied < 6
    assert summary.failed == 0


async def test_dry_run_touches_nothing(tmp_path):
    src, dst = gateways(3)
    sink = CollectingSink()
    m = await make(src, dst, tmp_path, sink=sink, dry_run=True)
    src.calls.clear()
    dst.calls.clear()
    summary = await m.run()
    assert src.calls == [] and dst.calls == []
    assert summary.mode == "dryrun" and summary.copied == 3 and summary.bytes_copied == 0
    done = [e for e in sink.events if isinstance(e, ItemDone)]
    assert done and all(e.ok and e.message.startswith("would copy") for e in done)
    stats = [e for e in sink.events if isinstance(e, Stats)]
    assert all(s.bytes_total == 0 for s in stats) and stats[-1].done == stats[-1].total == 3
    assert m.journal.entries == []


async def test_retry_on_non_http_error_then_fail_on_m365_error(tmp_path):
    src, dst = gateways(2)
    src.download_failures["/sites/src/Docs/Klanten/Sub/f0.pdf"] = [OSError("disk hiccup")]
    src.download_failures["/sites/src/Docs/Klanten/Sub/f1.pdf"] = [TransferError("gave up")]
    summary = await (await make(src, dst, tmp_path)).run()
    assert summary.copied == 1 and summary.failed == 1
    assert dst.file("Klanten/Sub/f0.pdf") is not None
    assert src.calls.count(("download_file", "/sites/src/Docs/Klanten/Sub/f1.pdf")) == 1


async def test_watchdog_warns_once_per_slow_file(tmp_path):
    src, dst = gateways(1)
    src.download_delay = 0.05
    sink = CollectingSink()
    m = await make(src, dst, tmp_path, sink=sink, watchdog_interval=0.01, watchdog_min_s=0.0)
    await m.run()
    warns = [e for e in sink.events if isinstance(e, LogLine) and "has been running" in e.text]
    assert len(warns) == 1


async def test_cancel_propagates(tmp_path):
    src, dst = gateways(2)
    src.download_delay = 0.05
    m = await make(src, dst, tmp_path)
    task = asyncio.create_task(m.run())
    await asyncio.sleep(0.01)
    task.cancel()
    try:
        await task
    except asyncio.CancelledError:
        pass
    assert not (tmp_path / "run" / "tmp").exists()


class FailingJournal(MemJournal):
    def __init__(self, failed=()):
        super().__init__()
        self.entries = [{"rel_path": p, "ok": False} for p in failed]

    def failed_paths(self):
        ok = {e["rel_path"] for e in self.entries if e["ok"]}
        return {e["rel_path"] for e in self.entries if not e["ok"]} - ok


async def test_sigint_handler_restored_to_previous(tmp_path):
    import signal

    def mine(signum, frame):  # pragma: no cover - never delivered
        pass

    before = signal.signal(signal.SIGINT, mine)
    try:
        src, dst = gateways(1)
        await (await make(src, dst, tmp_path)).run()
        assert signal.getsignal(signal.SIGINT) is mine
    finally:
        signal.signal(signal.SIGINT, before)


async def test_folder_failure_counts_as_failed(tmp_path):
    src, dst = gateways(0)
    src.add_folder("Leeg")

    async def boom(path):
        raise OSError("folder create failed")

    dst.ensure_folder = boom
    summary = await (await make(src, dst, tmp_path)).run()
    assert summary.failed == 1  # an empty folder that could not be created is still a failure (no lock, exit 1)


async def test_previously_failed_file_retried_even_if_plan_skips(tmp_path):
    # Run 1 uploaded f0 but failed on metadata: the destination copy has Modified = upload time (newer than the
    # source), so the re-built plan says SKIP. The journal's failure must make the engine copy it again.
    src, dst = gateways(1)
    dst.add_file("Klanten/Sub/f0.pdf", b"content-0", modified=datetime(2026, 10, 21, 10, 0, 0))
    m = await make(src, dst, tmp_path, journal=FailingJournal(["Klanten/Sub/f0.pdf"]))
    assert m.plan.entries[-1].action.value == "skip"
    summary = await m.run()
    assert summary.copied == 1 and summary.skipped == 0
    assert dst.file("Klanten/Sub/f0.pdf").modified == T


async def test_worker_exception_cancels_other_workers(tmp_path):
    src, dst = gateways(4)
    original = src.download_file

    async def download(url, dest):
        if not url.endswith("f0.pdf"):
            await asyncio.sleep(0.5)  # still in flight when f0's journal write fails
        await original(url, dest)

    src.download_file = download

    class BrokenJournal(MemJournal):
        def append(self, record):
            raise OSError("disk full")

    m = await make(src, dst, tmp_path, journal=BrokenJournal(), workers=4)
    try:
        await m.run()
    except OSError:
        pass
    else:  # pragma: no cover
        raise AssertionError("expected OSError")
    others = [t for t in asyncio.all_tasks() if t is not asyncio.current_task()]
    assert all(t.done() for t in others)
    assert not (tmp_path / "run" / "tmp").exists()
