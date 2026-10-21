"""Migration engine: executes a :class:`Plan` (folders -> files -> finalize).

Per file: download to ``run_dir/tmp/<uuid>`` -> upload -> metadata -> journal line -> delete tmp.
Resumable through the journal, stoppable (SIGINT / :meth:`Migrator.request_stop`), watched by a watchdog that
warns about files that take suspiciously long. All SharePoint access goes through :class:`SpGateway`.
"""

from __future__ import annotations

import asyncio
import csv
import shutil
import signal
import threading
import time
import uuid
from collections.abc import Callable
from datetime import datetime
from pathlib import Path
from typing import Any, Protocol

from m365tools.core.errors import M365Error
from m365tools.core.events import (
    EventSink,
    ItemDone,
    ItemStarted,
    ItemStep,
    LogLine,
    PhaseChanged,
    Stats,
    Step,
    null_sink,
)

from .gateway import SpGateway, UserNotFoundError
from .models import Action, JobOptions, LibraryRef, Plan, PlanEntry, RunSummary
from .usermap import UserMap

MB = 1024 * 1024
RESULTS_COLUMNS = ["time", "result", "path", "size_mb", "seconds", "message"]


class JournalLike(Protocol):
    def done_paths(self) -> set[str]: ...
    def append(self, record: dict[str, Any]) -> None: ...


class StatusLike(Protocol):
    def update(self, *, force: bool = False, **fields: Any) -> None: ...


def write_results_csv(path: Path, rows: list[dict[str, Any]], append: bool = False) -> None:
    """``;``-delimited UTF-8 (BOM) like the old PowerShell scripts."""
    new = not (append and path.exists())
    with path.open("w" if new else "a", encoding="utf-8-sig" if new else "utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=RESULTS_COLUMNS, delimiter=";", extrasaction="ignore")
        if new:
            w.writeheader()
        w.writerows(rows)


class Migrator:
    def __init__(
        self,
        src_gw: SpGateway,
        dst_gw: SpGateway,
        plan: Plan,
        opts: JobOptions,
        run_dir: Path,
        sink: EventSink = null_sink,
        journal: JournalLike | None = None,
        status: StatusLike | None = None,
        usermap: UserMap | None = None,
        *,
        dry_run: bool = False,
        route_targets: dict[str, tuple[SpGateway, LibraryRef]] | None = None,
        retries: int = 2,
        retry_pause: float = 5.0,
        watchdog_interval: float = 30.0,
        watchdog_min_s: float = 900.0,
        watchdog_min_bps: float = 0.2 * MB,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self.src_gw = src_gw
        self.dst_gw = dst_gw
        self.plan = plan
        self.opts = opts
        self.run_dir = Path(run_dir)
        self.sink = sink
        self.journal = journal
        self.status = status
        self.usermap = usermap
        self.dry_run = dry_run
        self.route_targets = {k.strip("/"): v for k, v in (route_targets or {}).items()}
        self.retries = retries
        self.retry_pause = retry_pause
        self.watchdog_interval = watchdog_interval
        self.watchdog_min_s = watchdog_min_s
        self.watchdog_min_bps = watchdog_min_bps
        self.clock = clock
        self.tmp_dir = self.run_dir / "tmp"
        self.results: list[dict[str, Any]] = []

        self._stop = asyncio.Event()
        self._loop: asyncio.AbstractEventLoop | None = None
        self._prev_sigint: Any = None  # Python-level SIGINT handler before run()
        self._prev_sigint_handle: Any = None  # an asyncio SIGINT handler on the same loop before run()
        self._inflight: dict[str, tuple[float, int]] = {}
        self._watchdog_warned: set[str] = set()
        self._t0 = 0.0
        self._phase = "init"
        self._current: tuple[str, str] = ("", "")
        self.total = 0
        self.done = 0
        self.bytes_total = 0
        self.bytes_done = 0
        self.copied = 0
        self.skipped = 0
        self.failed = 0
        self.metadata_warnings = 0
        self.bytes_copied = 0

    # ------------------------------------------------------------------------------------------- control

    def request_stop(self) -> None:
        """Finish in-flight files, dispatch nothing new. Safe to call from any thread."""
        loop = self._loop
        if loop is not None and loop.is_running():
            try:
                running = asyncio.get_running_loop()
            except RuntimeError:
                running = None
            if running is not loop:
                loop.call_soon_threadsafe(self._set_stop)
                return
        self._set_stop()

    def _set_stop(self) -> None:
        if not self._stop.is_set():
            self._stop.set()
            self.sink(LogLine("warn", "stop requested: finishing files in progress, then stopping"))

    @property
    def stopping(self) -> bool:
        return self._stop.is_set()

    # ------------------------------------------------------------------------------------------- helpers

    def _target(self, entry: PlanEntry) -> tuple[SpGateway, LibraryRef]:
        if entry.route and entry.route.strip("/") in self.route_targets:
            return self.route_targets[entry.route.strip("/")]
        return self.dst_gw, self.plan.dst

    def _step(self, rel_path: str, step: Step) -> None:
        self._current = (rel_path, step)
        self.sink(ItemStep(rel_path, step))
        self._status()

    def _stats(self) -> Stats:
        elapsed = max(self.clock() - self._t0, 1e-6)
        if self.dry_run:
            return Stats(self.done, self.total, 0, 0, 0.0, None, self.failed, self.metadata_warnings,
                         self.copied, self.skipped)
        mbps = self.bytes_done / elapsed / MB
        eta: float | None = None
        if self.bytes_total > 0 and self.bytes_done > 0:
            eta = (self.bytes_total - self.bytes_done) / (self.bytes_done / elapsed)
        elif self.done > 0:
            eta = (self.total - self.done) * elapsed / self.done
        return Stats(self.done, self.total, self.bytes_done, self.bytes_total, mbps, eta, self.failed,
                     self.metadata_warnings, self.copied, self.skipped)

    def _status(self, force: bool = False) -> None:
        if self.status is None:
            return
        st = self._stats()
        self.status.update(
            force=force, phase=self._phase, mode="dryrun" if self.dry_run else "migrate",
            done=st.done, total=st.total, bytes_done=st.bytes_done, bytes_total=st.bytes_total,
            mbps=round(st.mbps, 3), eta_s=st.eta_s, failed=st.failed, warnings=st.warnings,
            copied=st.copied, skipped=st.skipped, current_file=self._current[0],
            current_step=self._current[1], stopping=self.stopping, run_dir=str(self.run_dir),
        )

    def _result(self, result: str, path: str, size: int, seconds: float, message: str) -> None:
        self.results.append({
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "result": result, "path": path,
            "size_mb": f"{size / MB:.2f}", "seconds": f"{seconds:.1f}", "message": message,
        })

    def _map_user(self, email: str | None) -> str | None:
        if not email:
            return None
        if self.usermap is not None:
            return self.usermap.resolve(email) or email
        return email

    # ------------------------------------------------------------------------------------------- run

    async def run(self) -> RunSummary:
        started = datetime.now()
        self._loop = asyncio.get_running_loop()
        self._t0 = self.clock()
        sigint = self._install_sigint()
        watchdog: asyncio.Task[None] | None = None
        try:
            await self._folders_phase()
            if not self.dry_run:
                watchdog = asyncio.create_task(self._watchdog())
            await self._files_phase()
        finally:
            if watchdog is not None:
                watchdog.cancel()
                try:
                    await watchdog
                except asyncio.CancelledError:
                    pass
            if sigint:
                self._restore_sigint()
            self._finalize()
        return RunSummary(
            mode="dryrun" if self.dry_run else "migrate", started=started, finished=datetime.now(),
            run_dir=self.run_dir, copied=self.copied, skipped=self.skipped, failed=self.failed,
            metadata_warnings=self.metadata_warnings, bytes_copied=self.bytes_copied,
            stopped_early=self.stopping,
        )

    def _install_sigint(self) -> bool:
        if threading.current_thread() is not threading.main_thread():
            return False
        try:
            assert self._loop is not None
            self._prev_sigint = signal.getsignal(signal.SIGINT)
            handlers = getattr(self._loop, "_signal_handlers", None)
            self._prev_sigint_handle = handlers.get(signal.SIGINT) if isinstance(handlers, dict) else None
            self._loop.add_signal_handler(signal.SIGINT, self.request_stop)
            return True
        except (NotImplementedError, RuntimeError, ValueError, AssertionError):
            return False

    def _restore_sigint(self) -> None:
        """Undo :meth:`_install_sigint`, giving SIGINT back to whoever had it before (not the default handler,
        which is what ``remove_signal_handler`` alone would leave behind)."""
        assert self._loop is not None
        try:
            self._loop.remove_signal_handler(signal.SIGINT)
        except (NotImplementedError, RuntimeError, ValueError):
            return
        try:
            prev_handle = self._prev_sigint_handle
            if prev_handle is not None and not prev_handle.cancelled():
                # Someone had registered an asyncio handler on this loop: put it back.
                self._loop.add_signal_handler(signal.SIGINT, prev_handle._callback, *prev_handle._args)
            elif self._prev_sigint is not None:
                signal.signal(signal.SIGINT, self._prev_sigint)
        except (NotImplementedError, RuntimeError, ValueError, TypeError, AttributeError):
            pass
        finally:
            self._prev_sigint = self._prev_sigint_handle = None

    async def _folders_phase(self) -> None:
        self._phase = "folders"
        folders = sorted(self.plan.folders_to_create(),
                         key=lambda e: (e.target_server_relative_url.count("/"), e.target_server_relative_url))
        self.sink(PhaseChanged("folders", f"{len(folders)} folder(s) to create"))
        for entry in folders:
            if self.stopping:
                return
            rel = entry.item.rel_path
            self._step(rel, "folder")
            if self.dry_run:
                continue
            gw, _lib = self._target(entry)
            try:
                await gw.ensure_folder(entry.target_server_relative_url)
            except Exception as exc:  # noqa: BLE001 - files below it will fail and be reported
                self.failed += 1
                self.sink(LogLine("error", f"folder {rel}: {exc}"))
                self._result("FAILED", rel + "/", 0, 0.0, f"folder: {exc}")

    async def _files_phase(self) -> None:
        self._phase = "files"
        done_paths = {p.lower() for p in self.journal.done_paths()} if self.journal and not self.dry_run else set()
        # A file that failed in an earlier run of this run folder may have been uploaded without its metadata
        # (or left empty by an interrupted chunked upload). The destination then looks "newer" than the source
        # (Modified = upload time), so a re-planned SKIP would hide it for good: retry it instead.
        failed_before: set[str] = set()
        failed_fn = getattr(self.journal, "failed_paths", None)
        if failed_fn is not None and not self.dry_run:
            failed_before = {p.lower() for p in failed_fn()} - done_paths
        todo: list[PlanEntry] = []
        for e in self.plan.entries:
            if e.item.is_folder:
                continue
            if e.action == Action.SKIP and e.item.rel_path.lower() in failed_before:
                self.sink(LogLine("info", f"{e.item.rel_path}: failed in an earlier run; copying again"))
                todo.append(e)
            elif e.action == Action.SKIP:
                self.skipped += 1
                self._result("SKIPPED", e.item.rel_path, e.item.size, 0.0, e.reason)
            elif e.action in (Action.COPY, Action.CONFLICT):
                if e.item.rel_path.lower() in done_paths:
                    self.skipped += 1
                    self.sink(ItemDone(e.item.rel_path, True, 0.0, "already done (journal)", skipped=True))
                    self._result("SKIPPED", e.item.rel_path, e.item.size, 0.0, "already done (journal)")
                else:
                    todo.append(e)
        self.total = len(todo)
        self.bytes_total = sum(e.item.size for e in todo)
        resumed = f", {len(done_paths)} already done" if done_paths else ""
        self.sink(PhaseChanged("files", f"{self.total} file(s), {self.bytes_total / MB:.1f} MB{resumed}"))
        self.sink(self._stats())
        if not self.dry_run and todo:
            self.tmp_dir.mkdir(parents=True, exist_ok=True)

        sem = asyncio.Semaphore(max(1, self.opts.workers))
        tasks: list[asyncio.Task[None]] = []

        async def worker(entry: PlanEntry, index: int) -> None:
            try:
                await self._process(entry, index)
            finally:
                sem.release()

        try:
            for index, entry in enumerate(todo, start=1):
                await sem.acquire()
                if self.stopping:
                    sem.release()
                    break
                tasks.append(asyncio.create_task(worker(entry, index)))
            if tasks:
                await asyncio.gather(*tasks)
        except BaseException:
            # CancelledError, or a worker that raised (e.g. journal write failed): don't leave the others running
            # unsupervised while _finalize() removes the tmp folder under them.
            for t in tasks:
                t.cancel()
            await asyncio.gather(*tasks, return_exceptions=True)
            raise

    async def _process(self, entry: PlanEntry, index: int) -> None:
        item = entry.item
        rel = item.rel_path
        t0 = self.clock()
        self._inflight[rel] = (t0, item.size)
        self._current = (rel, "")
        self.sink(ItemStarted(rel, item.size, index, self.total))
        ok = False
        message = ""
        try:
            if self.dry_run:
                ok, message = True, "would copy" + (f" ({entry.reason})" if entry.action == Action.CONFLICT else "")
            else:
                attempt = 0
                while True:
                    try:
                        message = await self._copy_one(entry)
                        ok = True
                        break
                    except asyncio.CancelledError:
                        raise
                    except M365Error as exc:  # HTTP-level retries already happened in SpClient
                        message = f"{type(exc).__name__}: {exc}"
                        break
                    except Exception as exc:  # noqa: BLE001
                        message = f"{type(exc).__name__}: {exc}"
                        if attempt >= self.retries:
                            break
                        attempt += 1
                        self.sink(LogLine("warn", f"{rel}: {message}; retry {attempt}/{self.retries} "
                                                  f"in {self.retry_pause:g}s"))
                        await asyncio.sleep(self.retry_pause)
        finally:
            self._inflight.pop(rel, None)
        seconds = self.clock() - t0

        self.done += 1
        self.bytes_done += item.size
        if ok:
            self.copied += 1
            if not self.dry_run:
                self.bytes_copied += item.size
        else:
            self.failed += 1
            self.sink(LogLine("error", f"{rel}: {message}"))
        if self.journal is not None and not self.dry_run:
            self.journal.append({"rel_path": rel, "ok": ok, "bytes": item.size, "seconds": round(seconds, 2),
                                 "message": message})
        result = "DRYRUN" if self.dry_run else ("OK" if ok else "FAILED")
        if ok and message.startswith("metadata:"):
            result = "WARN"
        self._result(result, rel, item.size, seconds, message or "ok")
        self.sink(ItemDone(rel, ok, seconds, message))
        self.sink(self._stats())
        self._status()

    async def _copy_one(self, entry: PlanEntry) -> str:
        """Copy one file. Returns "" or a metadata warning message (prefixed ``metadata:``)."""
        item = entry.item
        rel = item.rel_path
        gw, lib = self._target(entry)
        tmp = self.tmp_dir / uuid.uuid4().hex
        try:
            self._step(rel, "download")
            await self.src_gw.download_file(item.server_relative_url, tmp)
            self._step(rel, "upload")
            folder, name = entry.target_server_relative_url.rsplit("/", 1)
            res = await gw.upload_file(folder, name, tmp, overwrite=True)
            if not self.opts.restore_metadata:
                return ""
            self._step(rel, "metadata")
            if res.item_id is None:
                self.metadata_warnings += 1
                msg = "metadata: upload returned no item id; dates/authors not restored"
                self.sink(LogLine("warn", f"{rel}: {msg}"))
                return msg
            author = self._map_user(item.author_email)
            editor = self._map_user(item.editor_email)
            try:
                await gw.set_item_metadata(lib, res.item_id, item.created, item.modified, author, editor)
            except UserNotFoundError as exc:
                await gw.set_item_metadata(lib, res.item_id, item.created, item.modified, None, None)
                self.metadata_warnings += 1
                msg = f"metadata: dates only, user not found ({exc})"
                self.sink(LogLine("warn", f"{rel}: {msg}"))
                return msg
            return ""
        finally:
            tmp.unlink(missing_ok=True)

    async def _watchdog(self) -> None:
        while True:
            await asyncio.sleep(self.watchdog_interval)
            now = self.clock()
            elapsed_total = max(now - self._t0, 1e-6)
            measured = self.bytes_copied / elapsed_total
            rate = max(self.watchdog_min_bps, measured)
            for rel, (t0, size) in list(self._inflight.items()):
                if rel in self._watchdog_warned:
                    continue
                limit = max(self.watchdog_min_s, 5 * size / rate)
                if now - t0 > limit:
                    self._watchdog_warned.add(rel)
                    self.sink(LogLine("warn", f"{rel} has been running for {now - t0:.0f}s "
                                              f"(expected < {limit:.0f}s); still waiting"))

    def _finalize(self) -> None:
        self._phase = "finalize"
        self._current = ("", "")
        self.sink(PhaseChanged("finalize"))
        if not self.dry_run:
            shutil.rmtree(self.tmp_dir, ignore_errors=True)
        try:
            self.run_dir.mkdir(parents=True, exist_ok=True)
            write_results_csv(self.run_dir / "results.csv", self.results, append=True)
        except OSError as exc:
            self.sink(LogLine("error", f"could not write results.csv: {exc}"))
        self.sink(self._stats())
        self._status(force=True)
        level = "warn" if self.failed or self.stopping else "ok"
        self.sink(LogLine(level, f"{'dry run' if self.dry_run else 'migration'} "
                                 f"{'stopped' if self.stopping else 'finished'}: {self.copied} copied, "
                                 f"{self.skipped} skipped, {self.failed} failed, "
                                 f"{self.metadata_warnings} metadata warning(s)"))
