import json
from datetime import datetime

import pytest

from m365tools.core.events import CollectingSink
from m365tools.tools.spmig.gateway import RoleAssignment
from m365tools.tools.spmig.job import BlockingFindingsError, Job
from m365tools.tools.spmig.models import Endpoint, JobOptions, Route

from .fake_gateway import FakeGateway, factory_for

SRC_URL = "https://contoso.sharepoint.com/sites/src"
DST_URL = "https://contoso.sharepoint.com/sites/dst"
T = datetime(2022, 1, 1)


def setup():
    src = FakeGateway(SRC_URL)
    src.add_library("Servicemanagement")
    src.add_file("Klanten/a.pdf", b"aaa", modified=T, author_email="jan@contoso.com")
    src.add_file("Klanten/b.docx", b"bbbb", modified=T, author_email="gone.user@contoso.com")
    src.add_file("0. Archief/2019/old.pdf", b"old", modified=T)
    src.lib().roles = [RoleAssignment(4, "Members", "8", ["Edit"], True)]
    dst = FakeGateway(DST_URL)
    dst.add_library("Documenten", "Gedeelde documenten", default=True)
    dst.add_library("Archief")
    dst.add_user("jan@contoso.com")
    dst.usage = (0, 10**9)
    return src, dst


def make_job(tmp_path, src, dst, **kw):
    opts = JobOptions(
        source=Endpoint("t", SRC_URL, "Servicemanagement"), destination=Endpoint("t", DST_URL),
        excludes=["0. Archief/**"],
        routes=[Route("0. Archief", Endpoint("t", DST_URL, "Archief"))], **kw,
    )
    n = iter(range(100))
    return Job(factory_for({SRC_URL: src, DST_URL: dst}), opts, CollectingSink(),
               lambda mode: tmp_path / f"{next(n)}-{mode}", migrator_kwargs={"retry_pause": 0.0})


async def test_audit_dryrun_migrate_verify_lock(tmp_path):
    src, dst = setup()
    job = make_job(tmp_path, src, dst, lock_source_after=True)

    audit = await job.audit()
    assert not audit.blocking
    assert (audit.run_dir / "findings.csv").exists() and (audit.run_dir / "plan.json").exists()
    assert (audit.run_dir / "report.html").exists()
    assert any(f.category == "users" and f.path == "gone.user@contoso.com" for f in audit.findings)

    plan, dry = await job.dryrun()
    assert dry.copied == 3 and len(plan.to_copy()) == 3 and not dst.write_calls()

    summary = await job.migrate()
    assert (summary.copied, summary.failed, summary.metadata_warnings) == (3, 0, 1)
    assert dst.file("Klanten/a.pdf").content == b"aaa"
    assert dst.file("2019/old.pdf", "Archief").content == b"old"  # routed
    assert dst.file("0. Archief") is None
    assert (summary.run_dir / "results.csv").exists() and (summary.run_dir / "journal.jsonl").exists()
    assert json.loads((summary.run_dir / "plan.json").read_text())["counts"]["copy"] == 3
    assert src.reader_set == [4]  # lock_source_after

    result, vs = await job.verify()
    assert vs.failed == 0
    assert result.counts().get("missing", 0) == 0

    changes = await job.lock(apply=False)
    assert [c.will_change for c in changes] == [False]  # already Read only


async def test_migrate_refuses_on_blocking_unless_forced(tmp_path):
    src, dst = setup()
    dst.write_denied = True
    job = make_job(tmp_path, src, dst)
    with pytest.raises(BlockingFindingsError) as ei:
        await job.migrate()
    assert ei.value.blocking[0].category == "write_access"
    assert not dst.write_calls()
    summary = await job.migrate(force=True)
    assert summary.copied == 3


async def test_resume_in_same_run_dir(tmp_path):
    src, dst = setup()
    job = make_job(tmp_path, src, dst)
    first = await job.migrate()
    src.calls.clear()
    second = await job.migrate(resume_run_dir=first.run_dir)
    assert second.copied == 0 and second.skipped == 3
    assert not [c for c in src.calls if c[0] == "download_file"]


async def test_unresolvable_route_destination_blocks_and_never_copies_to_main(tmp_path):
    src, dst = setup()
    del dst.libraries["Archief"]  # route destination library missing
    job = make_job(tmp_path, src, dst)
    job.opts.excludes = []  # nothing else would keep the routed files out of the main destination
    with pytest.raises(BlockingFindingsError) as ei:
        await job.migrate()
    assert any(f.category == "destination" and f.path == "0. Archief" for f in ei.value.blocking)
    summary = await job.migrate(force=True)
    assert summary.copied == 2
    assert dst.file("0. Archief/2019/old.pdf") is None  # not dumped into the main library
