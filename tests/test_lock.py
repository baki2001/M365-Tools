from m365tools.core.events import CollectingSink
from m365tools.tools.spmig.gateway import RoleAssignment
from m365tools.tools.spmig.lock import apply_lock, preview_lock

from .fake_gateway import FakeGateway


def site():
    gw = FakeGateway()
    lib = gw.add_library("Docs")
    gw.owner_group_id = 3
    lib.roles = [
        RoleAssignment(3, "Site Owners", "8", ["Full Control"], True),
        RoleAssignment(4, "Site Members", "8", ["Edit"], True),
        RoleAssignment(5, "Site Visitors", "8", ["Read"], False),
        RoleAssignment(6, "Guest", "1", ["Limited Access"], False),
        RoleAssignment(7, "Jan", "1", ["Contribute", "Limited Access"], True),
    ]
    return gw, lib


async def test_preview_never_widens_and_skips_owners():
    gw, lib = site()
    changes = {c.principal_id: c for c in await preview_lock(gw, lib.ref)}
    assert changes[3].skipped_reason and "owner" in changes[3].skipped_reason
    assert changes[5].skipped_reason and changes[6].skipped_reason
    assert changes[4].will_change and changes[7].will_change
    assert changes[7].from_roles == ["Contribute", "Limited Access"] and changes[7].to_role == "Read"
    assert not [c for c in gw.calls if c[0] in ("break_inheritance", "set_reader_only")]  # preview is read-only

    with_owners = {c.principal_id: c for c in await preview_lock(gw, lib.ref, include_owners=True)}
    assert with_owners[3].will_change


async def test_apply_breaks_inheritance_once_and_leaves_limited_access():
    gw, lib = site()
    changes = await preview_lock(gw, lib.ref)
    applied = await apply_lock(gw, lib.ref, changes, CollectingSink())
    assert gw.break_count == 1
    assert sorted(gw.reader_set) == [4, 7]
    assert [c.principal_id for c in applied] == [4, 7]
    roles = {r.principal_id: r.role_names for r in lib.roles}
    assert roles[6] == ["Limited Access"] and roles[5] == ["Read"] and roles[3] == ["Full Control"]
    # already unique: no second break
    lib.roles.append(RoleAssignment(8, "Late", "1", ["Edit"], True))
    await apply_lock(gw, lib.ref, await preview_lock(gw, lib.ref))
    assert gw.break_count == 1 and gw.reader_set[-1] == 8


async def test_apply_nothing_to_do_does_not_break_inheritance():
    gw, lib = site()
    lib.roles = [RoleAssignment(6, "Guest", "1", ["Limited Access"], False)]
    assert await apply_lock(gw, lib.ref, await preview_lock(gw, lib.ref)) == []
    assert gw.break_count == 0
