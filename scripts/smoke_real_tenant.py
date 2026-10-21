"""Read-only smoke test against a real tenant (see docs/REAL-TENANT-CHECKLIST.md).

    uv run python scripts/smoke_real_tenant.py --tenant NAME --site https://contoso.sharepoint.com/sites/x

Performs GET requests only; nothing is created, changed or deleted.
"""

from __future__ import annotations

import argparse
import asyncio
import sys
from collections.abc import Awaitable
from typing import Any

from m365tools.core import sharepoint as sp
from m365tools.core.auth import TokenProvider, sharepoint_scope
from m365tools.core.config import get_tenant, load_config
from m365tools.core.errors import M365Error, NotFoundError
from m365tools.core.http import SpClient


def header(title: str, items: str) -> None:
    print(f"\n=== {title}\n    validates checklist item(s): {items}")


async def step(title: str, items: str, coro: Awaitable[Any]) -> Any:
    header(title, items)
    try:
        result = await coro
    except M365Error as exc:
        print(f"    FAILED: {type(exc).__name__}: {exc}")
        return None
    return result


async def run(tenant_name: str, site: str, email: str | None) -> int:
    cfg = load_config()
    tenant = get_tenant(cfg, tenant_name)
    site = site.rstrip("/")
    provider = TokenProvider(tenant, prompt=lambda p: print(f"[auth] {p.message} {p.url or ''} {p.code or ''}"))
    async with SpClient(provider, sharepoint_scope(tenant), sink=lambda ev: print(f"    [event] {ev}")) as client:
        web = await step("get_web", "6 (basic REST, nometadata)", sp.get_web(client, site))
        if web:
            print(f"    title={web.title!r} server_relative_url={web.server_relative_url!r} language={web.language}")

        lib = await step("get_default_library", "6, 14", sp.get_default_library(client, site))
        if lib:
            print(f"    {lib.title!r} id={lib.list_id} items={lib.item_count} root={lib.root_server_relative_url}")
            print(f"    versioning={lib.enable_versioning} limit={lib.major_version_limit}")

        header("list libraries (get_library with a bogus name)", "6")
        try:
            await sp.get_library(client, site, "__no_such_library__")
        except NotFoundError as exc:
            print(f"    available libraries: {exc.available}")
        except M365Error as exc:
            print(f"    FAILED: {type(exc).__name__}: {exc}")

        if lib:
            header("first 5 items (iter_items + build_item)", "5 keyset paging, 6 fields/expand, 7 ProgId, 14 encoding")
            try:
                shown = 0
                async for raw in sp.iter_items(client, lib, page_size=50):
                    item = sp.build_item(raw, lib.root_server_relative_url)
                    if item is None:
                        continue
                    kind = "dir " if item.is_folder else "file"
                    print(f"    {kind} {item.rel_path!r} size={item.size} version={item.version!r} "
                          f"author={item.author_email!r} editor={item.editor_email!r} prog_id={item.prog_id!r} "
                          f"modified={item.modified}")
                    shown += 1
                    if shown >= 5:
                        break
                if shown == 0:
                    print("    (library is empty)")
            except M365Error as exc:
                print(f"    FAILED: {type(exc).__name__}: {exc}")

        usage = await step("get_site_usage", "8 StoragePercentageUsed fraction vs percent", sp.get_site_usage(client, site))
        if usage:
            used, quota = usage
            q = f"{quota / 1024**3:.1f} GB" if quota else "unknown"
            print(f"    used={used / 1024**3:.2f} GB derived quota={q} (compare with the admin center)")

        if lib:
            fields = await step("get_fields", "6", sp.get_fields(client, lib))
            for f in fields or []:
                print(f"    {f.internal_name:<30} {f.type_as_string:<14} {f.title}")
            if fields is not None and not fields:
                print("    (no custom columns)")

            ras = await step("get_role_assignments", "9 BasePermissions Low/High, 10 roles",
                             sp.get_role_assignments(client, lib))
            for ra in ras or []:
                print(f"    {ra.principal_title!r:<40} type={ra.principal_type} can_write={ra.can_write} "
                      f"roles={ra.role_names}")

        if email is None:
            try:
                accounts = provider.accounts()
                email = accounts[0]["username"] if accounts else None
            except Exception:  # noqa: BLE001 - informational only
                email = None
        if email:
            user = await step(f"find_site_user({email})", "12 user lookup / claims", sp.find_site_user(client, site, email))
            print(f"    {user}" if user else "    not found in the site's user list (normal if never visited)")
        else:
            print("\n(no signed-in account found; pass --email to test find_site_user)")
    print("\nDone. Remaining checklist items need writes: see docs/REAL-TENANT-CHECKLIST.md")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Read-only smoke test of the SharePoint REST layer.")
    ap.add_argument("--tenant", required=True, help="tenant name from config.toml")
    ap.add_argument("--site", required=True, help="site URL, e.g. https://contoso.sharepoint.com/sites/x")
    ap.add_argument("--email", help="account to look up (default: the signed-in account)")
    args = ap.parse_args()
    return asyncio.run(run(args.tenant, args.site, args.email))


if __name__ == "__main__":
    sys.exit(main())
