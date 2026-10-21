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

from rich.console import Console
from rich.markup import escape
from rich.table import Table

from m365tools.core import sharepoint as sp
from m365tools.core.auth import TokenProvider, sharepoint_scope
from m365tools.core.config import get_tenant, load_config
from m365tools.core.errors import M365Error, NotFoundError
from m365tools.core.http import SpClient

console = Console(highlight=False)
RESULTS: list[tuple[str, bool]] = []


def facts(*lines: str) -> None:
    for line in lines:
        console.print(f"    [dim]{escape(line)}[/]")


def done(title: str, items: str, ok: bool, error: str | None = None) -> None:
    RESULTS.append((title, ok))
    if ok:
        console.print(f"[green]✓ {escape(title)}[/]  [dim](checks: {escape(items)})[/]")
    else:
        console.print(f"[red]✗ {escape(title)}[/]  [dim](checks: {escape(items)})[/]")
        if error:
            console.print(f"    [red]{escape(error)}[/]")


async def step(title: str, items: str, coro: Awaitable[Any]) -> tuple[bool, Any]:
    """Await ``coro``; returns (ok, result). Facts must be printed by the caller *after* the status line."""
    try:
        result = await coro
    except M365Error as exc:
        done(title, items, False, f"{type(exc).__name__}: {exc}")
        return False, None
    done(title, items, True)
    return True, result


def auth_prompt(p: Any) -> None:
    console.print(f"[yellow]![/] {escape(str(p.message))} {escape(str(p.url or ''))} [bold]{escape(str(p.code or ''))}[/]")


async def run(tenant_name: str, site: str, email: str | None) -> int:
    cfg = load_config()
    tenant = get_tenant(cfg, tenant_name)
    site = site.rstrip("/")
    provider = TokenProvider(tenant, prompt=auth_prompt)
    console.print(f"[bold cyan]Smoke test[/] tenant [bold]{escape(tenant_name)}[/]  site {escape(site)}  [dim](read-only)[/]\n")
    async with SpClient(provider, sharepoint_scope(tenant), sink=lambda ev: console.print(f"    [dim]event {escape(str(ev))}[/]")) as client:
        ok, web = await step("get_web", "6", sp.get_web(client, site))
        if ok and web:
            facts(f"title={web.title!r} server_relative_url={web.server_relative_url!r} language={web.language}")

        ok, lib = await step("get_default_library", "6, 14", sp.get_default_library(client, site))
        if ok and lib:
            facts(f"{lib.title!r} id={lib.list_id} items={lib.item_count} root={lib.root_server_relative_url}",
                  f"versioning={lib.enable_versioning} limit={lib.major_version_limit}")

        try:
            await sp.get_library(client, site, "__no_such_library__")
            done("list libraries", "6", False, "bogus library name was accepted")
        except NotFoundError as exc:
            done("list libraries", "6", True)
            facts(f"available libraries: {exc.available}")
        except M365Error as exc:
            done("list libraries", "6", False, f"{type(exc).__name__}: {exc}")

        if lib:
            title, items = "first 5 items (iter_items + build_item)", "5, 6, 7, 14"
            try:
                lines: list[str] = []
                async for raw in sp.iter_items(client, lib, page_size=50):
                    item = sp.build_item(raw, lib.root_server_relative_url)
                    if item is None:
                        continue
                    kind = "dir " if item.is_folder else "file"
                    lines.append(f"{kind} {item.rel_path!r} size={item.size} version={item.version!r} "
                                 f"author={item.author_email!r} editor={item.editor_email!r} "
                                 f"prog_id={item.prog_id!r} modified={item.modified}")
                    if len(lines) >= 5:
                        break
                done(title, items, True)
                facts(*(lines or ["(library is empty)"]))
            except M365Error as exc:
                done(title, items, False, f"{type(exc).__name__}: {exc}")

        ok, usage = await step("get_site_usage", "8", sp.get_site_usage(client, site))
        if ok and usage:
            used, quota = usage
            q = f"{quota / 1024**3:.1f} GB" if quota else "unknown"
            facts(f"used={used / 1024**3:.2f} GB derived quota={q} (compare with the admin center)")

        if lib:
            ok, fields = await step("get_fields", "6", sp.get_fields(client, lib))
            if ok:
                facts(*[f"{f.internal_name:<30} {f.type_as_string:<14} {f.title}" for f in fields or []]
                      or ["(no custom columns)"])

            ok, ras = await step("get_role_assignments", "9, 10", sp.get_role_assignments(client, lib))
            if ok:
                facts(*[f"{ra.principal_title!r:<40} type={ra.principal_type} can_write={ra.can_write} "
                        f"roles={ra.role_names}" for ra in ras or []])

        if email is None:
            try:
                accounts = provider.accounts()
                email = accounts[0]["username"] if accounts else None
            except Exception:  # noqa: BLE001 - informational only
                email = None
        if email:
            ok, user = await step(f"find_site_user({email})", "12", sp.find_site_user(client, site, email))
            if ok:
                facts(str(user) if user else "not found in the site's user list (normal if never visited)")
        else:
            console.print("[yellow]![/] no signed-in account found; pass --email to test find_site_user")

    passed = sum(1 for _, ok in RESULTS if ok)
    failed = len(RESULTS) - passed
    table = Table(title="Summary", show_header=False, border_style="dim")
    table.add_column(style="bold")
    table.add_column(justify="right")
    table.add_row("passed", f"[green]{passed}[/]")
    table.add_row("failed", f"[red]{failed}[/]" if failed else "[dim]0[/]")
    console.print()
    console.print(table)
    console.print("[dim]Remaining checklist items need writes: see docs/REAL-TENANT-CHECKLIST.md[/]")
    return 1 if failed else 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Read-only smoke test of the SharePoint REST layer.")
    ap.add_argument("--tenant", required=True, help="tenant name from config.toml")
    ap.add_argument("--site", required=True, help="site URL, e.g. https://contoso.sharepoint.com/sites/x")
    ap.add_argument("--email", help="account to look up (default: the signed-in account)")
    args = ap.parse_args()
    return asyncio.run(run(args.tenant, args.site, args.email))


if __name__ == "__main__":
    sys.exit(main())
