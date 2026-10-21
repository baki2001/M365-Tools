"""Command palette (``:``): every action by name, plus one entry per profile, tenant and section."""

from __future__ import annotations

from collections.abc import Callable
from functools import partial
from typing import Any

from textual.command import DiscoveryHit, Hit, Hits, Provider


class M365Commands(Provider):
    def _commands(self) -> list[tuple[str, Callable[[], Any], str]]:
        app: Any = self.app
        ws = app.workspace
        cmds: list[tuple[str, Callable[[], Any], str]] = [
            ("Audit", ws.action_audit, "inventory both sides and run every check; writes nothing  (a)"),
            ("Dry run", ws.action_dryrun, "compute the full plan; writes nothing  (d)"),
            ("Migrate", ws.action_migrate, "copy files to the destination, asks first  (m)"),
            ("Resume a migration run", ws.action_resume, "continue an earlier run from its journal  (4, r)"),
            ("Verify", ws.action_verify, "compare source and destination  (v)"),
            ("Lock: preview", ws.action_lock_preview, "who would be set to Read on the source  (6, l)"),
            ("Lock: apply", ws.action_lock_apply, "downgrade writers on the source to Read  (6, A)"),
            ("Sign in", app.action_login, "log in / refresh the token for a tenant  (ctrl+l)"),
            ("Open runs folder", app.action_open_runs_dir, "show and copy the runs folder path"),
            ("Help", ws.action_help, "keys for the current section  (?)"),
            ("Quit", app.action_quit_app, "(q)"),
        ]
        for name in sorted(app.cfg.profiles):
            cmds.append((f"Load profile {name}", partial(app.load_profile, name), "replace the job with it"))
        for name in sorted(app.cfg.tenants):
            cmds.append((f"Switch tenant {name}", partial(app.switch_tenant, name),
                         "tenant used for sign-in and new pickers"))
        for i, title in enumerate(("Job", "Checks", "Plan", "Run", "Verify", "Lock", "Runs"), 1):
            cmds.append((f"Go to {title}", partial(ws.action_show, i), f"({i})"))
        return cmds

    async def discover(self) -> Hits:
        for name, cb, help_text in self._commands():
            yield DiscoveryHit(name, cb, help=help_text)

    async def search(self, query: str) -> Hits:
        matcher = self.matcher(query)
        for name, cb, help_text in self._commands():
            score = matcher.match(name)
            if score > 0:
                yield Hit(score, matcher.highlight(name), cb, text=name, help=help_text)
