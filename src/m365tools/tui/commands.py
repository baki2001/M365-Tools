"""Command palette (``:``): the actions of the current screen by name, then the toolbox-wide ones (open a tool,
home, sign in, switch tenant, quit). A screen contributes its own entries through ``palette_commands()``."""

from __future__ import annotations

from collections.abc import Callable
from functools import partial
from typing import Any

from textual.command import DiscoveryHit, Hit, Hits, Provider


class M365Commands(Provider):
    def _commands(self) -> list[tuple[str, Callable[[], Any], str]]:
        app: Any = self.app
        own = getattr(self.screen, "palette_commands", None)
        cmds: list[tuple[str, Callable[[], Any], str]] = list(own()) if own is not None else []
        cmds += [
            ("Open spmig", partial(app.open_tool, "spmig"), "SharePoint library migration"),
            ("Open groups", partial(app.open_tool, "groups"), "access groups, drift and onboarding"),
            ("Home", app.action_go_home, "the toolbox home screen  (g)"),
            ("Sign in", app.action_login, "log in / refresh the token for a tenant  (ctrl+l)"),
            ("Open runs folder", app.action_open_runs_dir, "show and copy the runs folder path"),
            ("Quit", app.action_quit_app, "(q)"),
        ]
        for name in sorted(app.cfg.tenants):
            cmds.append((f"Switch tenant {name}", partial(app.switch_tenant, name),
                         "tenant used for sign-in, pickers and groups"))
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
