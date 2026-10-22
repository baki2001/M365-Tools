"""Delegated token acquisition with MSAL (public client), one provider per tenant.

Strategy: silent (persistent cache) first; otherwise interactive sign-in in the system browser when one can be
opened (desktop session, WSL or ``$BROWSER``), else the device-code flow whose instructions are surfaced through
the ``prompt`` callback as an :class:`~m365tools.core.events.AuthPrompt`. Blocking MSAL calls run in a worker
thread so the event loop stays responsive. Core never prints: all user-facing text goes through ``prompt``.
"""

from __future__ import annotations

import asyncio
import os
import re
import shutil
import tempfile
import threading
import time
import webbrowser
from collections.abc import Callable
from pathlib import Path
from typing import Any

import msal

from m365tools.core.config import TenantConfig
from m365tools.core.errors import AuthError
from m365tools.core.events import AuthPrompt
from m365tools.core.paths import cache_dir

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
AUTHORITY_BASE = "https://login.microsoftonline.com"
INTERACTIVE_TIMEOUT_S = 300
_EXPIRY_MARGIN_S = 120

PromptCallback = Callable[[AuthPrompt], None]


def sharepoint_scope(tenant: TenantConfig) -> str:
    """``https://<sharepoint_host>/.default`` (all delegated SharePoint permissions consented for the app)."""
    return f"https://{tenant.sharepoint_host}/.default"


def is_wsl() -> bool:
    """True when running under Windows Subsystem for Linux."""
    try:
        return "microsoft" in Path("/proc/version").read_text(encoding="utf-8", errors="ignore").lower()
    except OSError:
        return False


def can_open_browser() -> bool:
    """Whether an interactive browser sign-in is likely to work in this environment."""
    return bool(
        os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY") or os.environ.get("BROWSER") or is_wsl()
    )


_wsl_browser_registered = False


def _register_wsl_browser() -> None:
    """On WSL, make ``webbrowser`` prefer ``wslview`` (opens the Windows default browser)."""
    global _wsl_browser_registered
    if _wsl_browser_registered or os.environ.get("BROWSER") or not is_wsl():
        return
    path = shutil.which("wslview")
    if path:
        webbrowser.register("wslview", None, webbrowser.GenericBrowser(path), preferred=True)
    _wsl_browser_registered = True


def cache_file_for(tenant: TenantConfig) -> Path:
    """Location of the persisted MSAL token cache for ``tenant``."""
    key = tenant.name or tenant.tenant_id
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", key) or "default"
    return cache_dir() / f"msal_{safe}.json"


# Errors that come from Entra itself (app registration / consent / user refused): the device-code flow would hit
# exactly the same wall, so they are reported instead of falling back.
_FINAL_ERRORS = frozenset({"access_denied", "invalid_client", "unauthorized_client", "invalid_scope",
                           "consent_required"})
# AADSTS65001 consent missing, 65004 user declined consent, 65005 misconfigured app permissions,
# 90094/90099 admin consent required, 700016 app not found in tenant, 7000218 public client flows disabled.
_FINAL_AADSTS = ("AADSTS65001", "AADSTS65004", "AADSTS65005", "AADSTS90094", "AADSTS90099", "AADSTS700016",
                 "AADSTS7000218")


def _is_final_error(result: dict[str, Any]) -> bool:
    if result.get("error") in _FINAL_ERRORS or result.get("suberror") == "consent_required":
        return True
    desc = str(result.get("error_description") or "")
    codes = [str(c) for c in result.get("error_codes") or []]
    return any(code in desc or code.removeprefix("AADSTS") in codes for code in _FINAL_AADSTS)


def _error_text(result: dict[str, Any] | None, fallback: str) -> str:
    if not result:
        return fallback
    desc = result.get("error_description") or ""
    err = result.get("error") or ""
    text = desc.strip() or err or fallback
    return text if not err or err in text else f"{err}: {text}"


class TokenProvider:
    """Hands out access tokens for one tenant; safe to share between concurrent tasks."""

    def __init__(self, tenant: TenantConfig, prompt: PromptCallback | None = None) -> None:
        self.tenant = tenant
        self.prompt = prompt
        self.cache_path = cache_file_for(tenant)
        self._cache = msal.SerializableTokenCache()
        self._file_lock = threading.Lock()
        self._lock: asyncio.Lock | None = None
        self._memo: dict[str, tuple[str, float]] = {}
        self._load_cache()
        self.app = msal.PublicClientApplication(
            tenant.client_id,
            authority=f"{AUTHORITY_BASE}/{tenant.tenant_id}",
            token_cache=self._cache,
        )

    # ------------------------------------------------------------------------------------------- cache file

    def _load_cache(self) -> None:
        try:
            self._cache.deserialize(self.cache_path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            return
        except (OSError, ValueError):
            # Corrupt or unreadable cache: start fresh; it will be overwritten on the next save.
            return

    def _save_cache(self) -> None:
        with self._file_lock:
            if not self._cache.has_state_changed:
                return
            self.cache_path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
            fd, tmp = tempfile.mkstemp(prefix=".msal-", suffix=".json", dir=self.cache_path.parent)
            try:
                os.fchmod(fd, 0o600)
                with os.fdopen(fd, "w", encoding="utf-8") as fh:
                    fh.write(self._cache.serialize())
                    fh.flush()
                    os.fsync(fh.fileno())
                os.replace(tmp, self.cache_path)
                os.chmod(self.cache_path, 0o600)
            except BaseException:
                Path(tmp).unlink(missing_ok=True)
                raise
            self._cache.has_state_changed = False

    # ------------------------------------------------------------------------------------------- public API

    async def token(self, scope: str, force_refresh: bool = False) -> str:
        """Return an access token for ``scope``; ``force_refresh`` bypasses cached access tokens."""
        if not force_refresh:
            memo = self._memo.get(scope)
            if memo and memo[1] > time.time():
                return memo[0]
        if self._lock is None:
            self._lock = asyncio.Lock()
        async with self._lock:
            if not force_refresh:
                memo = self._memo.get(scope)
                if memo and memo[1] > time.time():
                    return memo[0]
            result = await asyncio.to_thread(self._acquire_silent, scope, force_refresh)
            if not result or "access_token" not in result:
                result = await asyncio.to_thread(self._acquire_interactive, scope)
            self._save_cache()
            if not result or "access_token" not in result:
                raise AuthError(
                    f"Sign-in to tenant '{self.tenant.name or self.tenant.tenant_id}' failed: "
                    + _error_text(result, "no token returned")
                )
            token = str(result["access_token"])
            expires_in = float(result.get("expires_in") or 3600)
            self._memo[scope] = (token, time.time() + max(0.0, expires_in - _EXPIRY_MARGIN_S))
            return token

    def accounts(self) -> list[dict[str, Any]]:
        """Accounts present in the token cache (``username`` holds the UPN)."""
        return list(self.app.get_accounts())

    def logout(self) -> None:
        """Forget all cached accounts/tokens for this tenant and delete the cache file."""
        for acc in self.app.get_accounts():
            self.app.remove_account(acc)
        self._memo.clear()
        with self._file_lock:
            self.cache_path.unlink(missing_ok=True)
            self._cache.has_state_changed = False

    # ------------------------------------------------------------------------------------------- internals

    def _emit(self, message: str, url: str | None = None, code: str | None = None) -> None:
        if self.prompt is not None:
            self.prompt(AuthPrompt(tenant=self.tenant.name or self.tenant.tenant_id, message=message,
                                   url=url, code=code))

    def _acquire_silent(self, scope: str, force_refresh: bool) -> dict[str, Any] | None:
        for account in self.app.get_accounts():
            result = self.app.acquire_token_silent([scope], account=account, force_refresh=force_refresh)
            if result and "access_token" in result:
                return result
        return None

    def _acquire_interactive(self, scope: str) -> dict[str, Any] | None:
        if can_open_browser():
            _register_wsl_browser()
            try:
                result = self.app.acquire_token_interactive(
                    [scope],
                    prompt="select_account",
                    port=0,
                    timeout=INTERACTIVE_TIMEOUT_S,
                    on_before_launching_ui=lambda **_: self._emit(
                        "Opening your browser to sign in. Complete the sign-in there "
                        f"(waiting up to {INTERACTIVE_TIMEOUT_S // 60} minutes)."
                    ),
                )
            except Exception as exc:  # noqa: BLE001 - browser/listener failures: fall back to device code
                self._emit(f"Browser sign-in failed ({exc}); falling back to device code.")
                return self._acquire_device_code(scope)
            if result and "access_token" in result:
                return result
            if result and _is_final_error(result):
                return result
            self._emit("Browser sign-in did not complete; falling back to device code.")
        return self._acquire_device_code(scope)

    def _acquire_device_code(self, scope: str) -> dict[str, Any]:
        flow = self.app.initiate_device_flow(scopes=[scope])
        if "user_code" not in flow:
            raise AuthError("Could not start device-code sign-in: " + _error_text(flow, "unknown error"))
        self._emit(
            str(flow.get("message") or "Open the URL and enter the code to sign in."),
            url=flow.get("verification_uri") or flow.get("verification_url"),
            code=flow.get("user_code"),
        )
        return self.app.acquire_token_by_device_flow(flow)


def tenant_setup_instructions(tenant_id: str) -> str:
    """Plain-text instructions for creating the Entra app registration this toolkit signs in with."""
    return f"""\
m365tools signs in with a *public client* (no secret) app registration in tenant {tenant_id}.

Option A - you already have one (recommended)
  If you created an app earlier with PnP PowerShell, e.g.
      Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP" -Tenant {tenant_id} -Interactive
  just reuse its Application (client) ID (looks like 12345678-xxxx-xxxx-xxxx-xxxxxxxxxxxx). That app already
  has the right redirect URI and delegated permissions, and it works with this toolkit. Check that Graph
  Sites.Read.All and User.Read are listed too (needed for site search fallback and user mapping).

Option B - create a new one in the Entra admin center (https://entra.microsoft.com)
  1. Identity > Applications > App registrations > New registration
       Name: m365tools
       Supported account types: Accounts in this organizational directory only (single tenant)
       Redirect URI: platform "Public client/native (mobile & desktop)", value  http://localhost
  2. Authentication: set "Allow public client flows" = Yes (needed for device-code sign-in).
  3. API permissions > Add a permission (all *Delegated*):
       SharePoint       AllSites.FullControl
       Microsoft Graph  Sites.Read.All
       Microsoft Graph  User.Read
     then click "Grant admin consent for <your tenant>" (requires a Global/Privileged Role Administrator).
  4. Copy the Application (client) ID from the Overview page.

Extra Graph permissions for the groups tool (m365 groups / onboard / offboard), also *Delegated*, granted once
with admin consent like above:
  reading (inventory, suggest, drift, explain):
       Microsoft Graph  Group.Read.All
       Microsoft Graph  User.Read.All
       Microsoft Graph  Directory.Read.All
       Microsoft Graph  Organization.Read.All   (licence names)
       Microsoft Graph  Application.Read.All
  writing (groups sync, onboard, offboard):
       Microsoft Graph  Group.ReadWrite.All
       Microsoft Graph  GroupMember.ReadWrite.All
  Licences and enterprise-app assignments are only read, never changed. After granting, sign in again
  ("m365 tenant login <name>" or ctrl+l in the app) so the cached token picks up the new permissions.

Then register the tenant with m365tools ("m365 tenant add"), giving the tenant id ({tenant_id}), the
Application (client) ID and the SharePoint host (<yourtenant>.sharepoint.com), and test it with
"m365 tenant login".

Sign-in opens your browser (on WSL via wslview / $BROWSER); without a browser a device code is shown that you
enter at https://microsoft.com/devicelogin. Tokens are cached in ~/.cache/m365tools (mode 600).
"""
