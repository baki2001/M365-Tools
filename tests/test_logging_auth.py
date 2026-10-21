from __future__ import annotations

import logging
import os
import stat
from pathlib import Path

from m365tools.core.auth import TokenProvider, sharepoint_scope, tenant_setup_instructions
from m365tools.core.config import TenantConfig
from m365tools.core.events import CollectingSink, LogLine
from m365tools.core.logging import SinkHandler, setup_logging


def test_setup_logging_and_sink(tmp_path: Path) -> None:
    logger = setup_logging(tmp_path, level="DEBUG")
    sink = CollectingSink()
    handler = SinkHandler(sink)
    logger.addHandler(handler)
    try:
        logger.warning("careful %s", "now")
        logger.debug("hidden from sink")
    finally:
        logger.removeHandler(handler)
    for h in logger.handlers:
        h.flush()
    text = (tmp_path / "run.log").read_text(encoding="utf-8")
    assert "[WARNING] careful now" in text and "hidden from sink" in text
    assert [(e.level, e.text) for e in sink.events if isinstance(e, LogLine)] == [("warn", "careful now")]
    assert not any(type(h) is logging.StreamHandler for h in logger.handlers)


def test_token_provider_cache_file(tmp_path: Path) -> None:
    tenant = TenantConfig(name="contoso", tenant_id="contoso.onmicrosoft.com",
                          client_id="11111111-2222-3333-4444-555555555555",
                          sharepoint_host="contoso.sharepoint.com")
    assert sharepoint_scope(tenant) == "https://contoso.sharepoint.com/.default"
    tp = TokenProvider(tenant)
    assert tp.cache_path == tmp_path / "cache" / "msal_contoso.json"
    tp._cache.has_state_changed = True
    tp._save_cache()
    assert stat.S_IMODE(os.stat(tp.cache_path).st_mode) == 0o600
    tp.logout()
    assert not tp.cache_path.exists()
    assert "AllSites.FullControl" in tenant_setup_instructions("contoso.onmicrosoft.com")


def test_consent_error_does_not_fall_back_to_device_code(tmp_path: Path, monkeypatch) -> None:
    from m365tools.core import auth

    tenant = TenantConfig(name="c", tenant_id="c.onmicrosoft.com", client_id="cid",
                          sharepoint_host="c.sharepoint.com")
    tp = TokenProvider(tenant)
    monkeypatch.setattr(auth, "can_open_browser", lambda: True)
    consent = {"error": "invalid_grant", "error_codes": [65001],
               "error_description": "AADSTS65001: The user or administrator has not consented to use the app."}
    monkeypatch.setattr(tp.app, "acquire_token_interactive", lambda *a, **kw: consent)
    called = []
    monkeypatch.setattr(tp, "_acquire_device_code", lambda scope: called.append(scope) or {})
    assert tp._acquire_interactive("s") is consent
    assert called == []
    # A non-Entra failure (e.g. the browser flow did not complete) still falls back.
    monkeypatch.setattr(tp.app, "acquire_token_interactive", lambda *a, **kw: {"error": "timeout"})
    tp._acquire_interactive("s")
    assert called == ["s"]
