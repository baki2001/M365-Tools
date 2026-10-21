"""Authenticated HTTP client for SharePoint REST (and the base for Graph).

``SpClient`` wraps ``httpx.AsyncClient`` and adds: bearer tokens from a :class:`TokenProvider`, OData JSON
defaults, a retry policy for throttling (429/502/503/504 honouring ``Retry-After``) and transient transport
errors with exponential backoff, ``Throttled`` events so the UI can show "throttled, waiting Ns", error mapping
to the toolkit's exception hierarchy, paging, streaming download/upload and a per-site request-digest cache.
"""

from __future__ import annotations

import asyncio
import json as _json
import logging
import random
import time
from collections.abc import AsyncIterator, Awaitable, Callable
from datetime import UTC, datetime
from email.utils import parsedate_to_datetime
from pathlib import Path
from types import TracebackType
from typing import Any, Protocol, Self
from urllib.parse import urljoin, urlsplit

import httpx

from m365tools.core.errors import (
    AccessDeniedError,
    AuthError,
    M365Error,
    NotFoundError,
    ThrottledError,
    TransferError,
)
from m365tools.core.events import EventSink, Throttled, null_sink

log = logging.getLogger("m365tools")

RETRY_STATUSES: frozenset[int] = frozenset({429, 502, 503, 504})
RETRY_EXCEPTIONS: tuple[type[Exception], ...] = (
    httpx.ConnectError,
    httpx.ConnectTimeout,
    httpx.ReadTimeout,
    httpx.WriteTimeout,
    httpx.ReadError,
    httpx.WriteError,
    httpx.RemoteProtocolError,
)
MAX_BACKOFF_S = 300.0
MAX_RETRY_AFTER_S = 900.0
ACCEPT_NOMETADATA = "application/json;odata=nometadata"
CONTENT_TYPE_VERBOSE = "application/json;odata=verbose"
_CHUNK = 1024 * 1024


class TokenSource(Protocol):
    """What SpClient needs from a token provider (satisfied by :class:`m365tools.core.auth.TokenProvider`)."""

    async def token(self, scope: str, force_refresh: bool = False) -> str: ...


def api(site_url: str, path: str) -> str:
    """``api("https://x/sites/a/", "/web")`` -> ``"https://x/sites/a/_api/web"``."""
    return f"{site_url.rstrip('/')}/_api/{path.lstrip('/')}"


def short_url(url: str, limit: int = 80) -> str:
    """Path of ``url`` (no host/query), shortened from the left to ``limit`` characters for messages."""
    path = urlsplit(url).path or url
    return path if len(path) <= limit else "…" + path[-(limit - 1):]


def odata_error_message(response: httpx.Response) -> str:
    """Extract the human message from a SharePoint/Graph error body, falling back to text/reason."""
    try:
        body = response.json()
    except (ValueError, UnicodeDecodeError):
        body = None
    if isinstance(body, dict):
        err = body.get("odata.error") or body.get("error")
        if isinstance(err, dict):
            msg = err.get("message")
            if isinstance(msg, dict):
                msg = msg.get("value")
            if msg:
                return str(msg)
            if err.get("code"):
                return str(err["code"])
        if isinstance(body.get("error_description"), str):
            return body["error_description"]
    text = (response.text or "").strip()
    if text and len(text) < 500 and not text.lstrip().startswith("<"):
        return text
    return response.reason_phrase or f"HTTP {response.status_code}"


def parse_retry_after(value: str | None) -> float | None:
    """Seconds from a ``Retry-After`` header (delta-seconds or HTTP date)."""
    if not value:
        return None
    value = value.strip()
    try:
        return max(0.0, float(value))
    except ValueError:
        pass
    try:
        when = parsedate_to_datetime(value)
    except (TypeError, ValueError):
        return None
    if when.tzinfo is None:
        when = when.replace(tzinfo=UTC)
    return max(0.0, (when - datetime.now(UTC)).total_seconds())


class SpClient:
    """Async HTTP client bound to one token provider + scope. Use as ``async with SpClient(...) as c:``."""

    default_accept: str = ACCEPT_NOMETADATA
    json_content_type: str = CONTENT_TYPE_VERBOSE

    def __init__(
        self,
        token_provider: TokenSource,
        scope: str,
        sink: EventSink = null_sink,
        max_attempts: int = 6,
        base_delay: float = 2.0,
    ) -> None:
        self.token_provider = token_provider
        self.scope = scope
        self.sink = sink
        self.max_attempts = max(1, max_attempts)
        self.base_delay = base_delay
        self.http = httpx.AsyncClient(
            timeout=httpx.Timeout(60, read=600), http2=False, follow_redirects=True
        )
        self._sleep: Callable[[float], Awaitable[None]] = asyncio.sleep
        self._digests: dict[str, tuple[str, float]] = {}
        self._digest_lock = asyncio.Lock()

    async def __aenter__(self) -> Self:
        return self

    async def __aexit__(
        self, exc_type: type[BaseException] | None, exc: BaseException | None, tb: TracebackType | None
    ) -> None:
        await self.aclose()

    async def aclose(self) -> None:
        await self.http.aclose()

    # ------------------------------------------------------------------------------------------------ core

    def _backoff(self, attempt: int, retry_after: float | None) -> float:
        if retry_after is not None:
            return min(MAX_RETRY_AFTER_S, retry_after)
        return min(MAX_BACKOFF_S, self.base_delay * 2**attempt) + random.uniform(0, max(0.0, self.base_delay / 2))

    async def _wait(self, seconds: float, attempt: int, what: str) -> None:
        self.sink(Throttled(seconds=seconds, attempt=attempt, what=what))
        log.info("throttled/transient error on %s, waiting %.1fs (attempt %d)", what, seconds, attempt)
        await self._sleep(seconds)

    def _headers(self, token: str, headers: dict[str, str] | None, has_json: bool) -> httpx.Headers:
        h = httpx.Headers({"Accept": self.default_accept})
        if has_json:
            h["Content-Type"] = self.json_content_type
        if headers:
            h.update(headers)
        h["Authorization"] = f"Bearer {token}"
        return h

    async def request(
        self,
        method: str,
        url: str,
        *,
        json: Any = None,
        data: Any = None,
        content: bytes | str | None = None,
        headers: dict[str, str] | None = None,
        params: dict[str, Any] | None = None,
        stream: bool = False,
    ) -> httpx.Response:
        """Send a request with auth, retries and error mapping. Returns only 2xx/3xx responses.

        With ``stream=True`` the caller must close the response (``await resp.aclose()``).
        """
        method = method.upper()
        what = f"{method} {short_url(url)}"
        has_json = json is not None
        if has_json:
            content = _json.dumps(json).encode("utf-8")
        refreshed = False
        just_refreshed = False
        attempt = 0
        token = ""
        while True:
            # Ask the provider every attempt (it memoises until shortly before expiry): Retry-After waits can add
            # up to more than a token's lifetime, and a stale token would burn the one forced refresh below.
            if not just_refreshed:
                token = await self.token_provider.token(self.scope)
            just_refreshed = False
            req = self.http.build_request(
                method, url, content=content, data=data, params=params,
                headers=self._headers(token, headers, has_json),
            )
            try:
                resp = await self.http.send(req, stream=stream)
            except RETRY_EXCEPTIONS as exc:
                attempt += 1
                if attempt >= self.max_attempts:
                    raise TransferError(f"{what}: {type(exc).__name__}: {exc} (gave up after {attempt} attempts)") \
                        from exc
                await self._wait(self._backoff(attempt - 1, None), attempt, f"{what} ({type(exc).__name__})")
                continue

            status = resp.status_code
            if status < 400:
                return resp

            if stream:
                await resp.aread()
                await resp.aclose()

            if status in RETRY_STATUSES:
                attempt += 1
                retry_after = parse_retry_after(resp.headers.get("Retry-After"))
                if attempt >= self.max_attempts:
                    raise ThrottledError(
                        f"{what}: HTTP {status} after {attempt} attempts: {odata_error_message(resp)}",
                        retry_after=retry_after,
                    )
                await self._wait(self._backoff(attempt - 1, retry_after), attempt, what)
                continue

            if status == 401:
                if not refreshed:
                    refreshed = just_refreshed = True
                    token = await self.token_provider.token(self.scope, force_refresh=True)
                    continue
                raise AuthError(f"{what}: unauthorized (401): {odata_error_message(resp)}")
            message = f"{what}: {odata_error_message(resp)}"
            if status == 403:
                raise AccessDeniedError(f"Access denied - {message}")
            if status == 404:
                raise NotFoundError(f"Not found - {message}")
            raise M365Error(f"HTTP {status} - {message}")

    # ------------------------------------------------------------------------------------------------ helpers

    @staticmethod
    def _json_or_empty(resp: httpx.Response) -> Any:
        if not resp.content:
            return {}
        try:
            return resp.json()
        except ValueError:
            return {}

    async def get_json(self, url: str, **kw: Any) -> Any:
        """GET and return parsed JSON."""
        resp = await self.request("GET", url, **kw)
        return self._json_or_empty(resp)

    async def post_json(self, url: str, json: Any = None, **kw: Any) -> Any:
        """POST (optionally with a JSON body) and return parsed JSON, ``{}`` for an empty body."""
        resp = await self.request("POST", url, json=json, **kw)
        return self._json_or_empty(resp)

    async def iter_odata(self, url: str, **kw: Any) -> AsyncIterator[dict[str, Any]]:
        """Yield every item of an OData collection, following ``odata.nextLink`` / ``@odata.nextLink``."""
        next_url: str | None = url
        params = kw.pop("params", None)
        while next_url:
            body = await self.get_json(next_url, params=params, **kw)
            params = None  # the next link already carries the query
            if isinstance(body, dict) and isinstance(body.get("d"), dict):
                body = body["d"]
            items = body.get("value", body.get("results", [])) if isinstance(body, dict) else []
            for item in items or []:
                yield item
            link = None
            if isinstance(body, dict):
                link = body.get("odata.nextLink") or body.get("@odata.nextLink") or body.get("__next")
            next_url = urljoin(next_url, link) if link else None

    async def download_to(self, url: str, dest: Path) -> int:
        """Stream ``url`` to ``dest`` (via ``dest.part`` + rename). Returns bytes written.

        Transport errors while reading the body restart the download, with the same backoff as ``request``.
        """
        dest.parent.mkdir(parents=True, exist_ok=True)
        part = dest.with_name(dest.name + ".part")
        what = f"GET {short_url(url)}"
        attempt = 0
        while True:
            resp = await self.request("GET", url, headers={"Accept": "*/*"}, stream=True)
            written = 0
            try:
                with part.open("wb") as fh:
                    async for chunk in resp.aiter_bytes(_CHUNK):
                        fh.write(chunk)
                        written += len(chunk)
            except RETRY_EXCEPTIONS as exc:
                attempt += 1
                part.unlink(missing_ok=True)
                if attempt >= self.max_attempts:
                    raise TransferError(f"{what}: download failed after {attempt} attempts: {exc}") from exc
                await self._wait(self._backoff(attempt - 1, None), attempt, f"{what} ({type(exc).__name__})")
                continue
            except BaseException:
                part.unlink(missing_ok=True)
                raise
            finally:
                await resp.aclose()
            part.replace(dest)
            return written

    async def upload_bytes(
        self, url: str, data: bytes | Path, headers: dict[str, str] | None = None, method: str = "POST"
    ) -> Any:
        """Send raw bytes (or a file's contents) as the request body; returns parsed JSON (``{}`` if none)."""
        payload = data.read_bytes() if isinstance(data, Path) else data
        h = {"Content-Type": "application/octet-stream"}
        if headers:
            h.update(headers)
        resp = await self.request(method, url, content=payload, headers=h)
        return self._json_or_empty(resp)

    async def request_digest(self, site_url: str) -> str:
        """``X-RequestDigest`` value for ``site_url``, cached until shortly before it expires."""
        key = site_url.rstrip("/").lower()
        async with self._digest_lock:
            cached = self._digests.get(key)
            if cached and cached[1] > time.monotonic():
                return cached[0]
            body = await self.post_json(api(site_url, "contextinfo"))
            if isinstance(body, dict) and isinstance(body.get("d"), dict):
                body = body["d"].get("GetContextWebInformation", body["d"])
            try:
                value = str(body["FormDigestValue"])
            except (KeyError, TypeError) as exc:
                raise M365Error(f"{site_url}: contextinfo returned no FormDigestValue") from exc
            timeout = float(body.get("FormDigestTimeoutSeconds") or 1800)
            self._digests[key] = (value, time.monotonic() + max(0.0, timeout - 60))
            return value
