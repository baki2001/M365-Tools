"""Shared fixtures: isolate XDG dirs so tests never touch the real home directory.

``FakeClient`` implements the ``SpClient`` interface (core/http.py) over a plain ``httpx.AsyncClient`` so domain
modules can be tested with respx without depending on auth/retry logic.
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from pathlib import Path
from typing import Any

import httpx
import pytest

from m365tools.core.errors import AccessDeniedError, M365Error, NotFoundError, ThrottledError

NOMETA = "application/json;odata=nometadata"


class FakeClient:
    """Minimal stand-in for core.http.SpClient: no auth, no retries, same method surface."""

    def __init__(self) -> None:
        self.http = httpx.AsyncClient(headers={"Accept": NOMETA})

    async def aclose(self) -> None:
        await self.http.aclose()

    async def request(
        self,
        method: str,
        url: str,
        *,
        json: Any = None,
        data: Any = None,
        content: Any = None,
        headers: dict[str, str] | None = None,
        params: dict[str, Any] | None = None,
        stream: bool = False,
    ) -> httpx.Response:
        resp = await self.http.request(
            method, url, json=json, data=data, content=content, headers=headers, params=params
        )
        if resp.status_code == 404:
            raise NotFoundError(f"404 {resp.text}")
        if resp.status_code == 403:
            raise AccessDeniedError(f"403 {resp.text}")
        if resp.status_code in (429, 503):
            raise ThrottledError(f"{resp.status_code} {resp.text}")
        if resp.status_code >= 400:
            raise M365Error(f"{resp.status_code} {resp.text}")
        return resp

    async def get_json(self, url: str, **kw: Any) -> dict:
        resp = await self.request("GET", url, **kw)
        return resp.json() if resp.content else {}

    async def post_json(self, url: str, json: Any = None, **kw: Any) -> dict:
        resp = await self.request("POST", url, json=json, **kw)
        return resp.json() if resp.content else {}

    async def iter_odata(self, url: str) -> AsyncIterator[dict]:
        next_url: str | None = url
        while next_url:
            data = await self.get_json(next_url)
            for row in data.get("value", []):
                yield row
            next_url = data.get("odata.nextLink")

    async def download_to(self, url: str, dest: Path) -> None:
        resp = await self.request("GET", url)
        Path(dest).write_bytes(resp.content)

    async def upload_bytes(self, url: str, data: bytes | Path, headers: dict[str, str] | None = None) -> Any:
        """Like SpClient.upload_bytes: returns parsed JSON ({} for an empty body)."""
        body = Path(data).read_bytes() if isinstance(data, Path) else data
        resp = await self.request("POST", url, content=body, headers=headers)
        return resp.json() if resp.content else {}

    async def request_digest(self, site_url: str) -> str:
        return "fake-digest"


@pytest.fixture
async def fake_client() -> AsyncIterator[FakeClient]:
    client = FakeClient()
    try:
        yield client
    finally:
        await client.aclose()


@pytest.fixture(autouse=True)
def isolated_dirs(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    for name in ("CONFIG", "DATA", "CACHE"):
        monkeypatch.setenv(f"M365TOOLS_{name}_DIR", str(tmp_path / name.lower()))
    return tmp_path
