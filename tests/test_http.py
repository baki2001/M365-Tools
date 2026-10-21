from __future__ import annotations

from pathlib import Path

import httpx
import pytest
import respx

from m365tools.core.errors import AccessDeniedError, AuthError, M365Error, NotFoundError, ThrottledError
from m365tools.core.events import CollectingSink, Throttled
from m365tools.core.graph import GraphClient
from m365tools.core.http import SpClient, api, parse_retry_after

SITE = "https://contoso.sharepoint.com/sites/a"


class FakeTokens:
    def __init__(self) -> None:
        self.calls: list[tuple[str, bool]] = []

    async def token(self, scope: str, force_refresh: bool = False) -> str:
        self.calls.append((scope, force_refresh))
        return f"tok{len(self.calls)}"


class Sleeps:
    def __init__(self) -> None:
        self.waits: list[float] = []

    async def __call__(self, seconds: float) -> None:
        self.waits.append(seconds)


@pytest.fixture
def tokens() -> FakeTokens:
    return FakeTokens()


@pytest.fixture
def sink() -> CollectingSink:
    return CollectingSink()


@pytest.fixture
async def client(tokens: FakeTokens, sink: CollectingSink):
    c = SpClient(tokens, "https://contoso.sharepoint.com/.default", sink=sink, max_attempts=3, base_delay=0.01)
    c._sleep = Sleeps()
    async with c:
        yield c


def test_api_join() -> None:
    assert api(SITE + "/", "/web/lists") == f"{SITE}/_api/web/lists"
    assert parse_retry_after("12") == 12.0
    assert parse_retry_after(None) is None


@respx.mock
async def test_headers(client: SpClient) -> None:
    route = respx.post(api(SITE, "web/folders")).mock(return_value=httpx.Response(200, json={"ok": 1}))
    assert await client.post_json(api(SITE, "web/folders"), json={"ServerRelativeUrl": "/x"}) == {"ok": 1}
    req = route.calls.last.request
    assert req.headers["Authorization"] == "Bearer tok1"
    assert req.headers["Accept"] == "application/json;odata=nometadata"
    assert req.headers["Content-Type"] == "application/json;odata=verbose"
    get = respx.get(api(SITE, "web")).mock(return_value=httpx.Response(200, json={}))
    await client.get_json(api(SITE, "web"), headers={"Accept": "application/json"})
    assert "Content-Type" not in get.calls.last.request.headers
    assert get.calls.last.request.headers["Accept"] == "application/json"


@respx.mock
async def test_retry_on_429_honours_retry_after(client: SpClient, sink: CollectingSink) -> None:
    route = respx.get(api(SITE, "web")).mock(side_effect=[
        httpx.Response(429, headers={"Retry-After": "7"}),
        httpx.Response(503, headers={"Retry-After": "2"}),
        httpx.Response(200, json={"Title": "A"}),
    ])
    assert await client.get_json(api(SITE, "web")) == {"Title": "A"}
    assert route.call_count == 3
    assert client._sleep.waits == [7.0, 2.0]  # type: ignore[attr-defined]
    throttled = [e for e in sink.events if isinstance(e, Throttled)]
    assert [(t.seconds, t.attempt) for t in throttled] == [(7.0, 1), (2.0, 2)]
    assert throttled[0].what == "GET /sites/a/_api/web"


@respx.mock
async def test_retry_exhausted_raises_throttled(client: SpClient) -> None:
    respx.get(api(SITE, "web")).mock(return_value=httpx.Response(429, headers={"Retry-After": "1"}))
    with pytest.raises(ThrottledError) as ei:
        await client.get_json(api(SITE, "web"))
    assert ei.value.retry_after == 1.0
    assert len(client._sleep.waits) == 2  # type: ignore[attr-defined]


@respx.mock
async def test_transport_error_retried(client: SpClient) -> None:
    respx.get(api(SITE, "web")).mock(side_effect=[httpx.ConnectError("boom"), httpx.Response(200, json={})])
    assert await client.get_json(api(SITE, "web")) == {}
    assert len(client._sleep.waits) == 1  # type: ignore[attr-defined]


@respx.mock
async def test_404_odata_message(client: SpClient) -> None:
    body = {"odata.error": {"code": "-2130575338", "message": {"lang": "en-US",
                                                              "value": "List 'Docs' does not exist."}}}
    respx.get(api(SITE, "web/lists")).mock(return_value=httpx.Response(404, json=body))
    with pytest.raises(NotFoundError, match="List 'Docs' does not exist"):
        await client.get_json(api(SITE, "web/lists"))


@respx.mock
async def test_403_and_500(client: SpClient) -> None:
    respx.get(api(SITE, "a")).mock(return_value=httpx.Response(403, json={"error": {"message": "nope"}}))
    respx.get(api(SITE, "b")).mock(return_value=httpx.Response(500, text="server broke"))
    with pytest.raises(AccessDeniedError, match="nope"):
        await client.get_json(api(SITE, "a"))
    with pytest.raises(M365Error, match="HTTP 500.*server broke"):
        await client.get_json(api(SITE, "b"))


@respx.mock
async def test_401_refreshes_once(client: SpClient, tokens: FakeTokens) -> None:
    route = respx.get(api(SITE, "web")).mock(side_effect=[httpx.Response(401), httpx.Response(200, json={})])
    await client.get_json(api(SITE, "web"))
    assert tokens.calls[-1][1] is True
    assert route.calls.last.request.headers["Authorization"] == "Bearer tok2"
    respx.get(api(SITE, "x")).mock(return_value=httpx.Response(401))
    with pytest.raises(AuthError):
        await client.get_json(api(SITE, "x"))


@respx.mock
async def test_iter_odata_follows_next_link(client: SpClient) -> None:
    first = api(SITE, "web/lists/items")
    second = first + "?$skiptoken=Paged%3dTRUE%26p_ID%3d2"
    respx.get(first, params={"$top": "2"}).mock(return_value=httpx.Response(
        200, json={"value": [{"Id": 1}, {"Id": 2}], "odata.nextLink": second}))
    respx.get(second).mock(return_value=httpx.Response(
        200, json={"value": [{"Id": 3}], "@odata.nextLink": "items?$skiptoken=3"}))
    respx.get(api(SITE, "web/lists/items?$skiptoken=3")).mock(return_value=httpx.Response(
        200, json={"value": [{"Id": 4}]}))
    ids = [i["Id"] async for i in client.iter_odata(first, params={"$top": "2"})]
    assert ids == [1, 2, 3, 4]


@respx.mock
async def test_digest_cached(client: SpClient) -> None:
    route = respx.post(api(SITE, "contextinfo")).mock(return_value=httpx.Response(
        200, json={"FormDigestValue": "0xABC", "FormDigestTimeoutSeconds": 1800}))
    assert await client.request_digest(SITE) == "0xABC"
    assert await client.request_digest(SITE + "/") == "0xABC"
    assert route.call_count == 1
    other = "https://contoso.sharepoint.com/sites/b"
    respx.post(api(other, "contextinfo")).mock(return_value=httpx.Response(
        200, json={"d": {"GetContextWebInformation": {"FormDigestValue": "0xDEF",
                                                      "FormDigestTimeoutSeconds": 30}}}))
    assert await client.request_digest(other) == "0xDEF"
    assert await client.request_digest(other) == "0xDEF"  # timeout-60 < 0 => refetch
    assert respx.calls.call_count == 3


@respx.mock
async def test_download_and_upload(client: SpClient, tmp_path: Path) -> None:
    url = api(SITE, "web/GetFileByServerRelativePath(decodedurl='/sites/a/D/x.bin')/$value")
    respx.get(url).mock(return_value=httpx.Response(200, content=b"x" * 3000))
    dest = tmp_path / "dl" / "x.bin"
    assert await client.download_to(url, dest) == 3000
    assert dest.read_bytes() == b"x" * 3000
    assert not (tmp_path / "dl" / "x.bin.part").exists()

    up = respx.post(api(SITE, "files/add")).mock(return_value=httpx.Response(200, json={"UniqueId": "u"}))
    assert await client.upload_bytes(api(SITE, "files/add"), dest) == {"UniqueId": "u"}
    assert up.calls.last.request.content == b"x" * 3000
    assert up.calls.last.request.headers["Content-Type"] == "application/octet-stream"


@respx.mock
async def test_graph_client(tokens: FakeTokens) -> None:
    respx.get("https://graph.microsoft.com/v1.0/sites", params={"search": "docs"}).mock(
        return_value=httpx.Response(200, json={"value": [
            {"displayName": "Docs", "webUrl": "https://c.sharepoint.com/sites/docs", "description": "d"}]}))
    users = respx.get("https://graph.microsoft.com/v1.0/users").mock(side_effect=[
        httpx.Response(200, json={"value": []}),
        httpx.Response(200, json={"value": [{"mail": None, "userPrincipalName": "o'neil@c.nl"}]}),
    ])
    async with GraphClient(tokens) as g:
        assert await g.search_sites("docs") == [
            {"title": "Docs", "url": "https://c.sharepoint.com/sites/docs", "description": "d"}]
        user = await g.find_user_by_mail("o'neil@c.nl")
    assert user is not None and user["userPrincipalName"] == "o'neil@c.nl"
    assert "o''neil" in str(users.calls.last.request.url.params["$filter"])
    assert tokens.calls[0][0] == "https://graph.microsoft.com/.default"
    assert users.calls.last.request.headers["Accept"] == "application/json"


@respx.mock
async def test_token_fetched_per_attempt(client: SpClient, tokens: FakeTokens) -> None:
    # A long Retry-After wait can outlive the token: every attempt asks the (memoising) provider again.
    route = respx.get(api(SITE, "web")).mock(side_effect=[
        httpx.Response(429, headers={"Retry-After": "1"}), httpx.Response(200, json={})])
    await client.get_json(api(SITE, "web"))
    assert tokens.calls == [(client.scope, False), (client.scope, False)]
    assert route.calls.last.request.headers["Authorization"] == "Bearer tok2"
