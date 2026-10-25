# engine/tests/test_retrieve_route.py
import pytest
from mcp.server.mcpserver.exceptions import ToolError

from smartview_mcp import svhttp, tools_query
from smartview_mcp.host import InputError


class _Http:
    def __init__(self, result=None, exc=None):
        self.result, self.exc = result, exc

    def retrieve(self, *a):
        if self.exc:
            raise self.exc
        return self.result


def _route(monkeypatch, http, com=("L", "P", 1.0)):
    calls = []

    async def fake_call(fn, timeout=None, what="call"):
        calls.append(what)
        return com
    monkeypatch.setattr(svhttp, "shared", lambda: http)
    monkeypatch.setattr(tools_query, "_call", fake_call)
    return calls


@pytest.mark.anyio
async def test_http_first(monkeypatch):
    calls = _route(monkeypatch, _Http(result=("l", "p", 0.2)))
    assert await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1) == ("l", "p", 0.2, "http", None)
    assert calls == []


@pytest.mark.anyio
async def test_auto_falls_back_to_com(monkeypatch):
    calls = _route(monkeypatch, _Http(exc=svhttp.SvHttpError("provider down")))
    lay, p, secs, used, note = await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1)
    assert (lay, used, calls) == ("L", "com", ["sv_query"]) and "provider down" in note


@pytest.mark.anyio
async def test_http_only_never_falls_back(monkeypatch):
    calls = _route(monkeypatch, _Http(exc=svhttp.SvHttpError("provider down")))
    with pytest.raises(ToolError, match="source='com'"):
        await tools_query._retrieve("FinMax", {}, [], [], True, False, "http", 1)
    assert calls == []


@pytest.mark.anyio
async def test_input_errors_are_not_retried_in_excel(monkeypatch):
    calls = _route(monkeypatch, _Http(exc=InputError("Invalid query: unknown dimension 'x'")))
    with pytest.raises(ToolError, match="unknown dimension"):
        await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1)
    assert calls == []


@pytest.mark.anyio
async def test_com_skips_http(monkeypatch):
    calls = _route(monkeypatch, _Http(exc=AssertionError("HTTP must not run")))
    assert (await tools_query._retrieve("FinMax", {}, [], [], True, False, "com", 1))[3] == "com"
    assert calls == ["sv_query"]
