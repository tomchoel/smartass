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


@pytest.mark.anyio
async def test_cancel_during_http_never_submits_an_excel_job(monkeypatch):
    import threading

    import anyio
    release = threading.Event()

    class Slow:
        def retrieve(self, *a):
            release.wait(5)
            raise svhttp.SvHttpError("late")
    calls = _route(monkeypatch, Slow())
    with anyio.move_on_after(0.2):
        await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1)
    release.set()
    assert calls == []


@pytest.mark.anyio
async def test_pending_cancel_stops_before_the_excel_fallback(monkeypatch):
    import anyio
    calls = _route(monkeypatch, _Http(exc=svhttp.SvHttpError("down")))
    with anyio.CancelScope() as scope:
        scope.cancel()
        await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1)
    assert calls == []


@pytest.mark.anyio
async def test_both_routes_failing_names_both(monkeypatch):
    _route(monkeypatch, _Http(exc=svhttp.SvHttpError("provider down")))

    async def bad_call(fn, timeout=None, what="call"):
        raise ToolError("Excel did not answer. Next step: call sv_restart.")
    monkeypatch.setattr(tools_query, "_call", bad_call)
    with pytest.raises(ToolError) as e:
        await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1)
    assert "Excel did not answer" in str(e.value) and "HTTP route also failed" in str(e.value)
    assert "provider down" in e.value.http_note and str(e.value).count("Next step:") == 1


@pytest.mark.anyio
async def test_note_message_is_capped(monkeypatch):
    _route(monkeypatch, _Http(exc=svhttp.SvHttpError("x" * 2000)))
    note = (await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1))[4]
    assert len(note) < 600


@pytest.mark.anyio
async def test_huge_grid_skips_http(monkeypatch):
    calls = _route(monkeypatch, _Http(exc=AssertionError("HTTP must not run")))
    big = int(tools_query.config.SVHTTP_TIMEOUT_S * 200) + 1000
    res = await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1, n_cells=big)
    assert res[3:] == ("com", None) and calls == ["sv_query"]


class _Ctx:
    def __init__(self):
        self.progress = []

    async def report_progress(self, a, b, msg=None):
        self.progress.append(msg)


@pytest.mark.anyio
async def test_progress_before_fallback(monkeypatch):
    _route(monkeypatch, _Http(exc=svhttp.SvHttpError("down")))
    ctx = _Ctx()
    await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1, ctx=ctx)
    assert any("using Smart View in Excel" in m for m in ctx.progress)


class _P:
    columns, rows, row_labels, unknown_members = ["Acc", "v"], [["x", "1"]], [[]], []
    pov_problems, no_access_cells, suppressed_rows, total_rows = [], 0, [], 1


class _L:
    pov = [("Z", "z"), ("A", "a")]
    n_cells, nr = 1, 1


_ROW = [{"dimension": "Account", "members": ["a"]}]


@pytest.mark.anyio
async def test_sv_query_http_route(monkeypatch):
    logged = []
    monkeypatch.setattr(tools_query.qlog, "write", logged.append)
    _route(monkeypatch, _Http(result=(_L(), _P(), 0.1)))
    r = await tools_query.sv_query("FinMax", _ROW, _ROW, _Ctx())
    assert r["route"] == "http" and not r["warnings"]
    assert logged[-1]["route"] == "http" and logged[-1]["http_note"] is None


@pytest.mark.anyio
async def test_sv_query_generic_error_falls_back_and_notes(monkeypatch):
    logged = []
    monkeypatch.setattr(tools_query.qlog, "write", logged.append)
    _route(monkeypatch, _Http(exc=RuntimeError("boom")), com=(_L(), _P(), 1.0))
    r = await tools_query.sv_query("FinMax", _ROW, _ROW, _Ctx())
    assert r["route"] == "com" and "boom" in r["warnings"][-1]
    assert logged[-1]["route"] == "com" and "boom" in logged[-1]["http_note"]   # why HTTP was not used, for audit


@pytest.mark.anyio
async def test_sv_query_input_error_has_one_next_step(monkeypatch):
    logged = []
    monkeypatch.setattr(tools_query.qlog, "write", logged.append)
    _route(monkeypatch, _Http(exc=InputError("Invalid query: unknown dimension 'x'")))
    with pytest.raises(ToolError) as e:
        await tools_query.sv_query("FinMax", _ROW, _ROW, _Ctx(), source="auto")
    assert str(e.value).count("Next step:") == 1
    assert logged[-1]["source"] == "auto" and "error" in logged[-1]
