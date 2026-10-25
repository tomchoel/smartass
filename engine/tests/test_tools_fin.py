"""tools_fin: area resolution (incl. the abolished areas), prepare() with the warehouse and with the live-pull
fallback, refresh notes, and the fin_pnl tool handing its request to _build_job with the reconciliation hook. Every
figure is invented."""
import sqlite3

import pytest
from fin_util import CAT, FC, WHS, http
from mcp.server.mcpserver.exceptions import ToolError

from smartview_mcp import atlas, extract, fin_pnl, refresh, templates, tools_fin, tools_ws
from smartview_mcp.warehouse import MemoryWarehouse, WarehouseUnavailable


@pytest.fixture
def env(tmp_path, monkeypatch):
    db = sqlite3.connect(tmp_path / "atlas.db")
    db.executescript(atlas.SCHEMA)
    db.executemany("INSERT INTO member VALUES ('FinMax','All Entities',?,?,?,1,0)",
                   [("un.SY.2WHS", "2WHS Wholesale (un.SY)", "un.SY"),
                    ("un.SY.3WHS_W", "3WHS_W W (un.SY)", "un.SY.2WHS"),
                    ("un.SY.U1", "U1 Unit one (un.SY)", "un.SY.3WHS_W"),
                    ("un.SY.U2", "U2 Unit two (un.SY)", "un.SY.3WHS_W")])
    db.commit()
    monkeypatch.setattr(templates, "closed_month", lambda h, c="FinMax": "Sep 26")
    monkeypatch.setattr(templates, "latest_forecast", lambda yy: FC if yy == 26 else None)
    return extract.route(http=http(tmp_path)), db


@pytest.mark.parametrize("text, msg", [("Narnia", "names no business area"),
                                       ("Wholesale vs Business", "names several areas"),
                                       ("Consumer", "abolished on 2026-06-01"),
                                       ("Consumer", "Security Solutions & Mobile .un.SY.2SSM., Entertainment"),
                                       ("Infrastructure Growth", "Successors: Business .*Technology"),
                                       ("2CON", "ask which successor")])
def test_resolve_area_explains(text, msg):
    with pytest.raises(ToolError, match=msg):
        tools_fin.resolve_area(CAT, text)


def test_resolve_area_finds_an_alias():
    assert tools_fin.resolve_area(CAT, "grossist").key == "wholesale"


def test_prepare_fills_the_warehouse_then_computes(env):
    retrieve, db = env
    res, mode, fills, notes = tools_fin.prepare(None, CAT, WHS, None, fin_pnl.COMPARISONS, store=MemoryWarehouse(),
                                                retrieve=retrieve, db=db)
    assert mode == "warehouse" and [f.fill.job for f in fills] == ["base", "top_up", "top_up", "top_up"]
    assert res.scope.units == (("un.SY.U1", "U1 Unit one"), ("un.SY.U2", "U2 Unit two"))
    assert res.unit_bridge.closes() and res.scope.forecast == FC and notes == []
    assert not res.missing()


def test_prepare_answers_from_live_pulls_when_the_warehouse_is_unavailable(env):
    class Locked(MemoryWarehouse):
        def state(self):
            raise WarehouseUnavailable("another smartass process is writing")
    retrieve, db = env
    res, mode, fills, _ = tools_fin.prepare(None, CAT, WHS, "Sep 26", ("ly",), store=Locked(), retrieve=retrieve,
                                            db=db)
    assert mode.startswith("live pulls (the warehouse was unavailable: another smartass")
    assert fills[0].fill.job == "live" and res.fm["ebitda"]["ytd"].current.value is not None


def test_a_skipped_refresh_is_surfaced(env, monkeypatch):
    retrieve, db = env
    store = MemoryWarehouse()
    extract.fill(store, retrieve, CAT, extract.base_grids(CAT, "Sep 26", {26: FC}.get), "base", "Sep 26")
    skip = "warehouse refresh skipped: another smartass process is refreshing; answering from the current data"

    def skipped(*a, notes=None, **k):
        notes.append(skip)
        return []
    monkeypatch.setattr(refresh, "refresh", skipped)
    res, mode, fills, notes = tools_fin.prepare(None, CAT, WHS, None, ("ly",), store=store, retrieve=retrieve, db=db)
    assert notes == [skip]
    s = fin_pnl.summary(res, mode, fills, notes)
    assert "refresh skipped" in s["freshness"] and s["refresh_notes"] == [skip]


@pytest.mark.anyio
async def test_the_tool_builds_through_build_job_with_the_hook(env, monkeypatch, tmp_path):
    seen = {}

    def build_job(h, tm, question, pulls, spec, target, allow, q, replace=False, before_answer=None):
        seen.update(question=question, ids=[p.id for p in pulls], hook=before_answer, spec=spec)
        return {"workbook": "C:/x/a.xlsx", "question": "Q1", "checks": {}, "timings": {}, "warnings": []}

    async def call_user(fn, what):
        return fn(None)
    monkeypatch.setattr(tools_ws, "_build_job", build_job)
    monkeypatch.setattr(tools_fin, "_call_user", call_user)
    monkeypatch.setattr(tools_fin, "Warehouse", MemoryWarehouse)
    retrieve, db = env
    real = atlas.connect                       # svhttp still opens its own test atlas by path
    monkeypatch.setattr(extract, "route", lambda http=None, host=None: retrieve)
    monkeypatch.setattr(atlas, "connect", lambda path=None: real(path) if path else real(tmp_path / "atlas.db"))
    out = await tools_fin.fin_pnl("Wholesale")
    assert seen["ids"] == ["pnl", "fte", "fpa", "fpp", "units"] and seen["hook"] is not None
    assert out["area"] == "Wholesale" and out["workbook"] == "C:/x/a.xlsx" and "freshness" in out
    assert seen["spec"]["key_points"][0].startswith("EBITDA YTD Sep 26:")
    assert seen["spec"]["notes"][0].startswith("Data: warehouse;")
    assert out["headline_kpi"] == "personnel_per_fte" and out["missing"] == []


@pytest.mark.anyio
async def test_an_abolished_area_gets_no_workbook(monkeypatch):
    called = []
    monkeypatch.setattr(tools_fin, "_call_user", lambda fn, what: called.append(what))
    with pytest.raises(ToolError, match="Successors: Security Solutions & Mobile"):
        await tools_fin.fin_pnl("Consumer")
    assert called == []


def test_a_failed_job_is_a_visible_warning():
    from smartview_mcp.extract import FillResult
    from smartview_mcp.warehouse import Fill
    ok = FillResult(Fill("top_up", "x", 0.0, 0.1, 1, "ok", {}))
    bad = FillResult(Fill("new_month", "x", 0.0, 0.1, 0, "error", {"errors": ["FinMax: provider down"]}))
    assert tools_fin.stale_note([ok]) is None
    assert "(new_month: FinMax: provider down)" in tools_fin.stale_note([ok, bad])
