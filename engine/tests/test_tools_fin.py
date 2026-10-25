"""tools_fin: area resolution (incl. the abolished areas), prepare() with the warehouse and with the live-pull
fallback, the empty-forecast fallback, labels and the sub-unit cap, refresh notes, and the fin_pnl tool handing its
request to _build_job with the reconciliation hook. Every figure is invented."""
import sqlite3

import pytest
from fin_util import CAT, FC, WHS, http, make_value
from mcp.server.mcpserver.exceptions import ToolError

from smartview_mcp import atlas, extract, fin_pnl, refresh, templates, tools_fin, tools_ws
from smartview_mcp.warehouse import MemoryWarehouse, WarehouseUnavailable

MEMBERS = [("FinMax", "All Entities", "un.SY.2WHS", "2WHS Wholesale (un.SY)", "un.SY"),
           ("FinMax", "All Entities", "un.SY.3WHS_W", "3WHS_W W (un.SY)", "un.SY.2WHS"),
           ("FinMax", "All Entities", "un.SY.U1", "U1 Unit one (un.SY)", "un.SY.3WHS_W"),
           ("FinMax", "All Entities", "un.SY.U2", "U2 Unit two (un.SY)", "un.SY.3WHS_W"),
           ("FinMax", "All Accounts", "ac.C30", "C30 REVENUES (ac)", "ac.C48")]


def _env(tmp_path, monkeypatch, forecasts=(FC,), value_fn=None, members=MEMBERS):
    db = sqlite3.connect(tmp_path / "atlas.db")
    db.executescript(atlas.SCHEMA)
    db.executemany("INSERT INTO member VALUES (?,?,?,?,?,1,0)", members)
    db.commit()
    monkeypatch.setattr(templates, "closed_month", lambda h, c="FinMax": "Sep 26")
    monkeypatch.setattr(templates, "forecasts", lambda yy: list(forecasts) if yy == 26 else [])
    monkeypatch.setattr(templates, "latest_forecast", lambda yy: forecasts[0] if yy == 26 and forecasts else None)
    return extract.route(http=http(tmp_path, value_fn)), db


@pytest.fixture
def env(tmp_path, monkeypatch):
    return _env(tmp_path, monkeypatch)


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
    assert res.unit_bridge.closes() and res.scope.forecast == FC and notes == [] and not res.missing()
    assert res.label("FinMax", "rev") == "C30 REVENUES (Revenue)" and res.label("FinPlan", "oth") == "Other items"


def test_an_empty_newest_forecast_falls_back_to_the_newest_with_data(tmp_path, monkeypatch):
    retrieve, db = _env(tmp_path, monkeypatch, forecasts=("Forecast Q426", FC),
                        value_fn=make_value(empty=("Forecast Q426",)))
    res, mode, fills, notes = tools_fin.prepare(None, CAT, WHS, None, fin_pnl.COMPARISONS, store=MemoryWarehouse(),
                                                retrieve=retrieve, db=db)
    note = "Forecast Q426 is empty; Forecast Q326 used"
    assert res.scope.forecast == FC and res.scope.forecast_note == note
    assert note in res.landing.method and res.landing.forecast.value is not None
    assert f"Warning: {note}." in fin_pnl.tool_warnings(res, fills, notes)
    from smartview_mcp import fin_spec
    req = fin_spec.build(res, CAT)
    assert req.spec["notes"][0] == f"Warning: {note}." and any(note in k for k in req.spec["key_points"])


def test_the_sub_unit_cap_is_noted(tmp_path, monkeypatch):
    many = MEMBERS[:2] + [("FinMax", "All Entities", f"un.SY.U{i}", f"U{i} Unit {i} (un.SY)", "un.SY.3WHS_W")
                          for i in range(1, 46)]
    _, db = _env(tmp_path, monkeypatch, members=many)
    units, note = tools_fin._units(db, WHS)
    assert len(units) == tools_fin.UNIT_CAP and "has 45 sub-units" in note and "other 5" in note


def test_prepare_answers_from_live_pulls_when_the_warehouse_is_unavailable(env):
    class Locked(MemoryWarehouse):
        def state(self):
            raise WarehouseUnavailable("another smartass process is writing")
    retrieve, db = env
    res, mode, fills, _ = tools_fin.prepare(None, CAT, WHS, "Sep 26", ("ly",), store=Locked(), retrieve=retrieve,
                                            db=db)
    assert mode.startswith("live pulls (the warehouse was unavailable: another smartass")
    assert fills[0].fill.job == "live" and res.fm["ebitda"]["ytd"].current.value is not None


def test_duckdb_that_will_not_load_is_a_warehouse_unavailable(monkeypatch):
    import builtins
    real = builtins.__import__

    def no_duckdb(name, *a, **k):
        if name == "duckdb":
            raise ImportError("DLL load failed")
        return real(name, *a, **k)
    monkeypatch.setattr(builtins, "__import__", no_duckdb)
    with pytest.raises(WarehouseUnavailable, match="DuckDB will not load"):
        tools_fin._open(None)


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
    s = fin_pnl.summary(res, mode, fills, notes, cat=CAT)
    assert "refresh skipped" in s["freshness"] and s["refresh_notes"] == [skip]
    assert skip in fin_pnl.tool_warnings(res, fills, notes)


def _tool_env(monkeypatch, tmp_path, retrieve):
    seen = {}

    def build_job(h, tm, question, pulls, spec, target, allow, q, replace=False, before_answer=None):
        seen.update(question=question, ids=[p.id for p in pulls], hook=before_answer, spec=spec)
        return {"workbook": "C:/x/a.xlsx", "question": "Q1", "checks": {}, "timings": {}, "warnings": ["sheet w"]}

    async def call_user(fn, what):
        return fn(None)
    monkeypatch.setattr(tools_ws, "_build_job", build_job)
    monkeypatch.setattr(tools_fin, "_call_user", call_user)
    monkeypatch.setattr(tools_fin, "_open", lambda store: MemoryWarehouse())
    real = atlas.connect                       # svhttp still opens its own test atlas by path
    monkeypatch.setattr(extract, "route", lambda http=None, host=None: retrieve)
    monkeypatch.setattr(atlas, "connect", lambda path=None: real(path) if path else real(tmp_path / "atlas.db"))
    return seen


@pytest.mark.anyio
async def test_the_tool_builds_through_build_job_with_the_hook(env, monkeypatch, tmp_path):
    retrieve, _ = env
    seen = _tool_env(monkeypatch, tmp_path, retrieve)
    out = await tools_fin.fin_pnl("Wholesale")
    assert seen["ids"] == ["pnl", "fte", "fpa", "fpp", "units"] and seen["hook"] is not None
    assert out["area"] == "Wholesale" and out["workbook"] == "C:/x/a.xlsx" and "freshness" in out
    assert seen["spec"]["key_points"][0].startswith("EBITDA YTD Sep 26:")
    assert seen["spec"]["notes"][0].startswith("Data: warehouse;")
    assert out["headline_kpi"] == "personnel_per_fte" and out["missing_count"] == 0 and "facts" not in out
    assert out["table"] and out["warnings"] == ["sheet w"]
    full = await tools_fin.fin_pnl("Wholesale", detail="full")
    assert full["facts"]


@pytest.mark.anyio
async def test_a_stale_forecast_and_an_open_period_are_tool_warnings(tmp_path, monkeypatch):
    retrieve, _ = _env(tmp_path, monkeypatch, forecasts=("Forecast Q126",))
    _tool_env(monkeypatch, tmp_path, retrieve)
    out = await tools_fin.fin_pnl("Wholesale", period="Oct 26")
    assert out["warnings"][0].startswith("Oct 26 is after the Current Reporting Month")
    assert out["warnings"][1].startswith("Warning: Forecast Q126 is") and out["warnings"][-1] == "sheet w"
    assert out["provisional"] and all(f["provisional"] for f in out["flags"])


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
