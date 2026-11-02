"""tools_fin: area resolution (incl. the abolished areas), prepare() with the warehouse and with the live-pull
fallback, the empty-forecast fallback, labels and the sub-unit cap, refresh notes, and the fin_pnl tool handing its
request to _build_job with the reconciliation hook. Every figure is invented."""
import sqlite3

import pytest
from fin_util import CAT, FC, WHS, http, make_value
from mcp.server.mcpserver.exceptions import ToolError

from smartview_mcp import atlas, extract, fin_pnl, refresh, templates, tools_fin
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


def test_a_failed_job_is_a_visible_warning():
    from smartview_mcp.extract import FillResult
    from smartview_mcp.warehouse import Fill
    ok = FillResult(Fill("top_up", "x", 0.0, 0.1, 1, "ok", {}))
    bad = FillResult(Fill("new_month", "x", 0.0, 0.1, 0, "error", {"errors": ["FinMax: provider down"]}))
    assert tools_fin.stale_note([ok]) is None
    assert "(new_month: FinMax: provider down)" in tools_fin.stale_note([ok, bad])


@pytest.mark.parametrize("period, msg", [
    ("Dec 26", "Dec 26 is not available yet: the last closed month is Sep 26; the open month is Oct 26."),
    ("Dec 24", "Dec 24 is older than the warehouse holds: it keeps the 36 months Oct 23..Sep 26, and Dec 24 "
               "against last year needs Dec 22. Ask for Jan 25 or later.")])
def test_a_period_the_warehouse_cannot_answer_is_refused(env, period, msg):
    from smartview_mcp.host import InputError
    retrieve, db = env
    with pytest.raises(InputError) as e:
        tools_fin.prepare(None, CAT, WHS, period, ("ly",), store=MemoryWarehouse(), retrieve=retrieve, db=db)
    assert str(e.value) == msg
    tools_fin.check_period("Oct 26", "Sep 26")                                            # the open month is fine
    tools_fin.check_period("Jan 25", "Sep 26")


def test_an_old_schema_warehouse_falls_back_to_live_pulls(env, tmp_path):
    import duckdb

    from smartview_mcp.warehouse import Warehouse
    p = tmp_path / "old.duckdb"
    db0 = duckdb.connect(str(p))
    db0.execute("CREATE TABLE fact (cube VARCHAR, scenario VARCHAR, version VARCHAR, entity VARCHAR, "
                "account VARCHAR, period VARCHAR, balance VARCHAR, value DOUBLE, source VARCHAR, fetched_at DOUBLE, "
                "atlas_build VARCHAR)")
    db0.execute("INSERT INTO fact VALUES ('FinMax','Actual','','un.SY','ac.C30','Sep 26','Periodic',1,'http',1,'')")
    db0.close()
    retrieve, db = env
    res, mode, fills, _ = tools_fin.prepare(None, CAT, WHS, None, ("ly",), store=Warehouse(p), retrieve=retrieve,
                                            db=db)
    assert mode.startswith("live pulls (the warehouse was unavailable: warehouse schema is older")
    assert fills[0].fill.job == "live" and res.fm["ebitda"]["ytd"].current.value is not None


def test_a_first_fill_running_elsewhere_means_live_pulls(env, tmp_path, monkeypatch):
    from smartview_mcp import config
    from smartview_mcp.drill_browser import ProfileLock
    from smartview_mcp.warehouse import Warehouse
    monkeypatch.setattr(config, "WAREHOUSE_LOCK_WAIT_S", 0.05)
    retrieve, db = env
    store = Warehouse(tmp_path / "w.duckdb")
    other = ProfileLock(str(store.path) + ".refresh")
    assert other.acquire()
    try:
        _, mode, _, _ = tools_fin.prepare(None, CAT, WHS, None, ("ly",), store=store, retrieve=retrieve, db=db)
    finally:
        other.release()
    assert "first warehouse fill is still running" in mode and mode.startswith("live pulls")


# ------------------------------------------------------------------ sub-units (any FinMax entity below an area)
UNIT_MEMBERS = MEMBERS + [("FinMax", "All Entities", "un.SY.U1A", "U1A Unit one A (un.SY)", "un.SY.U1"),
                          ("FinMax", "All Entities", "un.SY.U1B", "U1B Unit one B (un.SY)", "un.SY.U1"),
                          ("FinMax", "All Entities", "un.SY.2CON", "2CON Consumer (un.SY)", "un.SY"),
                          ("FinMax", "All Entities", "un.SY.C1", "C1 Old unit (un.SY)", "un.SY.2CON"),
                          ("FinPlan", "Entity", "U1 Unit one", "U1 Unit one", "2WHS Wholesale"),
                          ("FinPlan", "Entity", "U10 Unit ten", "U10 Unit ten", "2WHS Wholesale")]


@pytest.fixture
def uenv(tmp_path, monkeypatch):
    return _env(tmp_path, monkeypatch, members=UNIT_MEMBERS)


def test_a_unit_below_an_area_resolves_to_a_synthetic_entity(uenv):
    _, db = uenv
    u1 = tools_fin.resolve_scope(CAT, "un.SY.U1", db)
    assert (u1.key, u1.label, u1.finmax, u1.finplan) == ("unit:un.SY.U1", "U1 Unit one", "un.SY.U1", "U1 Unit one")
    assert tools_fin.is_unit(u1) and not tools_fin.is_unit(WHS) and "Wholesale (un.SY.2WHS)" in u1.note
    u2 = tools_fin.resolve_scope(CAT, "u2", db)                     # the bare code, any case; no FinPlan member
    assert (u2.finmax, u2.finplan) == ("un.SY.U2", None)
    for text in ("un.SY.2WHS", "2whs", "Wholesale"):                # an area's own member is the area
        assert tools_fin.resolve_scope(CAT, text, db) == WHS


@pytest.mark.parametrize("text, msg", [("un.SY.C1", "below Consumer Norway .un.SY.2CON., abolished on 2026-06-01"),
                                       ("XYZ_9", "sv_find_member"), ("Narnia", "names no business area")])
def test_a_unit_that_cannot_be_answered_is_refused_readably(uenv, text, msg):
    _, db = uenv
    with pytest.raises(ToolError, match=msg):
        tools_fin.resolve_scope(CAT, text, db)


def test_a_unit_is_read_live_and_never_stored_in_the_warehouse(uenv):
    retrieve, db = uenv
    store = MemoryWarehouse()
    u1 = tools_fin.resolve_scope(CAT, "un.SY.U1", db)
    res, mode, fills, notes = tools_fin.prepare(None, CAT, u1, None, fin_pnl.COMPARISONS, store=store,
                                                retrieve=retrieve, db=db)
    assert mode == tools_fin.LIVE_UNIT == "live pulls (sub-unit: not kept in the warehouse)"
    assert not store.rows and not store.state() and "unavailable" not in mode
    assert fills[0].fill.job == "live" and all(f.ok for f in fills) and not res.missing()
    assert res.scope.units == (("un.SY.U1A", "U1A Unit one A"), ("un.SY.U1B", "U1B Unit one B"))
    assert res.unit_bridge.closes() and res.budget and res.landing is not None and res.scope.forecast == FC
    assert any("unit below Wholesale" in n for n in res.notes)


def test_a_leaf_unit_without_finplan_has_last_year_only_and_no_bridge(uenv):
    retrieve, db = uenv
    u2 = tools_fin.resolve_scope(CAT, "un.SY.U2", db)
    res, mode, fills, notes = tools_fin.prepare(None, CAT, u2, None, fin_pnl.COMPARISONS, store=MemoryWarehouse(),
                                                retrieve=retrieve, db=db)
    assert res.scope.units == () and res.unit_bridge is None and res.scope.forecast is None
    assert not res.budget and res.landing is None and not res.missing()
    assert any("no sub-units" in n for n in res.notes)
    assert any("compared with last year only" in w for w in fin_pnl.tool_warnings(res, fills, notes))


