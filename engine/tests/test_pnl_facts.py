"""pnl_facts: area and unit resolution (incl. the abolished areas), prepare() reading live into a FactStore, the
top-up, the empty-forecast fallback, labels and the sub-unit cap. Every figure is invented."""
import importlib.util
import sqlite3

import pytest
from analytics_util import closes
from fin_util import CAT, FC, WHS, http, make_value
from mcp.server.mcpserver.exceptions import ToolError
from svxml_util import EchoTransport, atlas_with_dimensions, pnl_value as value

from smartass.knowledge import atlas
from smartview_mcp import extract, fin_pnl, pnl_facts, svhttp, templates
from smartview_mcp.factstore import OPEN, FactStore

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


@pytest.fixture
def retrieve(tmp_path):
    return extract.route(http=svhttp.SvHttp(EchoTransport(value), home="https://plan1.example",
                                            db_path=atlas_with_dimensions(tmp_path / "a.db")))


@pytest.mark.parametrize("text, msg", [("Narnia", "names no business area"),
                                       ("Wholesale vs Business", "names several areas"),
                                       ("Consumer", "abolished on 2026-06-01"),
                                       ("Consumer", "Security Solutions & Mobile .un.SY.2SSM., Entertainment"),
                                       ("Infrastructure Growth", "Successors: Business .*Technology"),
                                       ("2CON", "ask which successor")])
def test_resolve_area_explains(text, msg):
    with pytest.raises(ToolError, match=msg):
        pnl_facts.resolve_area(CAT, text)


def test_resolve_area_finds_an_alias():
    assert pnl_facts.resolve_area(CAT, "grossist").key == "wholesale"


def test_prepare_never_touches_a_warehouse(env):
    retrieve, db = env
    res, mode, fills, notes = pnl_facts.prepare(None, CAT, WHS, None, fin_pnl.COMPARISONS, retrieve=retrieve, db=db)
    assert mode == pnl_facts.LIVE and mode.startswith("live")
    assert importlib.util.find_spec("smartview_mcp.warehouse") is None
    assert [f.fill.job for f in fills] == ["live", "top_up", "top_up", "top_up"] and all(f.ok for f in fills)
    assert res.scope.units == (("un.SY.U1", "U1 Unit one"), ("un.SY.U2", "U2 Unit two"))
    assert closes(res.unit_bridge) and res.scope.forecast == FC and notes == [] and not res.missing()
    assert res.label("FinMax", "rev") == "C30 REVENUES (Revenue)" and res.label("FinPlan", "oth") == "Other items"


def test_an_empty_newest_forecast_falls_back_to_the_newest_with_data(tmp_path, monkeypatch):
    retrieve, db = _env(tmp_path, monkeypatch, forecasts=("Forecast Q426", FC),
                        value_fn=make_value(empty=("Forecast Q426",)))
    res, mode, fills, notes = pnl_facts.prepare(None, CAT, WHS, None, fin_pnl.COMPARISONS, retrieve=retrieve, db=db)
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
    units, note = pnl_facts._units(db, WHS)
    assert len(units) == pnl_facts.UNIT_CAP and "has 45 sub-units" in note and "other 5" in note


def test_sub_units_skip_single_child_levels():
    tree = {"un.SY.2WHS": ["un.SY.3WHS_W"], "un.SY.3WHS_W": ["a", "b", "c"], "un.SY.2BUS": ["x", "y"]}
    kids = lambda m: tree.get(m, [])  # noqa: E731
    assert pnl_facts.sub_units(kids, "un.SY.2WHS") == ["a", "b", "c"]
    assert pnl_facts.sub_units(kids, "un.SY.2BUS", limit=1) == ["x"] and pnl_facts.sub_units(kids, "leaf") == []


def test_a_failed_read_is_a_visible_warning():
    from smartview_mcp.extract import FillResult
    from smartview_mcp.factstore import Fill
    ok = FillResult(Fill("top_up", "x", 0.0, 0.1, 1, "ok", {}))
    bad = FillResult(Fill("live", "x", 0.0, 0.1, 0, "error", {"errors": ["FinMax: provider down"]}))
    assert pnl_facts.stale_note([ok]) is None
    assert "(live: FinMax: provider down)" in pnl_facts.stale_note([ok, bad])


@pytest.mark.parametrize("period, msg", [
    ("Dec 26", "Dec 26 is not available yet: the last closed month is Sep 26; the open month is Oct 26."),
    ("Dec 24", "Dec 24 is outside the live read window: it covers the 36 months Oct 23..Sep 26, and Dec 24 "
               "against last year needs Dec 22. Ask for Jan 25 or later.")])
def test_a_period_outside_the_live_read_window_is_refused(env, period, msg):
    from smartview_mcp.host import InputError
    retrieve, db = env
    with pytest.raises(InputError) as e:
        pnl_facts.prepare(None, CAT, WHS, period, ("ly",), retrieve=retrieve, db=db)
    assert str(e.value) == msg
    pnl_facts.check_period("Oct 26", "Sep 26")                                            # the open month is fine
    pnl_facts.check_period("Jan 25", "Sep 26")


# ------------------------------------------------------------------ top-up
def test_top_up_reads_forecast_and_open_months_live_and_units_as_history(retrieve):
    store = FactStore()
    res = pnl_facts.top_up(store, retrieve, CAT, WHS, "Oct 26", "Sep 26", "Forecast Q326", ["u1", "u2"])
    assert [r.ok for r in res] == [True, True, True]               # live, units now, units last year
    fc = ("FinPlan", "Forecast Q326", "Final", "2WHS Wholesale", "EBITDA", "Dec 26", "Periodic")
    assert store.rows[fc].closed_month == OPEN
    assert store.rows[("FinMax", "Actual", "", "un.SY.2WHS", "ac.C30", "Oct 26", "Periodic")].closed_month == OPEN
    unit = store.rows[("FinMax", "Actual", "", "u1", "ac.EBITDA", "Oct 25", "Accumulated")]
    assert unit.closed_month == "Sep 26" and unit.value == 290.0 * 10
    assert store.rows[("FinMax", "Actual", "", "u1", "ac.EBITDA", "Oct 26", "Accumulated")].closed_month == OPEN
    assert pnl_facts.top_up(store, retrieve, CAT, WHS, "Sep 26", "Sep 26", None, []) == []


def test_top_up_reads_only_open_months_and_never_relabels_closed_history(retrieve):
    from smartview_mcp.factstore import protects
    store = FactStore()
    extract.fill(store, retrieve, CAT, extract.base_grids(CAT, "Sep 26", {26: FC}.get), "live", "Sep 26")
    before = {k: r.closed_month for k, r in store.rows.items() if protects(k[5], r.closed_month)}
    assert set(before.values()) == {"Sep 26"} and len(before) > 1000
    # FinPlan carries blank actuals for the months after the closed one: not closed history
    assert store.rows[("FinPlan", "Actual", "Working", "2WHS Wholesale", "EBITDA", "Oct 26", "Periodic")
                      ].closed_month == OPEN
    res = pnl_facts.top_up(store, retrieve, CAT, WHS, "Oct 26", "Sep 26", "Forecast Q326", ["u1"])
    assert all(r.ok for r in res)
    assert {k: store.rows[k].closed_month for k in before} == before            # every closed key kept its label
    assert len(store.get("FinPlan", ["2WHS Wholesale"], ["EBITDA"], ["Jan 26", "Mar 26"], "Actual", "Working",
                         "Periodic", closed_only=True)) == 2
    top = [g for g in res[0].rows]
    assert {k.key[5] for k in top if k.key[0] == "FinMax"} == {"Oct 26"}        # no opening level, no old months
    assert not any(k.key[4] == "ac.STA20Y" and k.key[5] == "Dec 25" for k in top)


def test_top_up_labels_open_period_units_open_and_takes_the_grand_line(retrieve):
    store = FactStore()
    grand = next(x.account for x in CAT.pnl["FinMax"] if x.kind == "grand")
    pnl_facts.top_up(store, retrieve, CAT, WHS, "Oct 26", "Sep 26", None, ["u1"])
    assert store.rows[("FinMax", "Actual", "", "u1", grand, "Oct 26", "Accumulated")].closed_month == OPEN
    assert store.rows[("FinMax", "Actual", "", "u1", grand, "Oct 25", "Accumulated")].closed_month == "Sep 26"
    pnl_facts.top_up(store, retrieve, CAT, WHS, "Sep 26", "Sep 26", None, ["u2"])
    assert store.rows[("FinMax", "Actual", "", "u2", grand, "Sep 26", "Accumulated")].closed_month == "Sep 26"


def test_a_forecast_top_up_replaces_the_base_forecast_after_the_closed_month(tmp_path):
    t = EchoTransport(value)
    retrieve = extract.route(http=svhttp.SvHttp(t, home="https://plan1.example",
                                                db_path=atlas_with_dimensions(tmp_path / "a.db")))
    store = FactStore()
    extract.fill(store, retrieve, CAT, extract.base_grids(CAT, "Sep 26", {26: FC}.get), "live", "Sep 26")
    key = ("FinPlan", FC, "Final", "2WHS Wholesale", "EBITDA", "Dec 26", "Periodic")
    assert store.get("FinPlan", ["2WHS Wholesale"], ["EBITDA"], ["Dec 26"], FC, "Final", "Periodic")[key] == 290.0
    t.value = lambda rows, cols, pov: 777.0 if FC in pov else value(rows, cols, pov)
    res = pnl_facts.top_up(store, retrieve, CAT, WHS, "Oct 26", "Sep 26", FC, [])
    assert res[0].fill.detail["refused"] > 0                          # the forecast months up to Sep 26 are history
    got = store.get("FinPlan", ["2WHS Wholesale"], ["EBITDA"], ["Mar 26", "Dec 26"], FC, "Final", "Periodic")
    assert got[key] == 777.0
    assert got[key[:5] + ("Mar 26", "Periodic")] == 290.0             # a closed period keeps its value


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
    u1 = pnl_facts.resolve_scope(CAT, "un.SY.U1", db)
    assert (u1.key, u1.label, u1.finmax, u1.finplan) == ("unit:un.SY.U1", "U1 Unit one", "un.SY.U1", "U1 Unit one")
    assert pnl_facts.is_unit(u1) and not pnl_facts.is_unit(WHS) and "Wholesale (un.SY.2WHS)" in u1.note
    u2 = pnl_facts.resolve_scope(CAT, "u2", db)                     # the bare code, any case; no FinPlan member
    assert (u2.finmax, u2.finplan) == ("un.SY.U2", None)
    for text in ("un.SY.2WHS", "2whs", "Wholesale"):                # an area's own member is the area
        assert pnl_facts.resolve_scope(CAT, text, db) == WHS


@pytest.mark.parametrize("text, msg", [("un.SY.C1", "below Consumer Norway .un.SY.2CON., abolished on 2026-06-01"),
                                       ("XYZ_9", "sv_find_member"), ("Narnia", "names no business area")])
def test_a_unit_that_cannot_be_answered_is_refused_readably(uenv, text, msg):
    _, db = uenv
    with pytest.raises(ToolError, match=msg):
        pnl_facts.resolve_scope(CAT, text, db)


def test_a_unit_is_read_live(uenv):
    retrieve, db = uenv
    u1 = pnl_facts.resolve_scope(CAT, "un.SY.U1", db)
    res, mode, fills, notes = pnl_facts.prepare(None, CAT, u1, None, fin_pnl.COMPARISONS, retrieve=retrieve, db=db)
    assert mode == pnl_facts.LIVE_UNIT == "live read (sub-unit)"
    assert fills[0].fill.job == "live" and all(f.ok for f in fills) and not res.missing()
    assert res.scope.units == (("un.SY.U1A", "U1A Unit one A"), ("un.SY.U1B", "U1B Unit one B"))
    assert closes(res.unit_bridge) and res.budget and res.landing is not None and res.scope.forecast == FC
    assert any("unit below Wholesale" in n for n in res.notes)


def test_a_leaf_unit_without_finplan_has_last_year_only_and_no_bridge(uenv):
    retrieve, db = uenv
    u2 = pnl_facts.resolve_scope(CAT, "un.SY.U2", db)
    res, mode, fills, notes = pnl_facts.prepare(None, CAT, u2, None, fin_pnl.COMPARISONS, retrieve=retrieve, db=db)
    assert res.scope.units == () and res.unit_bridge is None and res.scope.forecast is None
    assert not res.budget and res.landing is None and not res.missing()
    assert any("no sub-units" in n for n in res.notes)
    assert any("compared with last year only" in w for w in fin_pnl.tool_warnings(res, fills, notes))
