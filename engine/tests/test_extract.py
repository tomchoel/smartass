"""extract.py: grids from the catalogue, warehouse keys, a fill over the svhttp fake transport, the COM fallback.
Every figure is invented by the value function below."""
import pytest
from svxml_util import EchoTransport, atlas_with_dimensions, pnl_value as value

from smartview_mcp import catalog, extract, ops, svhttp
from smartview_mcp.grid import Axis
from smartview_mcp.host import InputError
from smartview_mcp.warehouse import MemoryWarehouse

CAT = catalog.load()


@pytest.fixture
def http(tmp_path):
    return svhttp.SvHttp(EchoTransport(value), home="https://plan1.example",
                         db_path=atlas_with_dimensions(tmp_path / "a.db"))


def test_areas_skip_abolished_and_put_the_company_first():
    fm = [e.key for e in extract.areas(CAT, "FinMax")]
    assert fm[0] == "company" and "consumer" not in fm and "ig" not in fm and "wholesale" in fm
    assert [e.finplan for e in extract.areas(CAT, "FinPlan")][0] == "BANORWAY"


def test_base_grids_cover_36_months_levels_spot_and_finplan():
    grids = extract.base_grids(CAT, "Sep 26", lambda yy: "Forecast Q326" if yy == 26 else None)
    fm = [g for g in grids if g.conn == "FinMax"]
    assert len(fm) == 3 + 1 + 1                                     # 36 months in 12s, FTE levels, YTD spot
    assert fm[0].cols[0].members[0] == "Oct 23" and fm[2].cols[0].members[-1] == "Sep 26"
    levels = fm[3].cols[0].members
    assert levels[0] == "Dec 22" and fm[3].pov["All Balance Types"] == "Accumulated"
    assert fm[4].cols[0].members == ["Sep 26"] and "ac.STA20Y" not in fm[4].rows[1].members
    fp = [(g.pov["Scenario"], g.pov["Version"], g.pov["Years"]) for g in grids if g.conn == "FinPlan"]
    assert fp == [("Actual", "Working", "FY25"), ("Actual", "Working", "FY26"), ("Target", "Final", "FY25"),
                  ("Target", "Final", "FY26"), ("Target", "Final", "FY27"), ("Forecast Q326", "Final", "FY26")]
    only = extract.base_grids(CAT, "Sep 26", lambda yy: None, keys={"wholesale"})
    assert {tuple(g.rows[0].members) for g in only} == {("un.SY.2WHS",), ("2WHS Wholesale",)}


def test_wh_key_for_both_cubes():
    assert extract.wh_key("FinMax", {"All Scenarios": "Actual", "All Entities": "un.SY.2WHS", "All Accounts": "ac.C30",
                                     "All Time Periods": "Sep 26", "All Balance Types": "Periodic",
                                     "All Currencies": "NOK"}) == (
        "FinMax", "Actual", "", "un.SY.2WHS", "ac.C30", "Sep 26", "Periodic")
    assert extract.wh_key("FinPlan", {"Scenario": "Target", "Version": "Final", "Entity": "2WHS Wholesale",
                                      "Account": "EBITDA", "Period": "Sep", "Years": "FY26",
                                      "All Balance Types": "Periodic"}) == (
        "FinPlan", "Target", "Final", "2WHS Wholesale", "EBITDA", "Sep 26", "Periodic")


def test_a_fill_over_http_stores_every_cell_and_reconciles(http):
    store = MemoryWarehouse()
    grids = extract.base_grids(CAT, "Sep 26", lambda yy: "Forecast Q326" if yy == 26 else None)
    res = extract.fill(store, extract.route(http=http), CAT, grids, "base", "Sep 26", now=lambda: 100.0)
    assert res.ok and res.fill.result == "ok", res.fill.detail
    assert res.fill.cells == sum(g.cells for g in grids) and res.fill.detail["routes"] == ["http"]
    assert res.fill.detail["checks"] > 1000 and res.fill.detail["problems"] == []
    assert "less the 8 areas" in res.fill.detail["info"][0]
    got = store.get("FinPlan", ["2WHS Wholesale"], ["EBITDA"], ["Sep 26"], "Target", "Final", "Periodic")
    assert got == {("FinPlan", "Target", "Final", "2WHS Wholesale", "EBITDA", "Sep 26", "Periodic"): 290.0}
    assert store.last_fills(1)[0]["job"] == "base"


def test_a_broken_identity_is_a_warning_and_a_failed_grid_an_error(http):
    def bad(rows, cols, pov):
        return 999.0 if rows[1] == "ac.C48" and cols == ("Mar 26",) else value(rows, cols, pov)
    http.t.value = bad
    store = MemoryWarehouse()
    grids = extract.finmax_grids(CAT, ["Jan 26", "Feb 26", "Mar 26"], ["un.SY.2WHS"])
    res = extract.fill(store, extract.route(http=http), CAT, grids, "new_month", "Mar 26")
    assert res.fill.result == "warn" and "un.SY.2WHS Mar 26 actual: C30 + C40S = C48" in res.fill.detail["problems"][0]

    def flaky(conn, pov, rows, cols):
        if pov["All Balance Types"] == "Accumulated":
            raise RuntimeError("provider down")
        return extract.route(http=http)(conn, pov, rows, cols)
    res = extract.fill(store, flaky, CAT, grids, "new_month", "Mar 26")
    assert res.fill.result == "error" and "provider down" in res.fill.detail["errors"][0] and res.fill.cells == 30


def test_route_falls_back_to_com_but_not_for_input_errors(monkeypatch):
    class Down:
        def retrieve(self, *a):
            raise svhttp.SvHttpError("signed out")

    class Host:
        def on_host(self, fn, timeout=None):
            return fn()
    monkeypatch.setattr(ops, "query", lambda h, c, pov, rows, cols, sm, sz: ("lay", "parsed", 1.0))
    r = extract.route(http=Down(), host=lambda: Host())
    assert r("FinMax", {}, [Axis("All Accounts", ["ac.C30"])], [Axis("All Time Periods", ["Sep 26"])])[3] == "com"
    with pytest.raises(svhttp.SvHttpError):
        extract.route(http=Down())("FinMax", {}, [], [])

    class Wrong:
        def retrieve(self, *a):
            raise InputError("dimension used twice")
    with pytest.raises(InputError):
        extract.route(http=Wrong(), host=lambda: Host())("FinMax", {}, [], [])


def test_a_fill_stores_blank_cells_as_none_and_never_skips_them(http):
    """The calc engine reads a stored None as 0 (e.g. 0 FTE) and a never-read key as missing data, so blanks must be
    stored; the fill reads with suppress_missing=False and suppress_zeros=False."""
    seen = []
    real = http.retrieve

    def spy(conn, pov, rows, cols, sm, sz):
        seen.append((sm, sz))
        return real(conn, pov, rows, cols, sm, sz)
    http.retrieve = spy
    http.t.value = lambda rows, cols, pov: None if rows[1] == "ac.STA20Y" else value(rows, cols, pov)
    store = MemoryWarehouse()
    grids = extract.finmax_grids(CAT, ["Jan 26", "Feb 26"], ["un.SY.2WHS"])
    res = extract.fill(store, extract.route(http=http), CAT, grids, "t", "Feb 26")
    assert seen and set(seen) == {(False, False)}
    keys = [k for k in store.rows if k[4] == "ac.STA20Y"]
    assert len(keys) == 3 and all(store.rows[k].value is None for k in keys)       # Dec 25, Jan 26, Feb 26
    assert res.fill.cells == sum(g.cells for g in grids)
