"""extract.py: grids from the catalogue, fact keys, a fill over the svhttp fake transport, the COM fallback.
Every figure is invented by the value function below."""
import pytest
from svxml_util import EchoTransport, atlas_with_dimensions, pnl_value as value

from smartview_mcp import catalog, extract
from smartass.epm import ops, svhttp
from smartass.epm.grid import Axis
from smartview_mcp.host import InputError
from smartview_mcp.factstore import FactStore

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
    store = FactStore()
    grids = extract.base_grids(CAT, "Sep 26", lambda yy: "Forecast Q326" if yy == 26 else None)
    res = extract.fill(store, extract.route(http=http), CAT, grids, "base", "Sep 26", now=lambda: 100.0)
    assert res.ok and res.fill.result == "ok", res.fill.detail
    assert res.fill.cells == sum(g.cells for g in grids) and res.fill.detail["routes"] == ["http"]
    assert res.fill.detail["checks"] > 1000 and res.fill.detail["problems"] == []
    assert "less the 8 areas" in res.fill.detail["info"][0]
    got = store.get("FinPlan", ["2WHS Wholesale"], ["EBITDA"], ["Sep 26"], "Target", "Final", "Periodic")
    assert got == {("FinPlan", "Target", "Final", "2WHS Wholesale", "EBITDA", "Sep 26", "Periodic"): 290.0}
    assert store.fills[-1].job == "base"


def test_a_broken_identity_is_a_warning_and_a_failed_grid_an_error(http):
    def bad(rows, cols, pov):
        return 999.0 if rows[1] == "ac.C48" and cols == ("Mar 26",) else value(rows, cols, pov)
    http.t.value = bad
    store = FactStore()
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
    store = FactStore()
    grids = extract.finmax_grids(CAT, ["Jan 26", "Feb 26"], ["un.SY.2WHS"])
    res = extract.fill(store, extract.route(http=http), CAT, grids, "t", "Feb 26")
    assert seen and set(seen) == {(False, False)}
    keys = [k for k in store.rows if k[4] == "ac.STA20Y"]
    assert len(keys) == 3 and all(store.rows[k].value is None for k in keys)       # Dec 25, Jan 26, Feb 26
    assert res.fill.cells == sum(g.cells for g in grids)


@pytest.fixture
def retrieve(http):
    return extract.route(http=http)


def test_rows_of_members_smart_view_did_not_recognise_are_not_stored(retrieve, tmp_path):
    from smartass.knowledge import atlas
    store = FactStore()
    grids = extract.finmax_grids(CAT, ["Jan 26"], ["un.SY.2WHS", "un.SY.BAD"])
    real = retrieve

    def unknown(conn, pov, rows, cols):
        lay, p, secs, how = real(conn, pov, rows, cols)
        p.unknown_members.append("All Entities: un.SY.BAD")
        return lay, p, secs, how
    adb = atlas.connect(tmp_path / "x.db")
    atlas.store(adb, "FinMax", "All Entities", [{"name": "All Entities", "label": "", "parent": None, "depth": 0,
                                                 "order": 0},
                                                {"name": "un.SY.2WHS", "label": "", "parent": None, "depth": 1,
                                                 "order": 1}], 0.1)
    res = extract.fill(store, unknown, CAT, grids, "t", "Jan 26", atlas_db=adb)
    assert res.fill.result == "warn" and "un.SY.BAD" in res.fill.detail["warnings"][0]    # the rest was usable
    assert not any(k[3] == "un.SY.BAD" for k in store.rows) and any(k[3] == "un.SY.2WHS" for k in store.rows)
    assert res.fill.detail["errors"] == []


def test_the_areas_check_uses_only_the_areas_in_the_batch(retrieve):
    store = FactStore()
    grids = extract.finmax_grids(CAT, ["Jan 26"], ["un.SY", "un.SY.2WHS"])
    res = extract.fill(store, retrieve, CAT, grids, "t", "Jan 26")
    assert res.fill.detail["info"] == [] or "less the 1 areas" in res.fill.detail["info"][0]
    wide = extract.fill(store, retrieve, CAT, extract.finmax_grids(CAT, ["Jan 26"], [
        e.finmax for e in extract.areas(CAT, "FinMax")]), "t", "Jan 26")
    assert "less the 8 areas" in wide.fill.detail["info"][0]


def test_a_big_grid_goes_straight_to_com(monkeypatch):
    from smartass.core import config
    calls = []

    class Http:
        def retrieve(self, *a):
            calls.append("http")
            return "lay", "parsed", 1.0

    class Host:
        def on_host(self, fn, timeout=None):
            calls.append(("com", timeout))
            return fn()
    monkeypatch.setattr(ops, "query", lambda h, c, pov, rows, cols, sm, sz: ("lay", "parsed", 1.0))
    monkeypatch.setattr(config, "SVHTTP_TIMEOUT_S", 10.0)
    r = extract.route(http=Http(), host=lambda: Host())
    small = r("FinMax", {}, [Axis("A", ["x"] * 10)], [Axis("T", ["m"] * 10)])
    big = r("FinMax", {}, [Axis("A", ["x"] * 100)], [Axis("T", ["m"] * 100)])          # 10 000 cells / 200 > 10 s
    assert small[3] == "http" and big[3] == "com" and calls[0] == "http" and calls[1][0] == "com"
    assert calls[1][1] >= 50


def test_cells_count_what_was_written_and_the_refused_are_in_the_detail(retrieve):
    store = FactStore()
    grids = extract.finmax_grids(CAT, ["Jan 26"], ["un.SY.2WHS"])
    first = extract.fill(store, retrieve, CAT, grids, "t", "Jan 26")
    assert first.fill.cells == sum(g.cells for g in grids) and first.fill.detail["refused"] == 0
    live = extract.fill(store, retrieve, CAT, grids, "top_up", "Jan 26", closed_label="open")   # closed period: refused
    assert live.fill.cells == 0 and live.fill.detail["refused"] == first.fill.cells


def _finplan_atlas(tmp_path, entity="2WHS Wholesale"):
    from smartass.knowledge import atlas
    db = atlas.connect(tmp_path / "fp.db")
    months = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    for dim, names in (("Entity", [entity]), ("Account", ["EBITDA", "Total Other Items"]), ("Period", list(months))):
        atlas.store(db, "FinPlan", dim, [{"name": dim, "label": "", "parent": None, "depth": 0, "order": 0},
                                         *[{"name": n, "label": n, "parent": dim, "depth": 1, "order": i + 1}
                                           for i, n in enumerate(names)]], 0.1)
    return db


def test_members_the_atlas_knows_are_kept_when_the_parser_misflags_them(retrieve, tmp_path):
    """FinPlan names equal their labels, so grid.parse flags a member with an empty row/column as unknown. The atlas
    decides: known members keep their cells (None stays None); a member the atlas lacks is dropped."""
    from smartass.knowledge import atlas
    store = FactStore()
    grids = extract.finplan_grids(CAT, [("Target", "Final", 27)], ["2WHS Wholesale", "2NOPE Missing"])
    real = retrieve

    def misflag(conn, pov, rows, cols):
        lay, p, secs, how = real(conn, pov, rows, cols)
        for row in p.rows:                                             # an empty budget year: nothing but blanks
            row[lay.nr:] = [None] * (len(row) - lay.nr)
        p.unknown_members.extend(["Entity: 2WHS Wholesale", "Account: Total Other Items", "Period: Oct",
                                  "Entity: 2NOPE Missing"])
        return lay, p, secs, how
    res = extract.fill(store, misflag, CAT, grids, "t", "Dec 26", atlas_db=_finplan_atlas(tmp_path))
    assert res.fill.result == "warn" and res.fill.detail["errors"] == []
    assert res.fill.detail["warnings"] == ["FinPlan Target/Final FY27: unknown members Entity: 2NOPE Missing"]
    kept = {k: v for k, v in store.rows.items() if k[3] == "2WHS Wholesale"}
    assert len(kept) == len(CAT.pnl["FinPlan"]) * 12 and all(r.value is None for r in kept.values())
    assert not any(k[3] == "2NOPE Missing" for k in store.rows)
    # no atlas dimension to judge with: nothing is dropped, nothing is reported
    again = extract.fill(FactStore(), misflag, CAT, grids, "t", "Dec 26", atlas_db=atlas.connect(tmp_path / "e.db"))
    assert again.fill.result == "ok" and again.fill.cells == sum(g.cells for g in grids)


def test_fill_retrieves_at_once_and_stores_in_grid_order(http):
    import random
    import threading
    import time as _t
    grids = extract.finmax_grids(CAT, ["Jan 26", "Feb 26", "Mar 26"], ["un.SY.2WHS", "un.SY.2BUS"])
    assert len(grids) >= 2
    answer = extract.route(http=http)
    seq = extract.fill(FactStore(), answer, CAT, grids, "t", "Mar 26")
    live, peak, lock = [0], [0], threading.Lock()
    two = threading.Barrier(2, timeout=5)                # the first two reads must be in flight together

    def slow(conn, pov, rows, cols):
        with lock:
            live[0] += 1
            peak[0] = max(peak[0], live[0])
        try:
            two.wait()
        except threading.BrokenBarrierError:             # later calls, or no overlap at all (peak tells)
            pass
        _t.sleep(random.uniform(0.0, 0.02))
        with lock:
            live[0] -= 1
        return answer(conn, pov, rows, cols)
    par = extract.fill(FactStore(), slow, CAT, grids, "t", "Mar 26")
    assert [r.key for r in par.rows] == [r.key for r in seq.rows]          # same rows, same order
    assert par.fill.result == seq.fill.result and par.fill.detail["checks"] == seq.fill.detail["checks"]
    assert peak[0] > 1


def test_fill_on_the_host_thread_stays_sequential(http, monkeypatch):
    import threading
    from smartview_mcp import hostproc
    monkeypatch.setattr(hostproc, "on_host_thread", lambda: True)
    grids = extract.finmax_grids(CAT, ["Jan 26", "Feb 26"], ["un.SY.2WHS"])
    answer = extract.route(http=http)
    seen = []

    def rec(conn, pov, rows, cols):
        seen.append(threading.current_thread().name)
        return answer(conn, pov, rows, cols)
    extract.fill(FactStore(), rec, CAT, grids, "t", "Feb 26")
    assert len(seen) == len(grids) and len(set(seen)) == 1       # no pool: a COM fallback could not deadlock



def test_fill_reads_http_at_once_and_falls_back_to_excel_one_grid_at_a_time(monkeypatch):
    import threading
    from smartview_mcp import extract as ex
    grids = ex.finmax_grids(CAT, ["Jan 26", "Feb 26", "Mar 26"], ["un.SY.2WHS", "un.SY.2BUS"])
    live, peak, lock = [0], [0], threading.Lock()

    class Down:
        def retrieve(self, *a):
            raise RuntimeError("HTTP down")

    class Host:
        def on_host(self, fn, timeout=None):
            with lock:
                live[0] += 1
                peak[0] = max(peak[0], live[0])
            try:
                raise RuntimeError("no Excel in tests")
            finally:
                with lock:
                    live[0] -= 1
    res = ex.fill(FactStore(), ex.route(http=Down(), host=lambda: Host()), CAT, grids, "t", "Mar 26")
    assert res.fill.result == "error" and len(res.fill.detail["errors"]) == len(grids)
    assert peak[0] == 1                                    # the Excel fallback never ran two at once
