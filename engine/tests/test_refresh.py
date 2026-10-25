"""refresh.py: which jobs run when (pure), restatement areas from atlas snapshots, a refresh and a top-up over the
svhttp fake transport. Every figure is invented."""
import sqlite3

import pytest
from svxml_util import EchoTransport, atlas_with_dimensions, pnl_value as value

from smartview_mcp import atlas, atlas_history, catalog, extract, refresh, svhttp
from smartview_mcp.warehouse import OPEN, MemoryWarehouse

CAT = catalog.load()
FC = {26: "Forecast Q326"}.get


def kinds(jobs):
    return [j.kind for j in jobs]


def test_first_call_is_the_base_fill():
    jobs = refresh.plan_refresh(CAT, {}, "Sep 26", 7, "Forecast Q326", set(), FC)
    assert kinds(jobs) == ["base"] and jobs[0].state == {"closed_month": "Sep 26", "entities_snapshot": "7",
                                                          "forecast": "Forecast Q326"}


def test_nothing_changed_means_no_job():
    state = {"closed_month": "Sep 26", "entities_snapshot": "7", "forecast": "Forecast Q326"}
    assert refresh.plan_refresh(CAT, state, "Sep 26", 7, "Forecast Q326", set(), FC) == []


def test_a_new_closed_month_rereads_the_two_before():
    state = {"closed_month": "Sep 26", "entities_snapshot": "7", "forecast": "Forecast Q326"}
    jobs = refresh.plan_refresh(CAT, state, "Oct 26", 7, "Forecast Q326", set(), FC)
    assert kinds(jobs) == ["new_month"] and jobs[0].state == {"closed_month": "Oct 26"}
    fm = [g for g in jobs[0].grids if g.conn == "FinMax"]
    assert fm[0].cols[0].members == ["Aug 26", "Sep 26", "Oct 26"] and fm[-1].cols[0].members == ["Oct 26"]
    two = refresh.plan_refresh(CAT, {**state, "closed_month": "Aug 26"}, "Oct 26", 7, "Forecast Q326", set(), FC)
    assert two[0].grids[0].cols[0].members == ["Aug 26", "Sep 26", "Oct 26"]


def test_restatement_and_forecast_change():
    state = {"closed_month": "Sep 26", "entities_snapshot": "7", "forecast": "Forecast Q226"}
    jobs = refresh.plan_refresh(CAT, state, "Sep 26", 8, "Forecast Q326", {"wholesale"}, FC)
    assert kinds(jobs) == ["restatement", "forecast"]
    ents = {m for g in jobs[0].grids for m in g.rows[0].members}
    assert ents == {"un.SY", "un.SY.2WHS", "BANORWAY", "2WHS Wholesale"}
    assert [g.pov["Scenario"] for g in jobs[1].grids] == ["Forecast Q326"]
    quiet = refresh.plan_refresh(CAT, {**state, "forecast": "Forecast Q326"}, "Sep 26", 8, "Forecast Q326", set(), FC)
    assert kinds(quiet) == ["restatement"] and quiet[0].grids == () and quiet[0].state == {"entities_snapshot": "8"}


@pytest.fixture
def snap_db(tmp_path):
    db = sqlite3.connect(tmp_path / "atlas.db")
    db.executescript(atlas.SCHEMA)

    def load(rows):
        db.execute("DELETE FROM member")
        db.executemany("INSERT INTO member VALUES ('FinMax','All Entities',?,?,?,0,0)", rows)
        db.commit()
        return atlas_history.take_snapshot(db, "FinMax", "All Entities")
    base = [("un.SY", "SY", ""), ("un.SY.2WHS", "WHS", "un.SY"), ("un.SY.3WHS_W", "W", "un.SY.2WHS"),
            ("un.SY.2BUS", "BUS", "un.SY"), ("un.SY.K1", "K1", "un.SY.3WHS_W")]
    s1 = load(base)
    s2 = load(base[:-1] + [("un.SY.K1", "K1", "un.SY.2BUS")])          # K1 moved from Wholesale to Business
    return db, s1, s2


def test_touched_areas_follow_moves_up_the_atlas(snap_db):
    db, s1, s2 = snap_db
    assert refresh.touched_areas(CAT, db, str(s1), s2) == {"wholesale", "business"}
    assert refresh.touched_areas(CAT, db, None, s2) == set() and refresh.touched_areas(CAT, db, str(s2), s2) == set()


def test_sub_units_skip_single_child_levels():
    tree = {"un.SY.2WHS": ["un.SY.3WHS_W"], "un.SY.3WHS_W": ["a", "b", "c"], "un.SY.2BUS": ["x", "y"]}
    kids = lambda m: tree.get(m, [])  # noqa: E731
    assert refresh.sub_units(kids, "un.SY.2WHS") == ["a", "b", "c"]
    assert refresh.sub_units(kids, "un.SY.2BUS", limit=1) == ["x"] and refresh.sub_units(kids, "leaf") == []


@pytest.fixture
def retrieve(tmp_path):
    http = svhttp.SvHttp(EchoTransport(value), home="https://plan1.example",
                         db_path=atlas_with_dimensions(tmp_path / "a.db"))
    return extract.route(http=http)


def test_refresh_runs_the_base_fill_once_then_nothing(retrieve, snap_db):
    db, _, _ = snap_db
    store = MemoryWarehouse()
    first = refresh.refresh(store, retrieve, CAT, "Sep 26", db, FC)
    assert [r.fill.job for r in first] == ["base"] and first[0].ok
    assert store.state()["closed_month"] == "Sep 26" and store.state()["forecast"] == "Forecast Q326"
    assert refresh.refresh(store, retrieve, CAT, "Sep 26", db, FC) == []


def test_a_failed_job_keeps_the_old_state(snap_db):
    db, _, _ = snap_db
    store = MemoryWarehouse()

    def down(*a):
        raise RuntimeError("no network")
    res = refresh.refresh(store, down, CAT, "Sep 26", db, FC)
    assert res[0].fill.result == "error" and store.state() == {}


def test_top_up_reads_forecast_and_open_months_live_and_units_as_history(retrieve):
    store = MemoryWarehouse()
    whs = next(e for e in CAT.entities if e.key == "wholesale")
    res = refresh.top_up(store, retrieve, CAT, whs, "Oct 26", "Sep 26", "Forecast Q326", ["u1", "u2"])
    assert [r.ok for r in res] == [True, True, True]               # live, units now, units last year
    fc = ("FinPlan", "Forecast Q326", "Final", "2WHS Wholesale", "EBITDA", "Dec 26", "Periodic")
    assert store.rows[fc].closed_month == OPEN
    assert store.rows[("FinMax", "Actual", "", "un.SY.2WHS", "ac.C30", "Oct 26", "Periodic")].closed_month == OPEN
    unit = store.rows[("FinMax", "Actual", "", "u1", "ac.EBITDA", "Oct 25", "Accumulated")]
    assert unit.closed_month == "Sep 26" and unit.value == 290.0 * 10
    assert store.rows[("FinMax", "Actual", "", "u1", "ac.EBITDA", "Oct 26", "Accumulated")].closed_month == OPEN
    assert refresh.top_up(store, retrieve, CAT, whs, "Sep 26", "Sep 26", None, []) == []


def test_an_unknown_or_mismatched_snapshot_means_refresh_every_area(snap_db):
    db, s1, s2 = snap_db
    everything = {e.key for e in extract.areas(CAT, "FinMax") if e.key != "company"}
    assert "wholesale" in everything
    assert refresh.touched_areas(CAT, db, "9999", s2) == everything                     # unknown old id
    assert refresh.touched_areas(CAT, db, "not-a-number", s2) == everything
    other = atlas_history.take_snapshot(db, "FinMax", "All Accounts")                    # different dimension
    assert refresh.touched_areas(CAT, db, str(s1), other or 12345) == everything


def test_refresh_after_an_unknown_snapshot_rereads_all_areas(retrieve, snap_db):
    db, _, s2 = snap_db
    store = MemoryWarehouse()
    store.set_state(closed_month="Sep 26", entities_snapshot="9999", forecast="Forecast Q326")
    res = refresh.refresh(store, retrieve, CAT, "Sep 26", db, FC)
    assert [r.fill.job for r in res] == ["restatement"] and res[0].ok
    assert store.state()["entities_snapshot"] == str(s2)
    assert store.get("FinMax", ["un.SY.2BUS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic")




def test_top_up_reads_only_open_months_and_never_relabels_closed_history(retrieve):
    store = MemoryWarehouse()
    whs = next(e for e in CAT.entities if e.key == "wholesale")
    extract.fill(store, retrieve, CAT, extract.base_grids(CAT, "Sep 26", FC), "base", "Sep 26")
    before = {k: r.closed_month for k, r in store.rows.items() if r.closed_month != OPEN}
    assert set(before.values()) == {"Sep 26"} and len(before) > 1000
    # FinPlan carries blank actuals for the months after the closed one: not closed history
    assert store.rows[("FinPlan", "Actual", "Working", "2WHS Wholesale", "EBITDA", "Oct 26", "Periodic")
                      ].closed_month == OPEN
    res = refresh.top_up(store, retrieve, CAT, whs, "Oct 26", "Sep 26", "Forecast Q326", ["u1"])
    assert all(r.ok for r in res)
    assert {k: store.rows[k].closed_month for k in before} == before            # every closed key kept its label
    assert store.rows[("FinPlan", "Actual", "Working", "2WHS Wholesale", "EBITDA", "Oct 26", "Periodic")
                      ].closed_month == OPEN
    assert len(store.get("FinPlan", ["2WHS Wholesale"], ["EBITDA"], ["Jan 26", "Mar 26"], "Actual", "Working",
                         "Periodic", closed_only=True)) == 2
    top = [g for g in res[0].rows]
    assert {k.key[5] for k in top if k.key[0] == "FinMax"} == {"Oct 26"}        # no opening level, no old months
    assert not any(k.key[4] == "ac.STA20Y" and k.key[5] == "Dec 25" for k in top)


def test_top_up_labels_open_period_units_open_and_takes_the_grand_line(retrieve):
    store = MemoryWarehouse()
    whs = next(e for e in CAT.entities if e.key == "wholesale")
    grand = next(x.account for x in CAT.pnl["FinMax"] if x.kind == "grand")
    refresh.top_up(store, retrieve, CAT, whs, "Oct 26", "Sep 26", None, ["u1"])
    assert store.rows[("FinMax", "Actual", "", "u1", grand, "Oct 26", "Accumulated")].closed_month == OPEN
    assert store.rows[("FinMax", "Actual", "", "u1", grand, "Oct 25", "Accumulated")].closed_month == "Sep 26"
    refresh.top_up(store, retrieve, CAT, whs, "Sep 26", "Sep 26", None, ["u2"])
    assert store.rows[("FinMax", "Actual", "", "u2", grand, "Sep 26", "Accumulated")].closed_month == "Sep 26"


def test_new_month_checks_the_whole_ytd_against_the_stored_months(retrieve):
    store = MemoryWarehouse()
    extract.fill(store, retrieve, CAT, extract.base_grids(CAT, "Sep 26", FC), "base", "Sep 26")
    store.set_state(closed_month="Sep 26", entities_snapshot="", forecast="Forecast Q326")
    jobs = refresh.plan_refresh(CAT, store.state(), "Oct 26", None, "Forecast Q326", set(), FC)
    res = extract.fill(store, retrieve, CAT, list(jobs[0].grids), "new_month", "Oct 26")
    assert res.fill.result == "ok", res.fill.detail
    assert res.fill.detail["checks"] > 100 and res.fill.detail["problems"] == []
    # the YTD spot was Oct 26; its ten months are in the store, only Oct was read: a corrupted stored month shows
    k = ("FinMax", "Actual", "", "un.SY.2WHS", "ac.C30", "Mar 26", "Periodic")
    store.rows[k] = type(store.rows[k])(k, 5.0, "http", 1.0, "Sep 26")
    bad = extract.fill(store, retrieve, CAT, list(jobs[0].grids), "new_month", "Oct 26")
    assert bad.fill.result == "warn" and any("Accumulated Oct 26 differs" in x for x in bad.fill.detail["problems"])


def test_a_year_roll_rereads_target_for_both_years_and_the_new_forecast():
    fc = {26: "Forecast Q426", 27: "Forecast Q127"}.get
    state = {"closed_month": "Dec 26", "entities_snapshot": "", "forecast": "Forecast Q426"}
    jobs = refresh.plan_refresh(CAT, state, "Jan 27", None, "Forecast Q127", set(), fc)
    nm = jobs[0]
    plans = {(g.pov["Scenario"], g.pov["Years"]) for g in nm.grids if g.conn == "FinPlan"}
    assert {("Target", "FY27"), ("Target", "FY28"), ("Actual", "FY26"), ("Actual", "FY27"),
            ("Forecast Q127", "FY27")} <= plans
    assert ("Forecast Q426", "FY26") not in plans and kinds(jobs) == ["new_month", "forecast"]


def test_rows_of_members_smart_view_did_not_recognise_are_not_stored(retrieve, monkeypatch):
    store = MemoryWarehouse()
    grids = extract.finmax_grids(CAT, ["Jan 26"], ["un.SY.2WHS", "un.SY.BAD"])
    real = retrieve

    def unknown(conn, pov, rows, cols):
        lay, p, secs, how = real(conn, pov, rows, cols)
        p.unknown_members.append("All Entities: un.SY.BAD")
        return lay, p, secs, how
    res = extract.fill(store, unknown, CAT, grids, "t", "Jan 26")
    assert res.fill.result == "error" and "un.SY.BAD" in res.fill.detail["errors"][0]
    assert not any(k[3] == "un.SY.BAD" for k in store.rows) and any(k[3] == "un.SY.2WHS" for k in store.rows)
    assert res.fill.detail["content_only"] is True


def _content_error(real):
    def f(conn, pov, rows, cols):
        lay, p, secs, how = real(conn, pov, rows, cols)
        p.rows[0][-1] = "#NoAccess"
        return lay, p, secs, how
    return f


def test_a_permanent_content_error_is_retried_after_15_minutes_then_daily(retrieve, snap_db):
    db, _, _ = snap_db
    store, clock, notes = MemoryWarehouse(), [1000.0], []
    bad = _content_error(retrieve)
    r1 = refresh.refresh(store, bad, CAT, "Sep 26", db, FC, now=lambda: clock[0], notes=notes)
    assert r1[0].fill.result == "error" and r1[0].fill.detail["content_only"] and store.state().get(
        "closed_month") is None
    assert float(store.state()["retry_after_base"]) == 1000.0 + refresh.RETRY_FIRST_S
    clock[0] += 60
    assert refresh.refresh(store, bad, CAT, "Sep 26", db, FC, now=lambda: clock[0], notes=notes) == []
    assert "paused" in notes[-1] and "#NoAccess" in notes[-1]
    clock[0] += refresh.RETRY_FIRST_S                                    # past the 15 minutes: tried again, fails again
    assert refresh.refresh(store, bad, CAT, "Sep 26", db, FC, now=lambda: clock[0])[0].fill.result == "error"
    assert float(store.state()["retry_after_base"]) - clock[0] == refresh.RETRY_LATER_S
    clock[0] += refresh.RETRY_LATER_S                                    # fixed upstream: succeeds, state cleared
    ok = refresh.refresh(store, retrieve, CAT, "Sep 26", db, FC, now=lambda: clock[0])
    assert ok[0].ok and store.state()["closed_month"] == "Sep 26" and float(store.state()["retry_after_base"]) == 0


def test_a_transport_error_is_retried_on_the_next_call(snap_db):
    db, _, _ = snap_db
    store = MemoryWarehouse()

    def down(*a):
        raise RuntimeError("no network")
    refresh.refresh(store, down, CAT, "Sep 26", db, FC, now=lambda: 1.0)
    assert "retry_after_base" not in store.state()
    assert len(refresh.refresh(store, down, CAT, "Sep 26", db, FC, now=lambda: 2.0)) == 1


class PathStore(MemoryWarehouse):
    def __init__(self, path):
        super().__init__()
        self.path = path


def test_a_second_process_holding_the_refresh_lock_means_no_refresh(retrieve, snap_db, tmp_path):
    from smartview_mcp.drill_browser import ProfileLock
    db, _, _ = snap_db
    store, notes = PathStore(tmp_path / "w.duckdb"), []
    other = ProfileLock(str(store.path) + ".refresh")
    assert other.acquire()
    try:
        assert refresh.refresh(store, retrieve, CAT, "Sep 26", db, FC, notes=notes) == []
        assert "another smartass process is refreshing" in notes[0] and store.state() == {}
    finally:
        other.release()
    assert refresh.refresh(store, retrieve, CAT, "Sep 26", db, FC)[0].ok                  # lock released: runs
    again = ProfileLock(str(store.path) + ".refresh")
    assert again.acquire()                                                                 # and it let go itself
    again.release()


def test_an_unavailable_warehouse_gives_an_error_result_not_a_silent_loss(retrieve, snap_db):
    from smartview_mcp.warehouse import WarehouseUnavailable
    db, _, _ = snap_db

    class Locked(MemoryWarehouse):
        def put(self, rows):
            raise WarehouseUnavailable("in use")
    store = Locked()
    res = refresh.refresh(store, retrieve, CAT, "Sep 26", db, FC)
    assert res[0].fill.result == "error" and "not stored" in res[0].fill.detail["errors"][0]
    assert store.state() == {} and res[0].rows == []


def test_a_change_only_under_the_root_touches_the_company(snap_db):
    db, s1, s2 = snap_db
    rows = [("un.SY", "SY", ""), ("un.SY.2WHS", "WHS", "un.SY"), ("un.SY.3WHS_W", "W", "un.SY.2WHS"),
            ("un.SY.2BUS", "BUS", "un.SY"), ("un.SY.K1", "K1", "un.SY.2BUS"), ("un.SY.NEW", "NEW", "un.SY")]
    db.execute("DELETE FROM member")
    db.executemany("INSERT INTO member VALUES ('FinMax','All Entities',?,?,?,0,0)", rows)
    db.commit()
    s3 = atlas_history.take_snapshot(db, "FinMax", "All Entities")
    assert refresh.touched_areas(CAT, db, str(s2), s3) == {"company"}
    jobs = refresh.plan_refresh(CAT, {"closed_month": "Sep 26", "entities_snapshot": str(s2), "forecast": ""},
                                "Sep 26", s3, None, {"company"}, FC)
    ents = {m for g in jobs[0].grids for m in g.rows[0].members}
    assert ents == {"un.SY", "BANORWAY"}


def test_the_areas_check_uses_only_the_areas_in_the_batch(retrieve):
    store = MemoryWarehouse()
    grids = extract.finmax_grids(CAT, ["Jan 26"], ["un.SY", "un.SY.2WHS"])
    res = extract.fill(store, retrieve, CAT, grids, "t", "Jan 26")
    assert res.fill.detail["info"] == [] or "less the 1 areas" in res.fill.detail["info"][0]
    wide = extract.fill(store, retrieve, CAT, extract.finmax_grids(CAT, ["Jan 26"], [
        e.finmax for e in extract.areas(CAT, "FinMax")]), "t", "Jan 26")
    assert "less the 8 areas" in wide.fill.detail["info"][0]


def test_a_big_grid_goes_straight_to_com(monkeypatch):
    from smartview_mcp import config, ops
    from smartview_mcp.grid import Axis
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
