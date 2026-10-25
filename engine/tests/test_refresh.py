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
    assert [r.ok for r in res] == [True, True]
    fc = ("FinPlan", "Forecast Q326", "Final", "2WHS Wholesale", "EBITDA", "Dec 26", "Periodic")
    assert store.rows[fc].closed_month == OPEN
    assert store.rows[("FinMax", "Actual", "", "un.SY.2WHS", "ac.C30", "Oct 26", "Periodic")].closed_month == OPEN
    unit = store.rows[("FinMax", "Actual", "", "u1", "ac.EBITDA", "Oct 25", "Accumulated")]
    assert unit.closed_month == "Sep 26" and unit.value == 290.0 * 10
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
