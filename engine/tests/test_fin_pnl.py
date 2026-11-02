"""fin_pnl.compute on a synthetic fact store: variances, flags, bridges, KPIs and landing (invented figures)."""
import dataclasses

import pytest
from fin_util import CAT, CAT0, FC, WHS, filled, make_value

from smartview_mcp import fin_pnl
from smartview_mcp.factstore import Row

YR = 1.1
AVG_FTE = (10 + 11 * 9) / 10                     # Dec 25 = 10 (2025), Jan..Sep 26 = 11: n+1 = 10 points


def test_finmax_month_and_ytd_vs_last_year(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)
    e = res.fm["ebitda"]
    assert e["m"].current.value == pytest.approx(290 * YR) and e["m"].comparison.value == 290
    assert e["ytd"].current.value == pytest.approx(290 * YR * 9) and e["ytd"].abs.value == pytest.approx(261)
    assert e["ytd"].favourable is True and len(e["ytd"].current.inputs) == 9
    fte = res.fm["fte"]
    assert fte["m"].current.value == pytest.approx(11.0) and fte["ytd"].current.value == pytest.approx(AVG_FTE)
    assert res.facts["fm.fte.lvl.Dec 25"].value == 10 and res.facts["fm.fte.lvl.Dec 24"].value == 10
    assert not res.missing() and not res.problems


def test_finplan_budget_forecast_and_landing(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)
    b = res.budget["ebitda"]["ytd"]
    assert b.current.value == pytest.approx(290 * YR * 9) and b.comparison.value == pytest.approx(290 * YR * 9 * 0.9)
    assert res.forecast["ebitda"]["ytd"].comparison.value == pytest.approx(290 * YR * 9 * 1.05)
    ld = res.landing
    assert ld.forecast.value == pytest.approx(290 * YR * 9 + 290 * YR * 1.05 * 3)
    assert res.facts["fp.ebitda.forecast.remaining"].value == pytest.approx(290 * YR * 1.05 * 3)
    # the seasonal share is last year's EBITDA BEFORE OTHER ITEMS: 300 x 9 / (300 x 12)
    assert ld.seasonal and ld.run_rate.value == pytest.approx(290 * YR * 9 / 0.75)
    assert set(ld.run_rate.inputs) >= set(res.facts["fp.boi.actual.ly_fy"].inputs)
    assert ld.target.value == pytest.approx(290 * YR * 0.9 * 12) and not ld.stale
    assert "Forecast Q326" in ld.method


def test_an_old_forecast_is_flagged_stale_with_the_run_rate_alongside(tmp_path):
    store, _, scope = filled(tmp_path, forecast="Forecast Q126")
    res = fin_pnl.compute(CAT, scope, store)
    assert res.landing.stale and "Forecast Q126" in res.landing.warning
    assert res.landing.forecast.value is not None and res.landing.run_rate.value is not None


def test_flags_judge_leaf_and_always_lines_with_month_swings(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)
    assert set(res.flags["ly"]) == {"rev", "cogs", "pers", "oth", "c73", "ebitda", "da", "fte"}   # no subtotals
    assert not any(f.material for f in res.flags["ly"].values())          # small invented figures
    assert {k for c, k, f in res.ranking() if c == "ly"} == {"rev", "pers", "ebitda"}   # the always-lines
    m = dataclasses.replace(CAT.materiality, abs_nok=100)
    low = fin_pnl.compute(dataclasses.replace(CAT, materiality=m), scope, store)
    assert low.flags["ly"]["rev"].basis == "ytd+month"                    # a month swing of 100 counts too
    assert low.flags["ly"]["ebitda"].basis == "ytd"                       # the month moved 29 only


def test_bridges_close(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)
    assert res.line_bridge.closes() and [s.key for s in res.line_bridge.steps] == ["rev", "cogs", "pers", "oth", "c73"]
    ub = res.unit_bridge
    assert ub.closes() and [s.key for s in ub.steps] == ["un.SY.U1", "un.SY.U2", "other"]
    assert ub.steps[-1].label == "Other units and eliminations"
    assert ub.steps[-1].fact.value == pytest.approx(261 - 2 * 261)


def test_kpis_signed_positive_lower_is_better_and_headline_first(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)
    assert list(res.kpis)[0] == "personnel_per_fte" == res.headline
    assert res.kpis["ebitda_margin"]["ytd.cur"].value == pytest.approx(290 / 1000)
    c = res.kpis["opex_per_fte"]["ytd.cur"]
    assert c.value == pytest.approx(300 * YR * 9 / AVG_FTE * 12 / 9) and c.unit == "nok_per_fte"    # shown positive
    ly = 300 * 9 / 10 * 12 / 9
    v = res.kpi_vars["opex_per_fte"]["ytd"]
    assert v.abs.value == pytest.approx(ly - c.value) and v.favourable is False    # cost per FTE rose: worse
    assert "m.cur" not in res.kpis["opex_per_fte"] and "m" in res.kpi_vars["ebitda_margin"]
    assert res.facts["kpi.opex_per_fte.ytd.var"] is v.abs


def test_compare_limits_the_flags_and_a_missing_month_is_explicit(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, dataclasses.replace(scope, compare=("ly",)), store)
    assert set(res.flags) == {"ly"}
    later = fin_pnl.compute(CAT, dataclasses.replace(scope, period="Nov 26"), store)
    assert later.fm["ebitda"]["ytd"].current.value is None and "Oct 26" in later.fm["ebitda"]["ytd"].current.reason
    assert any(f.id == "fm.ebitda.ytd.cur" for f in later.missing())


def test_a_cell_never_read_is_missing_but_a_blank_cell_is_zero(tmp_path):
    store, _, scope = filled(tmp_path)
    key = ("FinMax", "Actual", "", WHS.finmax, "ac.C60", "Mar 26", "Periodic")
    blank = store.rows[key]
    store.rows[key] = Row(key, None, blank.source, blank.fetched_at, blank.closed_month)
    res = fin_pnl.compute(CAT, scope, store)
    assert res.fm["oth"]["ytd"].current.value == pytest.approx(-100 * YR * 8)        # blank = 0
    del store.rows[key]
    res = fin_pnl.compute(CAT, scope, store)
    gone = res.fm["oth"]["ytd"]
    assert gone.current.value is None and "Mar 26" in gone.current.reason and gone.abs.value is None
    assert res.kpis["opex_per_fte"]["ytd.cur"].value is not None                       # opex is its own line
    assert {f.id for f in res.missing()} >= {"fm.oth.ytd.cur", "fm.oth.ytd.var"}


def test_an_unexplained_bridge_residual_is_a_problem(tmp_path):
    store, _, scope = filled(tmp_path)
    key = ("FinMax", "Actual", "", WHS.finmax, "ac.EBITDA", "Mar 26", "Periodic")
    r = store.rows[key]
    store.rows[key] = Row(key, r.value + 50, r.source, r.fetched_at, r.closed_month)
    res = fin_pnl.compute(CAT, scope, store)
    assert res.line_bridge.steps[-1].key == "unexplained" and res.line_bridge.unexplained == pytest.approx(50)
    assert res.problems and res.problems[0].startswith("Unexplained: ") and "NOK 50 " in res.problems[0]


def test_an_abolished_area_is_refused(tmp_path):
    store, _, scope = filled(tmp_path)
    con = next(e for e in CAT.entities if e.key == "consumer")
    with pytest.raises(ValueError, match="abolished"):
        fin_pnl.compute(CAT, dataclasses.replace(scope, entity=con), store)


def test_an_open_month_is_read_from_the_top_up_and_its_flags_are_provisional(tmp_path):
    store, _, scope = filled(tmp_path, period="Oct 26")
    res = fin_pnl.compute(CAT, scope, store)
    assert res.fm["ebitda"]["m"].current.value == pytest.approx(290 * YR)
    assert res.freshness["FinMax"]["open_rows"] > 0 and res.provisional
    assert "open month read live" in fin_pnl.freshness_text(res, "live read")
    s = fin_pnl.summary(res, "live read", [], cat=CAT)
    assert s["area"] == "Wholesale" and s["landing"]["method"].startswith("seasonality")
    assert "facts" not in s and s["headline_kpi"] == "personnel_per_fte" and s["missing_count"] == 0
    assert s["provisional"] and all(f["provisional"] for f in s["flags"])
    assert fin_pnl.flag_text(res, "ly", "rev") == "comment - provisional"
    assert any("still open" in w for w in fin_pnl.tool_warnings(res, []))
    full = fin_pnl.summary(res, "live read", [], cat=CAT, detail="full")
    assert full["facts"] and any(f["id"].endswith(".pct") for f in full["facts"])


def test_the_tool_output_is_compact_signed_and_labelled_per_cube(tmp_path):
    store, _, scope = filled(tmp_path)
    labels = fin_pnl.line_labels(CAT, lambda cube, acct: "C30 REVENUES" if acct == "ac.C30" else None)
    assert labels[("FinMax", "rev")] == "C30 REVENUES (Revenue)" and labels[("FinPlan", "oth")] == "Other items"
    res = fin_pnl.compute(CAT0, dataclasses.replace(scope, labels=labels), store)
    s = fin_pnl.summary(res, "live read", [], cat=CAT0)
    rev = next(f for f in s["flags"] if f["comparison"] == "ly" and f["line"] == "rev")
    assert rev["label"] == "C30 REVENUES (Revenue)" and rev["ytd_var"] == pytest.approx(900)
    assert rev["month_var"] == pytest.approx(100) and rev["favourable"] is True and "variance" not in rev
    row = next(r for r in s["table"] if r["comparison"] == "forecast" and r["line"] == "oth")
    assert row["label"] == "Other items" and row["ytd"]["pct"] == pytest.approx(
        res.forecast["oth"]["ytd"].pct.value)
    assert any(r["line"] == "kpi.personnel_per_fte" and r["label"].endswith("(annualised)") for r in s["table"])
    assert res.line_bridge.steps[0].label == "C30 REVENUES (Revenue)"


def test_the_variance_pct_is_blank_under_the_floor(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)                       # invented figures: all far below NOK 500k
    assert res.fm["rev"]["ytd"].pct.value is not None
    assert fin_pnl.shown_pct(res.fm["rev"]["ytd"], CAT.materiality) is None
    assert fin_pnl.shown_pct(res.fm["fte"]["ytd"], CAT.materiality) is not None    # FTE is not an amount
    s = fin_pnl.summary(res, "live read", [], cat=CAT)
    assert next(r for r in s["table"] if r["line"] == "rev")["ytd"]["pct"] is None


def test_a_finished_year_has_no_landing(tmp_path):
    store, _, scope = filled(tmp_path, period="Dec 25", forecast="Forecast Q325")
    res = fin_pnl.compute(CAT, scope, store)
    assert res.past_year and res.landing is None and not fin_pnl.landing_warnings(res)
    assert res.budget["ebitda"]["ytd"].current.value == pytest.approx(290 * 12)
    assert any("finished year" in n for n in res.notes) and not res.provisional


def test_a_forecast_without_data_is_detected(tmp_path):
    store, _, _ = filled(tmp_path, value_fn=make_value(empty=("Forecast Q426",)), forecast="Forecast Q426")
    assert not fin_pnl.forecast_has_data(store, CAT, WHS, "Forecast Q426", 26)
    store2, _, _ = filled(tmp_path / "b")
    assert fin_pnl.forecast_has_data(store2, CAT, WHS, FC, 26)


def test_warned_fills_and_refresh_notes_are_tool_warnings(tmp_path):
    from smartview_mcp.extract import FillResult
    from smartview_mcp.factstore import Fill
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)
    warn = FillResult(Fill("new_month", "x", 0.0, 0.1, 5, "warn", {"problems": ["C30 + C40S = C48 is off by 3"]}))
    w = fin_pnl.tool_warnings(res, [warn], ["a live read was slow: retried once"])
    assert any("C30 + C40S = C48 is off by 3" in x for x in w) and w[-1].startswith("a live read was slow")



def test_refresh_notes_are_in_the_freshness_line(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)
    note = "a live read was slow: retried once"
    assert note in fin_pnl.freshness_text(res, "live read", [note])
    assert fin_pnl.summary(res, "live read", [], [note])["refresh_notes"] == [note]
    assert "slow" not in fin_pnl.freshness_text(res, "live read")


def test_kpi_rows_have_no_variance_pct(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT0, scope, store)
    rows = [r for r in fin_pnl.summary(res, "live read", [], cat=CAT0)["table"] if r["line"].startswith("kpi.")]
    assert rows and all(r["ytd"]["pct"] is None and (r["month"] is None or r["month"]["pct"] is None) for r in rows)
    assert res.kpi_vars["ebitda_margin"]["ytd"].pct.value is not None                 # the engine still has it
