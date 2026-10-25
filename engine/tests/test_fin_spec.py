"""fin_spec: the sv_ws_build request for fin_pnl. The source sheets are simulated by retrieving each pull from the
same fake provider; the spec evaluated in Python (specs.evaluate: the sheet's formula rules) must give the engine's
facts, and warehouse = sheet must hold. Every figure is invented."""
import re

import pytest
from fin_util import CAT, WHS, filled, http

from smartview_mcp import fin_pnl, fin_spec, specs
from smartview_mcp.grid import Axis
from smartview_mcp.userexcel import _make_pull
from smartview_mcp.warehouse import Row


def sheets(tmp_path, req):
    """{pull id: Pull} as the source sheets would hold them (same provider, suppression off)."""
    sv, out = http(tmp_path / "sheets"), {}
    for pl in req.pulls:
        rows = [Axis(a["dimension"], a["members"]) for a in pl["rows"]]
        cols = [Axis(a["dimension"], a["members"]) for a in pl["columns"]]
        lay, parsed, _ = sv.retrieve(pl["connection"], pl["pov"], rows, cols, False, False)
        out[pl["id"]] = _make_pull(f"Q1 {pl['connection']} {pl['id']}", pl["connection"], lay, parsed, 0.0)
    return out


def evaluated(req, pulls):
    ids = {pid: p.sheet for pid, p in pulls.items()}
    sess = {p.sheet: p for p in pulls.values()}
    s = specs.expand(req.spec, ids, sess)
    return s, specs.evaluate(s, sess)


def _built(tmp_path, store, scope):
    res = fin_pnl.compute(CAT, scope, store)
    req = fin_spec.build(res, CAT, now="2026-10-25 12:00")
    return res, req, sheets(tmp_path, req)


@pytest.fixture
def built(tmp_path):
    store, _, scope = filled(tmp_path)
    return _built(tmp_path, store, scope)


def _all_cells(spec):
    for sec in spec["sections"]:
        for ln in sec["lines"]:
            for c in ln["cells"]:
                yield ln, c


def test_the_request_has_the_five_source_sheets_and_formula_cells_only(built):
    res, req, pulls = built
    assert [p["id"] for p in req.pulls] == ["pnl", "fte", "fpa", "fpp", "units"]
    assert req.spec["auto_points"] is False and req.spec["chart"]["type"] == "waterfall"
    titles = [s["title"] for s in req.spec["sections"]]
    assert titles[0].startswith("P&L vs last year (FinMax)") and any(t.startswith("Year-end landing") for t in titles)
    for ln, c in _all_cells(req.spec):
        if isinstance(c, dict) and "value" in c:                      # only the labelled inputs hold typed numbers
            assert ln["key"] in fin_spec.INPUT_KEYS
        if isinstance(c, dict) and "expr" in c:                       # no constants inside formulas
            assert not re.search(r"(?<![\w])\d", c["expr"]), c["expr"]
    labels = [c["label"] for c in req.spec["checks"]]
    assert "C30 + C40S = C48" in labels and "Revenue + COGS = Gross profit" in labels
    assert "Line bridge closes" in labels and "Sub-unit bridge closes" in labels and labels[-1] == "Period is closed"
    kpis = next(s for s in req.spec["sections"] if s["title"].startswith("KPIs"))
    assert kpis["lines"][0]["key"] == "k_personnel_per_fte"                         # the headline KPI first


def test_every_line_evaluates_to_the_engine_facts(built):
    res, req, pulls = built
    _, v = evaluated(req, pulls)
    cols = {0: "m.cur", 1: "m.ly", 2: "m.var", 3: "m.pct", 4: "ytd.cur", 5: "ytd.ly", 6: "ytd.var", 7: "ytd.pct"}
    for x in CAT.pnl["FinMax"]:
        for j, part in cols.items():
            assert v[x.key][j] == pytest.approx(res.facts[f"fm.{x.key}.{part}"].value), (x.key, part)
    for x in CAT.pnl["FinPlan"]:
        for name, pre in (("budget", "b_"), ("forecast", "f_")):
            b = f"fp.{x.key}"
            assert v[pre + x.key][0] == pytest.approx(res.facts[f"{b}.actual.m"].value)
            assert v[pre + x.key][1] == pytest.approx(res.facts[f"{b}.{name}.m"].value)
            assert v[pre + x.key][5] == pytest.approx(res.facts[f"{b}.{name}.ytd"].value)
            assert v[pre + x.key][6] == pytest.approx(res.facts[f"{b}.{name}.ytd.var"].value)
            assert v[pre + x.key][7] == pytest.approx(res.facts[f"{b}.{name}.ytd.pct"].value)
    assert v["fte"][4] == pytest.approx(res.fm["fte"]["ytd"].current.value)          # the n+1 average


def test_kpi_value_and_variance_formulas_equal_the_engine(built):
    res, req, pulls = built
    _, v = evaluated(req, pulls)
    for key, f in res.kpis.items():
        row = v[f"k_{key}"]
        assert row[4] == pytest.approx(f["ytd.cur"].value) and row[5] == pytest.approx(f["ytd.ly"].value), key
        assert row[6] == pytest.approx(res.kpi_vars[key]["ytd"].abs.value), key
        if "m.cur" in f:
            assert row[0] == pytest.approx(f["m.cur"].value) and row[2] == pytest.approx(
                res.kpi_vars[key]["m"].abs.value)
        else:
            assert row[0] is None
    assert v["k_personnel_per_fte"][4] > 0                                             # cost per FTE shown positive
    assert v["k_opex_per_fte"][6] < 0                                                  # it rose: worse, negative


def test_bridges_and_landing_evaluate_to_the_engine(built):
    res, req, pulls = built
    _, v = evaluated(req, pulls)
    assert v["l_fc"][4] == pytest.approx(res.landing.forecast.value)
    assert v["l_rr"][4] == pytest.approx(res.landing.run_rate.value)
    assert v["l_gap"][4] == pytest.approx(res.landing.gap_forecast.value)
    assert v["l_rr_gap"][4] == pytest.approx(res.landing.gap_run_rate.value)
    assert v["l_share"][4] == pytest.approx(0.75)                                     # EBITDA before other items
    assert v["lb_sum"][4] == pytest.approx(v["lb_end"][4]) and v["ub_sum"][4] == pytest.approx(v["ub_end"][4])
    for st in res.line_bridge.steps:
        assert v[f"lb_{st.key}"][4] == pytest.approx(st.fact.value)
    assert v["u_other"][6] == pytest.approx(res.unit_bridge.steps[-1].fact.value)
    s, _ = evaluated(req, pulls)
    by_line = {ln["key"]: ln for sec in s["sections"] for ln in sec["lines"]}
    assert by_line["rev"]["cells"][8] == {"text": "comment"}                          # an always-line
    assert by_line["gm"]["cells"][8] is None                                          # a subtotal is not flagged


def test_a_plain_run_rate_uses_the_labelled_month_inputs(tmp_path):
    store, _, scope = filled(tmp_path)
    for k in [k for k in store.rows if k[0] == "FinPlan" and k[1] == "Actual" and k[4] == "EBITDA before Other Items"
              and k[5] == "Jan 25"]:
        del store.rows[k]                                                              # last year's share: unknown
    res, req, pulls = _built(tmp_path, store, scope)
    assert not res.landing.seasonal
    _, v = evaluated(req, pulls)
    assert v["l_rr"][4] == pytest.approx(res.landing.run_rate.value) and "l_share" not in v
    assert v["months_ytd"][4] == 9 and v["months_year"][4] == 12


def test_an_unexplained_residual_and_a_stale_forecast_are_shown(tmp_path):
    store, _, scope = filled(tmp_path, forecast="Forecast Q126")
    key = ("FinMax", "Actual", "", WHS.finmax, "ac.EBITDA", "Mar 26", "Periodic")
    r = store.rows[key]
    store.rows[key] = Row(key, r.value + 50, r.source, r.fetched_at, r.closed_month)
    res = fin_pnl.compute(CAT, scope, store)
    req = fin_spec.build(res, CAT, now="2026-10-25 12:00")
    un = next(c for c in req.spec["checks"] if c["label"].startswith("Unexplained"))
    assert un["static"] == "DIFF" and "50" in un["detail"]
    assert req.spec["checks"][-1]["label"] == "Period is closed"
    assert req.spec["notes"][0].startswith("Warning: Forecast Q126") and "run-rate" in req.spec["notes"][0]
    assert any(n.startswith("Problem: Unexplained") for n in req.spec["notes"])
    assert any(k.startswith("Unexplained:") for k in req.spec["key_points"])


def test_warehouse_equals_every_source_sheet_and_a_difference_is_named(built):
    res, req, pulls = built
    spec = fin_spec.with_reconciliation(req.spec, req.expected, pulls)
    recon = [c for c in spec["checks"] if c["label"].startswith("Warehouse = source sheet")]
    assert len(recon) == 5 and all(c["static"] == "OK" for c in recon), recon
    assert sum(len(v) for v in req.expected.values()) > 60
    p = pulls["pnl"]
    row = next(r for r in p.parsed.rows if r[0] == "ac.C30")
    row[p.layout.nr + 2] += 5.0                                                       # Accumulated, Sep 26
    again = fin_spec.with_reconciliation(req.spec, req.expected, pulls)
    bad = next(c for c in again["checks"] if c["label"] == f"Warehouse = source sheet '{p.sheet}'")
    assert bad["static"] == "DIFF" and "ac.C30 Accumulated/Sep 26: sheet" in bad["detail"]


def test_a_missing_warehouse_figure_is_a_diff_never_a_silent_zero(tmp_path):
    store, _, scope = filled(tmp_path)
    del store.rows[("FinMax", "Actual", "", WHS.finmax, "ac.C60", "Sep 26", "Periodic")]
    res, req, pulls = _built(tmp_path, store, scope)
    spec = fin_spec.with_reconciliation(req.spec, req.expected, pulls)
    bad = next(c for c in spec["checks"] if c["label"] == f"Warehouse = source sheet '{pulls['pnl'].sheet}'")
    assert bad["static"] == "DIFF" and "no warehouse value (fm.oth.m.cur)" in bad["detail"]
    assert any("could not be computed" in n for n in req.spec["notes"])


def test_key_points_are_factual(built):
    res, req, _ = built
    kp = req.spec["key_points"]
    assert kp[0].startswith("EBITDA YTD Sep 26: 2 871 vs 2 610 YTD Sep 25 (variance 261,")
    assert kp[1].startswith("Personnel cost per FTE YTD Sep 26 (annualised):")
    assert any(k.startswith("Year-end landing FY26:") for k in kp)
