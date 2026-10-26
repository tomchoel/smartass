"""fin_spec: the sv_ws_build request for fin_pnl. The source sheets are simulated by retrieving each pull from the
same fake provider; the spec evaluated in Python (specs.evaluate: the sheet's formula rules) must give the engine's
facts, and warehouse = sheet must hold. Every figure is invented. CAT0 has no % floor, so every % is shown."""
import dataclasses
import re

import pytest
from fin_util import CAT, CAT0, WHS, filled, http, make_value

from smartview_mcp import fin_pnl, fin_spec, specs
from smartview_mcp.grid import Axis
from smartview_mcp.userexcel import _make_pull
from smartview_mcp.warehouse import Row

PARTS = {0: "m.cur", 1: "m.ly", 2: "m.var", 3: "m.pct", 4: "ytd.cur", 5: "ytd.ly", 6: "ytd.var", 7: "ytd.pct"}


def sheets(tmp_path, req, value_fn=None):
    """{pull id: Pull} as the source sheets would hold them (same provider, suppression off)."""
    sv, out = http(tmp_path / "sheets", value_fn), {}
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


def build(tmp_path, cat=CAT0, value_fn=None, store_fix=None, **kw):
    store, _, scope = filled(tmp_path, value_fn=value_fn, **kw)
    if store_fix:
        store_fix(store)
    res = fin_pnl.compute(cat, scope, store)
    req = fin_spec.build(res, cat, now="2026-10-25 12:00")
    return res, req, sheets(tmp_path, req, value_fn)


def same(a, b):
    return (a is None and b is None) or (a is not None and b is not None and a == pytest.approx(b))


def assert_sheet_is_engine(res, req, pulls):
    """Every number the sheet shows equals the engine's fact (FinMax, FinPlan, KPIs, bridges, landing)."""
    s, v = evaluated(req, pulls)
    lines = {ln["key"]: ln for sec in s["sections"] for ln in sec["lines"]}
    for x in CAT.pnl["FinMax"]:
        for j, part in PARTS.items():
            if lines[x.key]["cells"][j] is not None:
                assert same(v[x.key][j], res.facts[f"fm.{x.key}.{part}"].value), (x.key, part)
    for x in CAT.pnl["FinPlan"]:
        for name, pre in (("budget", "b_"), ("forecast", "f_")):
            if pre + x.key not in v:
                continue
            b = f"fp.{x.key}"
            want = {0: f"{b}.actual.m", 1: f"{b}.{name}.m", 2: f"{b}.{name}.m.var", 3: f"{b}.{name}.m.pct",
                    4: f"{b}.actual.ytd", 5: f"{b}.{name}.ytd", 6: f"{b}.{name}.ytd.var", 7: f"{b}.{name}.ytd.pct"}
            for j, fid in want.items():
                if lines[pre + x.key]["cells"][j] is not None:
                    assert same(v[pre + x.key][j], res.facts[fid].value), (pre + x.key, j)
    for key, f in res.kpis.items():
        row = v[f"k_{key}"]
        for j, part in ((0, "m.cur"), (1, "m.ly"), (4, "ytd.cur"), (5, "ytd.ly")):
            if part in f and f[part].value is not None:
                assert row[j] == pytest.approx(f[part].value), (key, part)
        if f["ytd.cur"].value is not None and f["ytd.ly"].value is not None:
            assert row[6] == pytest.approx(res.kpi_vars[key]["ytd"].abs.value), key
    if res.line_bridge is not None and res.line_bridge.closes():
        for st in res.line_bridge.steps:
            assert same(v[f"lb_{st.key}"][4], st.fact.value), st.key
        assert v["lb_sum"][4] == pytest.approx(v["lb_end"][4])
    ub = res.unit_bridge
    if ub is not None:
        top = [st for st in ub.steps if st.key != "other"]
        for i, st in enumerate(top, 1):
            assert v[f"u{i}"][4] == pytest.approx(res.facts[f"unit.{st.key}.cur"].value)
            assert v[f"u{i}"][5] == pytest.approx(res.facts[f"unit.{st.key}.ly"].value)
            assert v[f"u{i}"][6] == pytest.approx(st.fact.value)
        if ub.steps[-1].key == "other":
            assert v["u_other"][6] == pytest.approx(ub.steps[-1].fact.value)
        assert v["ub_sum"][4] == pytest.approx(v["ub_end"][4])
    ld = res.landing
    if ld is None:
        assert not any(k.startswith("l_") for k in v)
        return v
    for key, fact in (("l_rr", ld.run_rate), ("l_rr_gap", ld.gap_run_rate), ("l_fc", ld.forecast),
                      ("l_gap", ld.gap_forecast)):
        if key in v and fact.value is not None:
            assert v[key][4] == pytest.approx(fact.value), key
    return v


def reconciled(req, pulls):
    spec = fin_spec.with_reconciliation(req.spec, req.expected, pulls)
    return [c for c in spec["checks"] if c["label"].startswith("Warehouse = source sheet")]


@pytest.fixture
def built(tmp_path):
    return build(tmp_path)


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
    assert "Revenue + COGS = Gross profit (forecast)" in labels                          # the forecast section too
    assert "Line bridge closes" in labels and "Sub-unit bridge closes" in labels and labels[-1] == "Period is closed"
    kpis = next(s for s in req.spec["sections"] if s["title"].startswith("KPIs"))
    assert kpis["lines"][0]["key"] == "k_personnel_per_fte"                         # the headline KPI first


def test_every_number_on_the_sheet_is_the_engine_fact(built):
    res, req, pulls = built
    v = assert_sheet_is_engine(res, req, pulls)
    assert v["fte"][4] == pytest.approx(res.fm["fte"]["ytd"].current.value)          # the n+1 average
    assert v["k_personnel_per_fte"][4] > 0 and v["k_opex_per_fte"][6] < 0             # positive; rose = worse
    assert v["l_share"][4] == pytest.approx(0.75)                                     # EBITDA before other items
    assert all(c["static"] == "OK" for c in reconciled(req, pulls))


@pytest.mark.parametrize("period", ["Jan 26", "Oct 26"])
def test_january_and_an_open_month(tmp_path, period):
    res, req, pulls = build(tmp_path, period=period)
    v = assert_sheet_is_engine(res, req, pulls)
    if period == "Jan 26":
        assert v["fte_avg"][4] == pytest.approx((10 + 11) / 2) and v["months_ytd"][4] == 1 if "months_ytd" in v \
            else v["fte_avg"][4] == pytest.approx(10.5)
    else:
        lines = {ln["key"]: ln for sec in req.spec["sections"] for ln in sec["lines"]}
        assert lines["rev"]["cells"][8]["text"].endswith(" - provisional")
        assert req.spec["notes"][0].startswith("Oct 26 is after the Current Reporting Month")
        assert any("provisional" in k for k in req.spec["key_points"])
    assert all(c["static"] == "OK" for c in reconciled(req, pulls))


def test_a_finished_year_shows_actuals_without_a_landing(tmp_path):
    res, req, pulls = build(tmp_path, period="Dec 25", forecast="Forecast Q325")
    assert res.landing is None
    assert not any(s["title"].startswith("Year-end landing") for s in req.spec["sections"])
    assert not any(n.startswith("Warning:") for n in req.spec["notes"])
    assert_sheet_is_engine(res, req, pulls)


def test_an_earlier_year_names_what_it_could_not_compute(tmp_path):
    res, req, pulls = build(tmp_path, period="Sep 25", forecast=None)          # FinPlan FY24 is not in the warehouse
    assert res.facts["fp.boi.actual.ly_ytd"].value is None and not res.landing.seasonal
    assert_sheet_is_engine(res, req, pulls)
    assert any(n.startswith("Not computed (") and "fp.boi.actual.ly_ytd" in n for n in req.spec["notes"])


def test_no_forecast(tmp_path):
    res, req, pulls = build(tmp_path, forecast=None)
    assert not res.forecast and res.landing.forecast.value is None
    fpp = next(p for p in req.pulls if p["id"] == "fpp")
    assert fpp["columns"][0]["members"] == ["Target"]
    assert not any(s["title"].startswith("P&L vs the latest forecast") for s in req.spec["sections"])
    v = assert_sheet_is_engine(res, req, pulls)
    assert "l_fc" not in v and v["l_rr"][4] == pytest.approx(res.landing.run_rate.value)


def test_compare_last_year_only_still_shows_every_section(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT0, dataclasses.replace(scope, compare=("ly",)), store)
    req = fin_spec.build(res, CAT0, now="2026-10-25 12:00")
    titles = [s["title"] for s in req.spec["sections"]]
    assert any(t.startswith("P&L vs budget") for t in titles) and any(t.startswith("Year-end landing") for t in titles)
    lines = {ln["key"]: ln for sec in req.spec["sections"] for ln in sec["lines"]}
    assert lines["b_rev"]["cells"][8] is None and lines["rev"]["cells"][8] == {"text": "material"}


def test_more_than_five_units_close_through_other_units(tmp_path):
    units = tuple((f"un.SY.U{i}", f"Unit {i}") for i in range(1, 8))
    fn = make_value(scale={f"un.SY.U{i}": i / 10 for i in range(1, 8)})
    res, req, pulls = build(tmp_path, units=units, value_fn=fn)
    assert [st.key for st in res.unit_bridge.steps] == ["un.SY.U7", "un.SY.U6", "un.SY.U5", "un.SY.U4", "un.SY.U3",
                                                        "other"]
    assert_sheet_is_engine(res, req, pulls)
    assert all(c["static"] == "OK" for c in reconciled(req, pulls))


def test_a_cost_unit_with_negative_ebitda_before_other_items(tmp_path):
    fn = make_value(scale={WHS.finplan: -1})                  # FinPlan: every figure negative
    res, req, pulls = build(tmp_path, value_fn=fn)
    assert res.facts["fp.boi.actual.ly_fy"].value < 0 and res.landing.seasonal
    v = assert_sheet_is_engine(res, req, pulls)
    assert v["l_share"][4] == pytest.approx(0.75)


def test_a_share_outside_the_band_uses_the_plain_run_rate(tmp_path):
    def shrink(store):
        for k, r in list(store.rows.items()):
            if k[0] == "FinPlan" and k[1] == "Actual" and k[4] == "EBITDA before Other Items" and k[5].endswith(" 25") \
                    and k[5][:3] not in ("Oct", "Nov", "Dec"):
                store.rows[k] = Row(k, r.value * 0.1, r.source, r.fetched_at, r.closed_month)
    res, req, pulls = build(tmp_path, store_fix=shrink)
    assert not res.landing.seasonal and "outside" in res.landing.method
    v = assert_sheet_is_engine(res, req, pulls)
    assert "l_share" not in v and v["months_year"][4] == 12


def test_the_variance_pct_is_blank_under_the_floor_on_the_sheet(tmp_path):
    res, req, pulls = build(tmp_path, cat=CAT)                # invented figures are far below NOK 500k
    lines = {ln["key"]: ln for sec in req.spec["sections"] for ln in sec["lines"]}
    assert lines["rev"]["cells"][3] is None and lines["rev"]["cells"][7] is None
    assert lines["fte"]["cells"][7] == {"varpct": [4, 5]}    # FTE is not an amount
    assert any("Variance % is left blank" in n for n in req.spec["notes"])
    assert req.spec["key_points"][0].startswith("EBITDA YTD Sep 26: 2 871 vs 2 610 YTD Sep 25 (variance 261).")


def test_labels_follow_each_comparisons_cube(tmp_path):
    store, _, scope = filled(tmp_path)
    labels = fin_pnl.line_labels(CAT, lambda cube, a: {"ac.C60": "C60 OTHER OPEX"}.get(a))
    res = fin_pnl.compute(CAT0, dataclasses.replace(scope, labels=labels), store)
    req = fin_spec.build(res, CAT0)
    lines = {ln["key"]: ln for sec in req.spec["sections"] for ln in sec["lines"]}
    assert lines["oth"]["label"] == "C60 OTHER OPEX (Other opex)" and lines["b_oth"]["label"] == "Other items"
    assert lines["f_oth"]["label"] == "Other items" and lines["lb_oth"]["label"] == "C60 OTHER OPEX (Other opex)"


def test_an_unexplained_residual_and_a_stale_forecast_are_shown(tmp_path):
    def bump(store):
        key = ("FinMax", "Actual", "", WHS.finmax, "ac.EBITDA", "Mar 26", "Periodic")
        r = store.rows[key]
        store.rows[key] = Row(key, r.value + 50, r.source, r.fetched_at, r.closed_month)
    res, req, _ = build(tmp_path, forecast="Forecast Q126", store_fix=bump)
    un = next(c for c in req.spec["checks"] if c["label"].startswith("Unexplained"))
    assert un["static"] == "DIFF" and "50" in un["detail"]
    assert req.spec["checks"][-1]["label"] == "Period is closed"
    notes = req.spec["notes"]
    assert notes[0].startswith("Warning: Forecast Q126") and "run-rate" in notes[0]
    assert notes[1].startswith("Problem: Unexplained") and notes[2].startswith("Data: warehouse;")
    assert any(k.startswith("Unexplained:") for k in req.spec["key_points"])
    assert any("static" in n and "do not re-run" in n for n in notes)


def test_warehouse_equals_every_source_sheet_and_a_difference_is_named(built):
    res, req, pulls = built
    recon = reconciled(req, pulls)
    assert len(recon) == 5 and all(c["static"] == "OK" for c in recon), recon
    assert sum(len(v) for v in req.expected.values()) > 60
    p = pulls["pnl"]
    row = next(r for r in p.parsed.rows if r[0] == "ac.C30")
    row[p.layout.nr + 2] += 5.0                                                       # Accumulated, Sep 26
    bad = next(c for c in reconciled(req, pulls) if c["label"] == f"Warehouse = source sheet '{p.sheet}'")
    assert bad["static"] == "DIFF" and "ac.C30 Accumulated/Sep 26: sheet" in bad["detail"]


def test_fte_reconciles_within_a_twentieth_of_an_fte(built):
    res, req, pulls = built
    p = pulls["fte"]
    row = next(r for r in p.parsed.rows if r[0] == "ac.STA20Y")
    row[p.layout.nr] += 0.03
    assert next(c for c in reconciled(req, pulls) if c["label"].endswith(f"'{p.sheet}'"))["static"] == "OK"
    row[p.layout.nr] += 0.05
    assert next(c for c in reconciled(req, pulls) if c["label"].endswith(f"'{p.sheet}'"))["static"] == "DIFF"


def test_a_missing_warehouse_figure_is_a_diff_never_a_silent_zero(tmp_path):
    def drop(store):
        del store.rows[("FinMax", "Actual", "", WHS.finmax, "ac.C60", "Sep 26", "Periodic")]
    res, req, pulls = build(tmp_path, store_fix=drop)
    bad = next(c for c in reconciled(req, pulls) if c["label"] == f"Warehouse = source sheet '{pulls['pnl'].sheet}'")
    assert bad["static"] == "DIFF" and "no warehouse value (fm.oth.m.cur)" in bad["detail"]
    notes = req.spec["notes"]
    sheet_note = next(n for n in notes if "source figure(s) have no warehouse value" in n)
    assert sheet_note.startswith("2 source figure(s)")                              # m.cur and ytd.cur
    other = next(n for n in notes if n.startswith("Not computed ("))
    listed = [seg.split(": ")[0] for seg in other.split("): ", 1)[1].split("; ")]
    assert "fm.oth.m.var" in listed and "fm.oth.ytd.cur" not in listed and "fm.oth.m.cur" not in listed
    assert notes.index(sheet_note) < next(i for i, n in enumerate(notes) if n.startswith("Data: "))


def test_key_points_are_factual(built):
    res, req, _ = built
    kp = req.spec["key_points"]
    assert kp[0].startswith("EBITDA YTD Sep 26: 2 871 vs 2 610 YTD Sep 25 (variance 261, +10,0 %)")
    assert kp[1].startswith("Personnel cost per FTE YTD Sep 26 (annualised):")
    assert any(k.startswith("Year-end landing FY26:") for k in kp)


def test_material_key_points_use_each_cubes_label_and_the_basis(tmp_path):
    m = dataclasses.replace(CAT0.materiality, abs_nok=40, pct=10.0)          # only the NOK rule can fire
    cat = dataclasses.replace(CAT0, materiality=m)
    res, req, _ = build(tmp_path, cat=cat)
    kp = next(k for k in req.spec["key_points"] if k.startswith("Material variances: "))
    assert "Revenue vs last year 900 (YTD and the month)" in kp
    from smartview_mcp.analytics.materiality import Flag
    res.flags = {"budget": {"oth": Flag("oth", True, False, "a single month moved", 10.0, "month")}}
    kp = next(k for k in fin_spec.key_points(res, cat) if k.startswith("Material variances: "))
    assert kp.startswith("Material variances: Other items vs budget: the month ") and "single-month swing" in kp


def test_a_unit_without_finplan_is_compared_with_last_year_only(tmp_path):
    from fin_util import FC

    from smartview_mcp import extract
    from smartview_mcp.fin_pnl import Scope
    from smartview_mcp.warehouse import MemoryWarehouse
    unit = dataclasses.replace(WHS, key="unit:un.SY.U1", label="U1 Unit one", aliases=(), finmax="un.SY.U1",
                               finplan=None)
    store, retrieve = MemoryWarehouse(), extract.route(http=http(tmp_path))
    grids = extract.base_grids(CAT, "Sep 26", {26: FC}.get, entities=[unit])
    assert grids and all(g.conn == "FinMax" for g in grids)
    extract.fill(store, retrieve, CAT, grids, "live", "Sep 26")
    res = fin_pnl.compute(CAT0, Scope(unit, "Sep 26", "Sep 26", None), store)
    assert not res.budget and not res.forecast and res.landing is None and not res.missing()
    warn = ("Warning: no FinPlan entity for U1 Unit one (un.SY.U1): budget, forecast and landing are not "
            "available; compared with last year only.")
    assert warn in fin_pnl.tool_warnings(res, [])
    req = fin_spec.build(res, CAT0, now="2026-10-25 12:00", mode="live pulls (sub-unit: not kept in the warehouse)")
    assert [p["id"] for p in req.pulls] == ["pnl", "fte"] and req.spec["notes"][0] == warn
    assert "budget" not in req.spec["title"] and req.spec["title"].endswith("vs last year")
    pulls = sheets(tmp_path, req)
    assert_sheet_is_engine(res, req, pulls)
    spec = fin_spec.with_reconciliation(req.spec, req.expected, pulls, source=req.source)
    recon = [c for c in spec["checks"] if c["label"].startswith("Live read = source sheet")]
    assert req.source == "Live read" and len(recon) == 2 and all(c["static"] == "OK" for c in recon)
    assert all(c["static"] == "OK" for c in spec["checks"] if "static" in c)
    assert not any("Warehouse = source" in c["label"] for c in spec["checks"])
    assert any("'Live read = source sheet' checks are static" in n for n in spec["notes"])
    assert fin_spec.build(res, CAT0).source == "Warehouse"
