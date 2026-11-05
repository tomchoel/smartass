"""Plan 5b: presentation fixes Tom found in real answers (05.11.2026) - clean concept labels with the code in the code
column, a concept waterfall of the total's items, one hue per metric with a legend, charts split by unit, empty series
and leading empty months left out, line labels only at the ends, trend totals only for NOK flows, short trend headers,
the subtitle's period, trend and concept key points, no placeholder cause, no variance against a missing base. Every
figure is invented."""
from functools import partial

from test_v2_block_concept import TINY, cfacts, cresolver, cspec
from test_v5_charts import draw, sets, spec
from v2util import CLOSED, monthly_facts, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.blocks.common import ck
from smartass.model.blocks.concept import sheet_text
from smartass.model.expr import Var
from smartass.model.frame import Fact, Frame
from smartass.model.resolver import MetricDef, StaticResolver
from smartass.model.spec import AnalysisSpec
from smartass.model.units import COUNT, NOK, OTHER, PCT, PER_UNIT, short_name, unit_class
from smartass.render import charts as C
from smartass.render.keypoints import key_points
from smartass.render.layout import answer_subtitle, layout, sheet_label
from smartass.xl.vendor import tn_rapportstil as tn


def run(d, r=None, frame=None, stock=()):
    s = AnalysisSpec.model_validate(d)
    res = compute(s, r or resolver(), CLOSED, frame or monthly_facts())
    kp = partial(key_points, stock=set(stock))
    return s, res, layout(1, s, res, stock=set(stock), subtitle="x", key_points=kp)


# --- 1. concept labels: the code in the code column, a clean label on the sheet, the model label kept -------------

def test_sheet_text_drops_booked_and_a_bracket_that_repeats_the_code():
    assert sheet_text("Revenue (C30) (booked)", "C30") == "Revenue"
    assert sheet_text("Memo: revenue Streamix Mobil (C30) (booked)", "C30") == "Memo: revenue Streamix Mobil"
    assert sheet_text("Interest income (C80311) (attributed 100 %)", "C80311") == "Interest income (attributed)"
    assert sheet_text("Opex on Splitt (C70, incl. commission C60140) (booked)", "C70") == \
        "Opex on Splitt (C70, incl. commission C60140)"
    assert sheet_text("Hardware COGS (estimated)", "C40S") == "Hardware COGS (estimated)"


def test_concept_lines_carry_the_code_and_a_sheet_label_while_key_values_keep_the_model_label():
    s, res, p = run(cspec(), r=cresolver(), frame=cfacts())
    out = next(o for o in res.outs if o.type == "concept")
    rev = next(ln for ln in out.lines if ln.key == "i_rev")
    assert rev.code == "C30" and sheet_label(rev) == "Revenue" and rev.label == "Revenue (booked)"
    fin = next(ln for ln in out.lines if ln.key == "i_fin")
    assert sheet_label(fin) == "Interest income (attributed)"
    assert f"{out.title} · Revenue (booked)" in res.key_values            # golden keys unchanged
    shown = {v for (r, c), v in p.cells.items() if c == 3 and isinstance(v, str)}
    assert "   Revenue" in shown and not any("(booked)" in v for v in shown)
    assert p.cells[p.at(ck(out.key, "i_rev", 0))[0], 2] == "C30"


# --- 2. a concept's chart: the waterfall of its total's items ----------------------------------------------------

def test_concept_implicit_chart_is_a_waterfall_of_the_total_items_only():
    s, res, p = run(cspec(), r=cresolver(), frame=cfacts())
    (cs,) = p.charts
    assert cs.kind == "waterfall" and cs.axis_min is None            # a concept's bars always start at zero
    out = next(o for o in res.outs if o.type == "concept")
    assert out.steps == ["i_rev", "i_cogs", "i_fin", "total"]
    rows = [p.at(ck(out.key, k, 0))[0] for k in out.steps]
    assert cs.cat_range == ",".join(f"C{r}" for r in rows) or cs.cat_range.startswith(f"C{rows[0]}")
    cats = [int(x.strip("C")) for part in cs.cat_range.split(",") for x in part.split(":")]
    heading = next(ln for ln in out.lines if ln.key == "g0")
    assert all(r != p.at(ck(out.key, "i_gp", 0))[0] for r in cats)          # no sum row, no heading, no missing
    assert heading.cells == [None] * len(heading.cells)


def test_an_explicit_column_chart_on_a_concept_becomes_its_waterfall():
    d = cspec()
    d["blocks"].append({"type": "chart", "kind": "column", "title": "Tiny", "source": {"block": 0}})
    s, res, p = run(d, r=cresolver(), frame=cfacts())
    assert [c.kind for c in p.charts] == ["waterfall"]


# --- 3. colours and legend ---------------------------------------------------------------------------------------

def test_each_metric_its_own_hue_py_a_tint_and_a_legend_from_two_series():
    cs = spec("column", [("Rev AC", "D", "ac"), ("Rev PY", "E", "py"), ("Opex AC", "F", "ac")])
    cs.hues = [0, 0, 1]
    log, _ = draw(cs)
    fore = dict(sets(log, ".Format.Fill.ForeColor.RGB"))
    assert fore["c1.s1.Format.Fill.ForeColor.RGB"] == tn.rgb(C.HUES[0])
    assert fore["c1.s2.Format.Fill.ForeColor.RGB"] == tn.rgb(C.tint(C.HUES[0]))
    assert fore["c1.s3.Format.Fill.ForeColor.RGB"] == tn.rgb(C.HUES[1])
    assert ("set", "c1.HasLegend", True) in log and ("set", "c1.Legend.Position", C.XL_LEGEND_TOP) in log
    log, _ = draw(spec("column", [("AC", "D", "ac")]))
    assert ("set", "c1.HasLegend", False) in log                    # one AC series: the title names it


def test_line_scenarios_dashed_budget_dotted_forecast_and_a_lone_budget_line_gets_a_legend():
    cs = spec("line", [("Subs BU", "D", "bu"), ("Subs FC", "E", "fc")])
    log, _ = draw(cs)
    dash = dict(sets(log, ".Format.Line.DashStyle"))
    assert dash["c1.s1.Format.Line.DashStyle"] == C.MSO_LINE_DASH
    assert dash["c1.s2.Format.Line.DashStyle"] == C.MSO_LINE_ROUND_DOT
    log, _ = draw(spec("line", [("Subs BU", "D", "bu")]))
    assert ("set", "c1.HasLegend", True) in log                     # a dashed line must say it is budget
    log, _ = draw(spec("waterfall", [("Base", "D", "base"), ("Up", "E", "up"), ("Total", "F", "total")]))
    assert ("set", "c1.HasLegend", False) in log


def test_short_name_caps_long_finplan_members_keeping_their_last_words():
    assert short_name("Fixed Streamix ARPU") == "Fixed Streamix ARPU"
    s = short_name("Fixed Access Independent Streamix Subscriptions")
    assert len(s) <= 32 and s.startswith("…") and s.endswith("Streamix Subscriptions")


# --- 4. units: classes and the split ------------------------------------------------------------------------------

def _units_resolver():
    return StaticResolver(metrics=[MetricDef("Revenue", {"FinMax": ("ac.C30",)}),
                                   MetricDef("Heads", {"FinMax": ("ac.H1",)}, unit="count"),
                                   MetricDef("Mobile ARPU", {"FinPlan": ("ARPU Mobile",)})],
                          extra_members={"Fixed Streamix Subscriptions", "Churn %", "Something"})


def test_unit_class_from_the_semantic_metric_or_the_finplan_name():
    r = _units_resolver()
    assert unit_class(r, "Revenue", "FinMax") == NOK
    assert unit_class(r, "Heads", "FinMax") == COUNT
    assert unit_class(r, "Mobile ARPU", "FinPlan") == PER_UNIT
    assert unit_class(r, "ac.C40S", "FinMax") == NOK
    assert unit_class(r, "Fixed Streamix Subscriptions", "FinPlan") == COUNT
    assert unit_class(r, "Fixed Streamix ARPU", "FinPlan") == PER_UNIT
    assert unit_class(r, "Churn %", "FinPlan") == PCT
    assert unit_class(r, "Something", "FinPlan") == OTHER


def _fp_frame(actual=True, budget=True):
    out = []
    for m, lab in enumerate(("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"), start=1):
        for (scen, ver), ok, k in ((("Actual", "Working"), actual, 1.0), (("Target", "Final"), budget, 1.1)):
            if ok:
                out += [Fact("FinPlan", "fp", "Fixed Streamix Subscriptions", "2EI", f"{lab} 26", scen, ver,
                             1000.0 * m * k),
                        Fact("FinPlan", "fp", "Fixed Streamix ARPU", "2EI", f"{lab} 26", scen, ver, 120.0 + m)]
    return Frame(out)


def _fp_spec():
    d = spec_dict(queries=[{"id": "fp", "cube": "FinPlan", "metrics": ["Fixed Streamix Subscriptions",
                                                                       "Fixed Streamix ARPU"],
                            "entity": "2EI", "periods": "closed-ytd", "compare": ["budget"]}],
                  blocks=[{"type": "trend", "query": "fp", "title": "Streamix",
                           "metrics": ["Fixed Streamix Subscriptions", "Fixed Streamix ARPU"],
                           "series": ["value", "budget"]}])
    return d


def _fp_resolver():
    return StaticResolver(metrics=[], extra_members={"Fixed Streamix Subscriptions", "Fixed Streamix ARPU", "2EI"})


def _fp_run(frame):
    from smartass.model.periods import Closed  # noqa: F401
    s = AnalysisSpec.model_validate(_fp_spec())
    closed = {"FinMax": CLOSED, "FinPlan": CLOSED, "FinLite": CLOSED}
    res = compute(s, _fp_resolver(), closed, frame)
    return s, res, layout(1, s, res, stock=set(), subtitle="x", key_points=partial(key_points, stock=set()))


def test_a_trend_of_subscriptions_and_arpu_is_split_into_one_chart_per_unit():
    s, res, p = _fp_run(_fp_frame())
    assert [c.title for c in p.charts] == ["Streamix — Fixed Streamix Subscriptions", "Streamix — Fixed Streamix ARPU"]
    assert [c.unit for c in p.charts] == [COUNT, PER_UNIT]
    assert all(not c.millions and c.whole for c in p.charts)
    assert p.charts[0].hues == [0, 0] and p.charts[1].hues == [1, 1]     # a metric keeps its hue
    assert p.charts[0].anchor != p.charts[1].anchor


def test_empty_series_are_left_out_with_a_lint_line_and_no_data_means_no_chart():
    s, res, p = _fp_run(_fp_frame(actual=False))
    assert [[r for _, _, r in c.series] for c in p.charts] == [["bu"], ["bu"]]
    assert sum("has no data; left out" in x for x in p.chart_lint) == 2
    s, res, p = _fp_run(_fp_frame(actual=False, budget=False))
    assert all(not c.series for c in p.charts) and any("no series has data" in x for x in p.chart_lint)


# --- 6-7. a trend's leading empty months and its line labels --------------------------------------------------------

def _launch_frame():
    """Revenue from Aug 26 only (launched in August): Jul 26 and the year before are empty."""
    out = []
    for mon, lab in ((8, "Aug"), (9, "Sep")):
        out.append(Fact("FinMax", "fm", "ac.C30", "un.SY.2BUS", f"{lab} 26", "Actual", "", 10.0 * mon))
    return Frame(out)


def _launch_spec():
    d = spec_dict(blocks=[{"type": "trend", "query": "fm", "metrics": ["Revenue"]}])
    d["queries"][0]["periods"] = "Jul 26..Sep 26"
    return d


def test_trend_chart_starts_at_the_first_month_with_data_and_labels_only_ends():
    s, res, p = run(_launch_spec(), frame=_launch_frame())
    (cs,) = p.charts
    out = next(o for o in res.outs if o.type == "trend")
    aug = p.at(ck(out.key, "p1", 0))[0]
    assert cs.cat_range == f"C{aug}:C{aug + 1}"                          # Jul 26 (empty) is not a category
    assert [r for _, _, r in cs.series] == ["ac"]                         # PY has no data
    assert cs.label_points == [[1, 2]]                                    # last + min/max of two points
    assert p.cells[(p.at(ck(out.key, "p0", 0))[0], 3)].strip() == "Jul 26"   # the table keeps the row


def test_line_labels_are_switched_off_except_the_chosen_points():
    cs = spec("line", [("Rev AC", "D", "ac")])
    cs.label_points = [[5]]
    log, _ = draw(cs)
    off = [n for n, v in sets(log, ".HasDataLabel") if v is False]
    assert len(off) == 4 and not any("Points(5)" in n for n in off)
    assert ("set", "c1.s1.Points(5).DataLabel.Position", C.XL_LABEL_RIGHT) in log


# --- 8-9. trend totals and headers ----------------------------------------------------------------------------------

def test_trend_total_only_for_nok_flows_and_short_headers_keep_old_keys():
    s, res, p = _fp_run(_fp_frame())
    out = next(o for o in res.outs if o.type == "trend")
    assert "total" not in [ln.key for ln in out.lines]                   # counts and ARPU never add up
    assert out.headers == ["Fixed Streamix Subscriptions AC", "Fixed Streamix Subscriptions BU",
                           "Fixed Streamix ARPU AC", "Fixed Streamix ARPU BU"]
    assert "Fixed Streamix Subscriptions Actual" in res.key_values["Streamix · Jan 26"]


# --- 10. the subtitle -----------------------------------------------------------------------------------------------

def test_subtitle_uses_the_trend_period_and_no_nok_for_counts():
    s, res, p = _fp_run(_fp_frame())
    sub = answer_subtitle(s, res, CLOSED)
    assert "YTD Sep 26" in sub and not sub.startswith("NOK") and "AC" not in sub
    s, res, p = run(_launch_spec(), frame=_launch_frame())
    assert answer_subtitle(s, res, CLOSED).startswith("NOK") and "Jul–Sep 26" in answer_subtitle(s, res, CLOSED)


# --- 11-13. key points ----------------------------------------------------------------------------------------------

def _points(p):
    start = next(r for name, r in p.zones if name == "key_points")
    end = next(r for name, r in p.zones if name == "charts")
    return [p.cells.get((r, 2)) for r in range(start + 1, end) if p.cells.get((r, 2))]


def test_trend_main_block_gets_a_point_per_metric_with_its_unit():
    s, res, p = _fp_run(_fp_frame())
    pts = _points(p)
    assert len(pts) == 2
    assert pts[0].startswith('="•  Fixed Streamix Subscriptions Sep 26: "&FIXED(') and ",0)" in pts[0]
    assert "vs BU" in pts[0] and "MNOK" not in pts[0] and "total" not in pts[0]
    assert '" NOK' in pts[1]                                             # ARPU in NOK
    s, res, p = _fp_run(_fp_frame(actual=False))
    pts = _points(p)
    assert "Subscriptions BU Sep 26" in pts[0] and "no AC" in pts[0]


def test_trend_point_of_a_nok_flow_has_its_total_and_says_no_py_for_a_new_product():
    s, res, p = run(_launch_spec(), frame=_launch_frame())
    (pt,) = _points(p)
    assert "Revenue Sep 26: " in pt and "; no PY (new)" in pt and " · total Jul–Sep 26: " in pt


def test_concept_points_are_the_total_then_the_largest_items_and_no_placeholder_cause():
    s, res, p = run(cspec(), r=cresolver(), frame=cfacts())
    pts = _points(p)
    assert pts[0].startswith('="•  Direct contribution Tiny ')
    assert any("Revenue" in x for x in pts[1:])
    assert not any("Cause not yet explained" in str(v) for v in p.cells.values())


def test_a_base_missing_on_every_line_gives_a_blank_variance_not_the_whole_value():
    frame = Frame([f for f in cfacts().facts if f.period != "Sep 25"])          # no last year at all
    s, res, p = run(cspec(), r=cresolver(), frame=frame)
    out = next(o for o in res.outs if o.type == "concept")
    vi = out.cols.index("var")
    rev = next(ln for ln in out.lines if ln.key == "i_rev")
    assert isinstance(rev.cells[vi], Var) and res.cells.get(ck(out.key, "i_rev", vi)) is None
    f = p.cells[p.at(ck(out.key, "i_rev", vi))]
    ly = p.addr[ck(out.key, "i_rev", out.cols.index("ly"))]
    assert f.startswith(f'=IF({ly}=0,"",')
    pts = _points(p)
    assert sum("no PY (new)" in x for x in pts) == 1 and not any("vs PY" in x for x in pts)


def test_a_base_missing_on_some_lines_keeps_current_minus_zero():
    frame = Frame([f for f in cfacts().facts if not (f.period == "Sep 25" and f.account == "ac.C80311")])
    s, res, p = run(cspec(), r=cresolver(), frame=frame)
    out = next(o for o in res.outs if o.type == "concept")
    vi = out.cols.index("var")
    assert res.cells.get(ck(out.key, "i_fin", vi)) == 7.0                 # as before: the golden rule


def test_tiny_concept_fixture_is_unchanged():
    assert TINY.id == "tiny"


def test_material_marker_reads_a_blank_variance_through_n():
    from v2util import facts
    frame = Frame([f for f in facts().facts if f.period != "Sep 25"])       # no last year on any line
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "title": "P&L"}])
    s, res, p = run(d, frame=frame)
    out = next(o for o in res.outs if o.type == "table")
    vi, mi = out.cols.index("var"), out.cols.index("material")
    ln = next(x for x in out.lines if x.cells[mi] is not None)
    var_at = p.addr[ck(out.key, ln.key, vi)]
    assert f"ABS(N({var_at}))" in p.cells[p.at(ck(out.key, ln.key, mi))]   # never ABS("") -> #VALUE!
    assert res.cells.get(ck(out.key, ln.key, vi)) is None and res.cells.get(ck(out.key, ln.key, mi)) == 0.0


def test_concept_key_points_never_pick_a_memo_item():
    from types import SimpleNamespace

    from smartass.model.blocks.common import BlockOut, Line
    from smartass.render.keypoints import _concept_chosen
    main = BlockOut("b0", "concept", "C", ["AC", "PY", "ΔPY"], cols=["value", "ly", "var"],
                    lines=[Line("i_rev", "Revenue", [None] * 3, metric="Revenue"),
                           Line("i_memo", "Memo: revenue", [None] * 3, metric="Revenue"),
                           Line("total", "Total", [None] * 3, kind="grand")])
    main.steps = ["i_rev", "total"]
    res = SimpleNamespace(material=[{"key": "b0.i_memo", "variance": 9e9, "base": "ly"},
                                    {"key": "b0.i_rev", "variance": 1e6, "base": "ly"}])
    assert [ln.key for ln, _, _ in _concept_chosen(main, res, {})] == ["total", "i_rev"]
