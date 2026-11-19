"""Plan 5b: presentation fixes Tom found in real answers (05.11.2026) - clean concept labels with the code in the code
column, a concept waterfall of the total's items, one hue per metric with a legend, charts split by unit, empty series
and leading empty months left out, line labels only at the ends, trend totals only for NOK flows, short trend headers,
the subtitle's period, trend and concept key points, no placeholder cause, no variance against a missing base. Every
figure is invented."""
from functools import partial

from test_v2_block_concept import TINY, cfacts, cresolver, cspec
from test_xlsx_charts import spec
from v2util import CLOSED, monthly_facts, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.blocks.common import ck
from smartass.model.blocks.concept import sheet_text
from smartass.model.frame import Fact, Frame
from smartass.model.resolver import MetricDef, StaticResolver
from smartass.model.spec import AnalysisSpec
from smartass.model.units import COUNT, NOK, OTHER, PCT, PER_UNIT, short_name, unit_class
from smartass.render import charts as C
from smartass.render.keypoints import key_points
from smartass.render.layout import answer_subtitle, layout, sheet_label


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
    """The rules; the drawn fills are in test_xlsx_charts.py."""
    cs = spec("column", [("Rev AC", "D", "ac"), ("Rev PY", "E", "py"), ("Opex AC", "F", "ac")])
    cs.hues = [0, 0, 1]
    assert [C.hue(cs, k) for k in range(3)] == [C.HUES[0], C.HUES[0], C.HUES[1]]
    assert C.tint(C.HUES[0]) != C.HUES[0] and C.legend_shown(cs)
    assert not C.legend_shown(spec("column", [("AC", "D", "ac")]))          # one AC series: the title names it


def test_a_lone_budget_line_gets_a_legend_and_a_waterfall_never():
    """The legend rule (the dashed budget and dotted forecast lines are drawn in test_xlsx_charts.py)."""
    assert C.legend_shown(spec("line", [("Subs BU", "D", "bu")]))          # a dashed line must say it is budget
    assert not C.legend_shown(spec("waterfall", [("Base", "D", "base"), ("Up", "E", "up"), ("Total", "F", "total")]))


def test_short_name_drops_whole_leading_words_never_an_ellipsis():
    assert short_name("Fixed Streamix ARPU") == "Fixed Streamix ARPU"
    assert short_name("Fixed Access Independent Streamix Subscriptions") == "Streamix Subscriptions"
    assert short_name("Supercalifragilisticexpialidocious-and-more") == "Supercalifragilisticexpialidocious-and-more"


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
    assert [c.title for c in p.charts] == ["Fixed Streamix Subscriptions vs budget", "Fixed Streamix ARPU vs budget"]
    assert [c.unit for c in p.charts] == [COUNT, PER_UNIT]
    assert all(not c.millions and c.whole for c in p.charts)
    assert p.charts[0].hues == [0, 0] and p.charts[1].hues == [1, 1]     # a metric keeps its hue
    assert p.charts[0].anchor != p.charts[1].anchor


def test_empty_series_are_left_out_with_a_lint_line_and_no_data_means_no_chart():
    s, res, p = _fp_run(_fp_frame(actual=False))
    assert [[r for _, _, r in c.series] for c in p.charts] == [["bu"], ["bu"]]
    assert sum("has no data; left out" in x for x in p.chart_lint) == 0     # 5c: the empty AC column is dropped first
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
    # 5c: the table starts at launch too (no empty pre-launch rows), the empty PY column is gone
    assert [ln.label for ln in out.lines] == ["Aug 26", "Sep 26", "Total"] and out.period == "Aug–Sep 26"
    assert out.cols == ["value"] and out.no_base == ["ly"]
    assert res.cells.get(ck(out.key, "total", 0)) == 170.0


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
    assert answer_subtitle(s, res, CLOSED).startswith("NOK") and "Aug–Sep 26" in answer_subtitle(s, res, CLOSED)


# --- 11-13. key points ----------------------------------------------------------------------------------------------

def _points(p):
    start = next(r for name, r in p.zones if name == "key_points")
    end = next(r for name, r in p.zones if name == "charts")
    return [p.cells.get((r, 2)) for r in range(start + 1, end) if p.cells.get((r, 2))]


def test_trend_main_block_gets_a_point_per_metric_with_its_unit():
    s, res, p = _fp_run(_fp_frame())
    pts = _points(p)
    assert len(pts) == 2
    assert pts[0].startswith('="•  Fixed Streamix Subscriptions Sep 26: "&SUBSTITUTE(FIXED(') and ",0)" in pts[0]
    assert "vs BU" in pts[0] and "MNOK" not in pts[0] and "total" not in pts[0]
    assert '" NOK' in pts[1]                                             # ARPU in NOK
    s, res, p = _fp_run(_fp_frame(actual=False))
    pts = _points(p)
    assert "Subscriptions BU Sep 26" in pts[0] and "no AC" in pts[0]
    out = next(o for o in res.outs if o.type == "trend")
    assert "value" not in out.cols and set(out.cols) == {"budget"}       # 5c: no empty AC columns on the sheet


def test_trend_point_of_a_nok_flow_has_its_total_and_says_no_py_for_a_new_product():
    s, res, p = run(_launch_spec(), frame=_launch_frame())
    (pt,) = _points(p)
    assert "Revenue Sep 26: " in pt and "; no PY (new)" in pt and " · total Aug–Sep 26: " in pt


def test_concept_points_are_the_total_then_the_largest_items_and_no_placeholder_cause():
    s, res, p = run(cspec(), r=cresolver(), frame=cfacts())
    pts = _points(p)
    assert pts[0].startswith('="•  Direct contribution Tiny ')
    assert any("Revenue" in x for x in pts[1:])
    assert not any("Cause not yet explained" in str(v) for v in p.cells.values())


def test_a_base_missing_on_every_line_drops_its_columns_and_says_no_py_once():
    frame = Frame([f for f in cfacts().facts if f.period != "Sep 25"])          # no last year at all
    s, res, p = run(cspec(), r=cresolver(), frame=frame)
    out = next(o for o in res.outs if o.type == "concept")
    assert not {"ly", "var", "var_pct"} & set(out.cols) and out.no_base == ["ly"]    # no empty PY / ΔPY columns
    assert not any("PY" in h for h in out.headers)
    assert "Last year" not in str(res.key_values) and "Variance" not in str(res.key_values)
    assert res.cells.get(ck(out.key, "total", 0)) is not None                       # the values stay
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


def test_a_table_without_any_last_year_has_no_py_variance_or_material_columns():
    from v2util import facts
    frame = Frame([f for f in facts().facts if f.period != "Sep 25"])       # no last year on any line
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "title": "P&L"}])
    s, res, p = run(d, frame=frame)
    out = next(o for o in res.outs if o.type == "table")
    assert out.cols == ["value"] and out.no_base == ["ly"]                # Material has no variance left to judge
    assert all(len(ln.cells) == 1 for ln in out.lines)
    assert sum("no PY (new)" in x for x in _points(p)) == 1


def test_a_base_empty_on_one_period_group_only_drops_that_group_and_rebuilds_material():
    from v2util import facts
    frame = facts()
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "title": "P&L", "periods": ["closed", "closed-ytd"]}])
    d["queries"][0]["periods"] = "closed-ytd"
    s = AnalysisSpec.model_validate(d)
    res = compute(s, resolver(), CLOSED, frame)
    out = next(o for o in res.outs if o.type == "table")
    if "material" in out.cols:                                        # every Ref of the block still resolves
        mi = out.cols.index("material")
        assert all(res.cells.get(ck(out.key, ln.key, mi)) in (0.0, 1.0, None) for ln in out.lines)
    p = layout(1, s, res, stock=set(), subtitle="x", key_points=key_points)
    assert p.cells                                                    # every formula addresses a live cell


def test_one_unit_for_all_key_points_of_a_concept():
    from smartass.render.keypoints import shared_unit
    assert shared_unit([38_151_172.86, -219_601.0]) == "MNOK"
    assert shared_unit([-219_601.0, 7_271.0]) == "NOK"
    assert shared_unit([None]) == "MNOK"
    s, res, p = run(cspec(), r=cresolver(), frame=cfacts())
    pts = _points(p)
    assert not (any('" NOK' in x for x in pts) and any('" MNOK' in x for x in pts))


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


def test_an_agent_chart_title_stays_and_a_split_adds_the_metric():
    d = _fp_spec()
    d["blocks"].append({"type": "chart", "kind": "line", "title": "Streamix runs ahead of plan", "source": {"block": 0}})
    s = AnalysisSpec.model_validate(d)
    closed = {"FinMax": CLOSED, "FinPlan": CLOSED, "FinLite": CLOSED}
    res = compute(s, _fp_resolver(), closed, _fp_frame())
    p = layout(1, s, res, stock=set(), subtitle="x")
    assert [c.title for c in p.charts] == ["Streamix runs ahead of plan — Fixed Streamix Subscriptions",
                                           "Streamix runs ahead of plan — Fixed Streamix ARPU"]


def test_implicit_trend_titles_name_the_metrics_per_month_without_a_base():
    s, res, p = run(_launch_spec(), frame=_launch_frame())
    assert [c.title for c in p.charts] == ["Revenue per month"]           # PY has no data: no base drawn
    s, res, p = run(cspec(), r=cresolver(), frame=cfacts())
    assert p.charts[0].title == "Direct contribution Tiny Sep 26"          # the concept waterfall: total + period


def _notes(p):
    start = next(r for name, r in p.zones if name == "notes")
    return " ".join(str(p.cells.get((r, 2), "")) for r in range(start, start + 4))


def test_conventions_fit_the_sheet():
    s, res, p = _fp_run(_fp_frame())
    note = _notes(p)
    assert "FinPlan KPIs read Periodic per month; AC = actual (Working), BU = Target (Final)." in note
    assert "FinMax" not in note and "NOK" not in note and "▲" not in note and "●" not in note
    s, res, p = run(_launch_spec(), frame=_launch_frame())                # FinMax trend: NOK, no Δ columns
    note = _notes(p)
    assert note.startswith("Notes NOK. FinMax signs: costs are negative.") and "▲" not in note


def test_thresholds_only_what_formulas_use():
    s, res, p = _fp_run(_fp_frame())
    assert not any(name == "Thresholds" for name in (p.cells.get((r, 2)) for r in range(1, p.last_row + 1)))
    assert p.params == {}
    s, res, p = run(cspec(), r=cresolver(), frame=cfacts())                # concept: a check line -> Tolerance
    assert "tolerance" in p.params


def test_concept_headings_are_subheads_with_a_spacer_before_all_but_the_first():
    s, res, p = run(cspec(), r=cresolver(), frame=cfacts())
    kinds = {d["row"]: d["kind"] for d in p.rows}
    heads = [r for r, k in kinds.items() if k == "subhead"]
    assert [p.cells[(r, 3)] for r in heads] == ["Revenue and COGS", "Financing"]
    assert heads[1] - 1 not in kinds and (heads[0] - 1) in kinds          # spacer before the second only
    assert all((r, c) not in p.cells for r in heads for c in range(4, 12))
