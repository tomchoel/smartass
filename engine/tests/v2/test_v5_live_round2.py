# engine/tests/v2/test_v5_live_round2.py
"""Live round 2 of plan 5 (rendered answers to PDF, 05.11.2026): readable bridges (truncated axis, white labels),
category labels clear of negative bars, units by magnitude, key points for a lines-only answer, charts inside the
print area, codes not repeated in labels, no FinPlan metric codes, Lines sheets at 100 %. Every figure is
invented."""
from types import SimpleNamespace

from test_v5_charts import draw, sets, spec
from test_v5_keypoints import HEADERS, COLS, _line, addr
from test_v5_layout import BRIDGE, GROUPED, run
from v2util import CLOSED, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.blocks.common import BlockOut, Line, ck, metric_code
from smartass.model.findings import key_values
from smartass.model.frame import Frame
from smartass.model.lines import LineSet
from smartass.model.spec import AnalysisSpec
from smartass.render import charts as C
from smartass.render.keypoints import _columns, key_points
from smartass.render.layout import (LABEL_COL, TEXT_COL, ChartSpec, _axis_notes, _scale, amount_text,
                                    bridge_axis, col_letter, col_number, layout, nice_floor, sheet_label, unit_of,
                                    width_pt)


# --- 1. the bridge axis -------------------------------------------------------------------------------------------

def test_nice_floor_keeps_two_significant_digits():
    assert nice_floor(3_330_000_000) == 3_300_000_000
    assert nice_floor(40_950_000) == 40_000_000
    assert nice_floor(990.0) == 990.0


def test_small_steps_truncate_the_axis_below_every_level():
    amounts = [3.80e9, 0.12e9, -0.09e9, -0.06e9, 3.77e9]       # steps < 25 % of the smallest total
    axis = bridge_axis(amounts)
    assert axis == 3.3e9                                         # 90 % of the lowest level 3.77 bn, floored
    assert axis < min([3.80e9, 3.92e9, 3.83e9, 3.77e9])


def test_large_steps_or_a_zero_crossing_keep_the_axis_at_zero():
    assert bridge_axis([100.0, 40.0, -20.0, 120.0]) is None     # a step of 40 % reads from zero
    assert bridge_axis([10.0, 1.0, -12.0, 1.0]) is None          # levels cross zero
    assert bridge_axis([100.0, 100.0]) is None                   # no steps


def test_a_bridge_below_zero_truncates_its_maximum():
    axis = bridge_axis([-3.8e9, 0.1e9, -3.7e9])
    assert axis == -3.3e9                                       # MaximumScale, above every (negative) level


def test_the_note_names_the_axis_start():
    data = BlockOut("b0", "bridge", "EBITDA bridge", ["Amount"], cols=["amount"],
                    lines=[Line(k, k, [None]) for k in ("start", "s0", "s1", "end")])
    res = SimpleNamespace(cells={"b0.start.0": 3.8e9, "b0.s0.0": 0.1e9, "b0.s1.0": -0.05e9, "b0.end.0": 3.85e9})
    o = SimpleNamespace(type="bridge", key="b0", chart="waterfall", chart_of=None, title="EBITDA bridge")
    assert _axis_notes([o], {"b0": data}, res, set()) == ["Bridge axis starts at 3 400 MNOK."]
    assert amount_text(45_000.0, False) == "45 000 NOK" and amount_text(1.5e6, True) == "1.5 MNOK"


def test_a_truncated_waterfall_sets_the_scale_and_white_labels():
    roles = [("Base", "P", "base"), ("Up", "R", "up"), ("Down", "T", "down"), ("Total", "V", "total")]
    log, _ = draw(spec("waterfall", roles, axis_min=3.3e9), spec("waterfall", roles, axis_min=-3.3e9))
    assert ("set", "c1.Axes(2).MinimumScale", 3.3e9) in log
    assert ("set", "c2.Axes(2).MaximumScale", -3.3e9) in log
    pos = dict(sets(log, "DataLabels().Position"))
    assert pos["c1.s4.DataLabels().Position"] == C.XL_LABEL_INSIDE_END                 # total: inside end
    assert pos["c1.s2.DataLabels().Position"] == C.XL_LABEL_CENTER == pos["c1.s3.DataLabels().Position"]
    colours = dict(sets(log, "DataLabels().Font.Color"))
    assert {colours[f"c1.s{i}.DataLabels().Font.Color"] for i in (2, 3, 4)} == {C.WHITE}   # never navy on navy
    assert ("set", "c1.s4.DataLabels().Font.Bold", True) in log
    assert not any(n.startswith("c1.s1.DataLabels") for n, _ in sets(log, "Position"))  # the base has none


def test_layout_waterfall_gets_the_axis_and_the_notes_line():
    s, res, p = run([GROUPED, BRIDGE], scale=1e6)
    (cs,) = [c for c in p.charts if c.kind == "waterfall"]
    bridge = next(o for o in res.outs if o.type == "bridge")
    amounts = [res.cells.get(ck(bridge.key, ln.key, 0)) for ln in bridge.lines]
    assert cs.axis_min == bridge_axis(amounts)
    notes = [v for (r, c), v in p.cells.items() if c == TEXT_COL and isinstance(v, str) and "Bridge axis" in v]
    assert len(notes) == (1 if cs.axis_min is not None else 0)


def test_a_thin_step_gets_a_dark_label():
    data = BlockOut("b0", "bridge", "EBITDA bridge", ["Amount"], cols=["amount"],
                    lines=[Line(k, k, [None]) for k in ("start", "s0", "s1", "other", "end")])
    res = SimpleNamespace(cells={"b0.start.0": 3.8e9, "b0.s0.0": 0.12e9, "b0.s1.0": -0.09e9, "b0.other.0": -3e6,
                                 "b0.end.0": 3.827e9})
    cs = ChartSpec("waterfall", "Bridge")
    _scale(cs, data, res, [(20 + i, ln) for i, ln in enumerate(data.lines)], [0])
    assert cs.axis_min == 3.4e9 and cs.thin_steps == [4]                  # 3 m on a ~600 m span
    roles = [("Base", "P", "base"), ("Up", "R", "up"), ("Down", "T", "down"), ("Total", "V", "total")]
    log, _ = draw(spec("waterfall", roles, axis_min=3.4e9, thin_steps=[4]))
    colours = dict(sets(log, "DataLabel.Font.Color"))
    assert colours["c1.s2.Points(4).DataLabel.Font.Color"] == C.INK == colours["c1.s3.Points(4).DataLabel.Font.Color"]
    assert not any(".s4.Points" in n for n in colours)                     # totals stay white


# --- 2. category labels and all-negative charts -------------------------------------------------------------------

def test_category_labels_sit_low_and_high_on_a_reversed_axis():
    log, _ = draw(spec("column", [("AC", "D", "ac")]), spec("bar", [("AC", "D", "ac")], reverse_values=True),
                  spec("line", [("AC", "D", "ac")], reverse_values=True))
    assert ("set", "c1.Axes(1).TickLabelPosition", C.XL_TICK_LOW) in log
    assert not any(n == "c1.Axes(2).ReversePlotOrder" for n, _ in sets(log, "ReversePlotOrder"))
    for c in ("c2", "c3"):
        assert ("set", f"{c}.Axes(2).ReversePlotOrder", True) in log
        assert ("set", f"{c}.Axes(1).TickLabelPosition", C.XL_TICK_HIGH) in log   # next to zero (hidden Excel)


def _out(cols, values):
    lines = [Line(f"l{i}", f"L{i}", [None] * len(cols)) for i in range(len(values))]
    data = BlockOut("b0", "table", "T", list(cols), cols=list(cols), lines=lines)
    cells = {ck("b0", f"l{i}", j): v for i, row in enumerate(values) for j, v in enumerate(row)}
    return data, SimpleNamespace(cells=cells), [(10 + i, ln) for i, ln in enumerate(lines)]


def test_costs_reverse_the_value_axis_and_mixed_values_do_not():
    data, res, rows = _out(["value", "ly"], [[-120e6, -100e6], [-95e6, 0.0]])
    cs = ChartSpec("column", "Costs")
    _scale(cs, data, res, rows, [0, 1])
    assert cs.reverse_values is True and cs.millions is True
    data, res, rows = _out(["value", "ly"], [[-120e6, 5e6], [-95e6, 0.0]])
    cs = ChartSpec("column", "Mixed")
    _scale(cs, data, res, rows, [0, 1])
    assert cs.reverse_values is False


# --- 3. units by magnitude ----------------------------------------------------------------------------------------

def test_small_amounts_are_whole_nok_on_charts():
    data, res, rows = _out(["value"], [[-7_271.0], [-3_000.0]])
    cs = ChartSpec("bar", "Suppliers")
    _scale(cs, data, res, rows, [0])
    assert (cs.millions, cs.whole) == (False, True)
    assert C.label_format("ac", False, True) == "#,##0;\"−\"#,##0;;"
    log, _ = draw(spec("bar", [("AC", "D", "ac")], millions=False, whole=True))
    assert sets(log, "DataLabels().NumberFormat") == [("c1.s1.DataLabels().NumberFormat", "#,##0;\"−\"#,##0;;")]


def test_fte_charts_keep_one_decimal():
    data, res, rows = _out(["value"], [[12.5], [3.0]])
    cs = ChartSpec("column", "FTE", millions=False)
    _scale(cs, data, res, rows, [0])
    assert (cs.millions, cs.whole) == (False, False)


def test_key_points_say_nok_under_one_million():
    assert unit_of(7_271.0) == "NOK" and unit_of(-2e6) == "MNOK" and unit_of(None) == "MNOK"
    main = BlockOut("b0", "table", "P&L", HEADERS, cols=list(COLS), lines=[_line("m0", "Revenue", kind="grand")])
    vi = _columns(main)[0]
    a = addr(f"b0.m0.{vi}")
    small = SimpleNamespace(material=[], cells={f"b0.m0.{vi}": -7_271.0})
    (f, _), = key_points(main, small, [], addr)
    assert f"FIXED({a},0)" in f and '" NOK' in f and "MNOK" not in f and "$D$40" not in f
    big = SimpleNamespace(material=[], cells={f"b0.m0.{vi}": -7.2e6})
    (f, _), = key_points(main, big, [], addr)
    assert f"FIXED({a}/$D$40,1)" in f and " MNOK" in f


# --- 4. key points of a lines-only answer -------------------------------------------------------------------------

ROWS = [{"AP_CUSTOMER_NAME": "Accenture AS", "AMOUNT": -2_000_000.0},
        {"AP_CUSTOMER_NAME": "Bouvet", "AMOUNT": -30_000.0},
        {"AP_CUSTOMER_NAME": "Sopra Steria", "AMOUNT": -7_271.0},
        {"AP_CUSTOMER_NAME": "Atea", "AMOUNT": -1_000.0},
        {"AP_CUSTOMER_NAME": "", "AMOUNT": -5.0}]
CELL = {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Yr 26"}


def _lines_answer():
    d = spec_dict(blocks=[{"type": "lines", "cell": CELL, "group_by": "supplier", "top": 4}])
    s = AnalysisSpec.model_validate(d)
    fr = Frame([])
    fr.lines["b0"] = LineSet("b0", ROWS, ["AP_CUSTOMER_NAME", "AMOUNT"], cube_value=-2_038_276.0)
    res = compute(s, resolver(), CLOSED, fr)
    return s, res, layout(3, s, res, stock=set(), subtitle="SUB", key_points=key_points)


def test_a_lines_only_answer_gets_its_top_three_and_the_total():
    _, res, p = _lines_answer()
    r = dict(p.zones)["key_points"]
    pts = [p.cells.get((r + i, TEXT_COL)) for i in range(1, 6)]
    assert all(isinstance(x, str) and x.startswith('="•  ') for x in pts[:4]) and pts[4] is None
    g0, g1, g2 = (p.addr[f"b0.g{i}.0"] for i in range(3))
    assert pts[0].startswith('="•  Accenture AS: "&SUBSTITUTE(FIXED(' + g0 + "/") and '" MNOK (' in pts[0]
    # one unit for every point (5c): MNOK, two decimals under 50 000 (never '0,0 MNOK'), never NOK next to MNOK
    assert pts[1].startswith('="•  Bouvet: "&SUBSTITUTE(FIXED(' + g1 + "/") and ",2)" in pts[1] and '" MNOK (' in pts[1]
    assert pts[2].startswith('="•  Sopra Steria: "&SUBSTITUTE(FIXED(' + g2 + "/") and " NOK" not in pts[2].replace(" MNOK", "")
    assert p.addr["b0.g0.1"] in pts[0] and '" %"' in pts[0]                     # the share, live
    tot, gap = p.addr["b0.total.0"], p.addr["b0.gap.0"]
    assert pts[3].startswith('="•  Ledger lines total "&SUBSTITUTE(FIXED(' + tot + "/")
    assert "; gap to cube " in pts[3] and f"FIXED({gap},0)" in pts[3]          # 1 NOK gap: whole NOK


# --- 5. charts inside the print area ------------------------------------------------------------------------------

def test_the_print_area_reaches_the_right_edge_of_every_chart():
    blocks = [GROUPED, BRIDGE, {"type": "trend", "query": "fm", "metrics": ["Opex"]}]
    _, _, p = run(blocks)
    left, x = {}, 0.0
    for c in range(1, 60):
        left[c] = x
        x += width_pt(p.widths.get(col_letter(c), 8.43))
    last = col_number(p.print_area_col)
    for cs in p.charts:
        assert left[last + 1] >= left[col_number(cs.anchor_col)] + cs.width          # not cut off
    assert last >= col_number(p.print_last_col)
    from smartass.render.excel import answer_roles
    p.footer = "F"
    assert answer_roles(p)["print_area_col"] == p.print_area_col


def test_the_print_area_never_takes_in_the_chart_helpers():
    _, _, p = run([BRIDGE, {"type": "trend", "query": "fm", "metrics": ["Opex"]}])      # narrow: no wide table
    assert p.helper_cols and [c.kind for c in p.charts] == ["waterfall", "line"]
    assert col_number(p.print_area_col) < min(col_number(c) for c in p.helper_cols)
    left, x = {}, 0.0
    for c in range(1, 60):
        left[c] = x
        x += width_pt(p.widths.get(col_letter(c), 8.43))
    for cs in p.charts:
        assert left[col_number(p.print_area_col) + 1] >= left[col_number(cs.anchor_col)] + cs.width


def test_a_variance_chart_never_reverses_its_axis():
    data, res, rows = _out(["var"], [[-5e6], [-2e6]])
    cs = ChartSpec("variance", "All worse")
    _scale(cs, data, res, rows, [0])
    assert cs.reverse_values is False


# --- 6. codes not repeated in labels ------------------------------------------------------------------------------

def test_sheet_label_drops_a_repeated_code():
    assert sheet_label(Line("e0", "4P&C_T People & Competence", [], code="4P&C_T")) == "People & Competence"
    assert sheet_label(Line("e0", "2WHS Wholesale", [], code="2WHS")) == "Wholesale"
    assert sheet_label(Line("e0", "Revenue", [], code="C30")) == "Revenue"
    assert sheet_label(Line("e0", "2WHS", [], code="2WHS")) == "2WHS"             # nothing left: kept


def test_the_sheet_strips_the_code_and_key_values_keep_the_model_label():
    s, res, _ = run([GROUPED])
    o = res.outs[0]
    ln = next(x for x in o.lines if x.code)
    ln.label = f"{ln.code} Wholesale"
    p = layout(3, s, res, stock=set(), subtitle="SUB")
    row = p.at(f"{o.key}.{ln.key}.{o.cols.index('value')}")[0]
    assert p.cells[(row, LABEL_COL)].strip() == "Wholesale"
    assert f"{o.title} · {ln.code} Wholesale" in key_values(res.outs, res.cells)


# --- 7. FinPlan metrics have no code ------------------------------------------------------------------------------

def test_finplan_metric_rows_have_no_code():
    r = SimpleNamespace(members=lambda metric, cube: ["Revenue"] if cube == "FinPlan" else ["ac.C30"])
    assert metric_code(r, "Revenue", "FinPlan") == ""
    assert metric_code(r, "Revenue", "FinMax") == "C30"


# --- 8. Lines sheets: see test_v5_lines_header (100 %, header row repeated) ----------------------------------------

def test_print_setup_without_fit_sets_zoom_and_title_rows():
    from test_v5_writer import _ws

    from smartass.xl import styling
    ws = _ws()
    styling.print_setup(ws, "AR", 900, "F", fit_wide=False, title_rows="$4:$4")
    ps = ws.PageSetup
    assert (ps.Orientation, ps.Zoom, ps.PrintTitleRows) == (2, 100, "$4:$4")
    assert ps.PrintArea == "$A$1:$AR$900"

