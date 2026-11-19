# engine/tests/v2/test_v5_charts.py
"""Plan 5 Task 7: ChartSpec v2 from the layout (last column group, roles by column code, helper columns, series
validation); the drawing itself is in test_xlsx_charts.py. Every figure is invented."""
from test_v5_headers import OCT, ytd_frame
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.spec import AnalysisSpec
from smartass.render import layout as L
from smartass.render.chartdata import WATERFALL_ROLES, stacked_spans, waterfall_bars
from smartass.render.layout import CHART_ROWS, ChartSpec, SheetPlan, col_number, layout

GROUPED = {"type": "table", "query": "fm", "title": "P&L", "periods": ["closed", "closed-ytd"]}
BRIDGE = {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}


def run(blocks, periods="closed-ytd", frame=None, closed=OCT):
    d = spec_dict(blocks=blocks)
    d["queries"][0]["periods"] = periods
    s = AnalysisSpec.model_validate(d)
    res = compute(s, resolver(), closed, frame or ytd_frame())
    return s, res, layout(3, s, res, stock=set(), subtitle="SUB")


def chart(kind, block, **kw):
    return {"type": "chart", "kind": kind, "title": kw.pop("title", f"{kind} chart"), "source": {"block": block},
            **kw}


# --- the waterfall maths ------------------------------------------------------------------------------------------

def _rebuild(amounts):
    """The visible bar of every point as Excel stacks it: a total from 0, a step from the previous end to the new."""
    return [stacked_spans(b, WATERFALL_ROLES) for b in waterfall_bars(amounts)]


def test_waterfall_bars_rebuild_the_running_total():
    spans = _rebuild([10, 5, -8, 2, 9])
    assert spans == [(0, 10), (10, 15), (7, 15), (7, 9), (0, 9)]


def test_waterfall_bars_cross_zero():
    spans = _rebuild([3, -8, 1, -4])          # 3 -> -5 crosses zero, -5 -> -4 stays below
    assert spans == [(0, 3), (-5, 3), (-5, -4), (-4, 0)]
    bars = waterfall_bars([3, -8, 1, -4])
    assert bars[1][5:7] == [3, -5]            # the decrease split at zero: 3 above, -5 below
    assert bars[2][2] == -4 and bars[2][4] == -1   # below-zero base, the increase drawn below zero


def test_waterfall_helper_formulas_point_at_the_amount_and_the_previous_end():
    from smartass.render.chartdata import waterfall_helpers
    cols = list("PQRSTUVW")
    cells = waterfall_helpers([10, 11, 12], "D", cols)
    assert cells[(10, "P")] == "=D10" and cells[(10, "W")] == "=D10" and cells[(10, "Q")] == 0
    assert cells[(11, "P")] == "=P10+D11"
    assert cells[(11, "Q")] == "=MAX(MIN(P10,P11),0)" and cells[(11, "R")] == "=MIN(MAX(P10,P11),0)"
    assert cells[(11, "S")] == "=IF(D11>=0,MAX(MAX(P10,P11),0)-MAX(MIN(P10,P11),0),0)"
    assert cells[(12, "W")] == "=D12" and cells[(12, "P")] == "=D12"


# --- ChartSpec v2 from the layout ---------------------------------------------------------------------------------

def test_bridge_waterfall_uses_helper_columns_right_of_the_widest_table():
    s, res, p = run([GROUPED, BRIDGE])
    (cs,) = [c for c in p.charts if c.kind == "waterfall"]
    roles = [r for _, _, r in cs.series]
    assert roles == ["base", "base", "up", "up_neg", "down", "down_neg", "total"]
    bridge = next(o for o in res.outs if o.type == "bridge")
    first_line = min(int(a[1:]) for k, a in p.addr.items() if k.startswith(bridge.key + "."))
    j = bridge.cols.index("amount")       # zero steps are no category (start and end always are)
    rows = [first_line + i for i, ln in enumerate(bridge.lines)
            if ln.kind in ("sum", "grand") or abs(res.cells.get(L.ck(bridge.key, ln.key, j)) or 0) >= 0.5]
    assert cs.cat_range == L._runs("C", rows) and len(rows) < len(bridge.lines)
    widest = max(col_number("".join(ch for ch in a if ch.isalpha())) for a in p.addr.values())
    helper_cols = [col_number(c) for c in p.helper_cols]
    assert len(helper_cols) == 8 and min(helper_cols) >= widest + 2       # one empty column between
    assert p.last_col == p.helper_cols[-1] and col_number(p.print_last_col) < min(helper_cols)
    assert p.cells[(first_line - 2, min(helper_cols))] == "Chart helper"
    assert all(v for v in p.addr.values() if col_number("".join(ch for ch in v if ch.isalpha())) in helper_cols) \
        is True
    assert not any(col_number("".join(ch for ch in v if ch.isalpha())) in helper_cols for v in p.addr.values())
    run_col = p.helper_cols[0]
    assert p.cells[(first_line + 1, min(helper_cols))] == f"={run_col}{first_line}+D{first_line + 1}"
    for (r, c), v in cs.helpers.items():
        assert p.cells[(r, c)] == v
    zone = dict(p.zones)
    assert all(r >= zone["tables"] for r, _ in cs.helpers)                # never in the Charts zone


def test_column_chart_takes_the_last_column_group():
    from test_v5_layout import scaled
    s, res, p = run([GROUPED, chart("column", 0, title="Opex up")], frame=scaled(ytd_frame(), 1e6))
    cs = next(c for c in p.charts if c.title == "Opex up")
    assert [r for _, _, r in cs.series] == ["ac", "py"]
    table = next(o for o in res.outs if o.key == "b0")
    gap = table.cols.index("gap")
    first_val = "".join(ch for ch in cs.series[0][1].split(":")[0] if ch.isalpha())
    assert col_number(first_val) > 4 + gap                                # right of the spacer: the YTD group
    assert cs.series[0][0].startswith("AC ") and cs.series[1][0].startswith("PY ")
    assert cs.millions and cs.sort == "none" and cs.anchor_col == "B"


def test_series_roles_follow_the_column_codes():
    s, res, p = run([GROUPED, chart("combo", 0, title="Mix", series=["value", "var", "var_pct"])])
    cs = next(c for c in p.charts if c.title == "Mix")
    assert [r for _, _, r in cs.series] == ["ac", "var"]       # var % is blank under the floor: no data
    assert any("has no data; left out" in x for x in p.chart_lint)


def test_a_missing_series_is_dropped_with_a_lint_line():
    s, res, p = run([GROUPED, chart("column", 0, title="AC vs BU", series=["value", "budget"])])
    cs = next(c for c in p.charts if c.title == "AC vs BU")
    assert [r for _, _, r in cs.series] == ["ac"]
    (msg,) = [m for m in p.chart_lint if "AC vs BU" in m]
    assert "series 'budget' is not in the source" in msg and "fix:" in msg


def test_no_usable_series_skips_the_chart_with_one_lint_line():
    s, res, p = run([GROUPED, chart("column", 0, title="Plan only", series=["budget", "forecast"])])
    cs = next(c for c in p.charts if c.title == "Plan only")
    assert cs.series == []
    assert len([m for m in p.chart_lint if "Plan only" in m and "no series" in m]) == 1


def test_variance_chart_gets_fav_and_unfav_helpers():
    s, res, p = run([GROUPED, chart("variance", 0, title="Where it moved")])
    cs = next(c for c in p.charts if c.title == "Where it moved")
    assert [r for _, _, r in cs.series] == ["fav", "unfav"]
    assert len(p.helper_cols) == 2
    assert any(isinstance(v, str) and v.startswith("=MAX(") for v in cs.helpers.values())


def test_rows_exclude_grand_and_check_lines():
    trend = {"type": "trend", "query": "fm", "metrics": ["Opex"]}
    from v2util import monthly_facts
    s, res, p = run([trend], periods="Jul 26..Sep 26", frame=monthly_facts(), closed=CLOSED)
    (cs,) = p.charts
    assert cs.kind == "line" and len(cs.series) == 2
    a, b = cs.cat_range.split(":")
    assert int(b[1:]) - int(a[1:]) == 2                                   # three months, no Total


def test_implicit_chart_on_a_ranking_is_a_sorted_bar_without_the_total():
    rank = {"type": "ranking", "query": "fm", "metric": "Opex", "top": 3}
    s, res, p = run([rank], periods="closed", frame=facts(), closed=CLOSED)
    (cs,) = p.charts
    assert cs.kind == "bar" and cs.sort == "desc" and [r for _, _, r in cs.series] == ["ac"]
    o = next(o for o in res.outs if o.type == "ranking")
    grand = [ln for ln in o.lines if ln.kind in ("grand", "check")]
    assert grand and len(cs.cat_range.split(",")) == 1
    a, b = cs.cat_range.split(":")
    assert int(b[1:]) - int(a[1:]) + 1 == len(o.lines) - len(grand)
    assert p.chart_lint == []


def test_implicit_chart_on_kpi_tiles_is_a_column_of_value_against_its_base():
    kpi = {"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}
    s, res, p = run([kpi], periods="closed", frame=facts(), closed=CLOSED)
    (cs,) = p.charts
    assert cs.kind == "column" and [r for _, _, r in cs.series] == ["ac", "py"]


def test_highlight_maps_labels_to_point_indexes():
    rank = {"type": "ranking", "query": "fm", "metric": "Opex", "top": 3}
    s, res, p = run([rank, {"type": "chart", "kind": "bar", "title": "Top", "source": {"block": 0}}],
                    periods="closed", frame=facts(), closed=CLOSED)
    o = next(o for o in res.outs if o.type == "ranking")
    label = o.lines[1].label
    s, res, p = run([rank, {"type": "chart", "kind": "bar", "title": "Top", "source": {"block": 0},
                            "highlight": [label, "Nobody"]}], periods="closed", frame=facts(), closed=CLOSED)
    cs = next(c for c in p.charts if c.title == "Top")
    assert cs.highlight == [2]
    assert any("highlight 'Nobody'" in m for m in p.chart_lint)


def test_two_charts_side_by_side_and_a_third_below():
    blocks = [GROUPED, BRIDGE, chart("column", 0, title="A"), chart("variance", 0, title="B")]
    s, res, p = run(blocks)
    a, b, c = p.charts
    top = dict(p.zones)["charts"]
    assert a.anchor_row == b.anchor_row == top and c.anchor_row == top + CHART_ROWS
    assert a.anchor_col == c.anchor_col == "B" and b.anchor_col not in ("B", "C")
    assert (a.width, a.height) == (480, 260) and a.height < CHART_ROWS * 15


# --- drawing through the COM fake ---------------------------------------------------------------------------------

def test_implicit_bar_on_drivers_shows_the_variance_in_tree_order_without_lint():
    from smartass.render.layout import _chart
    from smartass.model.blocks.common import BlockOut, Line

    o = BlockOut("chart0", "chart", "Drivers", [], chart="bar", chart_of="b0")
    data = BlockOut("b0", "drivers", "Drivers", ["AC", "PY", "ΔPY", "ΔPY %"], cols=["value", "ly", "var", "var_pct"])
    data.lines = [Line("a", "A", [None] * 4), Line("b", "B", [None] * 4)]
    lint = []

    class Res:
        cells = {"b0.a.2": -1.0, "b0.b.2": 5.0}
    cs = _chart(o, AnalysisSpec.model_validate(spec_dict()), Res(), {"b0": data},
                {"b0": (9, [(10, data.lines[0]), (11, data.lines[1])], ["D", "E", "F", "G"])}, set(), lint)
    assert cs.sort == "none" and lint == []
    assert cs.series == [("ΔPY", "F10:F11", "var")]


def test_variance_helper_series_follow_the_category_runs():
    from smartass.render.layout import _place_helpers
    p = SheetPlan(rows=[{"row": r, "kind": "line"} for r in (9, 10, 11, 12, 13)])
    p.rows[0]["kind"] = "header"
    cs = ChartSpec("variance", "V", cat_range="C10:C11,C13", helpers={
        "kind": "variance", "rows": [10, 11, 13], "col": "F", "header_row": 9, "name": "ΔPY"})
    _place_helpers(p, [cs], 9, {})
    assert [rng for _, rng, _ in cs.series] == ["I10:I11,I13", "J10:J11,J13"]
    assert all(rng.replace("I", "C").replace("J", "C") == cs.cat_range for _, rng, _ in cs.series)


def test_highlight_only_on_column_and_bar():
    s, res, p = run([GROUPED, chart("variance", 0, title="V", highlight=["Opex"])])
    assert any("highlight is drawn on column and bar charts only" in m for m in p.chart_lint)
    assert next(c for c in p.charts if c.title == "V").highlight == []


def test_a_chart_block_on_a_bridge_replaces_the_bridges_own_waterfall():
    """Live 05.11.2026: a bridge (chart true by default) plus a waterfall chart block on it drew two identical
    waterfalls. The chart block wins (its title is the message); the bridge's sugar is not drawn."""
    _, _, p = run([GROUPED, BRIDGE, chart("waterfall", 1, title="EBITDA up on revenue")])
    assert [(c.kind, c.title) for c in p.charts] == [("waterfall", "EBITDA up on revenue")]
    assert len(p.helper_cols) == 8                                        # one set of helper columns


def test_a_chart_block_on_a_trend_replaces_the_trends_own_line():
    _, _, p = run([{"type": "trend", "query": "fm", "metrics": ["Opex"], "chart": True},
                   chart("column", 0, title="Opex by month")], periods="closed-ytd")
    assert [c.title for c in p.charts] == ["Opex by month"]


def test_bridge_sugar_alone_still_draws():
    _, _, p = run([GROUPED, BRIDGE])
    assert [c.kind for c in p.charts] == ["waterfall"]
