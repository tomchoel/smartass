# engine/tests/v2/test_v5_charts.py
"""Plan 5 Task 7: ChartSpec v2 from the layout (last column group, roles by column code, helper columns, series
validation) and the IBCS drawing through a COM fake that records every call. Every figure is invented."""
from test_v5_headers import OCT, ytd_frame
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.spec import AnalysisSpec
from smartass.render import charts as C
from smartass.render import layout as L
from smartass.render.chartdata import WATERFALL_ROLES, stacked_spans, waterfall_bars
from smartass.render.layout import CHART_ROWS, ChartSpec, SheetPlan, col_number, layout
from smartass.xl.vendor import tn_rapportstil as tn

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

class Rec:
    """Records attribute sets and calls; any attribute (or call result) is another Rec, cached."""
    def __init__(self, log, name):
        object.__setattr__(self, "_log", log)
        object.__setattr__(self, "_name", name)
        object.__setattr__(self, "_kids", {})

    def __getattr__(self, k):
        if k.startswith("__"):
            raise AttributeError(k)
        kids = object.__getattribute__(self, "_kids")
        if k not in kids:
            kids[k] = Rec(self._log, f"{self._name}.{k}")
        return kids[k]

    def __setattr__(self, k, v):
        self._log.append(("set", f"{self._name}.{k}", v))

    def __call__(self, *a):
        self._log.append(("call", self._name, a))
        key = f"({','.join(map(str, a))})"
        kids = object.__getattribute__(self, "_kids")
        if key not in kids:
            kids[key] = Rec(self._log, f"{self._name}{key}")
        return kids[key]


class Coll:
    def __init__(self, log, chart, auto):
        self.log, self.chart, self.n, self.made = log, chart, auto, 0

    @property
    def Count(self):
        return self.n

    def NewSeries(self):
        self.made += 1
        self.log.append(("new_series", self.chart, self.made))
        return Rec(self.log, f"{self.chart}.s{self.made}")


class FakeChart(Rec):
    def __init__(self, log, name, auto=1):
        super().__init__(log, name)
        object.__setattr__(self, "coll", Coll(log, name, auto))

    def SeriesCollection(self, i=None):
        if i is None:
            return self.coll
        coll, log = self.coll, self._log

        class Item:
            def Delete(_):
                log.append(("delete_auto",))
                coll.n -= 1
        return Item()


class FakeShapes:
    def __init__(self, log, fail_on=()):
        self.log, self.n, self.fail_on = log, 0, set(fail_on)

    def AddChart2(self, style, xl_type, left, top, w, h):
        self.n += 1
        if self.n in self.fail_on:
            raise RuntimeError("Excel said no")
        self.log.append(("add", xl_type, w, h))
        shape = Rec(self.log, f"shape{self.n}")
        object.__setattr__(shape, "Chart", FakeChart(self.log, f"c{self.n}"))
        return shape


class FakeWS(Rec):
    def __init__(self, log, fail_on=()):
        super().__init__(log, "ws")
        object.__setattr__(self, "Shapes", FakeShapes(log, fail_on))

    def Range(self, a):
        self._log.append(("range", a))
        return Rec(self._log, f"Range({a})")


def sets(log, suffix):
    return [(x[1], x[2]) for x in log if x[0] == "set" and x[1].endswith(suffix)]


def spec(kind, series, **kw):
    return ChartSpec(kind, kw.pop("title", f"{kind} chart"), cat_range="C10:C14",
                     series=[(n, f"{c}10:{c}14", r) for n, c, r in series], anchor=kw.pop("anchor", "B5"), **kw)


def draw(*specs, fail_on=(), lint=()):
    log = []
    ws = FakeWS(log, fail_on)
    out = C.write_charts(ws, SheetPlan(charts=list(specs), chart_lint=list(lint)))
    return log, out


NAVY = tn.rgb(tn.NAVY)


def test_chart_types_per_kind():
    kinds = ["column", "bar", "line", "stacked", "stacked100", "combo", "variance", "waterfall"]
    log, out = draw(*[spec(k, [("AC", "D", "ac")]) for k in kinds])
    assert out == []
    assert [e[1] for e in log if e[0] == "add"] == [51, 57, 4, 52, 53, 51, 57, 52]
    assert all((e[2], e[3]) == (480, 260) for e in log if e[0] == "add")
    names = [v for n, v in sets(log, ".Name") if n.startswith("shape")]
    assert names == ["smartass chart"] * len(kinds)


def test_series_are_added_explicitly_before_any_format():
    log, _ = draw(spec("column", [("AC", "D", "ac"), ("PY", "E", "py")]))
    assert ("delete_auto",) in log
    news = [i for i, e in enumerate(log) if e[0] == "new_series"]
    fmt = [i for i, e in enumerate(log) if e[0] == "set" and (".Format." in e[1] or "HasTitle" in e[1]
                                                               or "GapWidth" in e[1])]
    assert len(news) == 2 and max(news) < min(fmt)
    assert ("range", "D10:D14") in log and ("range", "E10:E14") in log and ("range", "C10:C14") in log
    assert [v for n, v in sets(log, ".Name") if ".s" in n] == ["AC", "PY"]


def test_ibcs_fills_navy_tint_outlined_hatched():
    log, _ = draw(spec("column", [("AC", "D", "ac"), ("PY", "E", "py"), ("BU", "F", "bu"), ("FC", "G", "fc")]))
    fore = dict(sets(log, ".Format.Fill.ForeColor.RGB"))
    assert fore["c1.s1.Format.Fill.ForeColor.RGB"] == NAVY
    assert fore["c1.s2.Format.Fill.ForeColor.RGB"] == tn.rgb(C.tint(tn.NAVY))      # PY: a lighter navy
    assert fore["c1.s3.Format.Fill.ForeColor.RGB"] == C.WHITE
    assert dict(sets(log, ".Format.Line.ForeColor.RGB"))["c1.s3.Format.Line.ForeColor.RGB"] == NAVY
    assert ("call", "c1.s4.Format.Fill.Patterned", (26,)) in log            # msoPatternWideUpwardDiagonal
    assert ("set", "c1.HasLegend", True) in log                            # 4 series: legend on top
    assert ("set", "c1.Legend.Position", C.XL_LEGEND_TOP) in log


def test_no_gridlines_and_mnok_labels():
    log, _ = draw(spec("column", [("AC", "D", "ac")]), spec("column", [("FTE", "D", "ac")], millions=False))
    assert ("set", "c1.Axes(2).HasMajorGridlines", False) in log
    assert ("set", "c1.Axes(2).MajorTickMark", C.XL_NONE) in log
    fmts = sets(log, "DataLabels().NumberFormat")
    assert fmts == [("c1.s1.DataLabels().NumberFormat", "#,##0.0,,;\"−\"#,##0.0,,;;"),
                    ("c2.s1.DataLabels().NumberFormat", "#,##0.0;\"−\"#,##0.0;;")]
    assert ("set", "c1.HasLegend", False) in log                           # one series: labelled directly
    assert ("set", "c1.ChartGroups(1).GapWidth", 50) in log
    assert ("set", "c1.ChartTitle.Text", "column chart") in log


def test_waterfall_bases_are_invisible_and_steps_coloured():
    roles = [("Base", "P", "base"), ("Base <0", "Q", "base"), ("Up", "R", "up"), ("Up <0", "S", "up_neg"),
             ("Down", "T", "down"), ("Down <0", "U", "down_neg"), ("Total", "V", "total")]
    log, _ = draw(spec("waterfall", roles))
    assert ("set", "c1.s1.Format.Fill.Visible", 0) in log and ("set", "c1.s2.Format.Fill.Visible", 0) in log
    fore = dict(sets(log, ".Format.Fill.ForeColor.RGB"))
    assert fore["c1.s3.Format.Fill.ForeColor.RGB"] == C.GOOD and fore["c1.s5.Format.Fill.ForeColor.RGB"] == C.BAD
    assert fore["c1.s7.Format.Fill.ForeColor.RGB"] == NAVY
    assert ("set", "c1.HasLegend", False) in log                           # 5 visible series, still no legend
    assert ("set", "c1.ChartGroups(1).Overlap", 100) in log
    assert not any(n.startswith("c1.s1.DataLabels") for n, _ in sets(log, "NumberFormat"))


def test_variance_overlaps_and_bar_reverses():
    log, _ = draw(spec("variance", [("ΔPY better", "P", "fav"), ("ΔPY worse", "Q", "unfav")]),
                  spec("bar", [("AC", "D", "ac")], sort="desc"), spec("bar", [("AC", "D", "ac")], sort="asc"))
    assert ("set", "c1.ChartGroups(1).Overlap", 100) in log
    fore = dict(sets(log, ".Format.Fill.ForeColor.RGB"))
    assert fore["c1.s1.Format.Fill.ForeColor.RGB"] == C.GOOD and fore["c1.s2.Format.Fill.ForeColor.RGB"] == C.BAD
    assert ("set", "c2.Axes(1).ReversePlotOrder", True) in log
    assert ("set", "c3.Axes(1).ReversePlotOrder", False) in log
    assert not any("InvertIfNegative" in n for n, _ in sets(log, "InvertIfNegative"))


def test_combo_puts_pct_on_a_secondary_line():
    log, _ = draw(spec("combo", [("AC", "D", "ac"), ("ΔPY %", "F", "pct")]))
    assert ("set", "c1.s2.ChartType", 4) in log and ("set", "c1.s2.AxisGroup", 2) in log
    i_type = log.index(("set", "c1.s2.ChartType", 4))
    i_fmt = min(i for i, e in enumerate(log) if e[0] == "set" and ".Format." in e[1])
    assert i_type < i_fmt
    assert ("set", "c1.s2.DataLabels().NumberFormat", '0.0 %;"−"0.0 %;;') in log
    assert ("set", "c1.s2.Format.Line.ForeColor.RGB", C.PCT_LINE) in log


def test_highlight_colours_points_accent_and_the_rest_grey():
    log, _ = draw(spec("bar", [("AC", "D", "ac")], highlight=[2]))
    pts = dict(sets(log, ".Format.Fill.ForeColor.RGB"))
    assert pts["c1.s1.Points(2).Format.Fill.ForeColor.RGB"] == C.ACCENT
    assert pts["c1.s1.Points(1).Format.Fill.ForeColor.RGB"] == C.GREY
    assert "c1.s1.Points(5).Format.Fill.ForeColor.RGB" in pts


def test_a_failing_chart_is_one_lint_line_and_the_others_are_drawn():
    log, out = draw(spec("column", [("AC", "D", "ac")], title="First"),
                    spec("bar", [("AC", "D", "ac")], title="Second"), fail_on=(1,), lint=["chart 'X': layout note"])
    assert out[0] == "chart 'X': layout note"
    (msg,) = out[1:]
    assert msg.startswith("chart 'First' failed: Excel said no; fix:")
    assert [e[1] for e in log if e[0] == "add"] == [57]


def test_a_chart_without_series_is_not_drawn():
    log, out = draw(ChartSpec("column", "Empty", anchor="B5"))
    assert out == [] and not any(e[0] == "add" for e in log)


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


def test_waterfall_labels_show_the_step_sign_above_and_below_zero():
    roles = [("Base", "P", "base"), ("Base <0", "Q", "base"), ("Up", "R", "up"), ("Up <0", "S", "up_neg"),
             ("Down", "T", "down"), ("Down <0", "U", "down_neg"), ("Total", "V", "total")]
    log, _ = draw(spec("waterfall", roles))
    fmts = dict(sets(log, "DataLabels().NumberFormat"))
    assert fmts == {"c1.s3.DataLabels().NumberFormat": "#,##0.0,,;#,##0.0,,;;",      # +1 above zero: 1.0
                    "c1.s4.DataLabels().NumberFormat": "#,##0.0,,;#,##0.0,,;;",      # +1 below zero: 1.0
                    "c1.s5.DataLabels().NumberFormat": "\"−\"#,##0.0,,;\"−\"#,##0.0,,;;",    # -8 above zero: -8.0
                    "c1.s6.DataLabels().NumberFormat": "\"−\"#,##0.0,,;\"−\"#,##0.0,,;;",    # -8 below zero: -8.0
                    "c1.s7.DataLabels().NumberFormat": "#,##0.0,,;\"−\"#,##0.0,,;;"}
    fore = dict(sets(log, ".Format.Fill.ForeColor.RGB"))
    assert fore["c1.s4.Format.Fill.ForeColor.RGB"] == C.GOOD and fore["c1.s6.Format.Fill.ForeColor.RGB"] == C.BAD


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
    log, _ = draw(spec("waterfall", [("Base", "P", "base"), ("Up", "R", "up")], highlight=[1]))
    assert not any("Points(" in x[1] for x in log if x[0] == "set")
    s, res, p = run([GROUPED, chart("variance", 0, title="V", highlight=["Opex"])])
    assert any("highlight is drawn on column and bar charts only" in m for m in p.chart_lint)
    assert next(c for c in p.charts if c.title == "V").highlight == []


def test_stacked_parts_get_a_legend():
    log, _ = draw(spec("stacked", [("A", "D", "part"), ("B", "E", "part")]))
    assert ("set", "c1.HasLegend", True) in log


def test_data_labels_get_the_chart_locale_format(monkeypatch):
    """The DataLabels object with a real International (nb-NO): the label format goes through chart_format."""
    from types import SimpleNamespace
    log = []
    ws = FakeWS(log)
    nb = SimpleNamespace(International=tuple([None, None, ",", " "] + [None] * 40))
    orig = Rec.__call__

    def call(self, *a):
        r = orig(self, *a)
        if self._name.endswith(".DataLabels"):
            object.__setattr__(r, "Application", nb)
        return r
    monkeypatch.setattr(Rec, "__call__", call)
    C.write_charts(ws, SheetPlan(charts=[spec("column", [("AC", "D", "ac")])]))
    fmts = [v for _, v in sets(log, "DataLabels().NumberFormat")]
    assert fmts == ['#\xa0##0,0\xa0\xa0;"−"#\xa0##0,0\xa0\xa0;;']


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
