"""Plan 5d: the visual audit of every test-set answer type (06.11.2026) - thresholds and parameters in their own
tables with no empty column, notes at the table width, text negatives with the typographic minus, a bridge residual
in whole kroner, no zero categories in charts, drivers by code with green/red bars of the leaf drivers, truthful
titles and conventions, a concept's Data header that counts its derived query ids. Every figure is invented."""
from functools import partial
from types import SimpleNamespace

from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.blocks.common import BlockOut, Line, ck
from smartass.model.blocks.drivers import leaf_keys
from smartass.model.expr import Num, Round
from smartass.model.spec import AnalysisSpec
from smartass.render import layout as L
from smartass.render.datasheet import _of
from smartass.render.keypoints import key_points


def run(blocks, frame=None):
    s = AnalysisSpec.model_validate(spec_dict(blocks=blocks))
    res = compute(s, resolver(), CLOSED, frame or facts())
    return s, res, L.layout(1, s, res, stock=set(), subtitle="x", key_points=partial(key_points, stock=set()))


def _row_of(p, label):
    return next(r for (r, c), v in p.cells.items() if c == L.LABEL_COL and v == label)


# --- thresholds and parameters ------------------------------------------------------------------------------

def test_thresholds_table_has_only_columns_with_values_and_parameters_get_their_own_table():
    s, res, p = run([{"type": "table", "query": "fm", "title": "P&L"}])
    head = _row_of(p, "Metric")
    cols = [c for (r, c), v in p.cells.items() if r == head and c >= L.FIRST_COL and v]
    assert cols == list(range(L.FIRST_COL, L.FIRST_COL + len(cols)))          # no gap between the headers
    body = [r for r in range(head + 1, head + 20) if isinstance(p.cells.get((r, L.LABEL_COL)), str)
            and p.cells[(r, L.LABEL_COL)].endswith("(NOK)")]
    for c in cols:                                                           # every drawn column has a value
        assert any(isinstance(p.cells.get((r, c)), (int, float)) for r in body), p.cells[(head, c)]
    if "unit_mnok" in p.params or "hundred" in p.params:
        prow = _row_of(p, "Parameter")
        assert prow > head and p.cells[(prow, L.FIRST_COL)] == "Value"
        assert not any(isinstance(v, str) and v.startswith("Unit:") for v in p.cells.values())


def test_notes_run_the_width_of_the_tables():
    s, res, p = run([{"type": "table", "query": "fm", "title": "P&L"}])
    notes = [p.cells[(r["row"], L.TEXT_COL)] for r in p.rows if r["kind"] == "note"]
    assert notes and max(len(n) for n in notes) > L.NOTE_WIDTH


# --- text negatives and the bridge residual ------------------------------------------------------------------

def test_key_point_numbers_use_the_typographic_minus():
    s, res, p = run([{"type": "table", "query": "fm", "title": "P&L"}])
    pts = [v for v in p.cells.values() if isinstance(v, str) and v.startswith('="•')]
    assert pts and all('SUBSTITUTE(FIXED(' in f and '"-","−")' in f for f in pts)


def test_round_is_whole_kroner_half_away_from_zero_like_excel():
    class C:  # a constant expression
        def __init__(self, v): self.v = v
        def evaluate(self, frame, cells): return self.v
        def to_excel(self, t, a, const=None): return "X"
    assert [Round(C(v)).evaluate(None, None) for v in (-0.3, 0.3, 2.5, -2.5, None)] == [0.0, 0.0, 3.0, -3.0, None]
    assert Round(C(1)).to_excel("t", str) == "ROUND(X,0)"
    assert str(Round(C(-0.3)).evaluate(None, None)) == "0.0"             # never -0.0 (it would print '(0)')


def test_bridge_residual_is_rounded():
    s, res, p = run([{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}])
    other = p.at("b0.other.0")
    assert p.cells[other].startswith("=ROUND(")


# --- charts: no zero categories ---------------------------------------------------------------------------------

def _data(vals, kind="line"):
    lines = [Line(f"l{i}", f"L{i}", [None], kind=kind) for i in range(len(vals))]
    data = BlockOut("b0", "ranking", "T", ["AC"], cols=["value"])
    data.lines = lines
    res = SimpleNamespace(cells={ck("b0", f"l{i}", 0): v for i, v in enumerate(vals)})
    return [(10 + i, ln) for i, ln in enumerate(lines)], data, res


def test_a_category_that_prints_zero_is_left_out_of_a_bar_chart():
    rows, data, res = _data([-418_000_000, -82_560, -0.3, None, 0.0])
    kept = L._no_zero_categories(rows, data, res, [0], "bar", True)
    assert [ln.key for _, ln in kept] == ["l0", "l1"]                      # -0,1 MNOK stays; -0,0 and – go


def test_small_charts_in_whole_kroner_keep_their_small_values():
    rows, data, res = _data([-900_000, -40_000, -0.2])
    kept = L._no_zero_categories(rows, data, res, [0], "bar", True)      # max < 1 m: labels in NOK
    assert [ln.key for _, ln in kept] == ["l0", "l1"]


def test_line_and_stacked_charts_keep_every_category_and_nothing_is_emptied():
    rows, data, res = _data([0.0, 0.0])
    assert L._no_zero_categories(rows, data, res, [0], "line", True) == rows
    assert L._no_zero_categories(rows, data, res, [0], "bar", True) == rows    # all zero: keep, never empty


def test_a_bridge_waterfall_drops_zero_steps_but_keeps_start_and_end():
    s, res, p = run([{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}])
    (cs,) = p.charts
    bridge = next(o for o in res.outs if o.type == "bridge")
    first = p.at("b0.start.0")[0]
    end = p.at("b0.end.0")[0]
    assert cs.cat_range.startswith(f"C{first}") and cs.cat_range.endswith(f"C{end}")
    if abs(res.cells.get(ck("b0", "other", 0)) or 0) < 0.5:
        assert f"C{p.at('b0.other.0')[0]}" not in cs.cat_range.replace(":", ",").split(",")


# --- drivers ----------------------------------------------------------------------------------------------------

def test_drivers_lines_carry_the_code_and_the_chart_shows_the_leaves_in_green_and_red():
    s, res, p = run([{"type": "drivers", "query": "fm", "metric": "Opex", "depth": 2}])
    out = next(o for o in res.outs if o.type == "drivers")
    assert all(ln.code for ln in out.lines)                                 # 2BUS, 4BUS_B, B1 ... in column B
    top = out.lines[0]
    assert L.sheet_label(top) == "Business Norway" and top.label == "2BUS Business Norway"   # key_values keep it
    (cs,) = p.charts
    assert cs.kind == "variance"                                            # bars green/red by effect
    leaves = leaf_keys(out)
    assert top.key not in leaves and leaves
    rows = {p.at(ck(out.key, k, 0))[0] for k in leaves}
    shown = {int(x[1:]) for part in cs.cat_range.split(",") for x in [part.split(":")[0]]}
    assert shown <= rows | {r + 1 for r in rows}


def test_leaf_keys_are_the_bottom_of_the_shown_tree():
    out = BlockOut("b0", "drivers", "D", ["AC"], cols=["value"])
    out.lines = [Line("top", "T", [None], level=0), Line("d0", "A", [None], level=1),
                 Line("d0_0", "A1", [None], level=2), Line("d0_1", "A2", [None], level=2),
                 Line("d1", "B", [None], level=1)]
    assert leaf_keys(out) == {"d0_0", "d0_1", "d1"}


# --- truthful texts ---------------------------------------------------------------------------------------------

def test_an_implicit_chart_title_never_says_vs_twice():
    cs = SimpleNamespace(auto_title=True, title="", series=[("AC", "D1", "ac"), ("BU", "E1", "bu")])
    data = SimpleNamespace(col_metrics=[], headers=[], title="Actual vs plan", type="table")
    assert L.chart_title(cs, data, [], None) == "Actual vs budget"


def test_a_concepts_derived_query_ids_count_as_the_query():
    assert _of("q1", "q1") and _of("q1:fam_mobil_own", "q1")
    assert not _of("q10", "q1") and not _of("q10:fam", "q1")


def test_a_rankings_key_points_put_the_parent_first_never_between_its_parts():
    from smartass.render.keypoints import _chosen
    main = BlockOut("b0", "ranking", "Top", ["AC", "PY", "dPY"], cols=["value", "ly", "var"])
    main.lines = [Line("e0", "A", [None] * 3), Line("e1", "B", [None] * 3), Line("total", "T", [None] * 3, kind="grand")]
    res = SimpleNamespace(material=[{"key": "b0.e0", "variance": -154.0, "base": "ly"},
                                    {"key": "b0.total", "variance": -60.0, "base": "ly"},
                                    {"key": "b0.e1", "variance": 29.0, "base": "ly"}])
    assert [ln.key for ln, _, _ in _chosen(main, res)] == ["total", "e0", "e1"]


def test_chart_labels_use_the_typographic_minus():
    from smartass.render.charts import label_format
    for role in ("ac", "down", "pct"):
        f = label_format(role, True)
        assert '"−"' in f and ";-" not in f and not f.startswith("-")


# --- print: a section is never cut by a page break --------------------------------------------------------------

def test_keep_breaks_moves_a_section_that_would_straddle_a_page_to_the_next():
    from smartass.xl.styling import PRINT_WIDTH_PT, keep_breaks
    heights = {r: 15.0 for r in range(1, 80)}
    # page at 100 %: ~30 rows of 15 pt. A section on rows 28-40 would be cut: it starts the next page
    assert keep_breaks(heights, 79, [(28, 40)], PRINT_WIDTH_PT) == [28]
    assert keep_breaks(heights, 79, [(5, 12)], PRINT_WIDTH_PT) == []           # fits where it is
    assert keep_breaks(heights, 79, [(10, 75)], PRINT_WIDTH_PT) == []          # taller than a page: it flows


def test_keep_breaks_follow_the_fit_to_width_zoom():
    from smartass.xl.styling import PRINT_WIDTH_PT, keep_breaks
    heights = {r: 15.0 for r in range(1, 80)}
    # twice as wide as the page: 50 % zoom, a page holds ~60 rows, so rows 28-40 fit on page 1
    assert keep_breaks(heights, 79, [(28, 40)], PRINT_WIDTH_PT * 2) == []


def test_answer_roles_keep_each_section_and_the_charts_zone():
    from smartass.render.excel import keep_spans
    s, res, p = run([{"type": "table", "query": "fm", "title": "P&L"}])
    p.footer = "f"
    spans = keep_spans(p)
    kinds = {r["row"]: r["kind"] for r in p.rows}
    heads = [r for r, k in kinds.items() if k == "section"]
    assert {a for a, _ in spans} >= set(heads)
    zones = dict(p.zones)
    if p.charts:
        assert any(a == zones["charts"] for a, _ in spans)
    for a, b in spans:
        assert a <= b <= p.last_row


def test_a_ranking_chart_title_counts_the_bars_it_draws():
    cs = SimpleNamespace(auto_title=True, title="", series=[("AC", "D1", "ac")], bars=11)
    data = SimpleNamespace(col_metrics=[], headers=[], title="Top 15 by Opex", type="ranking")
    assert L.chart_title(cs, data, [], None) == "Top 11 by Opex"


def test_a_stack_of_revenue_and_costs_gets_a_lint_line():
    s, res, p = run([{"type": "table", "query": "fm", "title": "P&L"},
                     {"type": "chart", "kind": "stacked", "title": "Parts", "source": {"block": 0},
                      "series": ["value", "ly"]}])
    assert any("both signs" in m for m in p.chart_lint)


def test_each_row_of_charts_is_kept_on_one_page():
    from smartass.render.excel import keep_spans
    s, res, p = run([{"type": "table", "query": "fm", "title": "P&L"},
                     {"type": "chart", "kind": "column", "title": "A", "source": {"block": 0}},
                     {"type": "chart", "kind": "variance", "title": "B", "source": {"block": 0}},
                     {"type": "chart", "kind": "column", "title": "C", "source": {"block": 0}}])
    rows = sorted({c.anchor_row for c in p.charts})
    assert len(rows) == 2
    spans = keep_spans(p)
    for r in rows:
        assert (r, min(r + L.CHART_ROWS - 1, p.last_row)) in spans


def test_key_points_name_every_base_of_the_group():
    from smartass.render.keypoints import _formula
    o = BlockOut("b0", "table", "T", ["AC", "BU", "dBU", "FC", "dFC"],
                 cols=["value", "budget", "var", "forecast", "var"])
    ln = Line("l0", "EBITDA", [Num(None)] * 5)          # any expression: the cell exists
    cells = {ck("b0", "l0", j): v for j, v in enumerate([-100.0, -90.0, -10.0, -95.0, -5.0])}
    f = _formula(o, ln, lambda k: "X" + k[-1], set(), "budget", cells, "MNOK")
    assert " vs BU; " in f and f.endswith(' vs FC"')


def test_a_child_with_only_rounding_noise_is_negligible_for_a_ranking():
    from smartass.model.blocks.common import negligible
    vals = {("value", "A"): -0.3, ("ly", "A"): None, ("value", "B"): -82_560.0, ("ly", "B"): 0.0,
            ("value", "C"): 0.0, ("ly", "C"): 0.4}
    ctx = SimpleNamespace(pick=lambda q, m, c, e=None, months=None, product=None: vals[(c, e)],
                          value=lambda x: x)
    cols = ["value", "ly", "var", "var_pct", "share"]
    assert negligible(ctx, None, "Opex", cols, entity="A") and negligible(ctx, None, "Opex", cols, entity="C")
    assert not negligible(ctx, None, "Opex", cols, entity="B")


def test_the_highlight_accent_is_the_tn_template_blue():
    from smartass.render import charts as C
    from smartass.xl.vendor import tn_rapportstil as tn
    assert C.ACCENT == tn.rgb((28, 22, 197))            # #1C16C5 (Tom 06.11.2026)


def test_two_lines_get_no_min_labels_one_line_keeps_its_low_point():
    from v2util import monthly_facts
    for metrics, want_min in ((["Revenue"], True), (["Revenue", "Opex"], False)):
        s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "trend", "query": "fm", "metrics": metrics}]))
        res = compute(s, resolver(), CLOSED, monthly_facts())
        p = L.layout(1, s, res, stock=set(), subtitle="x")
        for cs in (c for c in p.charts if c.kind == "line"):
            n = len(cs.label_points)
            if want_min and n == 1 and cs.label_min[0]:
                assert cs.label_min[0] in cs.label_points[0]
            if n > 1:
                for k in range(n):
                    assert not (cs.label_min[k] in cs.label_points[k] and cs.label_min[k]
                                not in (max(cs.label_points[k]),)) or len(cs.label_points[k]) <= 2


def test_matrix_and_entity_tables_carry_codes_and_a_real_title():
    s, res, p = run([{"type": "matrix", "query": "fm", "metric": "Opex"},
                     {"type": "table", "query": "fm", "rows": "entities", "metric": "Opex"}])
    mx = next(o for o in res.outs if o.type == "matrix")
    assert all(ln.code for ln in mx.lines if ln.kind != "check")
    tb = next(o for o in res.outs if o.type == "table")
    assert tb.title == "Opex by unit"


def test_a_member_without_a_label_shows_its_code():
    assert L.sheet_label(Line("e0", "un.SY.4BUS_B", [None], code="4BUS_B")) == "4BUS_B"
    assert L.sheet_label(Line("e0", "4BUS_B Business", [None], code="4BUS_B")) == "Business"


def test_a_matrix_leaves_out_a_child_with_nothing_in_any_month():
    from v2util import CLOSED as _C
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "matrix", "query": "fm", "metric": "Opex"}]))
    res = compute(s, resolver(), _C, facts())
    mx = next(o for o in res.outs if o.type == "matrix")
    assert all(any(isinstance(res.cells.get(ck(mx.key, ln.key, j)), (int, float))
                   and abs(res.cells.get(ck(mx.key, ln.key, j))) >= 0.5 for j in range(len(mx.cols)))
               for ln in mx.lines if ln.kind == "line")
    assert any(ln.kind == "check" for ln in mx.lines)
