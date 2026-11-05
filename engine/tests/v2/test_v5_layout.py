# engine/tests/v2/test_v5_layout.py
"""Plan 5 Task 5: the answer's fixed zones, the code column, the scenario number kinds (layout, SheetPlan only, no
COM) and the new number formats through the fake Excel. Every figure is invented."""
import pytest
from pydantic import ValidationError
from test_v5_headers import OCT, ytd_frame
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.spec import AnalysisSpec
from smartass.render import layout as L
from smartass.render.layout import CHART_ROWS, CODE_COL, FIRST_COL, LABEL_COL, TEXT_COL, answer_subtitle, layout

GROUPED = {"type": "table", "query": "fm", "title": "P&L", "periods": ["closed", "closed-ytd"]}
BRIDGE = {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}
CONVENTIONS = ("NOK. FinMax signs: costs are negative. Δ = current − comparison; ▲ better, ▼ worse. Variance % "
               "blank when |base| < floor. ● = material (thresholds below).")


def scaled(frame, k: float):
    """The frame with every value times k (MNOK-sized amounts: units by magnitude)."""
    from dataclasses import replace

    from smartass.model.frame import Frame
    return Frame([replace(f, value=f.value * k) for f in frame.facts])


def run(blocks, periods="closed-ytd", stock=(), metrics=None, scale=1.0, **kw):
    d = spec_dict(blocks=blocks)
    d["queries"][0]["periods"] = periods
    if metrics:
        d["queries"][0]["metrics"] = metrics
    s = AnalysisSpec.model_validate(d)
    res = compute(s, resolver(), OCT, ytd_frame() if scale == 1.0 else scaled(ytd_frame(), scale))
    return s, res, layout(3, s, res, stock=set(stock), subtitle="SUB", **kw)


def texts(p, kind):
    return [p.cells[(r["row"], TEXT_COL)] for r in p.rows if r["kind"] == kind]


def test_columns():
    assert (TEXT_COL, CODE_COL, LABEL_COL, FIRST_COL, CHART_ROWS) == (2, 2, 3, 4, 18)


def test_zone_order_and_title_in_b1():
    s, res, p = run([GROUPED, BRIDGE, {"type": "commentary"}])
    assert p.cells[(1, 2)] == "Explain Business opex vs LY" and p.cells[(2, 2)] == "SUB"
    names = [z for z, _ in p.zones]
    assert names == ["title", "subtitle", "key_points", "charts", "tables", "checks", "notes", "source"]
    rows = [r for _, r in p.zones]
    assert rows == sorted(rows) and rows[0] == 1
    z = dict(p.zones)
    assert p.cells[(z["key_points"], 2)] == "Key points"
    assert p.cells[(z["checks"], 2)] == "Checks" and p.cells[(z["notes"], 2)] == "Notes"
    assert z["tables"] - z["charts"] == CHART_ROWS                     # one chart (the bridge waterfall): one row
    assert p.cells[(z["tables"], 2)] == "P&L"                          # the main table first
    source = " ".join(str(p.cells[(r, 2)]) for r in range(z["source"], p.last_row + 1))
    assert source == "Source: Oracle EPM Cloud (FinMax) · table Data_Q3 (sheet Q3 Data) · Sources sheet"


def test_plan_title_wins():
    d = spec_dict(blocks=[GROUPED])
    d["plan"]["title"] = "Opex is up 9 MNOK"
    d["queries"][0]["periods"] = "closed-ytd"
    s = AnalysisSpec.model_validate(d)
    p = layout(3, s, compute(s, resolver(), OCT, ytd_frame()), stock=set(), subtitle="x")
    assert p.cells[(1, 2)] == "Opex is up 9 MNOK"


def test_no_plan_zone_and_no_commentary_text():
    _, _, p = run([GROUPED, {"type": "commentary", "key_points": ["kp one"],
                             "drivers": [{"line": "Opex", "text": "consultants", "basis": "hypothesis"}]}])
    vals = [v for v in p.cells.values() if isinstance(v, str)]
    assert "Plan" not in vals and not any(v.startswith(("Question:", "•  Hypothesis")) for v in vals)
    assert not any("Hypothesis" in v or "consultants" in v or "kp one" in v for v in vals)
    assert "Commentary" not in vals


def test_month_and_ytd_columns_gap_and_material():
    _, res, p = run([GROUPED])
    o = res.outs[0]
    hdr = next(r for r in p.rows if r["kind"] == "header")
    row = hdr["row"]
    got = [p.cells.get((row, c)) for c in range(FIRST_COL, FIRST_COL + 10)]
    assert got == ["AC Oct 26", "PY Oct 25", "ΔPY", "ΔPY %", "", "AC YTD Oct 26", "PY YTD Oct 25", "ΔPY", "ΔPY %",
                   "Material"]
    assert hdr["num_cols"] == ["D", "E", "F", "G", "I", "J", "K", "L", "M"]      # H: the gap
    assert p.at(f"b0.{o.lines[0].key}.9")[1] == 13                                # Material in M
    assert p.widths["H"] == 2 and p.widths["M"] == 9 and p.widths["D"] == 16
    assert p.widths["A"] == 2 and p.widths["B"] == 9 and p.widths["C"] >= 46


def test_num_kinds_per_column():
    _, _, p = run([GROUPED])
    line = next(r for r in p.rows if r["kind"] == "line")
    assert line["num"] == {"D": "heltall", "E": "heltall", "F": "avvik", "G": "avvikpst",
                           "I": "heltall", "J": "heltall", "K": "avvik", "L": "avvikpst", "M": "markor"}


def test_num_rules():
    assert L._num("var", False) == "avvik" and L._num("var", True) == "tall1"
    assert L._num("var_pct", False) == "avvikpst" and L._num("material", False) == "markor"
    assert L._num("share", False) == "prosent1" and L._num("key", True) == "prosent1"
    assert L._num("value", False) == "heltall" and L._num("ly", True) == "tall1"


def test_code_in_b_and_label_in_c():
    _, res, p = run([GROUPED])
    o = res.outs[0]
    r0, _ = p.at(f"b0.{o.lines[0].key}.0")
    assert p.cells[(r0, CODE_COL)] == o.lines[0].code != ""
    assert p.cells[(r0, LABEL_COL)].strip() == o.lines[0].label


def test_checks_rows_use_the_tolerance_param():
    _, res, p = run([{"type": "ranking", "query": "fm", "metric": "Opex", "top": 1}])
    z = dict(p.zones)
    rows = [r for r in p.rows if z["checks"] < r["row"] < z["notes"] and r["kind"] == "check"]
    first = rows[0]
    assert p.cells[(first["row"], LABEL_COL)] == "Sheet = model"
    assert p.status_cell == f"D{first['row']}" and (first["row"], FIRST_COL) not in p.cells
    tol = p.params["tolerance"]
    o = res.outs[0]
    ck = next(f"b0.{ln.key}.{j}" for ln in o.lines if ln.kind == "check" for j, e in enumerate(ln.cells) if e)
    f = p.cells[(rows[1]["row"], FIRST_COL)]
    assert f == f'=IF(ABS({p.addr[ck]})<={tol},"OK","DIFF")'
    assert rows[1]["status_col"] == "D"


def test_thresholds_hold_every_param_with_addresses():
    _, res, p = run([GROUPED], stock={"FTE"})
    for key in [*res.params, "unit_mnok", "tolerance", "hundred"]:
        r, c = _cell(p.params[key])
        assert isinstance(p.cells[(r, c)], (int, float)) and c >= FIRST_COL
    head = next(r for (r, c), v in p.cells.items() if c == LABEL_COL and v == "Metric")
    heads = [p.cells.get((head, c)) for c in range(FIRST_COL, FIRST_COL + 4)]
    assert heads == ["Var % floor", "Materiality", "Materiality %", "Materiality with %"]
    metrics = {k.split(".", 1)[1] for k in res.params}
    rows = {_cell(p.params[k])[0] for k in res.params}
    assert len(rows) == len(metrics) and min(rows) == head + 1     # one row per metric, under the header
    for k in res.params:                                           # each param under its own column header
        r, c = _cell(p.params[k])
        want = {"floor": "Var % floor", "mat_abs": "Materiality", "mat_pct": "Materiality %",
                "mat_small": "Materiality with %"}[k.split(".", 1)[0]]
        assert p.cells[(head, c)] == want
        assert p.cells[(r, LABEL_COL)].split(" (")[0] == k.split(".", 1)[1]
    for key, label, value in (("unit_mnok", "Unit: MNOK", 1_000_000), ("tolerance", "Tolerance (NOK)", 1.0),
                              ("hundred", "Percent factor", 100)):
        r, c = _cell(p.params[key])
        assert (c, p.cells[(r, LABEL_COL)], p.cells[(r, c)]) == (FIRST_COL, label, value)
        assert r > max(rows)
    assert any("(NOK)" in str(p.cells[(_cell(p.params[k])[0], LABEL_COL)]) for k in p.params
               if k.startswith("mat_abs."))


def _cell(a):
    """'$E$40' -> (40, 5)."""
    _, col, row = a.split("$")
    return int(row), L.col_number(col)


def test_stock_threshold_labels_have_no_nok():
    _, res, p = run([{"type": "table", "query": "fm", "metrics": ["FTE", "Opex"], "chain": False}],
                    stock={"FTE"}, metrics=["Opex", "FTE"])
    rows = {k: int(a.split("$")[2]) for k, a in p.params.items()}
    fte = [k for k in rows if k.endswith(".FTE")]
    assert fte
    for k in fte:
        assert "(NOK)" not in p.cells[(rows[k], LABEL_COL)]
    assert any("(NOK)" in p.cells[(rows[k], LABEL_COL)] for k in rows if k.endswith(".Opex"))


def test_conventions_note_and_departure_note():
    _, _, p = run([{"type": "table", "query": "fm", "title": "Subset", "metrics": ["Opex"], "chain": False}])
    notes = " ".join(t.strip() for t in texts(p, "note"))
    assert CONVENTIONS in notes
    assert "Subset" in notes and "chain" in notes


def test_every_model_cell_has_an_address():
    _, res, p = run([GROUPED, BRIDGE, {"type": "ranking", "query": "fm", "metric": "Opex", "top": 1}])
    for o in res.outs:
        for ln in o.lines:
            for j, e in enumerate(ln.cells):
                k = f"{o.key}.{ln.key}.{j}"
                assert (k in p.addr) == (e is not None)
                if e is not None:
                    assert p.at(k)[1] == FIRST_COL + j


def test_chart_data_after_notes():
    chart = {"type": "chart", "kind": "column", "title": "Opex rose", "source": {"query": "fm", "metrics": ["Opex"]}}
    _, res, p = run([GROUPED, chart])
    z = dict(p.zones)
    assert [n for n, _ in p.zones][-2:] == ["chart_data", "source"] and z["chart_data"] > z["notes"]
    assert p.cells[(z["chart_data"], 2)] == "Chart data: Opex rose"
    assert z["tables"] - z["charts"] == CHART_ROWS


def test_charts_zone_two_per_row_and_anchors():
    blocks = [GROUPED, BRIDGE, {"type": "trend", "query": "fm", "metrics": ["Opex"]},
              {"type": "chart", "kind": "column", "title": "Opex rose", "source": {"block": 0}}]
    _, _, p = run(blocks)
    z = dict(p.zones)
    assert z["tables"] - z["charts"] == 2 * CHART_ROWS                 # three charts: two zone rows
    a, b, c = p.charts
    assert (a.kind, b.kind, c.kind) == ("waterfall", "line", "column")     # spec order
    assert a.anchor_col == "B" and c.anchor_col == "B" and b.anchor_col not in ("B", "C")
    assert a.anchor_row == b.anchor_row == z["charts"] and c.anchor_row == z["charts"] + CHART_ROWS
    assert a.height_pt == 260 < CHART_ROWS * 15                         # 480 x 260 pt (spec 3.3) in its zone row
    assert a.series and b.series and c.series                           # every kind is drawn (Task 7)


def test_implicit_chart_reserves_the_zone():
    _, res, p = run([GROUPED])
    assert res.outs[-1].type == "chart" and len(p.charts) == 1
    z = dict(p.zones)
    assert z["tables"] - z["charts"] == CHART_ROWS


def test_key_points_hook_rows():
    calls = []

    def kp(main, res, drivers, addr):
        calls.append(main.key)
        return [(f'="•  x "&{addr(f"{main.key}.{main.lines[0].key}.0")}', "   because"), ("=\"•  y\"", None)]
    _, _, p = run([GROUPED], key_points=kp)
    z = dict(p.zones)
    r = z["key_points"]
    assert calls == ["b0", "b0"]
    first = p.at(f"b0.{_first_key(p)}.0")
    assert p.cells[(r + 1, TEXT_COL)] == f'="•  x "&{p.addr[f"b0.{_first_key(p)}.0"]}' and first[0] > r
    assert p.cells[(r + 2, TEXT_COL)] == "   because" and p.cells[(r + 3, TEXT_COL)] == '="•  y"'


def _first_key(p):
    return next(k.split(".")[1] for k in p.addr if k.startswith("b0."))


def test_label_width_measures_column_c():
    from smartass.render.excel import label_width
    from smartass.render.layout import SheetPlan
    p = SheetPlan(cells={(2, 3): "x" * 60, (2, 2): "y" * 90}, rows=[{"row": 2, "kind": "line", "num": {}}])
    assert label_width(p) == 62


def test_subtitle():
    s, res, _ = run([GROUPED])
    assert answer_subtitle(s, res, {"FinMax": (2026, 10)}) == \
        "NOK · FinMax · 2BUS Business Norway · YTD Oct 26 · closed Oct 26"


def test_old_closed_frame_still_lays_out():
    d = spec_dict()
    s = AnalysisSpec.model_validate(d)
    p = layout(1, s, compute(s, resolver(), CLOSED, facts()), stock=set(), subtitle="x")
    assert p.last_row == max(r["row"] for r in p.rows)


def test_waterfall_with_own_data_is_refused():
    d = spec_dict(blocks=[GROUPED, {"type": "chart", "kind": "waterfall", "title": "Bridge",
                                    "source": {"query": "fm", "metrics": ["Opex"]}, "series": ["amount"]}])
    with pytest.raises(ValidationError, match=r"point a waterfall at a bridge block: source \{block: <bridge index>\}"):
        AnalysisSpec.model_validate(d)


def test_code_column_widens_for_a_long_bold_code():
    from smartass.render.layout import SheetPlan, code_width
    p = SheetPlan(cells={(2, 2): "C30"}, rows=[{"row": 2, "kind": "line", "num": {}}])
    assert code_width(p) == 9
    p.cells[(3, 2)] = "EBITDA_BOI"
    p.rows.append({"row": 3, "kind": "sum", "num": {}})
    assert code_width(p) == 13


# --- the number formats through the fake Excel --------------------------------------------------------------

def test_a_laid_out_answer_styles_without_lint_findings():
    from fakexl import Workbook

    from smartass.render.excel import answer_grid, answer_roles
    from smartass.xl import styling
    _, _, p = run([GROUPED, BRIDGE, {"type": "ranking", "query": "fm", "metric": "Opex", "top": 1}])
    wb = Workbook(["Q3 Answer"])
    ws = wb.order[0]
    r0, c0, g = answer_grid(p)
    for i, row in enumerate(g):
        for j, v in enumerate(row):
            ws.put(r0 + i, c0 + j, v)
    roles = answer_roles(p)
    assert styling.apply_report(ws, roles) == []
    assert roles["col_widths"]["H"] == 16          # the ranking uses H for numbers: a shared column stays 16
    assert roles["col_widths"]["M"] == 9 and roles["code_cols"] == ["B"]
    r, c = p.at(next(k for k in p.addr if k.startswith("b0.") and k.endswith(".7")))     # YTD ΔPY
    assert ws.cells[(r, c)]["NumberFormatLocal"] == _local(styling, "avvik")


def _local(styling, kind):
    """The kind's en-US code in the fake's locale: through late-bound COM Excel reads number formats locally."""
    from fakexl import INTERNATIONAL
    return styling.tn.to_local(styling.NUM_FORMATS[kind], INTERNATIONAL[2], INTERNATIONAL[3])


def test_new_kinds_set_number_format_en_us():
    from fakexl import Workbook

    from smartass.xl import styling
    wb = Workbook(["Q1 Answer"])
    ws = wb.order[0]
    ws.put(1, 2, "T")
    ws.put(2, 2, "C30")
    ws.put(2, 3, "Revenue")
    for c, v in zip(range(4, 8), (1.0, -2.0, 0.05, 1)):
        ws.put(2, c, v)
    roles = {"last_col": "G", "label_cols": ["B", "C"], "code_cols": ["B"], "last_row": 2,
             "col_widths": {"A": 2, "B": 9, "C": 46, "D:G": 16},
             "rows": [{"row": 1, "kind": "title"},
                      {"row": 2, "kind": "line", "num": {"D": "heltall", "E": "avvik", "F": "avvikpst",
                                                          "G": "markor"}}]}
    styling.apply_report(ws, roles)
    cell = ws.cells
    assert cell[(2, 5)]["NumberFormatLocal"] == _local(styling, "avvik")
    assert cell[(2, 6)]["NumberFormatLocal"] == _local(styling, "avvikpst")
    assert cell[(2, 7)]["NumberFormatLocal"] == _local(styling, "markor")
    assert "NumberFormatLocal" not in cell[(2, 4)]
    assert cell[(2, 5)]["Style"] == styling._style_for("heltall", "")       # the base TN style of the kind
    assert cell[(2, 6)]["Style"] == styling._style_for("prosent1", "")
    assert cell[(2, 2)]["Font.Color"] == styling._rgb(styling.tn.GREY)
    assert cell[(2, 2)]["HorizontalAlignment"] == styling.tn.H_LEFT
    assert "Font.Color" not in cell.get((1, 2), {})


def test_new_formats_are_en_us_codes():
    from smartass.xl import styling
    f = styling.NUM_FORMATS
    assert "▲" in f["avvik"] and "▼" in f["avvik"] and "[Red]" in f["avvik"] and "#,##0" in f["avvik"]
    assert "0.0 %" in f["avvikpst"] and "●" in f["markor"]


def test_thresholds_table_never_lands_on_a_gap_marker_or_helper_column():
    """A grouped table has a gap (width 2) and a Material column (width 9): the thresholds skip them. A narrow
    answer (a bridge) is widened so the thresholds sit inside the print area, left of the chart helpers."""
    for blocks in ([GROUPED], [BRIDGE], [{"type": "ranking", "query": "fm", "metric": "Opex", "top": 1}]):
        _, res, p = run(blocks)
        if not res.params:
            continue
        last = L.col_number(p.print_last_col)
        for k in res.params:
            r, c = _cell(p.params[k])
            letter = L.col_letter(c)
            assert p.widths[letter] == L.W_NUM and c <= last and letter not in p.helper_cols, (blocks, k)
