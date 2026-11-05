# engine/tests/v2/test_v5_check.py
"""Plan 5 Task 13: the fast layout check - each rule fires on a crafted fake and stays silent on a clean draw from
the real layout(). No Excel: a tiny sheet fake that answers the check's few reads. Every figure is invented."""
import time
from types import SimpleNamespace

import pytest

from fakexl import parse_area
from test_v5_headers import OCT, ytd_frame
from v2util import resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.spec import AnalysisSpec
from smartass.render.check import check
from smartass.render.layout import FIRST_COL, TEXT_COL, SheetPlan, col_letter, layout

GROUPED = {"type": "table", "query": "fm", "title": "P&L", "periods": ["closed", "closed-ytd"]}
BLOCKS = {
    "grouped": [GROUPED],
    "bridge": [{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}],
    "ranking": [{"type": "ranking", "query": "fm", "metric": "Opex", "top": 1}],
    "drivers": [{"type": "drivers", "query": "fm", "metric": "Opex"}],
    "matrix": [{"type": "matrix", "query": "fm", "metric": "Opex"}],
    "kpi_tiles": [{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}],
    "trend": [{"type": "trend", "query": "fm", "metrics": ["Opex"]}],
    "commentary": [GROUPED, {"type": "commentary", "key_points": ["Opex up"],
                             "drivers": [{"line": "Opex", "text": "Probably consultants", "basis": "hypothesis"}]}],
    "chart": [GROUPED, {"type": "chart", "kind": "column", "title": "Opex", "source": {"block": 0}}],
}


class Sheet:
    """Value2 of a rectangle from a {(row, col): value} dict, ChartObjects().Count, Range(cell).Value."""

    def __init__(self, grid=None, charts=1, fail=False):
        self.grid, self.n, self.fail, self.reads = dict(grid or {}), charts, fail, 0

    def ChartObjects(self):
        self.reads += 1
        return SimpleNamespace(Count=self.n)

    def Range(self, a):
        self.reads += 1
        if self.fail:
            raise RuntimeError("COM busy")
        r1, c1, r2, c2 = parse_area(a)
        if (r1, c1) == (r2, c2):
            return SimpleNamespace(Value=self.grid.get((r1, c1)), Value2=self.grid.get((r1, c1)))
        v = tuple(tuple(self.grid.get((r, c)) for c in range(c1, c2 + 1)) for r in range(r1, r2 + 1))
        return SimpleNamespace(Value2=v, Value=v)


def drawn(blocks=None, value=1234.5):
    """A clean answer: the real layout, every formula shown as a number (value), key points as text."""
    d = spec_dict(blocks=blocks or [GROUPED])
    d["queries"][0]["periods"] = "closed-ytd"
    s = AnalysisSpec.model_validate(d)
    p = layout(3, s, compute(s, resolver(), OCT, ytd_frame()), stock=set(), subtitle="SUB")
    grid = {}
    for (r, c), v in p.cells.items():
        grid[(r, c)] = (value if c >= FIRST_COL else "Opex fell 3 MNOK") if isinstance(v, str) and v.startswith("=") \
            else v
    for r in [d["row"] for d in p.rows if "status_col" in d]:
        grid[(r, FIRST_COL)] = "OK"
    return p, Sheet(grid), Sheet({(13, 4): "OK"})


def run(p, ws, data, ok=True, cells=None):
    return check(p, ws, data, ok, {"fm": "D13"} if cells is None else cells)


@pytest.mark.parametrize("kind", sorted(BLOCKS))
def test_a_clean_draw_has_no_findings(kind):
    p, ws, data = drawn(BLOCKS[kind])
    assert run(p, ws, data) == []


def test_a_large_nok_total_fits_a_number_column():
    """-20 123 456 789 (a YTD revenue in NOK) in a 16 wide column, bold on sum rows, a ▼ (2 bn) variance, -15.3 %:
    commas and parentheses are narrow."""
    from smartass.render.layout import col_number
    p, ws, data = drawn()
    big = {"heltall": -20_123_456_789.0, "avvik": -2_123_456_789.0, "prosent1": -0.153, "avvikpst": -0.153}
    for d in p.rows:
        for letter, num in (d.get("num") or {}).items():
            if num in big and (d["row"], col_number(letter)) in ws.grid:
                ws.grid[(d["row"], col_number(letter))] = big[num]
    assert not [f for f in run(p, ws, data) if "too narrow" in f]


def test_a_clean_draw_reads_little():
    p, ws, data = drawn()
    run(p, ws, data)
    assert ws.reads <= 2 and data.reads == 1                         # ChartObjects + one Value2; one status cell


def test_zones_out_of_order_or_a_plan_zone_fire():
    p, ws, data = drawn()
    p.zones[2], p.zones[3] = p.zones[3], p.zones[2]
    assert any("zones out of order" in f for f in run(p, ws, data))
    p, ws, data = drawn()
    p.zones.insert(2, ("plan", 3))
    out = run(p, ws, data)
    assert any("Plan zone" in f for f in out)


def test_no_chart_block_or_no_chart_drawn_fires():
    p, ws, data = drawn()
    p.charts = []
    assert any("no chart block" in f for f in run(p, ws, data))
    p, ws, data = drawn()
    ws.n = 0
    assert any("no chart was drawn" in f for f in run(p, ws, data))


def test_a_line_row_without_a_number_format_fires_and_e_plus_is_named():
    p, ws, data = drawn()
    line = next(d for d in p.rows if d.get("kind") == "line" and d.get("num") and "status_col" not in d)
    letter = sorted(line["num"])[0]
    del line["num"][letter]
    ws.grid[(line["row"], FIRST_COL + ord(letter) - ord(col_letter(FIRST_COL)))] = 3.2e12
    out = run(p, ws, data)
    assert any(f"row {line['row']} has no number format in {letter}" in f for f in out)
    assert any("E+" in f and f"{letter}{line['row']}" in f for f in out)


def test_a_number_too_wide_for_its_column_fires():
    p, ws, data = drawn()
    line = next(d for d in p.rows if d.get("kind") == "line" and d.get("num") and "status_col" not in d)
    letter = sorted(line["num"])[0]
    from smartass.render.layout import col_number
    ws.grid[(line["row"], col_number(letter))] = 123_456_789_012_345.0
    out = run(p, ws, data)
    assert any(f"column {letter} too narrow" in f and "###" in f for f in out)


def test_a_key_point_showing_an_excel_error_fires():
    p, ws, data = drawn()
    kp = next(r for z, r in p.zones if z == "key_points") + 1
    ws.grid[(kp, TEXT_COL)] = -2146826265                           # pywin32's #REF!
    assert any("#REF!" in f and f"key point B{kp}" in f for f in run(p, ws, data))
    ws.grid[(kp, TEXT_COL)] = "#N/A"
    assert any("#N/A" in f for f in run(p, ws, data))


def test_a_data_header_status_that_is_not_ok_fires():
    p, ws, data = drawn()
    data.grid[(13, 4)] = "DIFF"
    assert any("query fm says DIFF" in f for f in run(p, ws, data))


def test_missing_sources_rows_fire():
    p, ws, data = drawn()
    assert any("Sources rows" in f for f in run(p, ws, data, ok=False))


def test_the_check_never_raises():
    p, _, data = drawn()
    out = run(p, Sheet(fail=True), Sheet(fail=True))
    assert out and all(f.startswith("layout:") for f in out)
    assert run(SheetPlan(), Sheet(), Sheet(), cells={}) == [
        "layout: no chart block; fix: add a chart block chosen by the message (skill reference charts.md)"]


def test_200_rows_check_fast():
    p = SheetPlan()
    p.zones = [("title", 1), ("subtitle", 2), ("key_points", 4), ("charts", 8), ("tables", 26), ("checks", 230),
               ("notes", 233), ("source", 240)]
    p.charts = [object()]
    grid = {}
    letters = [col_letter(FIRST_COL + j) for j in range(8)]
    for r in range(27, 227):
        p.rows.append({"row": r, "kind": "line", "num": {x: "heltall" for x in letters}})
        for j in range(8):
            p.cells[(r, FIRST_COL + j)] = "=SUMIFS(Data_Q3[Value],Data_Q3[Metric],\"Opex\")"
            grid[(r, FIRST_COL + j)] = -1_234_567.0
    p.widths = {x: 16 for x in letters}
    p.last_col, p.last_row = letters[-1], 240
    t0 = time.perf_counter()
    out = check(p, Sheet(grid), Sheet({(13, 4): "OK"}), True, {"fm": "D13"})
    took = time.perf_counter() - t0
    print(f"layout check, 200 rows x 8 columns: {took * 1000:.1f} ms")
    assert took < 0.5 and out == []
