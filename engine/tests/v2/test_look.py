# engine/tests/v2/test_look.py
"""Phase 3 Task 1: the Answer's look as a pure function (render/look.py). The oracle is today's COM styling,
styling.apply_report on the fake Excel: every cell's style, number format, font colour, alignment and wrap, the row
heights, column widths, merges, check colours and page breaks must come out the same. Every figure is invented."""
import datetime
import re
from functools import partial

import pytest
from fakexl import Workbook
from test_v5_layout import BRIDGE, GROUPED, run
from test_v5_live_round2 import _lines_answer

from smartass.render.excel import answer_footer, answer_grid, answer_roles, keep_spans
from smartass.render.keypoints import key_points
from smartass.render.layout import SheetPlan, col_letter, col_number, width_pt
from smartass.render import look as look_mod
from smartass.render.look import answer_look, style_props, text_heights
from smartass.xl import styling
from smartass.xl.vendor import tn_rapportstil as tn

FOOTER = answer_footer(datetime.datetime(2026, 11, 5))
LONG_POINT = "Opex rose on consultants, licences and the new data centre while revenue held up on wholesale volumes"
H = {tn.H_LEFT: "left", tn.H_RIGHT: "right", tn.H_GEN: None}
V = {tn.V_TOP: "top", tn.V_CENTER: "vcenter", tn.V_BOTTOM: "bottom"}


# ------------------------------------------------------------------ the cases: a plan and its calculated values
def _formula(p: SheetPlan, rc) -> bool:
    v = p.cells.get(rc)
    return isinstance(v, str) and v.startswith("=") and rc not in p.literal


def _calculated(p: SheetPlan, res) -> dict:
    """What Excel would show in every cell: model cells their model value, a Key point formula a sentence longer
    than its literals (so it wraps further), a check formula "OK", any other formula (chart helpers) a number."""
    model = {p.at(k): k for k in p.addr}
    out = {}
    for rc, v in p.cells.items():
        if rc in model:
            out[rc] = res.cells.get(model[rc])
        elif _formula(p, rc) and rc[1] == 2:
            words = "".join(re.findall(r'"((?:[^"]|"")*)"', v)).replace('""', '"')
            out[rc] = words + " 1 234,5 MNOK; " + "−12,3 MNOK vs PY (−4,5 %); " * 12
        elif _formula(p, rc):
            out[rc] = "OK" if '"OK"' in v else 0.0
        else:
            out[rc] = v
    return out


def _laid_out(blocks, **kw):
    s, res, p = run(blocks, key_points=partial(key_points, free=[LONG_POINT]), **kw)
    p.footer = FOOTER
    return p, _calculated(p, res)


def pl_table():
    """A P&L table (sum and grand rows, Δ NOK, Δ %, Material) with a bridge: a chart zone and chart helper columns."""
    return _laid_out([GROUPED, BRIDGE])


def fte_table():
    """A table with a stock metric (FTE): one decimal with a thousands separator over TN Tall 1d."""
    return _laid_out([{"type": "table", "query": "fm", "metrics": ["FTE", "Opex"], "chain": False}],
                     stock={"FTE"}, metrics=["Opex", "FTE"])


def trend():
    """A trend with its line chart in the Charts zone."""
    return _laid_out([{"type": "trend", "query": "fm", "metrics": ["Opex"]}])


def ranking_with_checks():
    """A ranking: three check rows (a status run), thresholds and notes."""
    return _laid_out([{"type": "ranking", "query": "fm", "metric": "Opex", "top": 1}])


def lines_ranking():
    """A ledger lines answer ranked by supplier."""
    _, res, p = _lines_answer()
    p.footer = FOOTER
    return p, _calculated(p, res)


def small_sheet():
    """The pinned answer of tests/test_excel_calls.py as a plan (no footer: v1 page setup), with a title that wraps,
    a sub-heading, a note, a Key point formula, a number and a text in columns without a number kind."""
    rows = [(1, "title", ["Business — EBITDA: revenue grew on wholesale volumes while opex rose with consultants and "
                          "licences across every quarter of the year"]),
            (2, "subtitle", ["Actual, NOK"]), (3, "subhead", ["", "Main lines"]),
            (4, "header", ["", "Line", "Sep 26", "Sep 25", "Var", "Var %"]),
            (5, "line", ["4REV", "Revenue", 100.0, 90.0, "=D5-E5", "=F5/E5"]),
            (6, "line", ["4OPX", "Opex", -40.0, -35.0, 3.0, '=IF(D6>0,"up","down")']),
            (7, "line", ["FTE", "FTE", 12.5, 11.0, 1.5, 0.13]),
            (8, "sum", ["EBITDA", "EBITDA", 60.0, 55.0, 5.0, 0.09]),
            (9, "check", ["", "Sum = cube", "", "", "", '=IF(ABS(D8-55)<1,"OK","DIFF")']),
            (10, "check", ["", "YTD = months", "", "", "", "OK"]),
            (11, "note", ["A note on the conventions."]), (12, "section", ["Notes"]),
            (13, "text", ['="•  EBITDA "&FIXED(D8,0)&" NOK"']),
            (14, "text", ["Revenue grew on wholesale volumes; opex rose with consultants. " * 3]),
            (16, "grand", ["TOTAL", "Total", 60.0, 55.0, 5.0, 0.09])]
    p = SheetPlan(last_col="G", last_row=17, widths={"F": 12})
    calc = {(5, 6): 10.0, (5, 7): 0.11, (6, 7): "down", (9, 7): "OK", (13, 2): "•  EBITDA 60 NOK " + LONG_POINT * 3}
    for r, kind, vals in rows:
        for c, v in enumerate(vals, 2):                          # B code (or title, note, text), C label, D..
            if v != "":
                p.cells[(r, c)] = v
        spec = {"row": r, "kind": kind}
        if kind == "header":
            spec["num_cols"] = ["D", "E", "F", "G"]
        if kind in ("line", "sum", "grand"):
            spec["num"] = {"D": "heltall", "E": "heltall"} if r == 6 else \
                {"D": "tall1", "E": "tall1", "F": "tall1", "G": "prosent1"} if r == 7 else \
                {"D": "heltall", "E": "heltall", "F": "avvik", "G": "avvikpst"}
        if kind == "check":
            spec["status_col"] = "G"
        p.rows.append(spec)
    return p, {rc: calc.get(rc, v) for rc, v in p.cells.items()}


CASES = [pl_table, fte_table, trend, ranking_with_checks, lines_ranking, small_sheet]


# ------------------------------------------------------------------ reading the fake Excel back
def fakexl_sheet_with(p: SheetPlan, values: dict):
    """The grid written as write_answer writes it (text with its apostrophe), then each cell's calculated value
    (the fake does not calculate)."""
    ws = Workbook(["Q3 Answer"]).order[0]
    r0, c0, g = answer_grid(p)
    for i, row in enumerate(g):
        for j, v in enumerate(row):
            ws.put(r0 + i, c0 + j, v)
    for (r, c), v in values.items():
        ws.put(r, c, v)
    return ws


def _props(ws, r, c) -> dict:
    return ws.cells.get((r, c), {})


def _en_us(code: str) -> str:
    """A NumberFormatLocal in the fake's nb-NO locale as en-US (outside quotes: NBSP -> ',', ',' -> '.')."""
    out, quoted = [], False
    for ch in code:
        if ch == '"':
            quoted = not quoted
        out.append(ch if quoted else {"\xa0": ",", ",": "."}.get(ch, ch))
    return "".join(out)


def shown_format(ws, r, c):
    p = _props(ws, r, c)
    if "NumberFormatLocal" in p:
        return _en_us(p["NumberFormatLocal"])
    return p.get("NumberFormat") or style_props(p["Style"]).get("num_format")


def _hex_bgr(v: int) -> str:
    return f"#{v & 255:02X}{v >> 8 & 255:02X}{v >> 16 & 255:02X}"


def font_colour(ws, r, c) -> str:
    p = _props(ws, r, c)
    return _hex_bgr(p["Font.Color"]) if "Font.Color" in p else style_props(p["Style"])["font_color"]


def shown(ws, r, c, prop: str, key: str, names: dict | None = None):
    """A cell's alignment or wrap: set on the cell itself, else its named style's."""
    p = _props(ws, r, c)
    if prop in p:
        return names[p[prop]] if names else bool(p[prop])
    return style_props(p["Style"]).get(key, None if names else False)


def merged_ranges(ws) -> list[tuple[int, int, int]]:
    out: list[list[int]] = []
    for r, c in sorted(rc for rc, p in ws.cells.items() if p.get("MergeCells")):
        if out and out[-1][0] == r and out[-1][2] == c - 1:
            out[-1][2] = c
        else:
            out.append([r, c, c])
    return [tuple(x) for x in out]


OK_RULES = [{"type": 1, "op": 3, "f": '="OK"', "Font.Color": tn.rgb(styling.GREEN), "Font.Bold": True},
            {"type": 1, "op": 4, "f": '="OK"', "Font.Color": tn.rgb(tn.BREACH_FONT), "Font.Bold": True,
             "Interior.Color": tn.rgb(tn.BREACH_FILL)}]


def status_formats(ws) -> list[tuple[int, int, int]]:
    """The cells with the check colours, as (col, first row, last row) runs; every one has the OK / not-OK pair."""
    cells = sorted(((c, r) for (r, c), p in ws.cells.items() if "FC" in p))
    out: list[list[int]] = []
    for c, r in cells:
        assert _props(ws, r, c)["FC"] == OK_RULES, (r, c)
        if out and out[-1][0] == c and out[-1][2] == r - 1:
            out[-1][2] = r
        else:
            out.append([c, r, r])
    return [tuple(x) for x in out]


def page_breaks(ws) -> list[int]:
    m = re.search(r"\|breaks ([\d,]+)", ws.Names("smartass_pagesetup").RefersTo)
    return [int(x) for x in m.group(1).split(",")] if m else []


# ------------------------------------------------------------------ the differential test
@pytest.mark.parametrize("case", CASES, ids=lambda f: f.__name__)
def test_look_equals_what_apply_report_does(case):
    plan, values = case()
    ws = fakexl_sheet_with(plan, values)
    styling.apply_report(ws, answer_roles(plan))
    look = answer_look(plan, texts={rc: v for rc, v in values.items() if isinstance(v, str)})
    for (r, c), cl in look.cells.items():
        props = style_props(cl.style)
        assert _props(ws, r, c)["Style"] == cl.style, (r, c)
        assert shown_format(ws, r, c) == (cl.num_format or props.get("num_format")), (r, c)
        assert font_colour(ws, r, c) == (cl.font_color or props["font_color"]), (r, c)
        assert shown(ws, r, c, "HorizontalAlignment", "align", H) == (cl.align or props.get("align")), (r, c)
        assert shown(ws, r, c, "VerticalAlignment", "valign", V) == (cl.valign or props["valign"]), (r, c)
        assert shown(ws, r, c, "WrapText", "text_wrap") == (props["text_wrap"] if cl.wrap is None else cl.wrap), \
            (r, c)
    assert {r: ws.Rows(r).RowHeight for r in look.heights} == look.heights
    assert {c: ws.Columns(c).ColumnWidth for c in look.widths} == look.widths
    assert merged_ranges(ws) == sorted(look.merges)
    assert status_formats(ws) == sorted(look.status_runs)
    if plan.footer:                                          # the v2 page setup (v1 has its own, xl_tidy's)
        assert ws.PageSetup.PrintArea == look.print_area and ws.PageSetup.CenterFooter == look.footer
        assert page_breaks(ws) == look.breaks
    bare = answer_look(plan)                                 # the file builder's path: no texts, then the read-back
    block = ws.Range("A1:" + plan.last_col + str(plan.last_row)).Value2
    assert bare.heights | text_heights(plan, bare, block) == look.heights
    if plan.footer:                                          # laid out: the texts change no style and no merge
        assert bare.cells == look.cells and bare.merges == look.merges


@pytest.mark.parametrize("case", CASES, ids=lambda f: f.__name__)
def test_the_breaks_after_the_read_back_are_the_breaks_of_the_look_with_its_texts(case):
    plan, values = case()
    texts = {rc: v for rc, v in values.items() if isinstance(v, str)}
    bare, full = answer_look(plan), answer_look(plan, texts=texts)
    ws = fakexl_sheet_with(plan, values)
    block = ws.Range("A1:" + plan.last_col + str(plan.last_row)).Value2
    assert look_mod.page_breaks(plan, bare) == bare.breaks
    assert look_mod.page_breaks(plan, bare, text_heights(plan, bare, block)) == full.breaks


def test_longer_key_points_can_move_the_page_breaks():
    plan, values = lines_ranking()
    texts = {rc: v for rc, v in values.items() if isinstance(v, str)}
    assert answer_look(plan).breaks != answer_look(plan, texts=texts).breaks      # the case above proves the move


def test_the_cases_reach_every_rule():
    """The differential test only proves what its cases contain: each rule of apply_report is hit at least once."""
    looks = [answer_look(p, texts={rc: v for rc, v in vals.items() if isinstance(v, str)})
             for p, vals in (case() for case in CASES)]
    cells = [cl for lk in looks for cl in lk.cells.values()]
    formats = {cl.num_format for cl in cells}
    assert {styling.TALL1_THOUSANDS, *styling.NUM_FORMATS.values()} <= formats
    styles = {cl.style for cl in cells}
    assert {"TN Heltall · Sum", "TN Heltall · Hovedsum", "TN Underseksjon", "TN Tekst tallkolonne",
            "TN Kolonne tall", "TN Note"} <= styles
    assert any(cl.align == "left" for cl in cells) and any(cl.font_color for cl in cells)
    assert any(lk.status_runs for lk in looks) and any(lk.breaks for lk in looks)
    merged = {(lk.cells[(r, c)].style) for lk in looks for r, c, _ in lk.merges}
    assert merged == {"TN Tittel", "TN Tekst"}               # a wrapped title and a merged Key point


# ------------------------------------------------------------------ the style findings (final review M1)
def _too_long(case):
    """The case with a label far wider than its column and a Key point that is one word longer than the table is
    wide: what the report-style lint finds."""
    def made():
        plan, values = case()
        lines = {r["row"] for r in plan.rows if r["kind"] == "line"}
        label = next(rc for rc, v in sorted(plan.cells.items()) if rc[0] in lines and rc[1] == 3
                     and isinstance(v, str) and not v.startswith("=") and (rc[0], 4) in plan.cells)
        point = next((r["row"] for r in plan.rows if r["kind"] == "text"), None)
        values[label] = plan.cells[label] = "A label far too long for its column, with words enough to spill over " * 2
        if point is not None:
            values[(point, 2)] = "W" * 300
        return plan, values
    made.__name__ = case.__name__ + "_too_long"
    return made


@pytest.mark.parametrize("case", CASES + [_too_long(c) for c in CASES], ids=lambda f: f.__name__)
def test_the_style_findings_are_the_text_findings_apply_report_gives(case):
    """The look's findings over the read-back block are apply_report's 'text does not fit' and 'text is cut off'
    findings, in its words, for the same sheet."""
    plan, values = case()
    ws = fakexl_sheet_with(plan, values)
    found = [f for f in styling.apply_report(ws, answer_roles(plan)) if ": text " in f]
    block = ws.Range("A1:" + plan.last_col + str(plan.last_row)).Value2
    assert look_mod.style_findings(plan, block) == found
    if case.__name__.endswith("_too_long"):
        assert any("text is cut off" in f for f in found)
        assert any("text does not fit" in f for f in found) == any(r["kind"] == "text" for r in plan.rows)


def test_the_look_gives_every_cell_arial_and_no_vertical_border():
    """What the look cannot get wrong, so the style findings need not check it: every format it writes is Arial,
    with a top and bottom edge at most (the font and vertical borders apply_report read from Excel)."""
    for case in CASES:
        plan, _ = case()
        for cl in answer_look(plan).cells.values():
            props = style_props(cl.style)
            assert props["font_name"] == "Arial" and not {"left", "right"} & set(props)


# ------------------------------------------------------------------ the pure rules
def test_every_cell_of_the_block_has_a_look_and_blank_cells_are_tn_tom():
    plan, _ = pl_table()
    look = answer_look(plan)
    assert len(look.cells) == plan.last_row * col_number(plan.last_col)
    listed = {r["row"] for r in plan.rows}
    blank = [r for r in range(1, plan.last_row + 1) if r not in listed]
    assert blank
    for r in blank:
        assert {look.cells[(r, c)].style for c in range(1, col_number(plan.last_col) + 1)} == {"TN Tom"}
        assert look.heights[r] == 15


def test_print_area_breaks_and_footer_come_from_the_plan():
    plan, _ = pl_table()
    look = answer_look(plan)
    assert look.print_area == "$A$1:$" + plan.print_area_col + "$" + str(plan.last_row)
    area = sum(width_pt(look.widths.get(col_letter(c), 8.43)) for c in range(1, col_number(plan.print_area_col) + 1))
    assert look.breaks == styling.keep_breaks(look.heights, plan.last_row, keep_spans(plan), area)
    assert look.footer == plan.footer == FOOTER


def test_a_number_or_a_formula_in_a_column_without_a_number_kind_is_tn_heltall_and_a_text_tn_etikett():
    plan, _ = small_sheet()
    look = answer_look(plan, texts={(6, 7): "down"})
    assert look.cells[(6, 6)].style == "TN Heltall"          # the constant 3.0
    assert look.cells[(6, 7)].style == "TN Etikett"          # its formula shows text
    assert answer_look(plan).cells[(6, 7)].style == "TN Heltall"     # without the text: a formula counts as a number


def _one_point_plan(*points: str) -> SheetPlan:
    """A narrow answer (A 2, B 9, C 46, D 16: Key points 71 characters wide) with one Key point formula per row."""
    p = SheetPlan(last_col="D", last_row=1 + len(points))
    p.cells[(1, 2)] = "Title"
    p.rows.append({"row": 1, "kind": "title"})
    for i, f in enumerate(points, 2):
        p.cells[(i, 2)] = f
        p.rows.append({"row": i, "kind": "text"})
    return p


def test_text_heights_only_returns_rows_that_change():
    p = _one_point_plan('="•  Opex "&FIXED(D9,0)&" NOK vs PY"', '="•  Revenue "&FIXED(D8,0)&" NOK"')
    look = answer_look(p)
    assert look.text_rows == [2, 3]
    long = "•  Opex 1 234 567 NOK vs PY; " + LONG_POINT + "; " + LONG_POINT[:40]
    assert tn.wrap_lines(long, 9 + 46 + 16) == 3
    block = ((None, "Title", None, None), (None, long, None, None), (None, "•  Revenue 12 NOK", None, None))
    assert text_heights(p, look, block) == {2: 3 * 13 + 2}
    assert look.heights[2] == look.heights[3] == 15         # the formulas' own text: one line (a lower bound)


def test_a_key_point_is_merged_over_the_table_width_with_its_text_at_the_top():
    p = _one_point_plan('="•  Opex "&FIXED(D9,0)&" NOK"')
    look = answer_look(p)
    assert look.merges == [(2, 2, 4)]
    first, rest = look.cells[(2, 2)], look.cells[(2, 3)]
    assert (first.style, first.wrap, first.valign) == ("TN Tekst", True, "top")
    assert (rest.style, rest.wrap) == ("TN Tekst", True)


def test_style_props_gives_the_concrete_format_of_a_named_style():
    p = style_props("TN Heltall · Hovedsum")
    assert (p["font_name"], p["font_size"], p["bold"], p["align"]) == ("Arial", 10, True, "right")
    assert p["num_format"] == '#,##0_);(#,##0);"–"_)' and p["top"] == 1 and p["bottom"] == 6   # thin, double
    assert p["top_color"] == p["bottom_color"] == "#404040"
    assert style_props("TN Tittel")["font_color"] == "#1F3864" and style_props("TN Tittel")["bottom"] == 2   # medium
    inp = style_props("TN Heltall · Input")                  # red, green and blue in their places
    assert (inp["font_color"], inp["bg_color"]) == ("#0000FF", "#FFF7DD")
    tom = style_props("TN Tom")
    assert "align" not in tom and tom["valign"] == "vcenter" and "num_format" not in tom and "top" not in tom
    note = style_props("TN Note")
    assert (note["italic"], note["indent"], note["font_size"], note["text_wrap"]) == (True, 1, 9, False)
    assert style_props("TN Tekst")["text_wrap"] is True and style_props("TN Tekst")["valign"] == "top"


def test_section_and_title_borders_do_not_bleed_into_col_a_helpers_or_narrow_empty_cols():
    from test_v5_layout import BRIDGE, GROUPED, run
    _, _, p = run([{"type": "kpi_tiles", "query": "fm", "metrics": ["Revenue"]}, GROUPED])
    lk = answer_look(p)
    title_row = 1
    assert lk.cells[(title_row, 1)].style == "TN Tom"
    assert lk.cells[(title_row, 2)].style == "TN Tittel"
    for hcol in p.helper_cols:
        assert lk.cells[(title_row, col_number(hcol))].style == "TN Tom"
    kpi_head = next(r["row"] for r in p.rows if r["kind"] == "header" and r.get("end_col") == "G")
    assert lk.cells[(kpi_head, col_number("G"))].style in ("TN Kolonne", "TN Kolonne tall")
    assert lk.cells[(kpi_head, col_number("H"))].style == "TN Tom"
    # Even a single-table narrow sheet with helpers skips Col A and helpers
    _, _, pb = run([BRIDGE])
    lkb = answer_look(pb)
    for hcol in pb.helper_cols:
        assert lkb.cells[(1, col_number(hcol))].style == "TN Tom"
