"""Plan 5 Task 10: provenance and tie-out on the 'Q<n> Lines <alias>' sheets (fakes only, no Excel). The header
sits in F1:K3: the lines table header is row 4 (drill_xl.write_lines), so rows 1-3 are the only free rows; A1, A2,
the cube cell and the table never move. Every figure is invented."""
from types import SimpleNamespace

from fakexl import Workbook
from test_v2_pipeline import LINE_CELL

from smartass import pipeline
from smartass.model.lines import AMOUNT, CUBE_CELL, LineSet
from smartass.render.datasheet import lines_header

COLS = ["ACCOUNT", "AP_CUSTOMER_NAME", "GL_DESCRIPTION_LINE", "PROJECT", "CC_DESC", "AMOUNT", "GL_POSTED_DATE"]
FETCHED = "2026-11-05 09:30 +0100 · tom"
TIE = '=IF(NOT(ISNUMBER($D$2)),"no cube value",IF(ABS($G$3-$D$2)<=$K$3,"OK","DIFF"))'


def _lines(cube=-30.0):
    rows = [{"ACCOUNT": "6700", "AP_CUSTOMER_NAME": "Bouvet", "GL_DESCRIPTION_LINE": "Consulting", "PROJECT": "P1",
             "CC_DESC": "Tech", "AMOUNT": -10.0 * (i + 1), "GL_POSTED_DATE": "2026-09-15"} for i in range(2)]
    return LineSet("b0", rows, COLS, cube_value=cube, cell=LINE_CELL)


def _draw(ls):
    wb = Workbook(["Q3 Answer"])
    t = SimpleNamespace(wb=wb, app=wb.Application, sheets_added=[])
    name = pipeline._write_lines(t, 3, ls, FETCHED)
    return wb.Worksheets(name)


def _v(ws, a):
    r, c = int(a[1:]), ord(a[0]) - 64
    return ws.cells.get((r, c), {}).get("Value")


def test_lines_header_is_label_value_pairs_with_the_exact_formulas():
    g = lines_header("Lines_Q3_b0", LINE_CELL, FETCHED)
    assert g[0][:2] == ["Drill cell", "All Accounts: ac.C60 · All Entities: un.SY.2TEC · All Time Periods: Yr 26"]
    assert g[1][:2] == ["Fetched", FETCHED]
    assert g[2] == ["Lines total", f"=SUM(Lines_Q3_b0[{AMOUNT}])", "Tie-out", TIE, "Tolerance (NOK)", 1.0]
    assert all(len(r) == 6 for r in g)


def test_the_lines_sheet_gets_the_header_in_f1_k3():
    ws = _draw(_lines())
    assert ws.Name == "Q3 Lines b0"
    assert _v(ws, "F1") == "Drill cell" and _v(ws, "G1").startswith("All Accounts: ac.C60")
    assert _v(ws, "F2") == "Fetched" and _v(ws, "G2") == FETCHED
    assert _v(ws, "F3") == "Lines total" and _v(ws, "G3") == "=SUM(Lines_Q3_b0[AMOUNT])"
    assert _v(ws, "H3") == "Tie-out" and _v(ws, "I3") == TIE
    assert _v(ws, "J3") == "Tolerance (NOK)" and _v(ws, "K3") == 1.0


def test_title_source_cube_cell_and_table_do_not_move():
    ws = _draw(_lines())
    assert CUBE_CELL == "$D$2"
    assert _v(ws, "C2") == "Cube value" and _v(ws, "D2") == -30.0
    assert _v(ws, "A1") == "Ledger lines behind ac.C60 / un.SY.2TEC / Yr 26"
    assert _v(ws, "A2").startswith("Source: Oracle EPM Cloud")
    lo = ws.ListObjects("Lines_Q3_b0")
    assert lo.HeaderRowRange.Row == 4 and lo.HeaderRowRange.Column == 1
    assert [ws.cells[(4, c)]["Value"] for c in range(1, len(COLS) + 1)] == COLS     # table header untouched


def test_a_missing_cube_value_says_so_instead_of_diff():
    ws = _draw(_lines(cube=None))
    assert _v(ws, "D2") is None and "no cube value" in _v(ws, "I3")


def test_a_redrill_rewrites_the_header_in_place():
    wb = Workbook(["Q3 Answer"])
    t = SimpleNamespace(wb=wb, app=wb.Application, sheets_added=[])
    pipeline._write_lines(t, 3, _lines(), FETCHED)
    name = pipeline._write_lines(t, 3, _lines(cube=-31.0), "2026-11-06 08:00 +0100 · tom")
    ws = wb.Worksheets(name)
    assert [s.Name for s in wb.order].count("Q3 Lines b0") == 1
    assert _v(ws, "G2").startswith("2026-11-06") and _v(ws, "D2") == -31.0
    assert ws.ListObjects("Lines_Q3_b0").HeaderRowRange.Row == 4


def test_the_lines_sheet_prints_landscape_one_page_wide_through_k():
    ws = _draw(_lines())                                  # 7 columns (A:G): the print area still reaches K
    ps = ws.PageSetup
    assert ps.PrintArea == "$A$1:$K$6"                    # header row 4 + 2 lines
    assert (ps.Orientation, ps.Zoom, ps.FitToPagesWide, ps.FitToPagesTall) == (2, False, 1, False)
    assert "&P" in ps.CenterFooter
