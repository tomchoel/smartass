"""Plan 5 Task 10: provenance and tie-out on the 'Q<n> Lines <alias>' sheets. The header sits in F1:K3: the lines
table header is row 4, so rows 1-3 are the only free rows; A1, A2, the cube cell and the table never move (the built
sheet itself: test_xlsx.py). Every figure is invented."""
from test_v2_pipeline import LINE_CELL

from smartass.model.lines import AMOUNT
from smartass.render.datasheet import lines_header

FETCHED = "2026-11-05 09:30 +0100 · tom"
TIE = '=IF(NOT(ISNUMBER($D$3)),"no cube value",IF(ABS($G$3-$D$3)<=$K$3,"OK","DIFF"))'


def test_lines_header_is_label_value_pairs_with_the_exact_formulas():
    g = lines_header("Lines_Q3_b0", LINE_CELL, FETCHED)
    assert g[0][:2] == ["Drill cell", "All Accounts: ac.C60 · All Entities: un.SY.2TEC · All Time Periods: Yr 26"]
    assert g[1][:2] == ["Fetched", FETCHED]
    assert g[2] == ["Lines total", f"=SUM(Lines_Q3_b0[{AMOUNT}])", "Tie-out", TIE, "Tolerance (NOK)", 1.0]
    assert all(len(r) == 6 for r in g)

