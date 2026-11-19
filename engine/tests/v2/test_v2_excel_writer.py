# engine/tests/v2/test_v2_excel_writer.py
"""What render/excel.py still gives the built file: the Data grid, the Answer's cells as one rectangle (text as text)
and the column widths of its apply_report roles."""
from smartass.model.frame import Fact, Frame
from smartass.render.excel import answer_roles, data_grid
from smartass.render.layout import SheetPlan


def test_data_grid_header_and_rows():
    fr = Frame([Fact("FinMax", "fm", "ac.C30", "E1", "Sep 26", "Actual", "", 5.0, "Periodic")])
    g = data_grid(fr)
    assert g[0] == ("Cube", "Query", "Account", "Entity", "Period", "Scenario", "Version", "Balance", "Value")
    assert g[1][4] == "Sep 26" and g[1][-1] == 5.0


def test_answer_cells_grid_is_one_rectangle():
    from smartass.render.excel import answer_grid
    p = SheetPlan(cells={(1, 2): "T", (4, 3): "=1+1"}, last_col="C", last_row=4)
    r0, c0, g = answer_grid(p)
    assert (r0, c0) == (1, 2) and g[0][0] == "'T" and g[3][1] == "=1+1" and g[1][0] == ""


def test_the_answer_roles_set_the_value_column_widths():
    p = SheetPlan(cells={(1, 2): "T", (2, 4): 1.0, (2, 6): 2.0}, last_col="F", last_row=2)
    roles = answer_roles(p)
    assert roles["col_widths"] == {"A": 2, "B": 9, "C": 46, "D": 16, "E": 16, "F": 16}
    assert roles["label_cols"] == ["B", "C"] and roles["code_cols"] == ["B"]
    p.widths = {"E": 2, "F": 9}                                   # the layout's gap and Material widths win
    assert answer_roles(p)["col_widths"]["E"] == 2 and answer_roles(p)["col_widths"]["F"] == 9


def test_answer_grid_stores_text_as_text_but_not_formulas_or_numbers():
    from smartass.render.excel import answer_grid
    p = SheetPlan(cells={(1, 2): "Sep 26", (1, 3): "=C2+1", (2, 3): 5.0, (2, 2): "1/2", (3, 2): ""},
                  last_col="C", last_row=3)
    _, _, g = answer_grid(p)
    assert g[0] == ["'Sep 26", "=C2+1"] and g[1] == ["'1/2", 5.0] and g[2][0] == ""


def _widths_for(cells, rows):
    return answer_roles(SheetPlan(cells=cells, rows=rows, last_col="D", last_row=3))["col_widths"]


def test_the_label_column_fits_a_long_line_label():
    """Plan 2D task 2 (plan 5: the label is in C): C = max(46, min(80, longest line label + 2))."""
    label = "   " + "x" * 57                                       # 60 characters with the level indent
    w = _widths_for({(1, 2): "T", (2, 3): label, (2, 4): 1.0, (3, 2): "y" * 120},
                    [{"row": 1, "kind": "title"}, {"row": 2, "kind": "line", "num": {}},
                     {"row": 3, "kind": "note"}])
    assert w == {"A": 2, "B": 9, "C": 62, "D": 16}                   # the 120-char note does not count


def test_the_label_column_stays_46_and_stops_at_80():
    short = _widths_for({(2, 3): "Revenue", (2, 4): 1.0}, [{"row": 2, "kind": "line", "num": {}}])
    assert short["C"] == 46
    long = _widths_for({(2, 3): "z" * 100, (2, 4): 1.0}, [{"row": 2, "kind": "grand", "num": {}}])
    assert long["C"] == 80
