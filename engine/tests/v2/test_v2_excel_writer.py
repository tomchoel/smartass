# engine/tests/v2/test_v2_excel_writer.py
from smartass.model.frame import Fact, Frame
from smartass.render.excel import data_grid, write_data
from smartass.render.layout import SheetPlan


class Rng:
    def __init__(self, ws, a, b):
        self.ws, self.a, self.b = ws, a, b

    def Clear(self):
        self.ws.log.append(("clear", self.a, self.b))

    def __setattr__(self, k, v):
        if k in ("ws", "a", "b"):
            object.__setattr__(self, k, v)
        else:
            self.ws.log.append((self.a, self.b, k, v))


class Lists:
    def __init__(self, ws):
        self.ws = ws

    def Add(self, src_type, rng, link, headers):
        self.ws.log.append(("table", rng.a, rng.b))
        ws = self.ws

        class T:
            def __setattr__(self, k, v):
                ws.log.append(("table." + k, v))
        return T()


class WS:
    def __init__(self):
        self.log = []
        self.ListObjects = Lists(self)

    @property
    def UsedRange(self):
        return Rng(self, "used", "used")

    def Cells(self, r, c):
        return (r, c)

    def Range(self, a, b=None):
        return Rng(self, a, b if b is not None else a)


def test_data_grid_header_and_rows():
    fr = Frame([Fact("FinMax", "fm", "ac.C30", "E1", "Sep 26", "Actual", "", 5.0, "Periodic")])
    g = data_grid(fr)
    assert g[0] == ("Cube", "Query", "Account", "Entity", "Period", "Scenario", "Version", "Balance", "Value")
    assert g[1][4] == "Sep 26" and g[1][-1] == 5.0


def test_write_data_sets_text_format_then_values_then_table():
    from fakexl import Workbook
    fr = Frame([Fact("FinMax", "fm", "ac.C30", "E1", "Sep 26", "Actual", "", 5.0, "Periodic")])
    ws = Workbook(["Q3 Data"]).order[0]
    write_data(ws, fr, 3)
    kinds = [x[0] for x in ws.log]
    assert kinds.index("NumberFormat") < kinds.index("Value2") < kinds.index("ListObjects.Add")
    assert [lo.Name for lo in ws.ListObjects] == ["Data_Q3"]


def test_answer_cells_grid_is_one_rectangle():
    from smartass.render.excel import answer_grid
    p = SheetPlan(cells={(1, 2): "T", (4, 3): "=1+1"}, last_col="C", last_row=4)
    r0, c0, g = answer_grid(p)
    assert (r0, c0) == (1, 2) and g[0][0] == "'T" and g[3][1] == "=1+1" and g[1][0] == ""


def test_write_answer_sets_value_column_widths(monkeypatch):
    from smartass.xl import styling

    from smartass.render import excel
    seen = {}
    monkeypatch.setattr(styling, "apply_report", lambda ws, roles: seen.update(roles) or [])
    p = SheetPlan(cells={(1, 2): "T", (2, 4): 1.0, (2, 6): 2.0}, last_col="F", last_row=2)
    excel.write_answer(WS(), p, [])
    assert seen["col_widths"] == {"A": 2, "B": 9, "C": 46, "D": 16, "E": 16, "F": 16}
    assert seen["label_cols"] == ["B", "C"] and seen["code_cols"] == ["B"]
    p.widths = {"E": 2, "F": 9}                                   # the layout's gap and Material widths win
    excel.write_answer(WS(), p, [])
    assert seen["col_widths"]["E"] == 2 and seen["col_widths"]["F"] == 9


def test_answer_grid_stores_text_as_text_but_not_formulas_or_numbers():
    from smartass.render.excel import answer_grid
    p = SheetPlan(cells={(1, 2): "Sep 26", (1, 3): "=C2+1", (2, 3): 5.0, (2, 2): "1/2", (3, 2): ""},
                  last_col="C", last_row=3)
    _, _, g = answer_grid(p)
    assert g[0] == ["'Sep 26", "=C2+1"] and g[1] == ["'1/2", 5.0] and g[2][0] == ""


class LO:
    def __init__(self, log, rows):
        self.log, self.rows = log, rows

    @property
    def HeaderRowRange(self):                  # the table's own origin; A1 here (row 14: test_v5_datasheet)
        return type("H", (), {"Row": 1, "Column": 1})()

    @property
    def Range(self):
        rows = self.rows

        class R:
            Rows = type("Rw", (), {"Count": rows})()
            Value2 = (("Cube", "Value"), ("FinMax", 1.0))
        return R()

    def Resize(self, rng):
        self.log.append(("resize", rng.a, rng.b))


class WSLO(WS):
    def __init__(self, rows=3, table=True):
        super().__init__()
        lo = LO(self.log, rows)

        class Lists2:
            def __call__(_, name):
                if not table:
                    raise RuntimeError("no such table")
                self.log.append(("lookup", name))
                return lo

            def Add(_, *a):
                return Lists(self).Add(*a)
        self.ListObjects = Lists2()


def test_rewrite_grid_resizes_the_existing_table_then_writes():
    from smartass.render.excel import rewrite_grid
    fr = Frame([Fact("FinMax", "fm", "ac.C30", "E1", "Sep 26", "Actual", "", 5.0, "Periodic"),
                Fact("FinMax", "fm", "ac.C30", "E2", "Sep 26", "Actual", "", 6.0, "Periodic")])
    ws = WSLO(rows=2)
    rewrite_grid(ws, data_grid(fr), 3)
    kinds = [x[2] if len(x) == 4 else x[0] for x in ws.log]
    assert ("lookup", "Data_Q3") in ws.log and ("resize", (1, 1), (3, 9)) in ws.log
    assert kinds.index("NumberFormat") < kinds.index("resize") < kinds.index("Value2")
    assert "table" not in kinds                                      # no new ListObject: the name never changes


def test_rewrite_grid_clears_leftover_rows_when_the_table_shrinks():
    from smartass.render.excel import rewrite_grid
    fr = Frame([Fact("FinMax", "fm", "ac.C30", "E1", "Sep 26", "Actual", "", 5.0, "Periodic")])
    ws = WSLO(rows=6)                                                # old table: header + 5 rows, new: header + 1
    rewrite_grid(ws, data_grid(fr), 3)
    assert ("clear", (3, 1), (6, 9)) in ws.log



def _value2_writes(ws):
    return [(a, b) for x in ws.log if len(x) == 4 and x[2] == "Value2" for a, b in [(x[0], x[1])]]


def test_rewrite_never_writes_header_and_body_in_one_rectangle():
    """Measured in Excel 365 (04.11.2026): Value2 on a rectangle covering a ListObject's header AND body deletes the
    table (ListObjects.Count 1 -> 0), even when the values are its own; every formula on it becomes #REF!. Header
    alone or body alone keeps it. So the header and the body are written separately."""
    from smartass.render.excel import rewrite_grid
    fr = Frame([Fact("FinMax", "fm", "ac.C30", "E1", "Sep 26", "Actual", "", 5.0, "Periodic"),
                Fact("FinMax", "fm", "ac.C30", "E2", "Sep 26", "Actual", "", 6.0, "Periodic")])
    ws = WSLO(rows=3)
    rewrite_grid(ws, data_grid(fr), 3)
    writes = _value2_writes(ws)
    assert writes, "nothing written"
    assert not [w for w in writes if w[0][0] == 1 and w[1][0] >= 2], writes     # never header + body together
    assert ((1, 1), (1, 9)) in writes and ((2, 1), (3, 9)) in writes


def test_rewrite_grid_without_the_table_refuses_and_leaves_the_sheet_alone():
    import pytest

    from smartass.model.errors import GuardError
    from smartass.render.excel import rewrite_grid
    fr = Frame([Fact("FinMax", "fm", "ac.C30", "E1", "Sep 26", "Actual", "", 5.0, "Periodic")])
    ws = WSLO(table=False)
    with pytest.raises(GuardError, match="Data_Q3"):
        rewrite_grid(ws, data_grid(fr), 3)
    assert ws.log == []                                              # no clear, no new table, nothing written


def _widths_for(monkeypatch, cells, rows):
    from smartass.xl import styling

    from smartass.render import excel
    seen = {}
    monkeypatch.setattr(styling, "apply_report", lambda ws, roles: seen.update(roles) or [])
    excel.write_answer(WS(), SheetPlan(cells=cells, rows=rows, last_col="D", last_row=3), [])
    return seen["col_widths"]


def test_the_label_column_fits_a_long_line_label(monkeypatch):
    """Plan 2D task 2 (plan 5: the label is in C): C = max(46, min(80, longest line label + 2))."""
    label = "   " + "x" * 57                                       # 60 characters with the level indent
    w = _widths_for(monkeypatch, {(1, 2): "T", (2, 3): label, (2, 4): 1.0, (3, 2): "y" * 120},
                    [{"row": 1, "kind": "title"}, {"row": 2, "kind": "line", "num": {}},
                     {"row": 3, "kind": "note"}])
    assert w == {"A": 2, "B": 9, "C": 62, "D": 16}                   # the 120-char note does not count


def test_the_label_column_stays_46_and_stops_at_80(monkeypatch):
    short = _widths_for(monkeypatch, {(2, 3): "Revenue", (2, 4): 1.0}, [{"row": 2, "kind": "line", "num": {}}])
    assert short["C"] == 46
    long = _widths_for(monkeypatch, {(2, 3): "z" * 100, (2, 4): 1.0}, [{"row": 2, "kind": "grand", "num": {}}])
    assert long["C"] == 80
