# engine/tests/v2/test_v2_excel_writer.py
from smartass.model.frame import Fact, Frame
from smartass.render.excel import data_grid, write_data
from smartass.render.layout import SheetPlan


class Rng:
    def __init__(self, ws, a, b):
        self.ws, self.a, self.b = ws, a, b

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
    fr = Frame([Fact("FinMax", "fm", "ac.C30", "E1", "Sep 26", "Actual", "", 5.0, "Periodic")])
    ws = WS()
    write_data(ws, fr, 3)
    kinds = [x[2] if len(x) == 4 else x[0] for x in ws.log]
    assert kinds.index("NumberFormat") < kinds.index("Value2") < kinds.index("table")
    assert ("table.Name", "Data_Q3") in ws.log


def test_answer_cells_grid_is_one_rectangle():
    from smartass.render.excel import answer_grid
    p = SheetPlan(cells={(1, 2): "T", (4, 3): "=1+1"}, last_col="C", last_row=4)
    r0, c0, g = answer_grid(p)
    assert (r0, c0) == (1, 2) and g[0][0] == "'T" and g[3][1] == "=1+1" and g[1][0] == ""


def test_write_answer_sets_value_column_widths(monkeypatch):
    from smartview_mcp import styling

    from smartass.render import excel
    seen = {}
    monkeypatch.setattr(styling, "apply_report", lambda ws, roles: seen.update(roles) or [])
    p = SheetPlan(cells={(1, 2): "T", (2, 3): 1.0, (2, 5): 2.0}, last_col="E", last_row=2)
    excel.write_answer(WS(), p, [])
    assert seen["col_widths"] == {"B": 46, "C": 16, "D": 16, "E": 16}


def test_answer_grid_stores_text_as_text_but_not_formulas_or_numbers():
    from smartass.render.excel import answer_grid
    p = SheetPlan(cells={(1, 2): "Sep 26", (1, 3): "=C2+1", (2, 3): 5.0, (2, 2): "1/2", (3, 2): ""},
                  last_col="C", last_row=3)
    _, _, g = answer_grid(p)
    assert g[0] == ["'Sep 26", "=C2+1"] and g[1] == ["'1/2", 5.0] and g[2][0] == ""
