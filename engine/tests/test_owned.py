"""The tool-owned range of an answer sheet: a redraw clears only what the tool drew; user content survives."""
from types import SimpleNamespace

from smartass.xl import owned


def test_growth_inserts_rows_only_when_the_band_below_holds_user_content():
    used = lambda r1, c1, r2, c2: (r1, c1, r2, c2) == (31, 1, 40, 9)         # user text in A31:I40
    assert owned.growth({"rows": 30, "cols": 9}, {"rows": 40, "cols": 9}, used) == {"rows": (31, 40), "cols": None}
    assert owned.growth({"rows": 30, "cols": 9}, {"rows": 40, "cols": 9}, lambda *a: False) == \
        {"rows": None, "cols": None}
    assert owned.growth({"rows": 30, "cols": 9}, {"rows": 20, "cols": 9}, used) == {"rows": None, "cols": None}
    assert owned.growth(None, {"rows": 40, "cols": 9}, used) == {"rows": None, "cols": None}


def test_growth_inserts_columns_when_the_band_to_the_right_holds_user_content():
    used = lambda r1, c1, r2, c2: c1 == 10
    g = owned.growth({"rows": 30, "cols": 9}, {"rows": 30, "cols": 12}, used)
    assert g == {"rows": None, "cols": (10, 12)}


class Shape:
    def __init__(self, name):
        self.Name, self.deleted = name, False

    def Delete(self):
        self.deleted = True


def test_only_smartass_charts_are_deleted():
    mine, users = Shape(owned.CHART_NAME), Shape("Chart 7")
    ws = SimpleNamespace(ChartObjects=lambda: [mine, users])
    owned.drop_charts(ws, com=lambda call, **kw: call())
    assert mine.deleted and not users.deleted


def test_without_an_owned_range_the_whole_sheet_is_cleared_as_before():
    calls = []
    ws = SimpleNamespace(Cells=SimpleNamespace(Clear=lambda: calls.append("all")),
                         Range=lambda a, b: SimpleNamespace(Clear=lambda: calls.append(("block", a, b))),
                         ChartObjects=lambda: [])
    ws_cells = lambda r, c: (r, c)
    owned.clear(ws, None, com=lambda call, **kw: call(), cell=ws_cells)
    owned.clear(ws, {"rows": 12, "cols": 6}, com=lambda call, **kw: call(), cell=ws_cells)
    assert calls == ["all", ("block", (1, 1), (12, 6))]


class _Sheet:
    def __init__(self, log):
        self.log = log

    def Cells(self, r, c):
        return SimpleNamespace(r=r, c=c)

    def Range(self, a, b):
        log = self.log

        class R:
            def __setattr__(self, k, v):
                log.append(("write", k))
        return R()
