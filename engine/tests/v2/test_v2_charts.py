# engine/tests/v2/test_v2_charts.py
from smartass.render.charts import CHART_NAME, XL_LINE, XL_WATERFALL, write_charts
from smartass.render.layout import ChartSpec, SheetPlan


class Obj:
    """Records attribute sets and calls; any attribute is another Obj."""
    def __init__(self, log, name):
        object.__setattr__(self, "_log", log)
        object.__setattr__(self, "_name", name)

    def __getattr__(self, k):
        child = Obj(self._log, f"{self._name}.{k}")
        object.__setattr__(self, k, child)
        return child

    def __setattr__(self, k, v):
        self._log.append(("set", f"{self._name}.{k}", v))

    def __call__(self, *a):
        self._log.append(("call", self._name, a))
        return Obj(self._log, f"{self._name}()")


class WS(Obj):
    def Range(self, a):
        self._log.append(("range", a))
        return Obj(self._log, f"Range({a})")

    def Cells(self, r, c):
        return Obj(self._log, f"Cells({r},{c})")


def test_line_and_waterfall_are_drawn_and_named():
    log = []
    ws = WS(log, "ws")
    p = SheetPlan(charts=[
        ChartSpec("line", "Trend", 10, [11, 12, 13], ["C", "D"], [], "G"),
        ChartSpec("waterfall", "EBITDA bridge", 20, [21, 22, 23, 24, 25], ["C"], [1, 5], "G")])
    warnings = write_charts(ws, p)
    assert warnings == []
    adds = [e for e in log if e[0] == "call" and e[1].endswith("Shapes.AddChart2")]
    assert [a[2][1] for a in adds] == [XL_LINE, XL_WATERFALL]
    assert ("range", "C11:C13") in log and ("range", "B11:B13") in log and ("range", "D11:D13") in log
    assert ("range", "B21:C25") in log
    assert any(e[0] == "set" and e[1].endswith(".Name") and e[2] == CHART_NAME for e in log)


class Coll:
    def __init__(self, log, n):
        self.log, self.n = log, n

    @property
    def Count(self):
        return self.n

    def NewSeries(self):
        self.log.append("new_series")
        self.n += 1
        return Obj(self.log, "series")


class FakeChart(Obj):
    def __init__(self, log, auto=2):
        super().__init__(log, "chart")
        object.__setattr__(self, "coll", Coll(log, auto))

    def SeriesCollection(self, i=None):
        if i is None:
            return self.coll
        coll = self.coll
        log = self._log

        class Item:
            def Delete(_):
                log.append("delete_auto")
                coll.n -= 1
        return Item()


def test_line_chart_drops_the_auto_added_series_before_its_own():
    log = []
    chart = FakeChart(log, auto=2)

    class ShapesFake:
        def AddChart2(self, *a):
            return type("Sh", (), {"Chart": chart})()

    class W(WS):
        Shapes = ShapesFake()

        def Range(self, a):
            return Obj(log, f"Range({a})")

        def Cells(self, r, c):
            return Obj(log, f"Cells({r},{c})")
    ws = W(log, "ws")
    p = SheetPlan(charts=[ChartSpec("line", "Trend", 10, [11, 12, 13], ["C", "D"], [], "G")])
    assert write_charts(ws, p) == []
    seq = [x for x in log if x in ("delete_auto", "new_series")]
    assert seq == ["delete_auto", "delete_auto", "new_series", "new_series"]
