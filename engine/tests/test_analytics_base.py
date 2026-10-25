"""analytics: facts, periods and frames (synthetic figures only)."""
import pytest

from smartview_mcp import catalog
from smartview_mcp.analytics import periods
from smartview_mcp.analytics.facts import Fact, combine, missing, total
from smartview_mcp.analytics.frame import from_rows


def test_combine_and_total_carry_inputs_and_reasons():
    a = Fact("a", 10.0, "nok", "A", (("k1",),))
    b = Fact("b", 4.0, "nok", "B", (("k2",), ("k1",)))
    d = combine("d", "nok", "a - b", lambda x, y: x - y, a, b)
    assert (d.value, d.inputs) == (6.0, (("k1",), ("k2",)))
    gone = missing("m", "nok", "M", "not read")
    assert combine("x", "nok", "a/m", lambda x, y: x / y, a, gone).reason == "m: not read"
    zero = Fact("z", 0.0, "nok", "Z")
    q = combine("q", "pct", "a/z", lambda x, y: x / y, a, zero)
    assert q.value is None and q.reason == "division by zero"
    assert total("t", [a, b], "nok", "sum").value == 14.0 and total("t", [], "nok", "none").value == 0.0
    assert total("t", [a, gone], "nok", "sum").value is None
    assert d.as_dict()["inputs"] == [["k1"], ["k2"]] and "reason" not in d.as_dict()


def test_periods():
    assert periods.shift("Jan 26", -1) == "Dec 25" and periods.shift("Nov 26", 3) == "Feb 27"
    assert periods.ytd("Mar 26") == ["Jan 26", "Feb 26", "Mar 26"]
    assert periods.ly("Sep 26") == "Sep 25" and periods.opening("Sep 26") == "Dec 25"
    assert periods.ltm("Sep 26")[0] == "Oct 25" and periods.ltm("Sep 26")[-1] == "Sep 26"
    assert periods.remaining("Sep 26") == ["Oct 26", "Nov 26", "Dec 26"] and periods.remaining("Dec 26") == []
    w = periods.window("Sep 26")
    assert len(w) == 36 and w[0] == "Oct 23" and w[-1] == "Sep 26"
    assert periods.between("Sep 26", "Nov 26") == ["Oct 26", "Nov 26"] and periods.between("Sep 26", "Sep 26") == []
    assert periods.after("Oct 26", "Sep 26") and not periods.after("Sep 26", "Sep 26")
    assert periods.finplan("Sep 26") == ("Sep", "FY26") and periods.from_finplan("Sep", "FY26") == "Sep 26"
    assert periods.year(25)[0] == "Jan 25" and len(periods.year(25)) == 12


def test_frame_takes_flows_periodic_and_stocks_accumulated():
    lines = catalog.load().pnl["FinMax"]
    rows = {("FinMax", "Actual", "", "E", "ac.C30", "Sep 26", "Periodic"): 100.0,
            ("FinMax", "Actual", "", "E", "ac.C30", "Sep 26", "Accumulated"): 900.0,     # not a flow's month
            ("FinMax", "Actual", "", "E", "ac.STA20Y", "Sep 26", "Accumulated"): 12.5,
            ("FinMax", "Actual", "", "E", "ac.STA20Y", "Sep 26", "Periodic"): 0.5,        # a stock's movement
            ("FinMax", "Actual", "", "E", "ac.C40S", "Sep 26", "Periodic"): None,
            ("FinMax", "Budget", "", "E", "ac.C30", "Sep 26", "Periodic"): 1.0,           # unmapped scenario
            ("FinMax", "Actual", "", "E", "ac.XYZ", "Sep 26", "Periodic"): 1.0}           # not a catalogue line
    f = from_rows("FinMax", lines, rows, {("Actual", ""): "actual"})
    assert f.get("E", "rev", "Sep 26", "actual") == (100.0, ("FinMax", "Actual", "", "E", "ac.C30", "Sep 26",
                                                             "Periodic"))
    assert f.get("E", "fte", "Sep 26", "actual")[0] == 12.5
    assert f.get("E", "cogs", "Sep 26", "actual")[0] == 0.0          # read, but blank: counts as 0
    assert f.get("E", "rev", "Aug 26", "actual") is None             # never read
    assert len(f.values) == 3
    with pytest.raises(KeyError, match="no line 'nope'"):
        f.line("nope")
