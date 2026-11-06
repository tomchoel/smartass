# engine/tests/v2/test_v5_headers.py
"""Plan 5 Task 1: scenario headers (AC/PY/BU/FC + period), month/YTD column groups and stable key_values keys."""
import pytest
from v2util import resolver, spec_dict

from smartass.model.blocks.common import ck, col_headers, evaluate
from smartass.model.blocks.table import build_table
from smartass.model.context import Ctx
from smartass.model.errors import GuardError
from smartass.model.findings import key_values, material
from smartass.model.frame import Fact, Frame
from smartass.model.periods import period_label
from smartass.model.spec import AnalysisSpec
from smartass.model.validate import guard_result, guard_spec

OCT = (2026, 10)
LEAVES = ("un.SY.B1", "un.SY.B2", "un.SY.G910")
MON = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct")


def ytd_frame() -> Frame:
    """Jan-Oct 26 and Jan-Oct 25. Every leaf: Revenue 7/month in 26 (Oct 26: 30), 9/month in 25; Opex -5/month.
    For 2BUS (3 leaves): Oct 26 vs Oct 25 is 90 vs 27 (variance 63), YTD is 279 vs 270 (variance 9)."""
    out = []
    for yy in (26, 25):
        for m in MON:
            per = f"{m} {yy}"
            leaf = {}
            for e in LEAVES:
                rev = (30.0 if m == "Oct" else 7.0) if yy == 26 else 9.0
                leaf[e] = {"ac.C30": rev, "ac.C50S": -5.0, "ac.C60": 0.0}
            leaf["un.SY.4BUS_B"] = {a: leaf["un.SY.B1"][a] + leaf["un.SY.B2"][a] for a in ("ac.C30", "ac.C50S",
                                                                                           "ac.C60")}
            leaf["un.SY.2BUS"] = {a: leaf["un.SY.4BUS_B"][a] + leaf["un.SY.G910"][a] for a in ("ac.C30", "ac.C50S",
                                                                                               "ac.C60")}
            for e, accs in leaf.items():
                accs["ac.EBITDA"] = sum(accs.values())
                out += [Fact("FinMax", "fm", a, e, per, "Actual", "", v) for a, v in accs.items()]
    return Frame(out)


def ctx(blocks, periods="closed-ytd", **res):
    d = spec_dict(blocks=blocks)
    d["queries"][0]["periods"] = periods
    s = AnalysisSpec.model_validate(d)
    return Ctx(s, resolver(**res), OCT, ytd_frame()), s


def test_period_label():
    assert period_label([(2026, 10)]) == "Oct 26"
    assert period_label([(2026, m) for m in range(1, 11)]) == "YTD Oct 26"
    assert period_label([(2026, m) for m in range(2, 10)]) == "Feb–Sep 26"
    assert period_label([(2025, 11), (2025, 12)] + [(2026, m) for m in range(1, 11)]) == "Nov 25–Oct 26"


def test_col_headers_texts():
    assert col_headers(["value", "ly", "var", "var_pct"], period="YTD Oct 26", ly_period="YTD Oct 25", fc="") == [
        "AC YTD Oct 26", "PY YTD Oct 25", "ΔPY", "ΔPY %"]
    assert col_headers(["value", "budget", "var", "forecast", "var", "var_pct"], period="Oct 26",
                       ly_period="Oct 25", fc="Q326") == [
        "AC Oct 26", "BU Oct 26", "ΔBU", "FC Oct 26 (Q326)", "ΔFC", "ΔFC %"]


def test_single_group_headers_and_keys_unchanged():
    c, s = ctx([{"type": "table", "query": "fm", "title": "P&L"}])
    out = build_table(s.blocks[0], c, "b0")
    assert out.headers == ["AC YTD Oct 26", "PY YTD Oct 25", "ΔPY", "ΔPY %", "Material"]
    assert out.keys == ["Actual", "Last year", "Variance", "Variance %", ""]
    kv = key_values([out], evaluate([out], c.frame))
    assert list(kv["P&L · Revenue"]) == ["Actual", "Last year", "Variance", "Variance %"]


def test_plan_headers_budget_and_forecast():
    d = spec_dict(blocks=[{"type": "table", "query": "fp", "metrics": ["EBITDA"],
                           "columns": ["value", "budget", "var", "forecast", "var", "var_pct"]}])
    d["queries"] = [{"id": "fp", "cube": "FinPlan", "metrics": ["EBITDA"], "entity": "un.SY.2BUS",
                     "periods": "closed-ytd", "compare": ["budget", "forecast"]}]
    s = AnalysisSpec.model_validate(d)
    c = Ctx(s, resolver(), OCT, Frame([]))
    out = build_table(s.blocks[0], c, "b0")
    assert out.headers == ["AC YTD Oct 26", "BU YTD Oct 26", "ΔBU", "FC YTD Oct 26 (Q3 26)", "ΔFC", "ΔFC %", "Material"]
    assert out.keys == ["Actual", "Budget", "Variance vs Budget", "Forecast", "Variance vs Forecast",
                        "Variance % vs Forecast", ""]


GROUPED = {"type": "table", "query": "fm", "title": "P&L", "periods": ["closed", "closed-ytd"]}


def test_month_and_ytd_groups():
    c, s = ctx([GROUPED])
    out = build_table(s.blocks[0], c, "b0")
    assert out.cols == ["value", "ly", "var", "var_pct", "gap", "value", "ly", "var", "var_pct", "material"]
    assert out.headers == ["AC Oct 26", "PY Oct 25", "ΔPY", "ΔPY %", "",
                           "AC YTD Oct 26", "PY YTD Oct 25", "ΔPY", "ΔPY %", "Material"]
    assert [ln.key for ln in out.lines] == ["m0", "m1", "m2"]
    cells = evaluate([out], c.frame)
    rev = [cells.get(ck("b0", "m0", j)) for j in range(9)]
    assert rev[:3] == [90.0, 27.0, 63.0]                     # Oct 26 vs Oct 25 only
    assert rev[4] is None and out.lines[0].cells[4] is None  # the gap
    assert rev[5:8] == [279.0, 270.0, 9.0]                   # Jan-Oct
    assert rev[8] == pytest.approx(9 / 270)


def test_var_in_group_two_references_group_two():
    c, s = ctx([GROUPED])
    out = build_table(s.blocks[0], c, "b0")
    var2 = out.lines[0].cells[7]
    refs = {r.key for r in (var2.a, var2.b)}
    assert refs == {ck("b0", "m0", 5), ck("b0", "m0", 6)}


def test_keys_month_group_prefixed_ytd_plain():
    c, s = ctx([GROUPED])
    out = build_table(s.blocks[0], c, "b0")
    kv = key_values([out], evaluate([out], c.frame))
    assert list(kv["P&L · Revenue"]) == ["closed · Actual", "closed · Last year", "closed · Variance",
                                          "closed · Variance %", "Actual", "Last year", "Variance", "Variance %"]
    assert kv["P&L · Revenue"]["Actual"] == 279.0 and kv["P&L · Revenue"]["closed · Actual"] == 90.0


def test_grouped_entity_rows_check_reconciles():
    c, s = ctx([{"type": "table", "query": "fm", "rows": "entities", "metric": "Revenue", "top": 1,
                 "columns": ["value", "ly", "var", "share"], "periods": ["closed", "closed-ytd"]}])
    out = build_table(s.blocks[0], c, "b0")
    cells = evaluate([out], c.frame)
    guard_result([out], cells)
    other = next(ln for ln in out.lines if ln.key == "other")
    assert other.cells[4] is None
    total = next(ln for ln in out.lines if ln.key == "total")
    assert [cells.get(ck("b0", "total", j)) for j in (0, 5)] == [90.0, 279.0]
    assert cells.get(ck("b0", "e0", 3)) == pytest.approx(60 / 90)   # share within the month group
    assert total.cells[4] is None


def test_group_outside_query_raises():
    c, s = ctx([{"type": "table", "query": "fm", "periods": ["closed", "closed-ytd"]}], periods="closed")
    with pytest.raises(GuardError) as e:
        guard_spec(s, c.resolver, OCT)
    assert "closed-ytd" in str(e.value) and "fm" in str(e.value)
    assert "widen the query periods" in e.value.fix


class Mat:
    """The resolver with Revenue's materiality replaced."""

    def __init__(self, r, rev):
        self.r, self.rev = r, rev

    def materiality(self, metric):
        return self.rev if metric == "Revenue" else self.r.materiality(metric)


def test_material_judged_on_ytd_group_once():
    c, s = ctx([GROUPED])
    out = build_table(s.blocks[0], c, "b0")
    cells = evaluate([out], c.frame)
    both = [h for h in material([out], cells, Mat(c.resolver, (5.0, 0.5))) if h["line"] == "Revenue"]
    assert len(both) == 1 and both[0]["variance"] == 9.0          # material in both groups: listed once, YTD
    month_only = [h for h in material([out], cells, Mat(c.resolver, (50.0, 0.5))) if h["line"] == "Revenue"]
    assert month_only == []                                       # month var 63 is material, YTD var 9 is not
