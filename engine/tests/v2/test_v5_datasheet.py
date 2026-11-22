"""Plan 5 Task 9: the Q<n> Data provenance header (its rows, labels and formulas; the built file's look of it is in
test_xlsx.py). Fakes only, no Excel. Every figure is invented."""
import datetime

import pytest
from fakexl import ComError, Workbook
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass import pipeline
from smartass.model.frame import Frame
from smartass.render.datasheet import REFRESH_TEXT, header_rows

WHEN = datetime.datetime(2026, 11, 5, 9, 30, tzinfo=datetime.timezone(datetime.timedelta(hours=1)))
LABELS = ["Title", "Question", "Query", "POV", "Periods", "Route", "Fetched", "Rows", "Control total", "Sources",
          "Tolerance (NOK)", "Refresh"]


def _pre(**over):
    return pipeline.prepare(spec_dict(**over), resolver(), FakeProvider(facts()))


def _header(q=3, **over):
    pre = _pre(**over)
    return pre, header_rows(q, pre.spec, pre, WHEN, "tom")


def _ws():
    return Workbook(["Q3 Data"]).order[0]


def _row(header, label):
    return next(i + 1 for i, (lab, _) in enumerate(header) if lab == label)


def test_header_rows_come_in_order_with_the_exact_labels():
    _, h = _header()
    assert [lab for lab, _ in h] == LABELS
    v = dict(h)
    assert v["Title"] == "Q3 Data — How is Business opex developing vs last year?"
    assert v["Question"] == "Q3: How is Business opex developing vs last year?"
    assert v["Query"] == "fm: FinMax · Actual"
    assert v["POV"] == "2BUS Business Norway (un.SY.2BUS) · NOK · Periodic"
    assert v["Periods"] == "Sep 26 (closed)"
    assert v["Route"] == "HTTP"
    assert v["Fetched"] == "2026-11-05 09:30 +0100 · tom"
    assert v["Sources"] == "" and v["Tolerance (NOK)"] == 1 and v["Refresh"] == REFRESH_TEXT
    assert REFRESH_TEXT == "Refresh: fin_refresh, or ask a follow-up. Do not edit this table by hand."


def test_finplan_query_lists_scenario_and_version_per_series_and_the_ytd_span():
    spec = spec_dict()
    spec["queries"] = [{"id": "fp", "cube": "FinPlan", "metrics": ["Revenue"], "entity": "un.SY.2BUS",
                        "periods": "closed-ytd", "compare": ["ly", "budget"]}]
    spec["blocks"] = [{"type": "table", "query": "fp", "title": "P&L", "columns": ["value", "budget"]}]
    pre = pipeline.prepare(spec, resolver(), FakeProvider(Frame()))
    v = dict(header_rows(3, pre.spec, pre, WHEN, "tom"))
    assert v["Query"] == "fp: FinPlan · Actual/Working, Target/Final"
    assert v["Periods"] == "Jan 26 – Sep 26 (closed-ytd)"


def test_route_names_the_ad_hoc_fallback_and_failed_grids():
    pre = _pre()
    first = next(iter(pre.fetched))
    pre.fetched[first] = type(pre.fetched[first])(first, [], [], "adhoc")
    assert dict(header_rows(3, pre.spec, pre, WHEN, "tom"))["Route"] == "HTTP + Smart View ad hoc"
    pre.failed = [first]
    assert dict(header_rows(3, pre.spec, pre, WHEN, "tom"))["Route"].endswith("HTTP failed for 1 grid(s)")


def test_route_names_the_second_engine():
    pre = _pre()
    first = next(iter(pre.fetched))
    pre.fetched[first] = type(pre.fetched[first])(first, [], [], "rest")
    assert dict(header_rows(3, pre.spec, pre, WHEN, "tom"))["Route"] == "HTTP + Planning REST"


def test_control_total_is_the_static_sum_beside_live_formulas():
    pre, h = _header()
    r, tol = _row(h, "Control total"), _row(h, "Tolerance (NOK)")
    static, live, status = dict(h)["Control total"]
    assert static == pytest.approx(sum(f.value for f in pre.frame.facts if f.query == "fm"))
    assert live == '=SUMIFS(Data_Q3[Value],Data_Q3[Query],"fm")'
    assert status == f'=IF(ABS(B{r}-C{r})<=$B${tol},"OK","DIFF")'
    assert dict(h)["Rows"] == '=COUNTIFS(Data_Q3[Query],"fm")'


def test_one_block_per_query_and_the_cells_task_11_needs():
    spec = spec_dict()
    spec["queries"].append({"id": "fm2", "cube": "FinMax", "metrics": ["Opex"], "entity": "un.SY.2BUS",
                            "periods": "closed"})
    pre = pipeline.prepare(spec, resolver(), FakeProvider(facts()))
    h = header_rows(3, pre.spec, pre, WHEN, "tom")
    labels = [lab for lab, _ in h]
    assert labels.count("Query") == 2 and labels.count("Sources") == 2 and labels.count("Tolerance (NOK)") == 1
    rows = [i + 1 for i, lab in enumerate(labels) if lab == "Sources"]
    assert h.sources == {"fm": f"B{rows[0]}", "fm2": f"B{rows[1]}"}
    cts = [i + 1 for i, lab in enumerate(labels) if lab == "Control total"]
    assert h.status == {"fm": f"D{cts[0]}", "fm2": f"D{cts[1]}"}
    assert dict(h)["Control total"][0] is not None


def test_the_fake_refuses_a_formula_on_a_missing_table():
    ws = _ws()
    with pytest.raises(ComError):
        ws.Range("B1").Formula = '=COUNTIFS(Data_Q3[Query],"fm")'

