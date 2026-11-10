"""Plan 5 Task 9: the Q<n> Data provenance header, the calm table style, and rewrites from the table's own origin
(fakes only, no Excel). Every figure is invented."""
import datetime

import pytest
from fakexl import ComError, Workbook
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass import pipeline
from smartass.model.frame import Fact, Frame
from smartass.render import datasheet
from smartass.render.datasheet import REFRESH_TEXT, header_rows
from smartass.render.excel import data_grid, rewrite_data, rewrite_grid, write_data

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


def test_write_data_puts_the_header_then_the_table_at_h_plus_2():
    pre, h = _header()
    ws = _ws()
    out = write_data(ws, pre.frame, 3, h)
    lo = ws.ListObjects("Data_Q3")
    assert lo.HeaderRowRange.Row == len(h) + 2 and lo.HeaderRowRange.Column == 1
    assert out["table_row"] == len(h) + 2 and out["sources"] == h.sources and out["status"] == h.status
    assert ws.cells[(1, 1)]["Value"] == dict(h)["Title"]
    assert ws.cells[(2, 1)]["Value"] == "Question"
    r = _row(h, "Rows")
    assert ws.cells[(r, 2)]["Value"] == '=COUNTIFS(Data_Q3[Query],"fm")'
    assert (len(h) + 1, 1) not in ws.cells or "Value" not in ws.cells[(len(h) + 1, 1)]   # one blank row
    assert ws.cells[(len(h) + 2, 1)]["Value"] == "Cube"
    assert ws.cells[(len(h) + 3, 1)]["Value"] == "FinMax"
    add = next(i for i, x in enumerate(ws.log) if x[0] == "ListObjects.Add")
    first_formula = next(i for i, x in enumerate(ws.log) if x[0] == "Formula")
    assert add < first_formula                                   # header formulas only once the table exists


def test_the_table_style_and_the_value_format():
    pre, h = _header()
    ws = _ws()
    write_data(ws, pre.frame, 3, h)
    lo = ws.ListObjects("Data_Q3")
    assert lo.TableStyle == "TableStyleLight1" and lo.ShowTableStyleRowStripes is False
    t = len(h) + 2
    value_col = len(data_grid(pre.frame)[0])
    cell = ws.cells[(t + 1, value_col)]
    assert cell["NumberFormat"] == '#,##0;(#,##0);"–"'
    assert cell["Font.Color"] == 0xC00000 and cell["HorizontalAlignment"] == -4152
    assert ws.cells[(t + 1, 1)]["NumberFormat"] == "@"           # label columns are text
    assert all(w <= 40 for w in ws.colw.values()) and ws.colw[1] >= 16
    assert ws.cells[(1, 1)]["Font.Bold"] is True


def test_rewrite_with_the_table_at_row_14_writes_from_row_14():
    ws = _ws()
    fr = Frame([Fact("FinMax", "fm", "ac.C30", "E1", "Sep 26", "Actual", "", 5.0, "Periodic")])
    g = data_grid(fr)
    rng = ws.Range(ws.Cells(14, 1), ws.Cells(15, len(g[0])))
    rng.Value2 = tuple(g)
    lo = ws.ListObjects.Add(1, rng, None, 1)
    lo.Name = "Data_Q3"
    ws.log.clear()
    g2 = [g[0], g[1], g[1]]
    rewrite_grid(ws, g2, 3)
    assert ws.ListObjects("Data_Q3") is lo and lo.rect == (14, 1, 16, len(g[0]))
    writes = [x[1] for x in ws.log if x[0] == "Value2"]
    assert writes == ["A14:I14", "A15:I16"]                    # header and body apart, from the table origin
    assert ws.cells[(16, 9)]["Value"] == 5.0 and (1, 1) not in ws.cells


def test_an_old_a1_table_gets_rows_inserted_then_the_header_and_keeps_its_name():
    pre, h = _header()
    ws = _ws()
    write_data(ws, pre.frame, 3)                                  # an old answer: table at A1, no header
    lo = ws.ListObjects("Data_Q3")
    assert lo.HeaderRowRange.Row == 1
    ws.log.clear()
    out = rewrite_data(ws, pre.frame, 3, h)
    assert ws.ListObjects("Data_Q3") is lo and lo.Name == "Data_Q3" and len(ws.ListObjects) == 1
    assert lo.HeaderRowRange.Row == len(h) + 2 == out["table_row"]
    kinds = [x[0] for x in ws.log]
    assert ("Insert", f"1:{len(h) + 1}") in ws.log
    assert kinds.index("Insert") < min(i for i, k in enumerate(kinds) if k == "Formula")
    assert "ListObjects.Add" not in kinds                         # never recreated
    t, n = len(h) + 2, len(data_grid(pre.frame))
    assert [x[1] for x in ws.log if x[0] == "Value2"] == [f"A{t}:I{t}", f"A{t + 1}:I{t + n - 1}"]   # never one rect
    assert ws.cells[(1, 1)]["Value"] == dict(h)["Title"]
    assert ws.cells[(len(h) + 2, 1)]["Value"] == "Cube"


def test_a_redraw_with_fewer_queries_moves_the_table_up_by_deleting_header_rows():
    spec = spec_dict()
    spec["queries"].append({"id": "fm2", "cube": "FinMax", "metrics": ["Opex"], "entity": "un.SY.2BUS",
                            "periods": "closed"})
    pre2 = pipeline.prepare(spec, resolver(), FakeProvider(facts()))
    big = header_rows(3, pre2.spec, pre2, WHEN, "tom")
    pre, small = _header()
    ws = _ws()
    write_data(ws, pre2.frame, 3, big)
    lo = ws.ListObjects("Data_Q3")
    rewrite_data(ws, pre.frame, 3, small)
    assert ws.ListObjects("Data_Q3") is lo and lo.HeaderRowRange.Row == len(small) + 2
    assert [x for x in ws.log if x[0] == "Delete"] == [("Delete", f"1:{len(big) - len(small)}")]
    labels = [ws.cells.get((r, 1), {}).get("Value") for r in range(2, len(small) + 1)]
    assert labels == LABELS[1:]                                   # the old second block is gone


def test_redraw_refreshes_fetched_and_control_total():
    pre, h = _header()
    ws = _ws()
    write_data(ws, pre.frame, 3, h)
    later = header_rows(3, pre.spec, pre, WHEN + datetime.timedelta(days=1), "tom")
    rewrite_data(ws, pre.frame, 3, later)
    assert ws.cells[(_row(later, "Fetched"), 2)]["Value"].startswith("2026-11-06 09:30")
    assert "Insert" not in [x[0] for x in ws.log] and "Delete" not in [x[0] for x in ws.log]


def test_rewrite_data_without_the_table_refuses_before_touching_the_sheet():
    from smartass.model.errors import GuardError
    pre, h = _header()
    ws = _ws()
    with pytest.raises(GuardError, match="Data_Q3"):
        rewrite_data(ws, pre.frame, 3, h)
    assert ws.log == [] and ws.cells == {}


def test_the_fake_refuses_a_formula_on_a_missing_table():
    ws = _ws()
    with pytest.raises(ComError):
        ws.Range("B1").Formula = '=COUNTIFS(Data_Q3[Query],"fm")'


def test_draw_stamps_one_time_for_the_header_and_the_footer(monkeypatch):
    from test_v2_pipeline_job import wire

    class Fixed(datetime.datetime):
        @classmethod
        def now(cls, tz=None):
            return cls(2026, 11, 4, 23, 59, 30, tzinfo=datetime.timezone(datetime.timedelta(hours=1)))

        def astimezone(self, tz=None):  # the machine's zone (UTC on CI) must not move the pinned Oslo time
            return self if tz is None else super().astimezone(tz)
    log, seen = [], {}
    wire(monkeypatch, log)
    monkeypatch.setattr(pipeline, "datetime", Fixed)
    monkeypatch.setattr(pipeline, "write_data", lambda ws, frame, q, header: seen.update(header=header) or {})
    monkeypatch.setattr(pipeline, "write_answer", lambda ws, plan, w: seen.update(plan=plan) or [])
    monkeypatch.setattr(datasheet, "windows_user", lambda: "tom")
    pipeline.excel_job(None, _pre(), "session")
    assert dict(seen["header"])["Fetched"] == "2026-11-04 23:59 +0100 · tom"
    assert seen["plan"].footer == "Source: Oracle EPM Cloud · fetched 2026-11-04 · page &P of &N"


def test_the_table_body_is_tn_font_9_and_the_sheet_prints_one_page_wide():
    pre, h = _header()
    ws = _ws()
    write_data(ws, pre.frame, 3, h)
    t = len(h) + 2
    g = data_grid(pre.frame)
    assert ws.cells[(t + 1, 1)]["Font.Size"] == 9 and ws.cells[(t + len(g) - 1, len(g[0]))]["Font.Size"] == 9
    assert ws.cells[(t + 1, 1)]["Font.Name"] == "Arial"
    ps = ws.PageSetup
    from smartass.render.layout import col_letter
    assert ps.PrintArea == f"$A$1:${col_letter(max(4, len(g[0])))}${t + len(g) - 1}"
    assert (ps.Orientation, ps.Zoom, ps.FitToPagesWide, ps.FitToPagesTall) == (2, False, 1, False)


def test_a_redraw_sets_the_font_and_print_setup_too():
    pre, h = _header()
    ws = _ws()
    write_data(ws, pre.frame, 3)                                  # an old answer: no header, default look
    ws.PageSetup.Orientation = 1
    rewrite_data(ws, pre.frame, 3, h)
    assert ws.PageSetup.Orientation == 2 and ws.cells[(len(h) + 3, 1)]["Font.Size"] == 9
