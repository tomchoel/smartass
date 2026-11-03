"""Fewer Excel calls, same workbook (smart flow spec D): each writer's exact effect on a fake workbook is pinned
(captured from the writer BEFORE batching) and its jobs.com call count is asserted."""
from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

import pytest
from fakexl import Workbook, count_com, freeze_time, snapshot

# ------------------------------------------------------------------ scenarios
SHEETS8 = ["Q1 Answer", "Q2 Answer", "Notes", "Q1 FinMax", "Q1 FinPlan", "Q1 Lines", "Q2 FinMax", "Sources"]


def arrange_case(moved: bool):
    from smartview_mcp import sheetorder
    names = list(SHEETS8)
    if moved:
        names.insert(1, names.pop(6))                 # 'Q2 FinMax' out of place
    wb = Workbook(names, styles_ver=False)
    for s in wb.order:
        kind = sheetorder.classify(s.Name)[0]
        s.Tab.Color = sheetorder.ANSWER_TAB if kind == "answer" else sheetorder.DATA_TAB if kind == "data" else False
    wb.active = wb.Worksheets("Q2 Answer")
    return wb, lambda: sheetorder.arrange_sheets(wb, "full", activate="Q2 Answer")


def fetched_5x14():
    from smartview_mcp import svhttp
    from smartview_mcp.grid import Axis, Layout
    periods = [f"{m} 26" for m in ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov",
                                   "Dec")] + ["FY26"]
    accounts = ["ac.C30", "ac.C70", "ac.STA20Y"]
    lay = Layout([("Entity", "un.SY.2WHS"), ("Scenario", "Actual")], [Axis("Account", accounts)],
                 [Axis("Period", periods)])
    values = [[""] + ["un.SY.2WHS", "Actual"] + [""] * 11, [""] + periods,
              ["C30 Revenues"] + [1000.0 * (k + 1) for k in range(13)],
              ["C70 Opex"] + [-250.5 * (k + 1) for k in range(12)] + ["#Missing"],
              ["STA20Y man-years"] + [12.25] * 13]
    return svhttp.Fetched(lay, values, 0.5)


def source_case():
    from smartview_mcp.core import source
    from smartview_mcp.core.contract import Pull
    wb = Workbook(["Q1 Answer"])
    t = SimpleNamespace(wb=wb, app=wb.Application, sheets_added=[])
    pin = Pull(id="m", connection="FinMax", rows=[{"dimension": "Account", "members": ["ac.C30"]}],
               columns=[{"dimension": "Period", "members": ["Jan 26"]}])
    return wb, lambda: source.write(t, 1, "Q1 FinMax", pin, fetched_5x14())


def answer_sheet():
    """A small answer sheet: title, subtitle, header, lines, a sum, a check, a text and a section."""
    wb = Workbook(["Q1 Answer"])
    ws = wb.order[0]
    rows = [(1, "title", ["Business — EBITDA"]), (2, "subtitle", ["Actual, NOK"]), (3, "blank", []),
            (4, "header", ["", "Line", "Sep 26", "Sep 25", "Var", "Var %", "Check"]),
            (5, "line", ["", "Revenue", 100.0, 90.0, 10.0, 0.11, ""]),
            (6, "line", ["", "Opex", -40.0, -35.0, -5.0, -0.14, ""]),
            (7, "line", ["", "FTE", 12.5, 11.0, 1.5, 0.13, ""]),
            (8, "sum", ["", "EBITDA", 60.0, 55.0, 5.0, 0.09, ""]),
            (9, "check", ["", "Sum = cube", "", "", "", "", "OK"]),
            (10, "check", ["", "YTD = months", "", "", "", "", "OK"]),
            (11, "blank", []), (12, "section", ["", "Notes"]),
            (13, "text", ["", "Revenue grew on wholesale volumes; opex rose with consultants. " * 3]),
            (14, "grand", ["", "Total", 60.0, 55.0, 5.0, 0.09, ""])]
    meta = []
    for r, kind, vals in rows:
        for c, v in enumerate(vals, 1):
            ws.put(r, c, v)
        spec = {"row": r, "kind": kind}
        if kind in ("line", "sum", "grand"):
            spec["num"] = {"F": "prosent1"} | ({"C": "tall1", "D": "tall1", "E": "tall1"} if r == 7 else {})
        if kind == "check":
            spec["status_col"] = "G"
        meta.append(spec)
    roles = {"last_col": "G", "label_cols": ["B"], "freeze_rows": 4, "last_row": 14,
             "col_widths": {"A": 2.5, "B": 46, "C:G": 17}, "rows": meta}
    return wb, ws, roles


def report_case(redraw: bool = False):
    from smartview_mcp import styling
    wb, ws, roles = answer_sheet()
    if redraw:
        styling.apply_report(ws, roles)
    return wb, lambda: styling.apply_report(ws, roles)


def lines_case():
    from smartview_mcp import drill_xl
    wb = Workbook(["Q1 Answer"])
    t = SimpleNamespace(wb=wb, app=wb.Application, sheets_added=[])
    cols = ["ACCOUNT", "GL_DESCRIPTION_LINE", "AMOUNT", "DEBET_AMOUNT", "CREDIT_AMOUNT", "GL_POSTED_DATE",
            "BOOKING_DATE", "PROJECT", "AP_CUSTOMER_NAME"]
    rows = [{"ACCOUNT": "6700", "GL_DESCRIPTION_LINE": f"Consulting {i}", "AMOUNT": 1000.0 + i,
             "DEBET_AMOUNT": 1000.0 + i, "CREDIT_AMOUNT": 0.0, "GL_POSTED_DATE": "2026-09-15",
             "BOOKING_DATE": "2026-09-14", "PROJECT": "P1", "AP_CUSTOMER_NAME": "Acme AS"} for i in range(6)]
    return wb, lambda: drill_xl.write_lines(t, 1, rows, cols, title="Q1 lines", source="FinLite drill", alias="d")


def run(case, monkeypatch, *a):
    from smartview_mcp import drill_xl
    from smartview_mcp.core import source
    freeze_time(monkeypatch, drill_xl, source)
    wb, go = case(*a)
    calls = count_com(monkeypatch)
    go()
    return snapshot(wb), len(calls)


# ------------------------------------------------------------------ pinned state + call counts
FIXTURES = Path(__file__).parent / "fixtures" / "excel_calls"


def pinned(name: str) -> dict:
    """The workbook state this scenario produced before batching (captured from the old writer)."""
    return json.loads((FIXTURES / f"{name}.json").read_text(encoding="utf-8"))


def same(snap: dict) -> dict:
    return json.loads(json.dumps(snap, ensure_ascii=False))


@pytest.mark.parametrize("moved, calls", [(False, 1), (True, 11)])          # before: 31, 50
def test_arrange_sheets_same_order_fewer_calls(monkeypatch, moved, calls):
    name = "arrange_moved" if moved else "arrange_ok"
    snap, n = run(arrange_case, monkeypatch, moved)
    assert same(snap) == pinned(name)["state"]
    assert n == calls


def test_source_sheet_same_look_fewer_calls(monkeypatch):
    snap, n = run(source_case, monkeypatch)
    assert same(snap) == pinned("source")["state"]
    assert n == 85


@pytest.mark.parametrize("redraw, calls", [(False, 150), (True, 150)])
def test_answer_styling_same_look_fewer_calls(monkeypatch, redraw, calls):
    snap, n = run(report_case, monkeypatch, redraw)
    assert same(snap) == pinned("report_redraw" if redraw else "report_first")["state"]
    assert n == calls


def test_lines_sheet_same_look_fewer_calls(monkeypatch):
    snap, n = run(lines_case, monkeypatch)
    assert same(snap) == pinned("lines")["state"]
    assert n == 39
