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
    from smartass.xl import sheetorder
    names = list(SHEETS8)
    if moved:
        names.insert(1, names.pop(6))                 # 'Q2 FinMax' out of place
    wb = Workbook(names, styles_ver=False)
    for s in wb.order:
        kind = sheetorder.classify(s.Name)[0]
        s.Tab.Color = sheetorder.ANSWER_TAB if kind == "answer" else sheetorder.DATA_TAB if kind == "data" else False
    wb.active = wb.Worksheets("Q2 Answer")
    return wb, lambda: sheetorder.arrange_sheets(wb, "full", activate="Q2 Answer")


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
    from smartass.xl import styling
    wb, ws, roles = answer_sheet()
    if redraw:
        styling.apply_report(ws, roles)
    return wb, lambda: styling.apply_report(ws, roles)


def lines_case():
    from smartass.xl import drill_xl
    wb = Workbook(["Q1 Answer"])
    t = SimpleNamespace(wb=wb, app=wb.Application, sheets_added=[])
    cols = ["ACCOUNT", "GL_DESCRIPTION_LINE", "AMOUNT", "DEBET_AMOUNT", "CREDIT_AMOUNT", "GL_POSTED_DATE",
            "BOOKING_DATE", "PROJECT", "AP_CUSTOMER_NAME"]
    rows = [{"ACCOUNT": "6700", "GL_DESCRIPTION_LINE": f"Consulting {i}", "AMOUNT": 1000.0 + i,
             "DEBET_AMOUNT": 1000.0 + i, "CREDIT_AMOUNT": 0.0, "GL_POSTED_DATE": "2026-09-15",
             "BOOKING_DATE": "2026-09-14", "PROJECT": "P1", "AP_CUSTOMER_NAME": "Acme AS"} for i in range(6)]
    return wb, lambda: drill_xl.write_lines(t, 1, rows, cols, title="Q1 lines", source="FinLite drill", alias="d")


def run(case, monkeypatch, *a):
    from smartass.xl import drill_xl
    freeze_time(monkeypatch, drill_xl)
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


@pytest.mark.parametrize("moved, calls", [(False, 1), (True, 10)])          # before: 31, 50
def test_arrange_sheets_same_order_fewer_calls(monkeypatch, moved, calls):
    name = "arrange_moved" if moved else "arrange_ok"
    snap, n = run(arrange_case, monkeypatch, moved)
    assert same(snap) == pinned(name)["state"]
    assert n == calls


@pytest.mark.parametrize("redraw, calls", [(False, 59), (True, 41)])          # before: 150, 150 (+1: CenterFooter read)
def test_answer_styling_same_look_fewer_calls(monkeypatch, redraw, calls):
    snap, n = run(report_case, monkeypatch, redraw)
    names = snap["sheets"]["Q1 Answer"].pop("names")
    want = pinned("report_redraw" if redraw else "report_first")["state"]
    assert want["sheets"]["Q1 Answer"].pop("names") == {}
    assert same(snap) == want
    assert names == {"smartass_pagesetup": ('="v1|$A$1:$G$14|Intern"', False)}     # the one addition
    assert n == calls


def test_page_setup_is_redone_when_the_print_area_changes(monkeypatch):
    from smartass.xl import styling
    wb, ws, roles = answer_sheet()
    styling.apply_report(ws, roles)
    ws.PageSetup.PrintArea = "changed by hand"           # only a changed setup key brings the setup back
    ws.put(15, 2, "Extra line")
    roles["rows"].append({"row": 15, "kind": "note"})
    roles["last_row"] = 15
    styling.apply_report(ws, roles)
    assert ws.PageSetup.PrintArea == "$A$1:$G$15"
    assert ws.Names("smartass_pagesetup").RefersTo == '="v1|$A$1:$G$15|Intern"'
    ws.PageSetup.PrintArea = "changed by hand"
    styling.apply_report(ws, roles)                      # same key: skipped
    assert ws.PageSetup.PrintArea == "changed by hand"


def test_lines_sheet_same_look_fewer_calls(monkeypatch):
    snap, n = run(lines_case, monkeypatch)
    assert same(snap) == pinned("lines")["state"]
    assert n == 22                                    # before: 39


def test_answer_lint_findings_unchanged():
    """The lint's findings as the unbatched lint returned them (captured before batching)."""
    from smartass.xl import styling
    wb, ws, roles = answer_sheet()
    assert styling.apply_report(ws, roles) == []
    wb, ws, roles = answer_sheet()
    ws.put(5, 2, "A very very long label that will not fit into the column at all, no way, really not")
    ws.put(3, 9, "outside")
    assert styling.apply_report(ws, roles) == ["content outside the table width (up to G)",
                                               "B5: text is cut off or spills over (~65 > 46 chars)"]


def lines_redrill_case(mod=None):
    """A re-drill with the same columns and fewer lines: the table is reused in place (cleared, resized)."""
    from smartass.xl import drill_xl
    mod = mod or drill_xl
    wb, first = lines_case()
    first()
    t = SimpleNamespace(wb=wb, app=wb.Application, sheets_added=[])
    cols = ["ACCOUNT", "GL_DESCRIPTION_LINE", "AMOUNT", "DEBET_AMOUNT", "CREDIT_AMOUNT", "GL_POSTED_DATE",
            "BOOKING_DATE", "PROJECT", "AP_CUSTOMER_NAME"]
    rows = [{"ACCOUNT": "6790", "GL_DESCRIPTION_LINE": f"Audit {i}", "AMOUNT": 50.0 * i, "DEBET_AMOUNT": 50.0 * i,
             "CREDIT_AMOUNT": 0.0, "GL_POSTED_DATE": "2026-10-01", "BOOKING_DATE": "2026-10-01", "PROJECT": "P2",
             "AP_CUSTOMER_NAME": "Revisor AS"} for i in range(3)]
    return wb, lambda: mod.write_lines(t, 1, rows, cols, title="Q1 lines again", source="FinLite drill", alias="d")


def test_lines_redrill_same_table_fewer_calls(monkeypatch):
    snap, n = run(lines_redrill_case, monkeypatch)
    assert same(snap) == pinned("lines_redrill")["state"]
    assert n == 22                                    # before: 40


def test_page_setup_marker_is_not_added_when_print_communication_stays_off(monkeypatch):
    from smartass.xl import styling
    wb, ws, roles = answer_sheet()

    class Flaky:
        def __init__(self, inner):
            object.__setattr__(self, "inner", inner)

        def __getattr__(self, k):
            return getattr(self.inner, k)

        def __setattr__(self, k, v):
            if k == "PrintCommunication" and v is True:
                raise RuntimeError("busy")
            setattr(self.inner, k, v)
    monkeypatch.setattr(styling, "_retry", lambda f, *a, **k: f())
    styling._page_setup(ws, Flaky(wb.Application), 14, 7, roles)
    assert not styling._has_page_setup(ws, styling._page_setup_key(14, 7, roles))
    styling._page_setup(ws, wb.Application, 14, 7, roles)
    assert styling._has_page_setup(ws, styling._page_setup_key(14, 7, roles))
