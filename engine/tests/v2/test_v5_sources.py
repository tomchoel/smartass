"""Plan 5 Task 11: the v2 Sources sheet - Known limits on top, then an append-only log of every query pull, linked
both ways with the Q<n> Data header (fakes only, no Excel). Every figure is invented."""
import datetime
from types import SimpleNamespace

from fakexl import Workbook
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass import pipeline
from smartass.render.datasheet import Header, header_rows
from smartass.render.excel import rewrite_data, write_data
from smartass.xl import sources

WHEN = datetime.datetime(2026, 11, 5, 9, 30, tzinfo=datetime.timezone(datetime.timedelta(hours=1)))
FETCHED = "2026-11-05 09:30 +0100 · tom"
LOG = ["Fetched", "Q", "Query id", "Question", "Cube", "Scenario/Version", "POV", "Periods", "Route", "Rows",
       "Control total", "Status", "Data sheet", "Supersedes"]
CAVEATS = [{"id": "a", "text": "Opex before 2025 excludes IFRS 16.", "severity": "warn"},
           {"id": "b", "text": "Opex before 2025 excludes IFRS 16.", "severity": "block"},
           {"id": "c", "text": "Just so you know.", "severity": "info"},
           {"id": "d", "text": "2BUS was reorganised in March 2026.", "severity": "block"}]


def _pre(spec=None, caveats=CAVEATS):
    pre = pipeline.prepare(spec or spec_dict(), resolver(), FakeProvider(facts()))
    pre.caveats = list(caveats)
    return pre


def _book(names=("Q3 Answer", "Q3 Data")):
    wb = Workbook(list(names))
    return wb, SimpleNamespace(wb=wb, app=wb.Application, sheets_added=[])


def _run(t, wb, pre, q=3, redraw=False, status=None):
    ws = wb.Worksheets(f"Q{q} Data")
    h = header_rows(q, pre.spec, pre, WHEN, "tom")
    (rewrite_data if redraw else write_data)(ws, pre.frame, q, h)
    return sources.write_sources(t, q, pre, ws, h, FETCHED, status or {}), h


def _v(ws, r, c):
    return ws.cells.get((r, c), {}).get("Value")


def _log_row(ws, r):
    return [_v(ws, r, c) for c in range(1, len(LOG) + 1)]


P = sources.LIMITS_COL        # Known limits: column P, right of the log


def test_first_call_creates_the_sheet_with_the_log_in_row_1_and_known_limits_to_the_right():
    wb, t = _book()
    pre = _pre()
    rows, _ = _run(t, wb, pre, status={"fm": "OK"})
    ws = wb.Worksheets("Sources")
    assert [s.Name for s in wb.order][-1] == "Sources"
    assert _log_row(ws, 1) == LOG and P == len(LOG) + 2
    assert _v(ws, 1, P) == "Known limits" and _v(ws, 1, P - 1) is None
    assert [_v(ws, r, P) for r in (2, 3)] == ["Opex before 2025 excludes IFRS 16.",
                                             "2BUS was reorganised in March 2026."]
    assert rows == {"fm": 2}
    row = _log_row(ws, 2)
    total = sum(f.value for f in pre.frame.facts if f.query == "fm")
    assert row[:9] == [FETCHED, "Q3", "fm", "How is Business opex developing vs last year?", "FinMax", "Actual",
                       "2BUS Business Norway (un.SY.2BUS) · NOK · Periodic", "Sep 26 (closed)", "HTTP"]
    assert row[9] == sum(1 for f in pre.frame.facts if f.query == "fm") and row[10] == total
    assert row[11] == "OK" and row[12] == "Q3 Data" and row[13] is None
    assert ws.cells[(1, 1)]["Font.Bold"] is True and ws.cells[(1, P)]["Font.Bold"] is True
    assert ws.AutoFilterMode and tuple(ws.autofilter[:2]) == (1, 1) and ws.autofilter[3] == len(LOG)


def test_a_second_call_for_the_same_q_appends_and_supersedes_the_first_row():
    wb, t = _book()
    pre = _pre()
    first, _ = _run(t, wb, pre)
    ws = wb.Worksheets("Sources")
    before = _log_row(ws, first["fm"])
    second, _ = _run(t, wb, pre, redraw=True)
    assert second["fm"] == first["fm"] + 1
    assert _log_row(ws, first["fm"]) == before                       # earlier rows are never edited
    assert _log_row(ws, second["fm"])[13] == f"row {first['fm']}"
    assert ws.AutoFilterMode and ws.autofilter[2] >= second["fm"]   # the filter stays on (AutoFilter toggles)
    assert len([s for s in wb.order if s.Name.startswith("Sources")]) == 1


def test_hyperlinks_go_both_ways():
    wb, t = _book()
    rows, h = _run(t, wb, _pre())
    ws, data = wb.Worksheets("Sources"), wb.Worksheets("Q3 Data")
    r = rows["fm"]
    link = ws.Hyperlinks.at(r, 13)
    assert link.SubAddress == "'Q3 Data'!A1" and _v(ws, r, 13) == "Q3 Data"
    src = h.sources["fm"]
    back = data.Hyperlinks.at(int(src[1:]), 2)
    assert back.SubAddress == f"'Sources'!A{r}" and _v(data, int(src[1:]), 2) == f"Sources row {r}"


def test_a_redraw_re_adds_the_data_sheet_link_its_header_rewrite_cleared():
    wb, t = _book()
    pre = _pre()
    _run(t, wb, pre)
    rows, h = _run(t, wb, pre, redraw=True)
    data = wb.Worksheets("Q3 Data")
    src = int(h.sources["fm"][1:])
    assert len(data.Hyperlinks) == 1 and data.Hyperlinks.at(src, 2).SubAddress == f"'Sources'!A{rows['fm']}"


def test_known_limits_are_deduplicated_and_merged_with_the_lines_already_there():
    wb, t = _book(("Q3 Answer", "Q3 Data", "Q4 Answer", "Q4 Data"))
    first, _ = _run(t, wb, _pre())
    ws = wb.Worksheets("Sources")
    more = [{"id": "e", "text": "FinPlan Working is a draft.", "severity": "warn"},
            {"id": "d", "text": "2BUS was reorganised in March 2026.", "severity": "block"}]
    second, _ = _run(t, wb, _pre(caveats=more), q=4)
    lines = [_v(ws, r, P) for r in range(2, 6)]
    assert lines == ["Opex before 2025 excludes IFRS 16.", "2BUS was reorganised in March 2026.",
                     "FinPlan Working is a draft.", None]
    assert second["fm"] == first["fm"] + 1


def test_rows_never_move_when_known_limits_grow():
    """The invariant behind every 'row <r>' and SubAddress: a row returned once holds the same content for good."""
    wb, t = _book(("Q3 Answer", "Q3 Data", "Q4 Answer", "Q4 Data"))
    first, h3 = _run(t, wb, _pre(caveats=[CAVEATS[0]]))
    ws = wb.Worksheets("Sources")
    old = _log_row(ws, first["fm"])
    many = [{"id": f"x{i}", "text": f"Limit number {i}.", "severity": "warn"} for i in range(8)]
    _run(t, wb, _pre(caveats=many), q=4)
    assert _log_row(ws, 1) == LOG and _log_row(ws, first["fm"]) == old
    assert ws.Hyperlinks.at(first["fm"], 13).SubAddress == "'Q3 Data'!A1"
    data3 = wb.Worksheets("Q3 Data")
    back = data3.Hyperlinks.at(int(h3.sources["fm"][1:]), 2).SubAddress
    assert back == f"'Sources'!A{first['fm']}" and _v(ws, int(back.split("A")[-1]), 3) == "fm"
    assert _v(ws, 10, P) == "Limit number 7."


def test_no_warn_or_block_caveat_says_none():
    wb, t = _book()
    rows, _ = _run(t, wb, _pre(caveats=[CAVEATS[2]]))
    ws = wb.Worksheets("Sources")
    assert _v(ws, 1, P) == "Known limits" and _v(ws, 2, P) == sources.NO_LIMITS
    assert _log_row(ws, 1) == LOG and rows == {"fm": 2}
    _run(t, wb, _pre(caveats=[CAVEATS[0]]), redraw=True)
    assert _v(ws, 2, P) == CAVEATS[0]["text"] and _v(ws, 3, P) is None   # the placeholder is not merged


def test_a_user_owned_sources_sheet_leads_to_sources_sv():
    wb, t = _book(("Q3 Answer", "Q3 Data", "Sources"))
    mine = wb.Worksheets("Sources")
    mine.put(1, 1, "My own sources")
    mine.put(2, 1, "a note")
    _run(t, wb, _pre())
    assert _v(mine, 1, 1) == "My own sources" and _v(mine, 2, 1) == "a note"
    assert _log_row(wb.Worksheets("Sources SV"), 1) == LOG


def test_the_empty_sources_sheet_of_a_new_workbook_is_ours():
    wb, t = _book(("Sources", "Q3 Answer", "Q3 Data"))
    _run(t, wb, _pre())
    _run(t, wb, _pre(), redraw=True)
    assert [s.Name for s in wb.order].count("Sources SV") == 0
    assert _log_row(wb.Worksheets("Sources"), 1) == LOG


def test_a_lines_only_answer_logs_no_rows():
    from test_v2_pipeline import lines_only_spec
    wb, t = _book()
    pre = _pre()
    pre.spec.queries = []
    assert sources.write_sources(t, 3, pre, wb.Worksheets("Q3 Data"), Header(), FETCHED, {}) == {}
    assert lines_only_spec()["queries"] == []
    ws = wb.Worksheets("Sources")
    assert _log_row(ws, 1) == LOG and _v(ws, 2, 1) is None


def test_sheetorder_keeps_a_v2_sources_sheet_last_in_the_users_workbook():
    from smartass.xl import sheetorder
    wb, t = _book(("Notes", "Q3 Answer", "Q3 Data"))
    _run(t, wb, _pre())
    wb.order.insert(0, wb.order.pop())                              # our Sources moved first
    sheetorder.arrange_sheets(wb, "active")
    assert [s.Name for s in wb.order][-1] == "Sources"


def test_draw_logs_the_sources_after_calculate_for_new_answers_and_redraws(monkeypatch):
    from test_v2_pipeline_job import wire
    log, seen = [], []
    wire(monkeypatch, log)
    monkeypatch.setattr(pipeline, "_write_sources",
                        lambda t, q, pre, ws, header, fetched: seen.append((q, fetched)) or log.append("sources"))
    pipeline.excel_job(None, _pre(), "session")
    assert log.index("calculate") < log.index("sources") < log.index("arrange")
    assert seen[0][0] == 4 and seen[0][1].startswith(f"{datetime.date.today():%Y-%m-%d}")


def test_a_failing_sources_log_is_a_warning_not_a_failed_answer(monkeypatch):
    from test_v2_pipeline_job import wire
    log, real = [], pipeline._write_sources
    wire(monkeypatch, log)
    monkeypatch.setattr(pipeline, "_write_sources", real)

    def boom(*a):
        raise RuntimeError("no Sources")
    monkeypatch.setattr(sources, "write_sources", boom)
    out = pipeline.excel_job(None, _pre(), "session")
    assert any("Sources" in w for w in out["warnings"]) and "rollback" not in log


def test_the_real_draw_seam_reads_the_status_and_links_the_row():
    wb, t = _book()
    pre = _pre()
    ws = wb.Worksheets("Q3 Data")
    h = header_rows(3, pre.spec, pre, WHEN, "tom")
    write_data(ws, pre.frame, 3, h)
    r, c = int(h.status["fm"][1:]), ord(h.status["fm"][0]) - 64
    ws.put(r, c, "OK")                                              # as Excel shows it after calculating
    assert pipeline._write_sources(t, 3, pre, ws, h, FETCHED) is None
    src = wb.Worksheets("Sources")
    row = next(rr for rr in range(1, 20) if _v(src, rr, 3) == "fm")
    assert _v(src, row, 12) == "OK"
    assert ws.Hyperlinks.at(int(h.sources["fm"][1:]), 2).SubAddress == f"'Sources'!A{row}"


def test_a_v1_log_is_left_as_it_is_and_v2_logs_on_sources_sv():
    wb, t = _book(("Q3 Answer", "Q3 Data", "Sources"))
    v1 = wb.Worksheets("Sources")
    head = ("Fetched", "Question no", "Question", "Cube", "POV", "Rows × columns", "Notes", "Sheet")
    for c, x in enumerate(head, 1):
        v1.put(1, c, x)
    v1.put(2, 1, "2026-10-01 10:00")
    before = {k: dict(x) for k, x in v1.cells.items()}
    pre = _pre()
    ws = wb.Worksheets("Q3 Data")
    h = header_rows(3, pre.spec, pre, WHEN, "tom")
    write_data(ws, pre.frame, 3, h)
    notes = []
    rows = sources.write_sources(t, 3, pre, ws, h, FETCHED, {}, None, notes)
    assert v1.cells == before and rows == {"fm": 2}
    assert _log_row(wb.Worksheets("Sources SV"), 1) == LOG and notes == [sources.V1_NOTE]
    sources.write_sources(t, 3, pre, ws, h, FETCHED, {}, None, notes)     # said once, when Sources SV is made
    assert notes == [sources.V1_NOTE] and v1.cells == before


def _lines_pre():
    from test_v2_pipeline import LINE_CELL, lines_only_spec

    from smartass.model.lines import LineSet
    rows = [{"AP_CUSTOMER_NAME": "Bouvet", "AMOUNT": -30.0}, {"AP_CUSTOMER_NAME": "Accenture", "AMOUNT": -70.0}]
    ls = LineSet("b0", rows, ["AP_CUSTOMER_NAME", "AMOUNT"], cube_value=-100.0, cell=LINE_CELL)
    pre = pipeline.prepare(lines_only_spec(), resolver(), FakeProvider(facts()), lines={"b0": ls})
    pre.caveats = []
    return pre


def test_a_drill_logs_its_own_row_with_the_lines_total_and_tie_out():
    from test_v2_pipeline import LINE_CELL
    wb, t = _book(("Q3 Answer", "Q3 Data", "Q3 Lines b0"))
    pre = _lines_pre()
    rows = sources.write_sources(t, 3, pre, wb.Worksheets("Q3 Data"), Header(), FETCHED, {}, {"b0": "OK"})
    ws = wb.Worksheets("Sources")
    assert rows == {"b0": 2}
    row = _log_row(ws, 2)
    assert row[1:3] == ["Q3", "b0"] and row[4] == "FinLite" and row[8] == "drill"
    assert row[7] == LINE_CELL["All Time Periods"] and LINE_CELL["All Entities"] in row[6]
    assert row[9] == 2 and row[10] == -100.0 and row[11] == "OK" and row[12] == "Q3 Lines b0"
    assert ws.Hyperlinks.at(2, 13).SubAddress == "'Q3 Lines b0'!A1"
    again = sources.write_sources(t, 3, pre, wb.Worksheets("Q3 Data"), Header(), FETCHED, {}, {"b0": "DIFF"})
    assert again == {"b0": 3} and _log_row(ws, 3)[13] == "row 2" and _log_row(ws, 3)[11] == "DIFF"


def test_the_draw_seam_reads_the_tie_out_of_each_lines_sheet():
    from smartass.render.datasheet import LINES_HEADER, LINES_STATUS, lines_header
    grid = lines_header("Lines_Q3_b0", {}, FETCHED)
    first = LINES_HEADER.split(":")[0]
    r, c = int(LINES_STATUS[1:]) - int(first[1:]), ord(LINES_STATUS[0]) - ord(first[0])
    assert grid[r][c - 1] == "Tie-out" and grid[r][c].startswith("=IF(")      # I3 is the tie-out formula
    wb, t = _book(("Q3 Answer", "Q3 Data", "Q3 Lines b0"))
    wb.Worksheets("Q3 Lines b0").put(3, 9, "OK")                         # as Excel shows it after calculating
    pre = _lines_pre()
    assert pipeline._write_sources(t, 3, pre, wb.Worksheets("Q3 Data"), Header(), FETCHED) is None
    assert _log_row(wb.Worksheets("Sources"), 2)[11] == "OK"


def test_a_pull_without_its_row_is_a_warning(monkeypatch):
    wb, t = _book()
    pre = _pre()
    ws = wb.Worksheets("Q3 Data")
    h = header_rows(3, pre.spec, pre, WHEN, "tom")
    write_data(ws, pre.frame, 3, h)
    monkeypatch.setattr(sources, "write_sources", lambda *a: {})
    assert "no row for fm" in pipeline._write_sources(t, 3, pre, ws, h, FETCHED)
