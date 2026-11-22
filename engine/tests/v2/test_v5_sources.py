"""Plan 5 Task 11: the v2 Sources sheet - Known limits on top, then an append-only log of every query pull, linked
both ways with the Q<n> Data header (fakes only, no Excel). Every figure is invented."""
import datetime
from types import SimpleNamespace

import fakexl
from fakexl import Workbook
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass import exceljob, pipeline
from smartass.render.datasheet import Header, header_rows
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


def _data_sheet(ws, h) -> None:
    """The Q<n> Data header as a moved-in file leaves it: a new sheet every draw (a follow-up replaces it), its texts
    and numbers in place (the formulas over the table are left out: the fake has no table)."""
    ws.cells.clear()
    for r, (lab, value) in enumerate(h, 1):
        for c, x in enumerate([lab, *(value if isinstance(value, tuple) else (value,))], 1):
            if x not in (None, "") and not (isinstance(x, str) and x.startswith("=")):
                ws.put(r, c, x)


def _run(t, wb, pre, q=3, redraw=False, status=None):
    """One draw's log rows; redraw (a follow-up) draws the same way: its Data sheet is new too."""
    ws = wb.Worksheets(f"Q{q} Data")
    h = header_rows(q, pre.spec, pre, WHEN, "tom")
    _data_sheet(ws, h)
    return sources.write_sources(t, q, pre, ws, h, FETCHED, status or {}), h


def _v(ws, r, c):
    return ws.cells.get((r, c), {}).get("Value")


def _log_row(ws, r):
    """A log row as Excel shows it (a HYPERLINK formula as its text)."""
    return [fakexl.shown(_v(ws, r, c)) for c in range(1, len(LOG) + 1)]


def _target(ws, r, c):
    """Where the HYPERLINK formula of a cell goes ('Q3 Data'!A1), None when it is not one."""
    link = fakexl.hyperlink(_v(ws, r, c))
    return link[0].lstrip("#") if link else None


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


def test_the_links_go_both_ways_as_hyperlink_formulas():
    wb, t = _book()
    rows, h = _run(t, wb, _pre())
    ws, data = wb.Worksheets("Sources"), wb.Worksheets("Q3 Data")
    r = rows["fm"]
    assert _v(ws, r, 13) == "=HYPERLINK(\"#'Q3 Data'!A1\",\"Q3 Data\")" and _log_row(ws, r)[12] == "Q3 Data"
    src = int(h.sources["fm"][1:])
    assert _v(data, src, 2) == f"=HYPERLINK(\"#'Sources'!A{r}\",\"Sources row {r}\")"
    assert len(ws.Hyperlinks) == len(data.Hyperlinks) == 0                  # no Hyperlinks.Add any more


def test_a_redraw_re_adds_the_data_sheet_link_its_header_rewrite_cleared():
    wb, t = _book()
    pre = _pre()
    _run(t, wb, pre)
    rows, h = _run(t, wb, pre, redraw=True)
    data = wb.Worksheets("Q3 Data")
    src = int(h.sources["fm"][1:])
    assert _target(data, src, 2) == f"'Sources'!A{rows['fm']}" and not data.Hyperlinks


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
    assert _target(ws, first["fm"], 13) == "'Q3 Data'!A1"
    data3 = wb.Worksheets("Q3 Data")
    back = _target(data3, int(h3.sources["fm"][1:]), 2)
    assert back == f"'Sources'!A{first['fm']}" and _v(ws, int(back.split("A")[-1]), 3) == "fm"
    assert _v(ws, 10, P) == "Limit number 7."


LEFT_OUT = {"id": "left-out:FinMax:ac.C60", "severity": "warn",
            "text": "FinMax: EPM refused All Accounts member ac.C60 over Planning REST ... is INCOMPLETE"}


def test_a_left_out_members_line_names_its_question_and_day_and_goes_when_that_question_no_longer_has_it():
    """A member left out over REST is true of one answer: its Known limits line says which question and which day, so
    an old line cannot be read as today's, and the question's next draw without the drop removes it - the lines of
    other questions and the static limits stay."""
    wb, t = _book(("Q3 Answer", "Q3 Data", "Q4 Answer", "Q4 Data"))
    _run(t, wb, _pre(caveats=[CAVEATS[0], LEFT_OUT]))
    _run(t, wb, _pre(caveats=[LEFT_OUT]), q=4)
    ws = wb.Worksheets("Sources")
    dated = [f"Q3 (2026-11-05): {LEFT_OUT['text']}", f"Q4 (2026-11-05): {LEFT_OUT['text']}"]
    assert [_v(ws, r, P) for r in range(2, 5)] == [CAVEATS[0]["text"], *dated]
    _run(t, wb, _pre(caveats=[CAVEATS[0]]), redraw=True)                 # Q3 again, the member back
    assert [_v(ws, r, P) for r in range(2, 5)] == [CAVEATS[0]["text"], dated[1], None]
    _run(t, wb, _pre(caveats=[]), q=4, redraw=True)
    assert [_v(ws, r, P) for r in range(2, 4)] == [CAVEATS[0]["text"], None]


def test_a_questions_own_dated_lines_are_matched_by_its_whole_number():
    old = ["Q10 (2026-11-01): a member left out", "Q1 (2026-11-01): a member left out", "A static limit."]
    assert sources.limits([], old, 1, FETCHED) == ["Q10 (2026-11-01): a member left out", "A static limit."]
    assert sources.limits([LEFT_OUT], old, 1, FETCHED) == [*old[::2], f"Q1 (2026-11-05): {LEFT_OUT['text']}"]


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


def test_the_status_cells_are_read_in_the_finish_batch():
    """The Data header's status cell of each query and the tie-out of each lines sheet, as Excel shows them after
    calculating (exceljob._record_status, part of the finish batch)."""
    from smartass.render.datasheet import LINES_STATUS
    from smartass.xl.inproc import runner
    from smartass.xl.inproc.batch import Batch
    wb, t = _book(("Q3 Answer", "Q3 Data", "Q3 Lines b0"))
    pre = _lines_pre()
    pre.spec.queries[:] = _pre().spec.queries
    ws = wb.Worksheets("Q3 Data")
    h = header_rows(3, pre.spec, pre, WHEN, "tom")
    _data_sheet(ws, h)
    r, c = int(h.status["fm"][1:]), ord(h.status["fm"][0]) - 64
    ws.put(r, c, "OK")
    wb.Worksheets("Q3 Lines b0").put(int(LINES_STATUS[1:]), ord(LINES_STATUS[0]) - 64, "DIFF")
    b = Batch()
    get = exceljob._record_status(b, b.workbook(wb.Name), 3, pre, h)
    assert get(runner.run(wb.Application, b)) == ({"fm": (h.status["fm"], "OK")}, {"b0": "DIFF"})


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
    _data_sheet(ws, h)
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
    assert _target(ws, 2, 13) == "'Q3 Lines b0'!A1"
    again = sources.write_sources(t, 3, pre, wb.Worksheets("Q3 Data"), Header(), FETCHED, {}, {"b0": "DIFF"})
    assert again == {"b0": 3} and _log_row(ws, 3)[13] == "row 2" and _log_row(ws, 3)[11] == "DIFF"


def test_a_pull_without_its_row_is_a_warning():
    from smartass.xl.inproc.runner import Result
    pre = _pre()
    written = sources.Written({}, range(0, 3), None, "Sources")
    assert "no row for fm" in exceljob._sources_warning(pre, (written, []), Result())
    refused = Result(skipped=[(1, "a set", "Excel said no")])
    assert exceljob._sources_warning(pre, (written, []), refused).startswith(
        "could not log the pull on the Sources sheet (Excel said no)")
    assert exceljob._sources_warning(pre, (sources.Written({"fm": 2}, range(0, 3)), [sources.V1_NOTE]),
                                     Result()) == sources.V1_NOTE


def test_the_log_is_tn_font_9_and_prints_a_to_n_landscape():
    wb, t = _book()
    pre = _pre()
    rows, _ = _run(t, wb, pre, status={"fm": "OK"})
    ws = wb.Worksheets("Sources")
    r = rows["fm"]
    assert ws.cells[(r, 1)]["Font.Size"] == 9 and ws.cells[(1, len(LOG))]["Font.Size"] == 9
    assert ws.cells[(2, P)]["Font.Size"] == 9                    # Known limits too
    ps = ws.PageSetup
    assert ps.PrintArea == "$A$1:$N$1048575" and (ps.Orientation, ps.FitToPagesWide) == (2, 1)   # A:N, set once



# ------------------------------------------------------------------ Task 8B: the log's look, pinned before the fast log
def _font(ws, r, c, k):
    return ws.cells.get((r, c), {}).get(f"Font.{k}")


def test_the_log_shows_the_tn_font_9_a_navy_bold_header_and_the_control_total_format():
    """What the log looks like on the used rows (characterisation: the same before and after Task 8B)."""
    from smartass.render.excel import VALUE_FORMAT
    from smartass.xl import styling
    navy = styling._rgb(styling.tn.NAVY)
    wb, t = _book()
    pre = _pre()
    _run(t, wb, pre, status={"fm": "OK"})
    rows, _ = _run(t, wb, pre, redraw=True, status={"fm": "OK"})
    ws = wb.Worksheets("Sources")
    end = rows["fm"]
    for r in range(1, end + 1):
        for c in range(1, len(LOG) + 1):
            assert (_font(ws, r, c, "Name"), _font(ws, r, c, "Size")) == (styling.tn.FONT, 9), (r, c)
    for r in range(1, 4):
        assert (_font(ws, r, P, "Name"), _font(ws, r, P, "Size")) == (styling.tn.FONT, 9), r
    assert all(_font(ws, 1, c, "Bold") is True and _font(ws, 1, c, "Color") == navy for c in range(1, len(LOG) + 1))
    assert _font(ws, 1, P, "Bold") is True and _font(ws, 1, P, "Color") == navy
    assert not _font(ws, 2, 1, "Bold") and not _font(ws, 2, P, "Bold")
    local = styling.tn.to_local(VALUE_FORMAT, fakexl.INTERNATIONAL[2], fakexl.INTERNATIONAL[3])   # nb-NO's codes
    assert all(ws.cells[(r, 11)].get("NumberFormat") == local for r in range(2, end + 1))
    assert ws.colw[sources.LIMITS_COL] == sources.LIMITS_WIDTH and ws.colw.get(4, 0) <= sources.QUESTION_MAX
    assert ws.AutoFilterMode and ws.wb.window_state["Sources"]["FreezePanes"] is True
    assert ws.wb.window_state["Sources"]["SplitRow"] == 1



# ------------------------------------------------------------------ Task 8B: a fast log, its furniture set once
def _second(t, wb, pre=None, **kw):
    """A log that is there (one answer logged), then the calls and the sheet log of a second write."""
    pre = pre or _pre()
    _run(t, wb, pre)
    ws = wb.Worksheets("Sources")
    ws.log.clear()
    return pre, ws


def _runs(monkeypatch) -> list:
    """Every batch run (xl/inproc/runner.run)."""
    from smartass.xl.inproc import runner
    runs, real = [], runner.run
    monkeypatch.setattr(runner, "run", lambda app, b, *a, **k: runs.append(b) or real(app, b, *a, **k))
    return runs


def test_logging_into_a_log_that_is_there_takes_one_write(monkeypatch):
    """Two reads (which sheet is the log; the log) and one write (rows, links, Known limits): in the file job these
    ride in its look, finish and close-out batches."""
    wb, t = _book(("Q3 Answer", "Q3 Data", "Q3 Lines b0"))
    pre = _pre()
    _run(t, wb, pre)
    ws = wb.Worksheets("Q3 Data")
    h = header_rows(3, pre.spec, pre, WHEN, "tom")
    _data_sheet(ws, h)
    runs = _runs(monkeypatch)
    sources.write_sources(t, 3, pre, ws, h, FETCHED, {})
    assert [b.quiet for b in runs] == [False, False, True]


def test_the_furniture_is_set_once_and_a_second_answer_only_appends():
    wb, t = _book()
    pre, ws = _second(t, wb)
    widths, setup, af = dict(ws.colw), dict(ws.PageSetup.__dict__), (ws.AutoFilterMode, ws.autofilter)
    rows, _ = _run(t, wb, pre, redraw=True)
    r = rows["fm"]
    assert {(k, a) for k, a in ws.log} == {("Formula2", f"A{r}:N{r}")}               # the row, in one rectangle
    assert all((_font(ws, r, c, "Name"), _font(ws, r, c, "Size")) == ("Arial", 9) for c in (13, 14))   # the links
    assert ws.colw == widths and ws.PageSetup.__dict__ == setup and (ws.AutoFilterMode, ws.autofilter) == af


def test_fixed_widths_replace_autofit():
    wb, t = _book()
    _run(t, wb, _pre())
    ws = wb.Worksheets("Sources")
    assert not [x for x in ws.log if x[0] == "AutoFit"]
    assert [ws.colw[c] for c in range(1, len(LOG) + 1)] == list(sources.WIDTHS) and ws.colw[P] == sources.LIMITS_WIDTH
    assert ws.colw[4] <= sources.QUESTION_MAX


def test_a_superseding_row_links_to_the_row_it_supersedes():
    wb, t = _book()
    pre = _pre()
    first, _ = _run(t, wb, pre)
    second, _ = _run(t, wb, pre, redraw=True)
    ws = wb.Worksheets("Sources")
    assert _log_row(ws, second["fm"])[13] == f"row {first['fm']}"
    assert _target(ws, second["fm"], 14) == f"'Sources'!A{first['fm']}"


def test_known_limits_are_written_only_when_an_answer_adds_a_line():
    wb, t = _book()
    pre, ws = _second(t, wb)
    _run(t, wb, pre, redraw=True)
    assert not [x for x in ws.log if x[1].startswith("P")]
    more = _pre(caveats=CAVEATS + [{"id": "e", "text": "FinPlan Working is a draft.", "severity": "warn"}])
    _run(t, wb, more, redraw=True)
    assert ("Value", "P1:P4") in ws.log and _v(ws, 4, P) == "FinPlan Working is a draft."


def test_the_placeholder_is_cleared_when_the_first_limit_comes():
    wb, t = _book()
    _run(t, wb, _pre(caveats=[]))
    ws = wb.Worksheets("Sources")
    assert _v(ws, 2, P) == sources.NO_LIMITS
    _run(t, wb, _pre(caveats=[CAVEATS[0]]), redraw=True)
    assert [_v(ws, r, P) for r in (1, 2, 3)] == ["Known limits", CAVEATS[0]["text"], None]


def test_a_link_doubles_the_apostrophes_of_a_sheet_name_and_the_quotes_of_both_texts():
    f = sources.link("O'Brien's \"Q3\"", "A1", 'O\'Brien\'s "Q3"')
    assert f == '=HYPERLINK("#\'O\'\'Brien\'\'s ""Q3""\'!A1","O\'Brien\'s ""Q3""")'
    assert fakexl.hyperlink(f) == ("#'O''Brien''s \"Q3\"'!A1", 'O\'Brien\'s "Q3"')


def _old_log(wb, rows: int = 2):
    """A v2 log as the version before Task 8B left it: the header, rows with REAL hyperlinks, the TN font and the
    control-total format on each used cell, AutoFit widths, the AutoFilter and the print area up to its last row,
    Known limits with the placeholder, and no furniture name."""
    from smartass.render.excel import VALUE_FORMAT
    from smartass.xl import styling
    ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Sources"
    end = rows + 1
    for c, x in enumerate(LOG, 1):
        ws.put(1, c, x)
    for i in range(rows):
        r = 2 + i
        for c, x in enumerate(("2026-11-01 10:00 · tom", "Q3", "fm", "An older question", "FinMax", "Actual", "pov",
                               "Sep 26", "HTTP", 12, 1000.0 * r, "OK"), 1):
            ws.put(r, c, x)
        ws.Hyperlinks.Add(ws.Cells(r, 13), "", "'Q3 Data'!A1", "Q3 Data", "Q3 Data")
        if i:
            ws.put(r, 14, f"row {r - 1}")
    ws.put(1, P, sources.LIMITS_TITLE)
    ws.put(2, P, sources.NO_LIMITS)
    rng = ws.Range(f"A1:N{end}")
    rng.Font.Name, rng.Font.Size = styling.tn.FONT, 9
    ws.Range("A1:N1").Font.Bold = True
    ws.Range(f"K2:K{end}").NumberFormat = VALUE_FORMAT
    for c, w in enumerate((16.4, 3.0, 7.9, 30.1), 1):
        ws.colw[c] = w
    ws.colw[P] = sources.LIMITS_WIDTH
    ws.Range(f"A1:N{end}").AutoFilter()
    ws.PageSetup.PrintArea = f"$A$1:$N${end}"
    ws.Names.Add(Name=styling.PAGE_SETUP_NAME, RefersTo=f'="v1|$A$1:$N${end}"', Visible=False)
    ws.log.clear()
    return ws


def test_a_log_of_the_earlier_version_keeps_its_rows_links_and_filter_and_gets_the_furniture_once(monkeypatch):
    from smartass.render.excel import VALUE_FORMAT
    from smartass.xl import styling
    wb, t = _book()
    ws = _old_log(wb)
    old = {rc: dict(p) for rc, p in ws.cells.items()}
    links = [(h.cell, h.SubAddress, h.TextToDisplay) for h in ws.Hyperlinks]
    widths, af = dict(ws.colw), (ws.AutoFilterMode, ws.autofilter)
    pre = _pre(caveats=[])
    rows, _ = _run(t, wb, pre)
    r = rows["fm"]
    assert r == 4 and _log_row(ws, 4)[13] == "row 3" and _target(ws, 4, 13) == "'Q3 Data'!A1"
    assert all(ws.cells[rc].get("Value") == p.get("Value") for rc, p in old.items())       # nothing rewritten
    assert [(h.cell, h.SubAddress, h.TextToDisplay) for h in ws.Hyperlinks] == links       # the real links stay
    assert ws.colw == widths and (ws.AutoFilterMode, ws.autofilter) == af                  # not toggled off
    assert ("Value", "A1:N1") not in ws.log and [s.Name for s in wb.order].count("Sources") == 1
    local = styling.tn.to_local(VALUE_FORMAT, fakexl.INTERNATIONAL[2], fakexl.INTERNATIONAL[3])   # nb-NO's codes
    assert (_font(ws, r, 1, "Name"), _font(ws, r, 1, "Size"), ws.cells[(r, 11)].get("NumberFormat")) == (
        "Arial", 9, local)
    assert ws.PageSetup.PrintArea == "$A$1:$N$1048575" and sources.FURNITURE_NAME in ws.Names.items
    data = wb.Worksheets("Q3 Data")
    h = header_rows(3, pre.spec, pre, WHEN, "tom")
    _data_sheet(data, h)
    runs = _runs(monkeypatch)
    sources.write_sources(t, 3, pre, data, h, FETCHED, {})
    assert len(runs) == 3 and not [op for op in runs[-1].to_dict()["ops"] if op.get("name") == "AutoFilter"]
    assert sources.FURNITURE_NAME not in str(runs[-1].to_dict()["ops"])     # the furniture is not set again


def test_the_log_kind_reader_without_callers_is_gone():
    """Task 8B review 5: sources.log_kind had no callers left; _kind (over what record_kind read) is the one."""
    assert not hasattr(sources, "log_kind") and callable(sources._kind)
    assert [sources._kind(*x) for x in (("Fetched", "Q", "$A$1:$N$9"), ("Fetched", "Question no", "$A$1:$H$3"),
                                        (None, None, "$A$1"), ("Notes", None, "$A$1:$B$2"))] == [
        "v2", "v1", "empty", "user"]
