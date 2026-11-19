# engine/tests/v2/test_file_redraw.py
"""Phase 3 Tasks 7 and 9: pipeline.redraw_job, a follow-up or refresh - the sheets of Q<n> parked, the new answer
moved in where they were, the old ones dropped; on a failure the new sheets removed and the old ones given their names
back, and once the old ones are gone the new answer kept and said so (never saving the user's workbook). Also
server.py using it for fin_analyze(patch=...) and fin_refresh, and the two carry-over fixes in _fit_text_rows. All on
fakes; every figure is invented."""
from __future__ import annotations

import asyncio
import json
import logging
from types import SimpleNamespace

import pytest
from test_file_job import SUPPLIERS, _calculated_sheet, _point_plan
from test_look import lines_ranking
from test_v2_pipeline import LINE_CELL, FakeProvider, lines_only_spec
from v2util import facts, resolver, spec_dict

import fakexl
from smartass import pipeline, server
from smartass.core import config
from smartass.model.lines import LineSet
from smartass.render import look
from smartass.xl import jobs, movein, sheets
from smartass.xl.host import HostError, InputError
from smartass.xl.target import Target

H = None
PATH = "C:/analyser/answer.xlsx"            # the stored path a follow-up targets (every _target here is replaced)


def _pre(spec=None, lines=None):
    return pipeline.prepare(spec or spec_dict(), resolver(), FakeProvider(facts()), lines=lines)


def _lines_pre():
    """An answer with a lines block: its own 'Q<n> Lines b0' sheet and table Lines_Q<n>_b0."""
    ls = LineSet("b0", list(SUPPLIERS), ["AP_CUSTOMER_NAME", "AMOUNT"], cube_value=-100.0, cell=LINE_CELL)
    return _pre(lines_only_spec(), lines={"b0": ls})


def _names(wb) -> list[str]:
    return [ws.Name for ws in wb.order]


def _tables(wb) -> list[str]:
    return [lo.Name for ws in wb.order for lo in ws.ListObjects]


def _temp_files(home) -> list:
    return list((home / "tmp").glob("*.xlsx")) if (home / "tmp").exists() else []


class Book:
    """A fake workbook the answers are drawn into, each job with a Target of its own (as every live job resolves
    one); the fake Excel does not calculate, so the read-back sees the model's values where the plan put them."""

    def __init__(self, monkeypatch, wb, ours: bool):
        self.wb, self.ours, self.targets, self.pre = wb, ours, [], None
        wb.saved = 0

        def save():
            wb.saved += 1
        wb.Save = save

        def target(h, target, title):
            self.targets.append(Target(app=wb.Application, wb=wb, pid=0, hwnd=0, created=ours))
            return self.targets[-1]
        real = pipeline.read_block

        def read_block(ws, plan):
            block = [list(row) for row in real(ws, plan)]
            for key in plan.addr:
                r, c = plan.at(key)
                block[r - 1][c - 1] = self.pre.result.cells.get(key)
            return tuple(tuple(row) for row in block)
        monkeypatch.setattr(pipeline, "_target", target)
        monkeypatch.setattr(pipeline, "read_block", read_block)
        monkeypatch.setattr(pipeline, "remember", lambda pre, out: None)

    def answer(self, pre) -> dict:
        self.pre = pre
        return pipeline.excel_job(H, pre)

    def follow_up(self, pre, q: int) -> dict:
        self.pre = pre
        return pipeline.redraw_job(H, pre, q, PATH)


@pytest.fixture
def tmp_home(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "HOME", tmp_path)
    return tmp_path


@pytest.fixture
def own_book(monkeypatch, tmp_home):
    """Our own analysis file: Q1 and Q2 (a table each), Q3 with ledger lines, and the user's 'Notes' and 'Q3 Notes'
    (a sheet named like a topic of Q3 that the engine did not draw). Sheet order 'full': the Answers first, the Data
    and Lines sheets after the user's sheets - so Q3's sheets are not side by side."""
    b = Book(monkeypatch, fakexl.Workbook(["Notes", "Q3 Notes", "Sources"]), ours=True)
    b.answer(_pre())
    b.answer(_pre())
    b.answer(_lines_pre())
    assert _names(b.wb) == ["Q1 Answer", "Q2 Answer", "Q3 Answer", "Notes", "Q3 Notes", "Q1 Data", "Q2 Data",
                            "Q3 Data", "Q3 Lines b0", "Sources", "_svmeta"]
    return b


@pytest.fixture
def moves(monkeypatch) -> list:
    """Every Worksheet.Move the job makes (sheetorder.arrange_sheets moves one sheet per call)."""
    seen, real = [], fakexl.Worksheet.Move

    def move(ws, *args, **kw):
        seen.append(ws.Name)
        return real(ws, *args, **kw)
    monkeypatch.setattr(fakexl.Worksheet, "Move", move)
    return seen


# ------------------------------------------------------------------ the sheets replaced
def test_a_follow_up_replaces_every_sheet_of_the_question(own_book, tmp_home, moves):
    wb = own_book.wb
    before = _names(wb)
    moves.clear()
    out = own_book.follow_up(_pre(), 3)
    names = _names(wb)
    assert "Q3 Lines b0" not in names and not [n for n in names if n.startswith("old ")]     # the dropped block's sheet goes
    assert names == [n for n in before if n != "Q3 Lines b0"]                                 # same place in the workbook
    assert names.index("Q3 Answer") == before.index("Q3 Answer")
    assert moves == ["Q3 Data"]                     # to its old place; the sheet order then has nothing to move
    assert _tables(wb) == ["Data_Q1", "Data_Q2", "Data_Q3"] and not [n for n in _tables(wb) if n.endswith("_old")]
    assert out["followup"] is True and out["q"] == 3 and out["sheets"] == ["Q3 Answer", "Q3 Data"]
    assert out["checks"]["sheet_equals_model"] == "OK"
    assert not [w for w in out["warnings"] if "stored range" in w]
    assert wb.saved == 4 and not _temp_files(tmp_home) and len(wb.Application.Workbooks) == 1
    app = wb.Application
    assert (app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts, app.Calculation) == (True, True, True, -4105)


def test_the_new_sheets_hold_the_new_answer(own_book):
    wb = own_book.wb
    old_data = fakexl.sheet_state(wb.Worksheets("Q1 Data"))
    own_book.follow_up(_pre(), 3)
    assert wb.Worksheets("Q3 Data").ListObjects(1).Name == "Data_Q3"
    assert fakexl.sheet_state(wb.Worksheets("Q3 Data"))["formats"] == old_data["formats"]   # Q1's look: same spec
    assert wb.Worksheets("Q3 Answer").Range("B1").Value == wb.Worksheets("Q1 Answer").Range("B1").Value


def test_a_follow_up_that_adds_a_lines_block_brings_its_sheet(own_book):
    wb = own_book.wb
    out = own_book.follow_up(_lines_pre(), 1)
    assert out["sheets"] == ["Q1 Answer", "Q1 Data", "Q1 Lines b0"]
    assert "Q1 Lines b0" in _names(wb) and "Lines_Q1_b0" in _tables(wb)
    assert not [n for n in _names(wb) if n.startswith("old ")]


def test_six_follow_ups_in_a_row_keep_the_order_and_leave_nothing_behind(own_book):
    wb = own_book.wb
    own_book.follow_up(_pre(), 3)
    after_first = _names(wb)
    for i in range(6):
        out = own_book.follow_up(_lines_pre() if i % 2 else _pre(), 2)
        assert out["checks"]["sheet_equals_model"] == "OK"
        assert not [n for n in _names(wb) if n.startswith("old ")]
        assert not [n for n in _tables(wb) if n.endswith(("_old", "_old2"))]
    assert [n for n in _names(wb) if n != "Q2 Lines b0"] == after_first


def test_a_move_excel_refuses_leaves_the_old_answer_exactly_as_it_was(monkeypatch, tmp_home):
    b = Book(monkeypatch, fakexl.Workbook(["Intro", "Sources", "Budget"]), ours=False)
    b.answer(_lines_pre())
    before = fakexl.snapshot(b.wb)

    def refuse(ws, *args, **kw):
        raise fakexl.ComError("the move was refused")
    monkeypatch.setattr(fakexl.Worksheet, "Move", refuse)
    with pytest.raises(fakexl.ComError, match="refused"):
        b.follow_up(_lines_pre(), 1)
    assert fakexl.snapshot(b.wb) == before


def test_a_question_whose_sheets_are_side_by_side_is_replaced_without_a_move(monkeypatch, tmp_home, moves):
    b = Book(monkeypatch, fakexl.Workbook(["Sources"]), ours=True)
    b.answer(_lines_pre())
    before = _names(b.wb)
    assert before == ["Q1 Answer", "Q1 Data", "Q1 Lines b0", "Sources", "_svmeta"]
    moves.clear()
    b.follow_up(_lines_pre(), 1)
    assert _names(b.wb) == before and moves == []


def test_in_the_users_workbook_the_question_keeps_its_places_between_the_users_sheets(monkeypatch, tmp_home, moves):
    b = Book(monkeypatch, fakexl.Workbook(["Intro", "Sources", "Budget"]), ours=False)
    b.answer(_lines_pre())
    before = _names(b.wb)
    assert before == ["Intro", "Q1 Answer", "Budget", "Q1 Data", "Q1 Lines b0", "Sources", "_svmeta"]
    moves.clear()
    b.follow_up(_lines_pre(), 1)
    assert _names(b.wb) == before and b.wb.saved == 0                       # the user's workbook is never saved
    assert moves == ["Q1 Lines b0", "Q1 Data"]                               # only new sheets move, to their places


def test_the_users_active_sheet_on_the_old_answer_lands_on_the_new_answer(own_book):
    wb = own_book.wb
    wb.Worksheets("Q3 Answer").Activate()
    own_book.follow_up(_pre(), 3)
    assert wb.active is wb.Worksheets("Q3 Answer") and wb.active in wb.order


# ------------------------------------------------------------------ what else the follow-up touches
def test_other_questions_and_the_users_own_sheets_are_untouched(own_book):
    wb = own_book.wb
    keep = [n for n in _names(wb) if n not in ("Q3 Answer", "Q3 Data", "Q3 Lines b0", "Sources", "_svmeta")]
    assert "Q3 Notes" in keep and "Notes" in keep
    before = {n: fakexl.sheet_state(wb.Worksheets(n)) for n in keep}
    own_book.follow_up(_pre(), 3)
    assert {n: fakexl.sheet_state(wb.Worksheets(n)) for n in keep} == before


def _log(ws) -> dict:
    """The Sources log, A:N of every row that holds a value: {row: values}."""
    return {r: tuple(ws.cells.get((r, c), {}).get("Value") for c in range(1, 15))
            for r in sorted({r for (r, c), p in ws.cells.items() if c <= 14 and "Value" in p})}


def test_the_sources_log_keeps_its_rows_and_logs_the_new_pull(own_book):
    ws = own_book.wb.Worksheets("Sources")
    old = _log(ws)
    q3 = [r for r, v in old.items() if v[1] == "Q3"]
    own_book.follow_up(_pre(), 3)
    new = _log(ws)
    assert {r: new[r] for r in old} == old                                   # earlier rows are never edited
    added = [v for r, v in new.items() if r not in old]
    assert len(added) == 1 and added[0][1:3] == ("Q3", "fm")                  # the new pull, query id 'fm'
    assert fakexl.hyperlink(added[0][12]) == ("#'Q3 Data'!A1", "Q3 Data") and not added[0][13]   # a first pull in Q3
    assert q3 and fakexl.hyperlink(new[q3[0]][12]) == ("#'Q3 Lines b0'!A1", "Q3 Lines b0")   # names as text


def test_the_stored_spec_of_the_question_is_rewritten_without_a_drawn_range(own_book):
    t = SimpleNamespace(wb=own_book.wb)
    before = sheets._stored_specs(t)
    pre = _pre(spec_dict(question="Business P&L, the follow-up"))
    own_book.follow_up(pre, 3)
    after = sheets._stored_specs(t)
    assert after[3] == {"spec": {"v2": pre.raw}, "ids": {}, "owned": None}
    assert {q: after[q] for q in (1, 2)} == {q: before[q] for q in (1, 2)}


def test_a_workbook_drawn_by_the_old_renderer_is_redrawn(monkeypatch, tmp_home):
    """The COM renderer's sheets have the same names and tables; the drawn range it stored with the spec is not
    used, and the rewritten spec stores none."""
    wb = fakexl.Workbook(["Q1 Answer", "Q1 Data", "Q1 Lines b0", "Sources", "Q1 Notes"])
    data, lines = wb.Worksheets("Q1 Data"), wb.Worksheets("Q1 Lines b0")
    data.ListObjects.Add(1, data.Range("A7:C9"), None, 1).Name = "Data_Q1"
    lines.ListObjects.Add(1, lines.Range("A4:B6"), None, 1).Name = "Lines_Q1_b0"
    wb.Worksheets("Q1 Answer").put(3, 2, "an old answer's text")
    meta = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
    sheets._name_hidden(meta, sheets.META_SHEET)
    old = SimpleNamespace(wb=wb)
    sheets.save_spec(old, 1, {"v2": spec_dict()}, ids={}, owned={"rows": 40, "cols": 9})
    wb.Worksheets("Q1 Notes").put(1, 1, "my own notes")
    mine = fakexl.sheet_state(wb.Worksheets("Q1 Notes"))
    b = Book(monkeypatch, wb, ours=False)
    out = b.follow_up(_pre(), 1)
    assert out["sheets"] == ["Q1 Answer", "Q1 Data"] and out["checks"]["sheet_equals_model"] == "OK"
    assert _names(wb) == ["Q1 Answer", "Q1 Data", "Sources", "Q1 Notes", "_svmeta"]
    assert fakexl.sheet_state(wb.Worksheets("Q1 Notes")) == mine                    # the user's sheet, not swept
    assert _tables(wb) == ["Data_Q1"] and wb.Worksheets("Q1 Answer").Range("B3").Value != "an old answer's text"
    assert sheets._stored_specs(old)[1]["owned"] is None


# ------------------------------------------------------------------ failures
def test_a_failed_follow_up_leaves_the_old_answer_exactly_as_it_was(own_book, monkeypatch, tmp_home):  # Focus 3
    wb, app = own_book.wb, own_book.wb.Application
    before = fakexl.snapshot(wb)

    def finish(t, built):
        raise HostError("Excel is busy")
    monkeypatch.setattr(movein, "finish", finish)
    with pytest.raises(HostError, match="busy"):
        own_book.follow_up(_pre(), 3)
    assert fakexl.snapshot(wb) == before and own_book.targets[-1].sheets_added == []
    assert (app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts, app.Calculation) == (True, True, True, -4105)
    assert not _temp_files(tmp_home) and len(app.Workbooks) == 1 and wb.saved == 3


def test_a_failed_drop_puts_the_old_sheets_back_in_their_places(own_book, monkeypatch):
    wb = own_book.wb
    order = _names(wb)
    q3 = {n: fakexl.sheet_state(wb.Worksheets(n)) for n in ("Q3 Answer", "Q3 Data", "Q3 Lines b0")}

    before = fakexl.snapshot(wb)

    def drop(t, parked, *a, **k):
        raise HostError("the delete was refused")
    monkeypatch.setattr(movein, "drop", drop)
    with pytest.raises(HostError, match="refused"):
        own_book.follow_up(_pre(), 3)
    assert _names(wb) == order and _tables(wb) == ["Data_Q1", "Data_Q2", "Data_Q3", "Lines_Q3_b0"]
    assert {n: fakexl.sheet_state(wb.Worksheets(n)) for n in q3} == q3
    assert fakexl.snapshot(wb) == before                # the drop comes before Sources and the spec: both as they were


def test_a_drop_that_deletes_and_then_fails_keeps_the_new_answer(own_book, monkeypatch):
    """The old sheets are gone once the Delete went through: a rollback then would leave neither answer."""
    wb, real = own_book.wb, movein.drop

    def drop(t, parked, *a, **k):
        real(t, parked, *a, **k)
        raise HostError("could not restore Application.DisplayAlerts")
    monkeypatch.setattr(movein, "drop", drop)
    with pytest.raises(HostError, match="DisplayAlerts") as e:
        own_book.follow_up(_pre(), 3)
    assert not isinstance(e.value, pipeline.RestoreError)
    assert {"Q3 Answer", "Q3 Data"} <= set(_names(wb)) and not [n for n in _names(wb) if n.startswith("old ")]
    assert "Data_Q3" in _tables(wb)


def test_a_failure_after_the_drop_keeps_the_new_answer_saves_our_file_and_says_so(own_book, monkeypatch):
    def check(*a):
        raise HostError("the job was abandoned")
    monkeypatch.setattr(pipeline, "_check", check)
    saved = own_book.wb.saved
    with pytest.raises(HostError, match="abandoned") as e:
        own_book.follow_up(_pre(), 3)
    assert not isinstance(e.value, pipeline.RestoreError)
    assert str(e.value).startswith("the job was abandoned") and "the new answer is in place" in str(e.value)
    assert own_book.wb.saved == saved + 1
    assert "Q3 Answer" in _names(own_book.wb) and "Q3 Lines b0" not in _names(own_book.wb)


def test_a_save_that_fails_after_the_drop_keeps_the_new_answer_and_says_so_without_saving_again(own_book,
                                                                                                monkeypatch):
    """Task 8B review 2: the save comes after the old sheets are gone; its failure must say the new answer is there
    (and the save is not tried a second time)."""
    def save(t):
        raise HostError("the file is locked")
    monkeypatch.setattr(pipeline, "_save", save)
    saved = own_book.wb.saved
    with pytest.raises(HostError, match="locked") as e:
        own_book.follow_up(_pre(), 3)
    assert not isinstance(e.value, pipeline.RestoreError)
    assert str(e.value) == "the file is locked" + pipeline.KEPT and own_book.wb.saved == saved
    assert "Q3 Answer" in _names(own_book.wb) and not [n for n in _names(own_book.wb) if n.startswith("old ")]


def test_a_failure_after_the_drop_never_saves_the_users_workbook(monkeypatch, tmp_home):
    """Task 8B review 6: _kept saves our own file only."""
    b = Book(monkeypatch, fakexl.Workbook(["Intro", "Sources", "Budget"]), ours=False)
    b.answer(_pre())

    def check(*a):
        raise HostError("the job was abandoned")
    monkeypatch.setattr(pipeline, "_check", check)
    with pytest.raises(HostError, match="abandoned") as e:
        b.follow_up(_pre(), 1)
    assert str(e.value) == "the job was abandoned" + pipeline.KEPT and "saved" not in str(e.value)
    assert b.wb.saved == 0 and "Q1 Answer" in _names(b.wb)


def test_a_follow_up_in_the_active_workbook_of_a_hidden_excel_ends_on_the_new_answer(own_book):
    """Task 8B review 1: quiet lands on the new Answer whenever the target is the active workbook, visible or not
    (the Book's active sheet, renamed by the park, is not trusted)."""
    wb = own_book.wb
    wb.Application.Visible = False
    wb.Worksheets("Notes").Activate()
    own_book.follow_up(_pre(), 3)
    assert wb.Application.ActiveWorkbook is wb and wb.active is wb.Worksheets("Q3 Answer") and wb.active in wb.order


def test_a_follow_up_makes_at_most_37_com_calls_and_reads_the_sheets_once(own_book, monkeypatch):
    from smartass.core.timing import COM_PREFIX
    from smartass.xl import sheetorder
    reads, real = [], movein._read_book
    monkeypatch.setattr(movein, "_read_book", lambda wb: reads.append(wb) or real(wb))
    monkeypatch.setattr(movein, "_inventory", lambda wb: pytest.fail("the sheets read again"))
    monkeypatch.setattr(sheetorder, "_read_sheets", lambda wb, busy: pytest.fail("the order read again"))
    calls = fakexl.count_com(monkeypatch)
    own_book.pre = _pre()
    out, timings = pipeline.with_timings(lambda h: pipeline.redraw_job(h, own_book.pre, 3, PATH))(H)
    assert len(calls) <= 37, {k: v for k, v in timings.items() if k.startswith(COM_PREFIX)}    # 35 (19.11.2026)
    assert len(reads) == 1 and out["checks"]["sheet_equals_model"] == "OK"


def test_a_failed_follow_up_rolls_back_before_it_restores_and_never_after(own_book, monkeypatch):
    seen = []
    real_rollback, real_restore = pipeline._rollback, movein.restore
    monkeypatch.setattr(pipeline, "_rollback", lambda t: seen.append(("rollback", list(t.sheets_added)))
                        or real_rollback(t))
    monkeypatch.setattr(movein, "restore", lambda t, parked: seen.append(("restore", dict(parked.sheets)))
                        or real_restore(t, parked))
    monkeypatch.setattr(movein, "finish", lambda t, built: (_ for _ in ()).throw(HostError("refused")))
    with pytest.raises(HostError):
        own_book.follow_up(_pre(), 3)
    assert seen == [("rollback", ["Q3 Answer", "Q3 Data"]),
                    ("restore", {"Q3 Answer": "old Q3 Answer", "Q3 Data": "old Q3 Data",
                                 "Q3 Lines b0": "old Q3 Lines b0"})]


def test_a_restore_that_fails_is_a_restore_error(own_book, monkeypatch):
    def finish(t, built):
        raise HostError("a formula was refused")

    def restore(t, parked):
        raise HostError("could not put back the names of sheet 'old Q3 Data' -> 'Q3 Data'")
    monkeypatch.setattr(movein, "finish", finish)
    monkeypatch.setattr(movein, "restore", restore)
    with pytest.raises(pipeline.RestoreError) as e:
        own_book.follow_up(_pre(), 3)
    assert "a formula was refused" in str(e.value) and "the restore was incomplete" in str(e.value)
    assert "'old Q3 Data' -> 'Q3 Data'" in str(e.value) and isinstance(e.value.__cause__, HostError)
    assert "Q3 Answer" not in _names(own_book.wb)                  # the new sheets are gone; the old stay parked


@pytest.mark.parametrize("missing", ["Q3 Answer", "Q3 Data"])
def test_a_question_without_its_answer_or_data_sheet_is_refused_by_name(own_book, missing, tmp_home):
    wb = own_book.wb
    wb.order.remove(wb.Worksheets(missing))
    before = fakexl.snapshot(wb)
    with pytest.raises(InputError, match=f"'{missing}'"):
        own_book.follow_up(_pre(), 3)
    assert fakexl.snapshot(wb) == before and not _temp_files(tmp_home) and len(wb.Application.Workbooks) == 1


def test_a_table_name_taken_elsewhere_is_refused_and_the_old_answer_stays(own_book):
    wb = own_book.wb
    notes = wb.Worksheets("Notes")
    notes.ListObjects.Add(1, notes.Range("A1:B3"), None, 1).Name = "Lines_Q2_b0"   # not on a sheet of Q2
    before = fakexl.snapshot(wb)
    with pytest.raises(movein.NameTaken, match="Lines_Q2_b0"):
        own_book.follow_up(_lines_pre(), 2)
    assert fakexl.snapshot(wb) == before


# ------------------------------------------------------------------ which sheets belong to the question
def test_each_new_sheet_takes_its_namesakes_place_and_an_added_one_follows_the_sheet_before_it():
    parked = {"Q3 Answer": "old Q3 Answer", "Q3 Data": "old Q3 Data", "Q3 Lines b1": "old Q3 Lines b1",
              "Q3 Lines b2": "old Q3 Lines b2"}
    names = ["Notes", "Q3 Answer", "Q3 Data", "Q3 Lines b0", "Q3 Lines b1", "old Q3 Answer", "Budget",
             "old Q3 Lines b1", "Q2 Data", "old Q3 Data", "old Q3 Lines b2", "Sources"]
    new = ["Q3 Answer", "Q3 Data", "Q3 Lines b0", "Q3 Lines b1"]
    assert pipeline.wanted_order(names, new, parked) == [
        "Notes", "Q3 Answer", "Budget", "Q3 Lines b1", "Q2 Data", "Q3 Data", "Q3 Lines b0", "Sources"]
    side_by_side = ["Q3 Answer", "Q3 Data", "old Q3 Answer", "old Q3 Data", "Sources"]
    assert pipeline.wanted_order(side_by_side, new[:2], parked) == ["Q3 Answer", "Q3 Data", "Sources"]



def test_the_sheets_of_a_question_are_its_answer_data_and_lines_blocks_only():
    names = ["Q3 Answer", "Q3 Notes", "Q3 Data", "Q3 Lines b0", "Q3 Lines notes", "Q3 Lines", "Q13 Data",
             "Q3 Lines b12", "Q3 FinMax", "old Q3 Data", "Q30 Lines b0"]
    assert sheets.answer_sheets(names, 3) == ["Q3 Answer", "Q3 Data", "Q3 Lines b0", "Q3 Lines b12"]
    assert sheets.answer_tables(3, ["Data_Q3", "Data_Q30", "Lines_Q3_b0", "Lines_Q30_b0", "Mine", "Lines_Q3"]) == [
        "Data_Q3", "Lines_Q3_b0"]


# ------------------------------------------------------------------ server.py uses redraw_job
def _wire_server(monkeypatch, stored: dict, job):
    async def user(fn, what):
        return fn(None)

    async def prepare(raw, provider, spec=None, target="session", complete=False):
        return SimpleNamespace(raw=raw, target=target), []
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "_provider", lambda: "provider")
    monkeypatch.setattr(server, "_prepare", prepare)
    monkeypatch.setattr(server, "_default_target", lambda target: target)
    monkeypatch.setattr(server, "read_stored", lambda h, target: (stored, "Q1 Answer", PATH))
    monkeypatch.setattr(server, "redraw_job", lambda h, pre, q, target: job(pre, q, target))


def _stored(*qs) -> dict:
    return {q: {"spec": {"v2": spec_dict(question=f"Question {q}")}, "ids": {}, "owned": {"rows": 3, "cols": 2}}
            for q in qs}


def test_fin_refresh_and_the_follow_up_path_use_redraw_job(monkeypatch):
    """A drawn range an earlier version stored with the spec does not reach the job."""
    seen = []

    def job(pre, q, target):
        seen.append((q, target))
        return {"q": q, "workbook": target, "closed": {}}
    _wire_server(monkeypatch, _stored(1, 2), job)
    out = asyncio.run(server.fin_analyze(patch={"question": "Question 1, by month"}, q=1, target=PATH))
    assert out["q"] == 1 and seen == [(1, PATH)]
    seen.clear()
    out = asyncio.run(server.fin_refresh(target=PATH))
    assert [d["ok"] for d in out["refreshed"]] == [True, True] and seen == [(1, PATH), (2, PATH)]


def test_fin_refresh_goes_on_after_a_failure_and_stops_after_a_failed_restore(monkeypatch):
    seen = []

    def job(pre, q, target):
        seen.append(q)
        if q == 1:
            raise HostError("Excel is busy")
        if q == 3:
            raise pipeline.RestoreError("refused; the restore was incomplete - sheet 'old Q3 Data' -> 'Q3 Data'")
        return {"q": q, "workbook": target, "closed": {}}
    _wire_server(monkeypatch, _stored(1, 2, 3, 4), job)
    out = asyncio.run(server.fin_refresh(target=PATH))
    got = {d["q"]: d for d in out["refreshed"]}
    assert seen == [1, 2, 3] and [got[q]["ok"] for q in (1, 2, 3, 4)] == [False, True, False, False]
    assert "busy" in got[1]["error"] and "restore was incomplete" in got[3]["error"]
    assert got[4]["skipped"] is True and "Q3" in got[4]["error"] and out["workbook"] == PATH


# ------------------------------------------------------------------ carry-overs in _fit_text_rows
def test_the_key_point_rows_are_addressed_inside_the_com_call(monkeypatch):
    """ws.Range is a COM call too: it must run inside jobs.com (busy retry, abort check, the count)."""
    long = "•  Opex 1 234 567 NOK vs PY; " + "consultants, licences and the new data centre " * 3
    plan, block = _point_plan(long)
    ws = fakexl.Workbook(["Q1 Answer"]).order[0]
    inside, outside = [], []
    real = jobs.com

    def com(call, *a, **k):
        inside.append(True)
        try:
            return real(call, *a, **k)
        finally:
            inside.pop()
    monkeypatch.setattr(jobs, "com", com)

    class Sheet:
        def __getattr__(self, k):
            return getattr(ws, k)

        def Range(self, a, b=None):
            if not inside:
                outside.append(a)
            return ws.Range(a, b)
    pipeline._fit_text_rows(Sheet(), plan, block)
    assert outside == [] and ws.rowh[2] == ws.rowh[3] > 0


def test_a_break_excel_refuses_after_the_reset_keeps_the_files_breaks(monkeypatch, caplog):
    plan, values, ws, block = _calculated_sheet(lines_ranking)
    file_breaks = look.answer_look(plan).breaks
    new = look.answer_look(plan, texts={rc: v for rc, v in values.items() if isinstance(v, str)}).breaks
    refused = next(r for r in new if r not in file_breaks)

    def add(sheet, rng):
        if rng.Row == refused:
            raise fakexl.ComError("Excel refused the page break")
        sheet.page_breaks.append(rng.Row)
    monkeypatch.setattr(fakexl.Worksheet, "HPageBreaks", property(lambda s: SimpleNamespace(Add=lambda r: add(s, r))))
    with caplog.at_level(logging.WARNING):
        pipeline._fit_text_rows(ws, plan, block)
    assert ws.page_breaks == file_breaks and any("page breaks" in r.getMessage() for r in caplog.records)


def test_the_spec_rows_parse_as_json(own_book):
    """_svmeta stays readable after many follow-ups: one row per question, no stale chunk left beside it."""
    for _ in range(3):
        own_book.follow_up(_pre(), 2)
    meta = own_book.wb.Worksheets(sheets.META_SHEET)
    rows = [r for (r, c), p in meta.cells.items() if c == 1 and str(p.get("Value", "")).startswith("SPEC Q")]
    assert sorted(meta.cells[(r, 1)]["Value"] for r in rows) == ["SPEC Q1", "SPEC Q2", "SPEC Q3"]
    for r in rows:
        json.loads("".join(str(meta.cells[(r, c)]["Value"]) for c in range(2, 4) if (r, c) in meta.cells
                           and "Value" in meta.cells[(r, c)]))
