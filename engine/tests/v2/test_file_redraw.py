# engine/tests/v2/test_file_redraw.py
"""Phase 3 Tasks 7 and 9: exceljob.redraw_job, a follow-up or refresh - the sheets of Q<n> parked, the new answer
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
from test_file_job import BATCHES, SUPPLIERS, _batches, _calculated_sheet, _point_plan, fit
from test_look import lines_ranking
from test_v2_pipeline import LINE_CELL, FakeProvider, lines_only_spec
from v2util import facts, resolver, spec_dict

import fakexl
from smartass import exceljob, pipeline, server
from smartass.core import config
from smartass.model.lines import LineSet
from smartass.render import look
from smartass.xl import movein, sheets
from smartass.xl.host import HostError, InputError
from smartass.xl.inproc import runner
from smartass.xl.inproc.batch import Batch
from smartass.xl.inproc.runner import BatchError
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

        def target(h, target, title, build_new=False):
            self.targets.append(Target(app=wb.Application, wb=wb, pid=0, hwnd=0, created=ours))
            return self.targets[-1]
        real = exceljob.read_block

        def read_block(ws, plan, b):
            get = real(ws, plan, b)

            def block(res):
                got = [list(row) for row in get(res)]
                for key in plan.addr:
                    r, c = plan.at(key)
                    got[r - 1][c - 1] = self.pre.result.cells.get(key)
                return tuple(tuple(row) for row in got)
            return block
        monkeypatch.setattr(exceljob, "_target", target)
        monkeypatch.setattr(exceljob, "read_block", read_block)
        monkeypatch.setattr(exceljob, "remember", lambda pre, out: None)

    def answer(self, pre) -> dict:
        self.pre = pre
        return exceljob.excel_job(H, pre)

    def follow_up(self, pre, q: int) -> dict:
        self.pre = pre
        return exceljob.redraw_job(H, pre, q, PATH)


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
    with pytest.raises(BatchError, match="Move.*refused"):
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
    old = SimpleNamespace(wb=wb, app=wb.Application)
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

    def refuse(self, *a):
        raise fakexl.ComError("a table was refused")
    monkeypatch.setattr(fakexl._ListObjects, "Add", refuse)
    with pytest.raises(BatchError, match="refused"):
        own_book.follow_up(_pre(), 3)
    assert fakexl.snapshot(wb) == before and own_book.targets[-1].sheets_added == []
    assert (app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts, app.Calculation) == (True, True, True, -4105)
    assert not _temp_files(tmp_home) and len(app.Workbooks) == 1 and wb.saved == 3


def test_a_follow_up_whose_insert_fails_half_way_puts_the_old_answer_back(own_book, monkeypatch, tmp_home):
    """Excel put some of the new sheets in, then the insert failed: they are removed first (rollback), then the
    parked old sheets get their names back (restore); the workbook is exactly as it was."""
    wb, app = own_book.wb, own_book.wb.Application
    before = fakexl.snapshot(wb)
    real = fakexl._Sheets.Add

    def half(self, *args):
        real(self, *args)
        self.wb.order.remove(self("Q3 Data"))
        raise fakexl.ComError("Add method of Sheets class failed")
    monkeypatch.setattr(fakexl._Sheets, "Add", half)
    with pytest.raises(BatchError, match="Add method"):
        own_book.follow_up(_pre(), 3)
    assert fakexl.snapshot(wb) == before and own_book.targets[-1].sheets_added == []
    assert not _temp_files(tmp_home) and len(app.Workbooks) == 1 and wb.saved == 3


def test_a_failed_drop_puts_the_old_sheets_back_in_their_places(own_book, monkeypatch):
    wb = own_book.wb
    order = _names(wb)
    q3 = {n: fakexl.sheet_state(wb.Worksheets(n)) for n in ("Q3 Answer", "Q3 Data", "Q3 Lines b0")}

    before = fakexl.snapshot(wb)

    real = fakexl._SheetGroup.Delete

    def refuse(group):
        if group.sheets[0].Name.startswith("old "):    # the drop; the rollback's deletes go through
            raise fakexl.ComError("the delete was refused")
        real(group)
    monkeypatch.setattr(fakexl._SheetGroup, "Delete", refuse)
    with pytest.raises(BatchError, match="Delete.*refused"):
        own_book.follow_up(_pre(), 3)
    assert _names(wb) == order and _tables(wb) == ["Data_Q1", "Data_Q2", "Data_Q3", "Lines_Q3_b0"]
    assert {n: fakexl.sheet_state(wb.Worksheets(n)) for n in q3} == q3
    assert fakexl.snapshot(wb) == before                # the drop comes before Sources and the spec: both as they were


def test_a_drop_that_deletes_and_then_fails_keeps_the_new_answer(own_book, monkeypatch):
    """The old sheets are gone once the Delete went through: a rollback then would leave neither answer."""
    wb, real = own_book.wb, fakexl._SheetGroup.Delete

    def delete(group):
        real(group)
        raise fakexl.ComError("the Delete went through and then failed")
    monkeypatch.setattr(fakexl._SheetGroup, "Delete", delete)
    with pytest.raises(BatchError, match="then failed") as e:
        own_book.follow_up(_pre(), 3)
    assert not isinstance(e.value, exceljob.RestoreError)
    assert {"Q3 Answer", "Q3 Data"} <= set(_names(wb)) and not [n for n in _names(wb) if n.startswith("old ")]
    assert "Data_Q3" in _tables(wb)


def test_a_failure_after_the_drop_keeps_the_new_answer_saves_our_file_and_says_so(own_book, monkeypatch):
    def check(*a):
        raise HostError("the job was abandoned")
    monkeypatch.setattr(exceljob, "_check", check)
    saved = own_book.wb.saved
    with pytest.raises(HostError, match="abandoned") as e:
        own_book.follow_up(_pre(), 3)
    assert not isinstance(e.value, exceljob.RestoreError)
    assert str(e.value).startswith("the job was abandoned") and "the new answer is in place" in str(e.value)
    assert own_book.wb.saved == saved + 1
    assert "Q3 Answer" in _names(own_book.wb) and "Q3 Lines b0" not in _names(own_book.wb)


def test_a_save_that_fails_after_the_drop_keeps_the_new_answer_and_says_so_without_saving_again(own_book,
                                                                                                monkeypatch):
    """Task 8B review 2: the save comes after the old sheets are gone; its failure must say the new answer is there
    (and the save is not tried a second time)."""
    def save(t):
        raise HostError("the file is locked")
    monkeypatch.setattr(exceljob, "_save", save)
    saved = own_book.wb.saved
    with pytest.raises(HostError, match="locked") as e:
        own_book.follow_up(_pre(), 3)
    assert not isinstance(e.value, exceljob.RestoreError)
    assert str(e.value) == "the file is locked" + exceljob.KEPT and own_book.wb.saved == saved
    assert "Q3 Answer" in _names(own_book.wb) and not [n for n in _names(own_book.wb) if n.startswith("old ")]


def test_a_failure_after_the_drop_never_saves_the_users_workbook(monkeypatch, tmp_home):
    """Task 8B review 6: _kept saves our own file only."""
    b = Book(monkeypatch, fakexl.Workbook(["Intro", "Sources", "Budget"]), ours=False)
    b.answer(_pre())

    def check(*a):
        raise HostError("the job was abandoned")
    monkeypatch.setattr(exceljob, "_check", check)
    with pytest.raises(HostError, match="abandoned") as e:
        b.follow_up(_pre(), 1)
    assert str(e.value) == "the job was abandoned" + exceljob.KEPT and "saved" not in str(e.value)
    assert b.wb.saved == 0 and "Q1 Answer" in _names(b.wb)


def _spec_row(wb, q: int):
    meta = wb.Worksheets(sheets.META_SHEET)
    return next(meta.Range(f"A{r}:C{r}").Value2 for r in range(1, 50) if meta.Range(f"A{r}").Value2 == f"SPEC Q{q}")


def test_a_spec_that_could_not_be_stored_after_the_drop_is_said_with_the_kept_answer(own_book, monkeypatch):
    """Final review M4: the close-out fails at the spec, after the old sheets are gone: the new answer is kept, and
    the error says that the spec stored is still the one before (a next follow-up would patch that one)."""
    from smartass.xl import jobs
    wb = own_book.wb
    old = _spec_row(wb, 3)

    def save_spec(t, q, raw, b, rows):                  # Excel refuses the spec's write
        b.workbook(jobs.book_name(t)).Worksheets("no such sheet").Range("A1").Value = "the spec"
        return False
    monkeypatch.setattr(exceljob, "_save_spec", save_spec)
    with pytest.raises(BatchError) as e:
        own_book.follow_up(_pre(spec_dict(question="Opex by month, again")), 3)
    assert exceljob.KEPT in str(e.value) and exceljob.SPEC_NOT_STORED.format(q=3) in str(e.value)
    assert _spec_row(wb, 3) == old and "Q3 Answer" in _names(wb)


def test_a_spec_whose_store_cannot_be_told_after_the_drop_is_said_as_such(own_book, monkeypatch):
    """The close-out's failure has no index (the job was abandoned in it): the spec may or may not be stored."""
    real = runner.run

    def run(app, b, *a, **kw):
        res = real(app, b, *a, **kw)
        if any(op.get("k") == "mark" and op.get("name") == "spec" for op in b.to_dict()["ops"]):
            raise BatchError(None, "the batch in Excel", "the batch's outcome is unknown: gone", where="inproc")
        return res
    monkeypatch.setattr(runner, "run", run)
    with pytest.raises(BatchError) as e:
        own_book.follow_up(_pre(), 3)
    assert exceljob.KEPT in str(e.value) and exceljob.SPEC_UNKNOWN.format(q=3) in str(e.value)


def test_a_follow_up_in_the_active_workbook_of_a_hidden_excel_ends_on_the_new_answer(own_book):
    """Task 8B review 1: quiet lands on the new Answer whenever the target is the active workbook, visible or not
    (the Book's active sheet, renamed by the park, is not trusted)."""
    wb = own_book.wb
    wb.Application.Visible = False
    wb.Worksheets("Notes").Activate()
    own_book.follow_up(_pre(), 3)
    assert wb.Application.ActiveWorkbook is wb and wb.active is wb.Worksheets("Q3 Answer") and wb.active in wb.order


def test_a_follow_up_is_four_batches_and_reads_the_sheets_once(own_book, monkeypatch):
    """Look, bring in (the park, the insert, the moves, the finish), close out (the drop among the rest), save: one
    Excel run each, one jobs.com each through the COM executor and one for the workbook's name."""
    reads, real = [], runner.VERBS["each"]
    monkeypatch.setitem(runner.VERBS, "each", lambda c, paths: reads.append(list(paths)) or real(c, paths))
    calls = fakexl.count_com(monkeypatch)
    own_book.pre = _pre()
    out, timings = exceljob.with_timings(lambda h: exceljob.redraw_job(h, own_book.pre, 3, PATH))(H)
    assert _batches(timings) == BATCHES and len(calls) == 5 and out["checks"]["sheet_equals_model"] == "OK"
    assert reads == [movein.BOOK_PATHS, ["Name"]]


def test_a_failed_follow_up_rolls_back_before_it_restores_and_never_after(own_book, monkeypatch):
    seen = []
    real_rollback, real_restore = exceljob._rollback, movein.restore
    monkeypatch.setattr(exceljob, "_rollback", lambda t: seen.append(("rollback", list(t.sheets_added)))
                        or real_rollback(t))
    monkeypatch.setattr(movein, "restore", lambda t, parked: seen.append(("restore", dict(parked.sheets)))
                        or real_restore(t, parked))
    monkeypatch.setattr(fakexl._ListObjects, "Add", lambda self, *a: (_ for _ in ()).throw(
        fakexl.ComError("refused")))              # in the batch, after the park and the insert
    with pytest.raises(BatchError):
        own_book.follow_up(_pre(), 3)
    assert seen == [("rollback", ["Q3 Answer", "Q3 Data"]),
                    ("restore", {"Q3 Answer": "old Q3 Answer", "Q3 Data": "old Q3 Data",
                                 "Q3 Lines b0": "old Q3 Lines b0"})]


def test_a_restore_that_fails_is_a_restore_error(own_book, monkeypatch):
    def refuse(self, *a):
        raise fakexl.ComError("a table was refused")

    def restore(t, parked):
        raise HostError("could not put back the names of sheet 'old Q3 Data' -> 'Q3 Data'")
    monkeypatch.setattr(fakexl._ListObjects, "Add", refuse)
    monkeypatch.setattr(movein, "restore", restore)
    with pytest.raises(exceljob.RestoreError) as e:
        own_book.follow_up(_pre(), 3)
    assert "a table was refused" in str(e.value) and "the restore was incomplete" in str(e.value)
    assert "'old Q3 Data' -> 'Q3 Data'" in str(e.value) and isinstance(e.value.__cause__, BatchError)
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


def test_a_defined_name_a_follow_ups_new_table_wants_is_refused_before_the_park(own_book, monkeypatch):
    """Re-review M5: the names check is the bring-in batch's first operation, before the park: the old answer keeps
    its names, nothing is rolled back or restored, and the error is the NameTaken of before."""
    wb = own_book.wb
    wb.Names.Add("Lines_Q2_b0", "=Notes!$A$1", False)
    before = fakexl.snapshot(wb)
    seen = []
    monkeypatch.setattr(movein, "restore", lambda t, parked: seen.append("restore"))
    with pytest.raises(movein.NameTaken, match="the name 'Lines_Q2_b0' already exists") as e:
        own_book.follow_up(_lines_pre(), 2)
    assert not isinstance(e.value, exceljob.RestoreError) and seen == []
    assert fakexl.snapshot(wb) == before and own_book.targets[-1].sheets_added == []


# ------------------------------------------------------------------ which sheets belong to the question
def test_each_new_sheet_takes_its_namesakes_place_and_an_added_one_follows_the_sheet_before_it():
    parked = {"Q3 Answer": "old Q3 Answer", "Q3 Data": "old Q3 Data", "Q3 Lines b1": "old Q3 Lines b1",
              "Q3 Lines b2": "old Q3 Lines b2"}
    names = ["Notes", "Q3 Answer", "Q3 Data", "Q3 Lines b0", "Q3 Lines b1", "old Q3 Answer", "Budget",
             "old Q3 Lines b1", "Q2 Data", "old Q3 Data", "old Q3 Lines b2", "Sources"]
    new = ["Q3 Answer", "Q3 Data", "Q3 Lines b0", "Q3 Lines b1"]
    assert exceljob.wanted_order(names, new, parked) == [
        "Notes", "Q3 Answer", "Budget", "Q3 Lines b1", "Q2 Data", "Q3 Data", "Q3 Lines b0", "Sources"]
    side_by_side = ["Q3 Answer", "Q3 Data", "old Q3 Answer", "old Q3 Data", "Sources"]
    assert exceljob.wanted_order(side_by_side, new[:2], parked) == ["Q3 Answer", "Q3 Data", "Sources"]



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

    async def prepare(raw, provider, spec=None, complete=False):     # server._prepare's signature: no target
        return SimpleNamespace(raw=raw), []
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


def test_fin_refresh_carries_each_questions_warnings_and_layout_findings(monkeypatch):
    """Final review I/F1: what a question's redraw could only say as a warning (a formula the re-pointing could not
    rewrite, the scan for such formulas failing, AutoSave left off) or a layout finding reaches the refresh's reply;
    a question without any keeps today's small entry."""
    from smartass.xl.movein import Mentions
    said = [*exceljob._mentions_notes(1, Mentions(failed="operation 83 (verb mentions): busy")),
            "could not point 1 formula(s) at the new Q1 (the cell is protected): Summary!A2 show #REF!",
            exceljob.AUTOSAVE_NOTE.format("book.xlsx")]
    found = ["Q1 Answer: a chart title is cut - shorten it (patch blocks[2].title)"]

    def job(pre, q, target):
        if q == 1:
            return {"q": q, "workbook": target, "closed": {}, "warnings": said, "layout": found, "key_values": {"x": 1},
                    "sheets": ["Q1 Answer", "Q1 Data"], "text_fixes": ["Kostnad"]}
        return {"q": q, "workbook": target, "closed": {}, "warnings": [], "layout": [], "key_values": {"y": 2}}
    _wire_server(monkeypatch, _stored(1, 2), job)
    got = {d["q"]: d for d in asyncio.run(server.fin_refresh(target=PATH))["refreshed"]}
    assert got[1] == {"q": 1, "ok": True, "closed": {}, "warnings": said, "layout": found, "text_fixes": ["Kostnad"]}
    assert "were not checked" in got[1]["warnings"][0] and "may show #REF!" in got[1]["warnings"][0]
    assert got[2] == {"q": 2, "ok": True, "closed": {}}


def test_fin_refresh_goes_on_after_a_failure_and_stops_after_a_failed_restore(monkeypatch):
    seen = []

    def job(pre, q, target):
        seen.append(q)
        if q == 1:
            raise HostError("Excel is busy")
        if q == 3:
            raise exceljob.RestoreError("refused; the restore was incomplete - sheet 'old Q3 Data' -> 'Q3 Data'")
        return {"q": q, "workbook": target, "closed": {}}
    _wire_server(monkeypatch, _stored(1, 2, 3, 4), job)
    out = asyncio.run(server.fin_refresh(target=PATH))
    got = {d["q"]: d for d in out["refreshed"]}
    assert seen == [1, 2, 3] and [got[q]["ok"] for q in (1, 2, 3, 4)] == [False, True, False, False]
    assert "busy" in got[1]["error"] and "restore was incomplete" in got[3]["error"]
    assert got[4]["skipped"] is True and "Q3" in got[4]["error"] and out["workbook"] == PATH


# ------------------------------------------------------------------ carry-overs in _fit_text_rows
def test_the_key_point_rows_are_operations_of_the_batch(monkeypatch):
    """ws.Range is an Excel call too: it is an operation of the batch (busy retry, abort check, the count), never a
    call of its own."""
    long = "•  Opex 1 234 567 NOK vs PY; " + "consultants, licences and the new data centre " * 3
    plan, block = _point_plan(long)
    ws = fakexl.Workbook(["Q1 Answer"]).order[0]
    calls = fakexl.count_com(monkeypatch)
    ops = fit(ws, plan, block)
    assert len(calls) == 1 and ws.rowh[2] == ws.rowh[3] > 0
    assert [op["args"] for op in ops if op.get("name") == "Range"] == [["2:3"]]


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
        fit(ws, plan, block)
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


# ------------------------------------------------------------------ phase 3b Task 2b: failures inside the batches
def test_a_park_that_fails_half_way_in_the_move_in_batch_leaves_the_old_answer_exactly_as_it_was(own_book,
                                                                                               monkeypatch):
    """The park, the insert and the moves are one batch: a rename Excel refuses stops it before the insert, so the
    new sheets' names (still the old sheets' own) leave t.sheets_added - the rollback must not delete the old answer
    - and the restore puts back what was renamed."""
    wb = own_book.wb
    before, n = fakexl.snapshot(wb), len(wb.Application.events)
    data = wb.Worksheets("Q3 Data")                 # the second rename: 'Q3 Answer' is parked, 'Q3 Data' is not

    class Stubborn(fakexl.Worksheet):
        def __setattr__(self, k, v):
            if k == "Name" and "Name" in self.__dict__:
                raise fakexl.ComError("cannot rename the sheet")
            super().__setattr__(k, v)
    data.__class__ = Stubborn
    with pytest.raises(BatchError, match="rename the sheet"):
        own_book.follow_up(_lines_pre(), 3)
    data.__class__ = fakexl.Worksheet
    assert fakexl.snapshot(wb) == before and own_book.targets[-1].sheets_added == []
    assert not [e for e in wb.Application.events[n:] if e[1] in ("Sheets.Add", "Delete")]   # nothing in, nothing out


def _move_in_batch(b) -> bool:
    return any(op.get("name") == "Add" and len(op.get("args", [])) == 4 for op in b.to_dict()["ops"])


@pytest.mark.parametrize("ran", [True, False])
def test_a_move_in_whose_outcome_is_unknown_is_told_by_the_workbook(own_book, monkeypatch, ran):
    """The add-in took the move-in batch and gave no result (BatchError without an index): the workbook says whether
    the park went through, so the rollback removes the new sheets when they came in and never the old ones."""
    wb = own_book.wb
    before = fakexl.snapshot(wb)
    real = runner.run

    def run(app, b, *a, **k):
        if _move_in_batch(b):
            if ran:
                real(app, b, *a, **k)
            raise BatchError(None, "the batch in Excel", "the batch's outcome is unknown: test", where="inproc")
        return real(app, b, *a, **k)
    monkeypatch.setattr(runner, "run", run)
    with pytest.raises(BatchError, match="unknown"):
        own_book.follow_up(_pre(), 3)
    assert fakexl.snapshot(wb) == before and own_book.targets[-1].sheets_added == []


def test_an_abandoned_follow_up_stops_at_its_next_batch_and_still_cleans_up(own_book, monkeypatch):
    """The job abandoned by a timeout while the bring-in batch ran: the close-out never starts (JobAborted), and the
    clean-up batches run for the abandoned job (check=False): the old answer is back, Excel as it was."""
    from smartass.xl import jobs
    wb, app = own_book.wb, own_book.wb.Application
    before, n = fakexl.snapshot(wb), len(app.events)
    token = jobs.JobToken()
    jobs.bind_job(token)
    real = runner.run

    def run(app_, b, *a, **k):
        out = real(app_, b, *a, **k)
        if _move_in_batch(b):
            token.aborted = True
        return out
    monkeypatch.setattr(runner, "run", run)
    try:
        with pytest.raises(jobs.JobAborted):
            own_book.follow_up(_pre(), 3)
    finally:
        jobs.bind_job()
    assert fakexl.snapshot(wb) == before and not [e for e in app.events[n:] if e[1] == "Delete"
                                                  and e[2][0].startswith("old ")]   # no drop: no close-out
    assert (app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts, app.Calculation) == (True, True, True, -4105)


def test_a_busy_excel_inside_a_batch_is_the_busy_error_and_the_old_answer_stays(own_book, monkeypatch):
    from smartass.xl import jobs
    wb = own_book.wb
    before = fakexl.snapshot(wb)
    real = runner.run

    def run(app, b, *a, **k):
        if any(op.get("name") == "Calculate" for op in b.to_dict()["ops"]):
            raise BatchError(9, "Workbooks(...).Worksheets(...).Calculate()",
                             f"operation 9 (call Calculate): {jobs.BUSY} (0x8001010A)")
        return real(app, b, *a, **k)
    monkeypatch.setattr(runner, "run", run)
    with pytest.raises(HostError) as e:
        own_book.follow_up(_pre(), 3)
    assert str(e.value) == jobs.BUSY and isinstance(e.value.__cause__, BatchError)
    assert fakexl.snapshot(wb) == before


# ------------------------------------------------------------------ formulas that point at the answer (Task 3)
SUMMARY = {(1, 1): "='Q3 Answer'!E12",                       # a cell
           (2, 1): "=SUM(Data_Q3[Value])",                    # a table
           (3, 1): "=SUM('Q3 Data'!B2:B9)",                   # a cross-sheet SUM
           (4, 1): "=\"'Q3 Answer'!E12\"",                    # text in a string literal: never a reference
           (5, 1): "='Q3 Notes'!A1"}                          # the user's sheet named like a topic of Q3
GONE = (6, 1), "='Q3 Lines b0'!A1"                            # a sheet the new answer does not have
SERIES = "=SERIES(,,'Q3 Data'!$B$2:$B$5,1)"


def _point_at_q3(wb):
    """The user's formulas that point at Q3: a 'Summary' sheet of their own (cells, a chart), a cell on Q1's Answer
    (another answer), a defined name. As Excel, the fake rewrites them when Q3's sheets are parked and turns them
    #REF! when those are deleted."""
    s = fakexl.Worksheet(wb, "Summary")
    wb.order.insert(wb.order.index(wb.Worksheets("Notes")) + 1, s)
    for rc, f in [*SUMMARY.items(), GONE]:
        s.cells[rc] = {"Value": f}
    s.charts.add(SERIES)
    wb.Worksheets("Q1 Answer").cells[(40, 20)] = {"Value": "='Q3 Answer'!E12*2"}
    wb.Names.Add("TotalQ3", "='Q3 Answer'!$E$12")
    return s


def _pointing(wb) -> dict:
    s = wb.Worksheets("Summary")
    return {"summary": {rc: s.cells[rc]["Value"] for rc in sorted(s.cells)},
            "series": s.charts[0].Chart.series[0].Formula, "q1": wb.Worksheets("Q1 Answer").cells[(40, 20)]["Value"],
            "name": wb.Names.items["TotalQ3"].RefersTo}


POINTING = {"summary": {**SUMMARY, GONE[0]: GONE[1]}, "series": SERIES, "q1": "='Q3 Answer'!E12*2",
            "name": "='Q3 Answer'!$E$12"}


def test_formulas_elsewhere_that_point_at_the_question_point_at_the_new_answer(own_book):
    """A follow-up parks Q3's sheets (the formulas follow the rename), brings the new ones in and drops the old ones:
    before the drop every reference to a parked sheet or table is named as the new one again. A reference to a sheet
    the new answer has no namesake for shows #REF! (that cell is gone), as Excel leaves it."""
    wb = own_book.wb
    _point_at_q3(wb)
    out = own_book.follow_up(_pre(), 3)
    assert _pointing(wb) == {**POINTING, "summary": {**SUMMARY, GONE[0]: "=#REF!A1"}}
    assert not [w for w in out["warnings"] if "point" in w] and out["checks"]["sheet_equals_model"] == "OK"
    assert wb.links == [] and not [n for n in _names(wb) if n.startswith("old ")]
    assert not [f for f in out["layout"] if "could not run" in f]          # the layout check saw the answer's plan


def test_a_new_answers_layout_check_sees_its_plan(own_book):
    """The close-out's layout check gets the answer's plan also when there is nothing to re-point."""
    out = own_book.answer(_pre())
    assert not [f for f in out["layout"] if "could not run" in f]


def test_six_follow_ups_in_a_row_keep_the_formulas_pointing_at_the_question(own_book):
    """With a lines block in every new answer, a reference to the Lines sheet and its table is re-pointed too."""
    wb = own_book.wb
    _point_at_q3(wb)
    wb.Worksheets("Summary").cells[(7, 1)] = {"Value": "=SUM(Lines_Q3_b0[AMOUNT])"}
    for _ in range(6):
        own_book.follow_up(_lines_pre(), 3)
    assert _pointing(wb) == {**POINTING, "summary": {**POINTING["summary"], (7, 1): "=SUM(Lines_Q3_b0[AMOUNT])"}}


def test_a_drop_refused_after_the_formulas_were_re_pointed_points_them_back(own_book, monkeypatch):
    """The re-pointing comes right before the drop in the close-out batch; when the drop fails, the formulas are
    pointed back at the old sheets before the new ones are removed (they would turn #REF! with them) and the old ones
    get their names back: the workbook exactly as it was."""
    wb = own_book.wb
    _point_at_q3(wb)
    before, real = fakexl.snapshot(wb), fakexl._SheetGroup.Delete

    def refuse(group):
        if group.sheets[0].Name.startswith("old "):
            raise fakexl.ComError("the delete was refused")
        real(group)
    monkeypatch.setattr(fakexl._SheetGroup, "Delete", refuse)
    with pytest.raises(BatchError, match="refused"):
        own_book.follow_up(_pre(), 3)
    assert fakexl.snapshot(wb) == before and _pointing(wb) == POINTING and wb.links == []


def test_a_drop_that_deletes_and_then_fails_keeps_the_formulas_on_the_new_answer(own_book, monkeypatch):
    wb, real = own_book.wb, fakexl._SheetGroup.Delete
    _point_at_q3(wb)

    def delete(group):
        real(group)
        raise fakexl.ComError("the Delete went through and then failed")
    monkeypatch.setattr(fakexl._SheetGroup, "Delete", delete)
    with pytest.raises(BatchError, match="then failed"):
        own_book.follow_up(_pre(), 3)
    assert _pointing(wb) == {**POINTING, "summary": {**SUMMARY, GONE[0]: "=#REF!A1"}}


def _close_out_batch(b) -> bool:
    """The batch that drops the parked sheets (Worksheets([old names]).Delete()), not the rollback's."""
    return any(op.get("name") == "Worksheets" and isinstance(op.get("args", [None])[0], dict)
               and str(op["args"][0].get("list", [""])[0]).startswith("old ") for op in b.to_dict()["ops"])


@pytest.mark.parametrize("ran", [True, False])
def test_a_close_out_whose_outcome_is_unknown_is_told_by_the_workbook(own_book, monkeypatch, ran):
    """The add-in took the close-out and gave no result: the workbook says whether the drop went through - then the
    new answer stays with the formulas on it, else they are pointed back and the old answer is restored."""
    wb = own_book.wb
    _point_at_q3(wb)
    before, real = fakexl.snapshot(wb), runner.run

    def run(app, b, *a, **k):
        if _close_out_batch(b):
            if ran:
                real(app, b, *a, **k)
            raise BatchError(None, "the batch in Excel", "the batch's outcome is unknown: test", where="inproc")
        return real(app, b, *a, **k)
    monkeypatch.setattr(runner, "run", run)
    with pytest.raises(BatchError, match="unknown"):
        own_book.follow_up(_pre(), 3)
    if ran:
        assert _pointing(wb) == {**POINTING, "summary": {**SUMMARY, GONE[0]: "=#REF!A1"}}
        assert "Q3 Answer" in _names(wb) and not [n for n in _names(wb) if n.startswith("old ")]
    else:
        assert fakexl.snapshot(wb) == before and _pointing(wb) == POINTING


def test_a_job_abandoned_between_the_re_pointing_and_the_drop_points_the_formulas_back(own_book, monkeypatch):
    """Over COM an abandoned job stops at its next operation: here after the first formula was re-pointed, before the
    drop. The clean-up (for the abandoned job too) points it back before the new sheets go."""
    from smartass.xl import jobs
    wb = own_book.wb
    _point_at_q3(wb)
    before = fakexl.snapshot(wb)
    token = jobs.JobToken()
    jobs.bind_job(token)
    real = fakexl.Range.__setattr__

    def setattr_(self, k, v):
        real(self, k, v)
        if self.ws.Name == "Summary" and k == "Formula2":
            token.aborted = True
    monkeypatch.setattr(fakexl.Range, "__setattr__", setattr_)
    try:
        with pytest.raises(jobs.JobAborted):
            own_book.follow_up(_pre(), 3)
    finally:
        jobs.bind_job()
    assert fakexl.snapshot(wb) == before and _pointing(wb) == POINTING


def test_a_close_out_that_fails_before_the_re_pointing_needs_no_pointing_back(own_book, monkeypatch):
    wb = own_book.wb
    _point_at_q3(wb)
    before, back = fakexl.snapshot(wb), []
    monkeypatch.setattr(movein, "unpoint", lambda t, plan: back.append(plan))
    real = exceljob._fit_text_rows

    def fit_(ws, plan, block, b):
        b.app.Workbooks("No such book.xlsx").Activate()             # fails before the re-pointing and the drop
        return real(ws, plan, block, b)
    monkeypatch.setattr(exceljob, "_fit_text_rows", fit_)
    with pytest.raises(BatchError):
        own_book.follow_up(_pre(), 3)
    assert back == [] and fakexl.snapshot(wb) == before and _pointing(wb) == POINTING


def test_a_formula_excel_refuses_to_re_point_is_said_and_the_rest_goes_on(own_book, monkeypatch):
    """A locked cell on a protected sheet, or a column the new table lacks: Excel refuses the write; the drop goes on,
    that formula shows #REF! and the reply says where."""
    wb = own_book.wb
    _point_at_q3(wb)
    real = fakexl.Range.__setattr__

    def setattr_(self, k, v):
        if self.ws.Name == "Summary" and k == "Formula2" and self.rects[0][:2] == (2, 1):
            raise fakexl.ComError("the cell is protected")
        real(self, k, v)
    monkeypatch.setattr(fakexl.Range, "__setattr__", setattr_)
    out = own_book.follow_up(_pre(), 3)
    got = _pointing(wb)
    assert got["summary"][(2, 1)] == "=SUM(#REF!)" and got["summary"][(1, 1)] == SUMMARY[(1, 1)]
    note = [w for w in out["warnings"] if "Summary!A2" in w]
    assert len(note) == 1 and "protected" in note[0] and "#REF!" in note[0] and "Q3" in note[0]


def test_a_look_for_the_formulas_that_fails_is_said_and_the_answer_is_drawn(own_book, monkeypatch):
    """An answer never fails because another sheet could not be scanned: the old sheets are dropped anyway, and the
    warning says plainly that the formulas pointing at the old answer were not checked (final review F3)."""
    wb = own_book.wb
    _point_at_q3(wb)

    def refuse(*a):
        raise fakexl.ComError("the workbook could not be searched")
    monkeypatch.setitem(runner.VERBS, "mentions", refuse)
    out = own_book.follow_up(_pre(), 3)
    assert _pointing(wb)["summary"][(1, 1)] == "=#REF!E12" and out["checks"]["sheet_equals_model"] == "OK"
    note = [w for w in out["warnings"] if "could not be searched" in w]
    assert len(note) == 1 and "point at the old Q3 were not checked" in note[0] and "dropped anyway" in note[0]
    assert "may show #REF!" in note[0]


def test_a_sheet_whose_formulas_cannot_be_read_is_said(own_book, monkeypatch):
    wb = own_book.wb
    _point_at_q3(wb)
    real = fakexl.Range.__getattr__

    def get(self, k):
        if k == "Formula2" and self.ws.Name == "Summary" and len(self.cells()) > 1:
            raise fakexl.ComError("the formulas are hidden")
        return real(self, k)
    monkeypatch.setattr(fakexl.Range, "__getattr__", get)
    out = own_book.follow_up(_pre(), 3)
    assert [w for w in out["warnings"] if "'Summary'" in w and "hidden" in w and "#REF!" in w]
    assert _pointing(wb)["q1"] == POINTING["q1"]                  # the other sheets are re-pointed


# ------------------------------------------------------------------ review fixes (Task 3)
def test_a_protected_sheet_whose_only_mention_is_a_hidden_formula_is_said(own_book):
    """Review Important 2: a hidden formula on a protected sheet cannot be read; the reply says so, never silence."""
    wb = own_book.wb
    lk = fakexl.Worksheet(wb, "Locked")
    wb.order.append(lk)
    lk.cells[(2, 2)] = {"Value": "='Q3 Answer'!E12"}
    lk.ProtectContents, lk.hidden = True, {(2, 2)}
    out = own_book.follow_up(_pre(), 3)
    assert [w for w in out["warnings"] if "'Locked'" in w and "hidden formula" in w and "Q3" in w]


def _locked(wb, formula: str):
    lk = fakexl.Worksheet(wb, "Locked")
    wb.order.append(lk)
    lk.cells[(2, 2)] = {"Value": formula}
    lk.cells[(3, 2)] = {"Value": "=#REF!A1"}                    # a #REF! the user had before: not ours
    lk.ProtectContents, lk.hidden = True, {(2, 2)}
    return lk


def test_a_protected_sheet_whose_hidden_formula_does_not_point_at_the_answer_is_not_warned_about(own_book):
    """Phase 3b review: a protected sheet with a hidden formula warned on every follow-up. Its values can be read:
    the warning comes only when a cell there shows #REF! after the drop that did not before."""
    _locked(own_book.wb, "=Notes!A1")
    out = own_book.follow_up(_pre(), 3)
    assert not [w for w in out["warnings"] if "'Locked'" in w]


def test_a_protected_sheet_whose_hidden_formula_pointed_at_the_answer_says_how_many_cells_show_ref(own_book):
    _locked(own_book.wb, "='Q3 Answer'!E12")
    out = own_book.follow_up(_pre(), 3)
    (note,) = [w for w in out["warnings"] if "'Locked'" in w]
    assert note == exceljob.UNREAD_REF.format(sheet="Locked", why="a hidden formula of a protected sheet cannot be read",
                                              n=1, q=3, cells="B2")


def test_a_new_ref_on_a_protected_sheet_is_said_also_when_the_follow_up_cleared_another_there(own_book, monkeypatch):
    """Task 3 review Minor 2: compared by where, not by how many - B3's #REF! gone and B2's new one is one new
    #REF!, though the count stays 1."""
    lk = _locked(own_book.wb, "='Q3 Answer'!E12")
    real = fakexl._SheetGroup.Delete

    def delete(group):
        real(group)
        lk.cells[(3, 2)] = {"Value": 5}                         # cleared in the same follow-up
    monkeypatch.setattr(fakexl._SheetGroup, "Delete", delete)
    out = own_book.follow_up(_pre(), 3)
    assert [w for w in out["warnings"] if "'Locked'" in w and "1 cell(s)" in w and "(B2)" in w]


def test_the_ref_is_told_by_its_error_code_not_its_text():
    """Task 3 review Minor 3: Value2 gives an error as its code in both executors (a number as a float), so the
    #REF! test does not depend on Excel's language."""
    wb = fakexl.Workbook(["S"])
    ws = wb.Worksheets("S")
    ws.cells[(1, 1)] = {"Value": float(exceljob.REF_ERROR)}            # the number, not the error
    ws.cells[(2, 1)] = {"Value": "=#REF!A1"}
    ws.cells[(3, 2)] = {"Value": "#REF!"}                                # the text, not the error
    b = Batch(quiet=False)
    b.ret(b.verb("errors", b.workbook(wb.Name).Worksheets("S"), exceljob.REF_ERROR), "refs")
    assert runner.run(wb.Application, b).values["refs"] == ("A2",)


def test_a_protected_sheet_whose_refs_cannot_be_read_is_warned_about_as_before(own_book, monkeypatch):
    _locked(own_book.wb, "=Notes!A1")

    def refuse(ws, code):
        raise fakexl.ComError("the values could not be read")
    monkeypatch.setitem(runner.VERBS, "errors", refuse)
    out = own_book.follow_up(_pre(), 3)
    assert [w for w in out["warnings"] if "'Locked'" in w and "if one there pointed at Q3" in w]


def test_the_verb_errors_runs_only_for_a_sheet_whose_formulas_could_not_be_read(own_book, monkeypatch):
    """Final review M4: a follow-up in a workbook whose sheets were all scanned makes no errors call at all; one with
    an unread sheet makes two for it (before and after the drop)."""
    calls, real = [], runner.VERBS["errors"]
    monkeypatch.setitem(runner.VERBS, "errors", lambda ws, code: calls.append(str(ws.Name)) or real(ws, code))
    _point_at_q3(own_book.wb)
    own_book.follow_up(_pre(), 3)
    assert calls == []
    _locked(own_book.wb, "=Notes!A1")
    own_book.follow_up(_pre(), 3)
    assert calls == ["Locked", "Locked"]


def test_in_the_users_manual_calculation_the_protected_sheet_is_warned_about_as_before(own_book):
    """Counting would need the sheet calculated, which the user keeps for later: today's warning, nothing read."""
    _locked(own_book.wb, "=Notes!A1")
    own_book.wb.Application.Calculation = -4135
    out = own_book.follow_up(_pre(), 3)
    assert [w for w in out["warnings"] if "'Locked'" in w and "if one there pointed at Q3" in w]


def test_a_formula_the_user_changed_while_the_answer_was_drawn_is_left_alone_and_said(own_book, monkeypatch):
    """Review ruling 3: each write checks the formula is still the text the bring-in read. Between the two batches
    the user edits a cell and adds a name that shifts the names' order: neither is written, the rest is."""
    wb = own_book.wb
    _point_at_q3(wb)
    real = exceljob._paint

    def paint(look):
        wb.Worksheets("Summary").cells[(1, 1)]["Value"] = "='old Q3 Answer'!E12+1"
        wb.Names.items = {"AAA": fakexl._Name(Name="AAA", RefersTo="=1", Visible=True), **wb.Names.items}
        real(look)
    monkeypatch.setattr(exceljob, "_paint", paint)
    out = own_book.follow_up(_pre(), 3)
    got = _pointing(wb)
    assert got["summary"][(1, 1)] == "=#REF!E12+1" and got["name"] == "=#REF!$E$12"     # not re-pointed
    assert got["summary"][(2, 1)] == SUMMARY[(2, 1)] and got["q1"] == POINTING["q1"]   # the rest is
    assert wb.Names.items["AAA"].RefersTo == "=1"
    note = [w for w in out["warnings"] if "changed while the answer was drawn" in w]
    assert len(note) == 1 and "Summary!A1" in note[0] and "the name TotalQ3" in note[0] and "2 formula(s)" in note[0]


def test_the_close_out_calculates_manually_only_when_it_re_points(own_book, monkeypatch):
    """Review ruling 2: with re-pointing writes the close-out batch sets Calculation to manual (each write would
    recalculate); without any it leaves Calculation alone, as before."""
    wb, seen, real = own_book.wb, [], runner.run

    def run(app, b, *a, **k):
        if _close_out_batch(b):
            seen.append(b.to_dict().get("calculation", True))
        return real(app, b, *a, **k)
    monkeypatch.setattr(runner, "run", run)
    own_book.follow_up(_pre(), 3)
    _point_at_q3(wb)
    own_book.follow_up(_pre(), 3)
    assert seen == [False, True]


def test_a_point_back_excel_refuses_is_said_in_the_error(own_book, monkeypatch):
    """Review Minor 5: the drop is refused after the re-pointing, and Excel refuses to point one formula back: the
    error names it (it shows #REF! once the new sheets are gone)."""
    wb = own_book.wb
    _point_at_q3(wb)
    real_delete, real_set = fakexl._SheetGroup.Delete, fakexl.Range.__setattr__

    def refuse(group):
        if group.sheets[0].Name.startswith("old "):
            raise fakexl.ComError("the delete was refused")
        real_delete(group)

    def setattr_(self, k, v):
        if self.ws.Name == "Summary" and k == "Formula2" and v == "=SUM(Data_Q3_old[Value])":
            raise fakexl.ComError("the cell is locked now")
        real_set(self, k, v)
    monkeypatch.setattr(fakexl._SheetGroup, "Delete", refuse)
    monkeypatch.setattr(fakexl.Range, "__setattr__", setattr_)
    with pytest.raises(BatchError) as e:
        own_book.follow_up(_pre(), 3)
    text = str(e.value)
    assert "could not point 1 formula(s) back at the old Q3" in text and "the cell is locked now" in text
    assert text.endswith("Summary!A2 show #REF!")
    assert _pointing(wb)["summary"][(2, 1)] == "=SUM(#REF!)" and _pointing(wb)["summary"][(1, 1)] == SUMMARY[(1, 1)]


def test_a_point_back_batch_that_fails_whole_names_every_formula(own_book, monkeypatch):
    from smartass.xl.repoint import Repoint
    plan = [Repoint("cell", "Summary", "A1", "='old Q3 Answer'!E12", "='Q3 Answer'!E12"),
            Repoint("name", "TotalQ3", "2", "='old Q3 Answer'!$E$12", "='Q3 Answer'!$E$12")]

    def fail(*a, **k):
        raise HostError("Excel is gone")
    monkeypatch.setattr(movein, "run", fail)
    t = Target(app=own_book.wb.Application, wb=own_book.wb, pid=0, hwnd=0)
    assert [(r.where_text, why) for r, why in movein.unpoint(t, plan)] == [("Summary!A1", "Excel is gone"),
                                                                       ("the name TotalQ3", "Excel is gone")]


def test_a_legacy_array_and_a_chart_sheet_are_re_pointed_too(own_book):
    """Review Minor 8: the two kinds the other tests do not write, on the fake."""
    wb = own_book.wb
    s = fakexl.Worksheet(wb, "Summary")
    wb.order.append(s)
    s.Range("A1:A2").FormulaArray = "='Q3 Data'!B2:B3*2"
    wb.Charts.append(fakexl.ChartSheet("Q3 Chart", "=SERIES(,,'Q3 Data'!$B$2:$B$5,1)"))
    own_book.follow_up(_pre(), 3)
    assert s.Range("A1:A2").FormulaArray == "='Q3 Data'!B2:B3*2" and s.Range("A2").HasArray
    assert wb.Charts("Q3 Chart").series[0].Formula == "=SERIES(,,'Q3 Data'!$B$2:$B$5,1)"
