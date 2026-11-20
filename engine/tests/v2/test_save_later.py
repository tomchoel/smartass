# engine/tests/v2/test_save_later.py
"""Phase 3c Task 2: the save happens after the reply. With the add-in in the user's Excel, the job's save step asks
Excel to save our own file once it has painted the answer (xl/inproc/later.py): the reply does not wait for the
save, AutoSave comes back on after it, the next job finds it done (or runs it first when it came too soon, once), a
failed one is a note of the next reply and that job saves in the foreground. The user's workbook is never saved; the
COM executor and SV_SAVE=sync save in the job as before. On the fake Excel with the fake add-in (fakeaddin.py); every
figure is invented."""
from __future__ import annotations

from pathlib import Path

import pytest
from fakeaddin import install
from test_file_job import _job, _pre

import test_file_job
import test_first_answer
from smartass import pipeline
from smartass.core import config
from smartass.core.timing import BATCH_PREFIX
from smartass.xl.inproc import later

tmp_home, calculated, fake_target = test_file_job.tmp_home, test_file_job.calculated, test_file_job.fake_target
new_book = test_first_answer.new_book


@pytest.fixture
def addin(monkeypatch, fake_target, tmp_path):
    """Our own analysis file with AutoSave on (a OneDrive file), in an Excel with the add-in; every Save and AutoSave
    switch logged in order."""
    wb = fake_target.wb
    steps = wb.steps = []

    def save():
        wb.saved += 1
        steps.append("save")
    wb.Save = save
    wb.__dict__["_autosave"] = True
    type(wb).AutoSaveOn = property(lambda self: self.__dict__.get("_autosave"),
                                   lambda self, v: (self.__dict__.__setitem__("_autosave", v),
                                                    getattr(self, "steps", []).append(f"autosave {v}")))
    monkeypatch.setattr(later, "_queued", {})
    monkeypatch.setattr(later, "_done", {})
    yield install(monkeypatch, fake_target.app, tmp_path)
    del type(wb).AutoSaveOn


def test_the_reply_comes_before_the_save_and_autosave_is_back_on_after_it(fake_target, calculated, addin):
    out, timings = _job(_pre(), calculated)
    wb = fake_target.wb
    assert out["checks"]["sheet_equals_model"] == "OK"
    assert wb.saved == 0 and wb.AutoSaveOn is False and wb.steps == ["autosave False"]
    assert len(addin.asked) == 1 and addin.asked[0][:2] == (wb.Name, True)
    addin.release()
    assert wb.saved == 1 and wb.steps == ["autosave False", "save", "autosave True"]


def test_the_next_job_records_the_saves_own_seconds_and_its_save_step_is_one_run(fake_target, calculated, addin):
    """(The first answer's save step also names the new Sources log's print setup: a batch before the request.)"""
    _job(_pre(), calculated)
    addin.release()
    out, timings = _job(_pre(), calculated)
    assert timings["previous save"] == 0.25 and not [w for w in out["warnings"] if "save" in w]
    assert timings[BATCH_PREFIX + "save"] == 1 and addin.log.count("batch") == 3 + 3 + 1


def test_a_job_that_comes_before_the_save_ran_has_it_run_first_and_once(fake_target, calculated, addin):
    """Its look's Run runs the waiting save before anything else: the first answer is saved and AutoSave back on
    before the second job switches it off; the timer finds nothing left; no save is lost or doubled."""
    _job(_pre(), calculated)
    _job(_pre(), calculated)
    wb = fake_target.wb
    assert wb.steps == ["autosave False", "save", "autosave True", "autosave False"]
    assert addin.release()["how"] == "after the reply"           # the second job's own save
    assert wb.saved == 2 and wb.AutoSaveOn is True
    assert wb.steps[-3:] == ["autosave False", "save", "autosave True"]


def test_a_failed_save_is_told_in_the_next_reply_and_that_job_saves_in_the_foreground(fake_target, calculated, addin):
    _job(_pre(), calculated)
    addin.save_error = "the file is locked by another user"
    addin.release()
    wb = fake_target.wb
    assert wb.saved == 0 and wb.AutoSaveOn is True                 # back on after the failed save too
    addin.save_error = None
    out, timings = _job(_pre(), calculated)
    note = [w for w in out["warnings"] if "after the last answer failed" in w]
    assert note == [f"the save of {wb.Name} after the last answer failed (the file is locked by another user); "
                    "it is saved now"]
    assert len(addin.asked) == 1 and wb.saved == 1                 # saved in the job, nothing asked for later
    assert wb.steps[-3:] == ["autosave False", "save", "autosave True"]


def test_a_failed_save_of_another_workbook_of_ours_is_saved_by_the_next_job_too(fake_target, calculated, addin):
    """The next job saves in the foreground: its own file and that workbook, in its save batch."""
    import fakexl
    other = fakexl.Workbook(["Q1 Answer"], app=fake_target.app, styles_ver=False)
    other.Name = "older answer.xlsx"
    fake_target.app.ActiveWorkbook = fake_target.wb
    addin.pending = {"workbook": other.Name, "autosave": False, "token": "0" * 32, "queued": 0.0}
    addin.save_error = "the network went away"
    addin.release()
    addin.save_error = None
    out, _ = _job(_pre(), calculated)
    assert other.saved == 1 and fake_target.wb.saved == 1 and addin.asked == []     # this job saves in the job
    assert [w for w in out["warnings"] if "older answer" in w] == [
        "the save of older answer.xlsx after the last answer failed (the network went away); it is saved now"]


def test_a_workbook_closed_before_its_save_is_a_note_only(fake_target, calculated, addin):
    addin.pending = {"workbook": "closed answer.xlsx", "autosave": True, "token": "1" * 32, "queued": 0.0}
    addin.release()
    out, _ = _job(_pre(), calculated)
    assert [w for w in out["warnings"] if "closed answer" in w] == [
        "closed answer.xlsx was closed before smartass saved it after the last answer (Excel asks whether to save "
        "it when it closes)"]
    assert len(addin.asked) == 1                                    # this job's own save is asked for later as usual


def test_sv_save_sync_saves_in_the_job_with_the_add_in_too(fake_target, calculated, addin, monkeypatch):
    monkeypatch.setattr(config, "SAVE", "sync")
    _job(_pre(), calculated)
    wb = fake_target.wb
    assert addin.asked == [] and wb.saved == 1 and wb.steps == ["autosave False", "save", "autosave True"]


def test_a_refused_request_saves_in_the_job(fake_target, calculated, addin):
    addin.answer = "not run: refused: no workbook"
    _job(_pre(), calculated)
    assert len(addin.asked) == 1 and fake_target.wb.saved == 1 and fake_target.wb.AutoSaveOn is True


def test_the_users_workbook_is_never_asked_to_be_saved(fake_target, calculated, addin, monkeypatch):
    monkeypatch.setattr(pipeline, "_is_ours", lambda t: False)
    _job(_pre(), calculated)
    _job(_pre(), calculated)
    assert addin.asked == [] and fake_target.wb.saved == 0


def test_a_follow_up_asks_for_its_save_after_the_reply_too(fake_target, calculated, addin):
    _job(_pre(), calculated)
    addin.release()
    calculated["pre"] = pre = _pre()
    out, _ = pipeline.with_timings(lambda h: pipeline.redraw_job(h, pre, 1))(None)
    assert out["followup"] is True and len(addin.asked) == 2 and fake_target.wb.saved == 1
    addin.release()
    assert fake_target.wb.saved == 2 and fake_target.wb.AutoSaveOn is True


def test_the_first_answer_of_a_new_workbook_asks_for_its_save_after_the_reply_too(new_book, calculated, monkeypatch,
                                                                                    tmp_path):
    """Task 1's built workbook, opened by Excel, ends in the same save step."""
    nb = new_book
    monkeypatch.setattr(later, "_queued", {})
    monkeypatch.setattr(later, "_done", {})
    addin = install(monkeypatch, nb.app, tmp_path)
    out, _ = _job(_pre(), calculated)
    wb = nb.app.Workbooks(Path(out["workbook"]).name)
    assert len(addin.asked) == 1 and wb.saved == 0 and nb.mine.saved == 0
    addin.release()
    assert wb.saved == 1 and nb.mine.saved == 0
