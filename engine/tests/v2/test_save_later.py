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

import fakexl
import test_file_job
import test_first_answer
from smartass import exceljob
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
    other.Name, other.FullName = "older answer.xlsx", r"C:\analyser\older answer.xlsx"
    fake_target.app.ActiveWorkbook = fake_target.wb
    addin.pending = {"workbook": other.Name, "full_name": other.FullName, "autosave": False, "token": "0" * 32,
                     "queued": 0.0}
    addin.save_error = "the network went away"
    addin.release()
    addin.save_error = None
    out, _ = _job(_pre(), calculated)
    assert other.saved == 1 and fake_target.wb.saved == 1 and addin.asked == []     # this job saves in the job
    assert [w for w in out["warnings"] if "older answer" in w] == [
        "the save of older answer.xlsx after the last answer failed (the network went away); it is saved now"]


def test_a_failed_save_is_kept_for_the_job_after_next_when_the_next_one_fails_first(fake_target, calculated, addin,
                                                                                     monkeypatch):
    """The next job took the failed outcome, then failed before its own save (its error says the earlier save
    failed): the outcome is put back, so the job after it says so again and saves in the foreground."""
    _job(_pre(), calculated)
    addin.save_error = "the file is locked by another user"
    addin.release()
    addin.save_error = None
    real = exceljob._close_out

    def broken(*a, **k):
        raise RuntimeError("the close-out broke")
    monkeypatch.setattr(exceljob, "_close_out", broken)
    with pytest.raises(RuntimeError, match="the close-out broke.*after the last answer failed"):
        _job(_pre(), calculated)
    assert fake_target.wb.saved == 0
    monkeypatch.setattr(exceljob, "_close_out", real)
    out, _ = _job(_pre(), calculated)
    assert [w for w in out["warnings"] if "after the last answer failed" in w] == [
        f"the save of {fake_target.wb.Name} after the last answer failed (the file is locked by another user); it is "
        "saved now"]
    assert fake_target.wb.saved == 1 and len(addin.asked) == 1


def test_another_file_open_by_that_name_now_is_not_saved_by_the_next_job(fake_target, calculated, addin):
    import fakexl
    other = fakexl.Workbook(["Q1 Answer"], app=fake_target.app, styles_ver=False)
    other.Name, other.FullName = "older answer.xlsx", r"C:\analyser\older answer.xlsx"
    fake_target.app.ActiveWorkbook = fake_target.wb
    addin.pending = {"workbook": other.Name, "full_name": other.FullName, "autosave": False, "token": "2" * 32,
                     "queued": 0.0}
    addin.save_error = "the network went away"
    addin.release()
    addin.save_error = None
    other.FullName = r"C:\Users\me\older answer.xlsx"              # the user's own file of that name now
    out, _ = _job(_pre(), calculated)
    assert other.saved == 0 and fake_target.wb.saved == 1
    assert [w for w in out["warnings"] if "older answer" in w] == [
        "the save of older answer.xlsx after the last answer failed (the network went away); the workbook open by "
        "that name now is another file, so it was not saved"]


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
    monkeypatch.setattr(exceljob, "_is_ours", lambda t: False)
    _job(_pre(), calculated)
    _job(_pre(), calculated)
    assert addin.asked == [] and fake_target.wb.saved == 0


def test_a_follow_up_asks_for_its_save_after_the_reply_too(fake_target, calculated, addin):
    _job(_pre(), calculated)
    addin.release()
    calculated["pre"] = pre = _pre()
    out, _ = exceljob.with_timings(lambda h: exceljob.redraw_job(h, pre, 1))(None)
    assert out["followup"] is True and len(addin.asked) == 2 and fake_target.wb.saved == 1
    addin.release()
    assert fake_target.wb.saved == 2 and fake_target.wb.AutoSaveOn is True


def test_the_first_answer_of_a_new_workbook_is_saved_in_the_job_and_the_next_one_after_the_reply(
        new_book, calculated, monkeypatch, tmp_path):
    """Final review I1 (the controller's ruling): a workbook the job created is saved inside the job, with the add-in
    loaded - its batches end in the save, nothing is asked for later, no outcome is written - so the file on disk
    holds the answer (formulas, tables, spec, Sources rows) before the session points at it; the next answer into that
    workbook asks for its save after the reply as any other."""
    from test_first_answer import _runs, _writes

    from smartass.xl import target as xltarget
    nb = new_book
    monkeypatch.setattr(later, "_queued", {})
    monkeypatch.setattr(later, "_done", {})
    addin = install(monkeypatch, nb.app, tmp_path)
    runs = _runs(monkeypatch)
    out, _ = _job(_pre(), calculated)
    wb = nb.app.Workbooks(Path(out["workbook"]).name)
    assert addin.asked == [] and wb.saved == 1 and nb.mine.saved == 0
    assert addin.log.count("batch") == len(runs) and any(x.endswith(".Save()") for x in _writes(runs[-1]))
    assert not list((config.INPROC_DIR / "saved").glob("*.json")) and nb.sessions == [out["workbook"]]
    t = xltarget.Target(app=nb.app, wb=wb, pid=7, hwnd=70, created=True)      # the session's workbook from now on
    monkeypatch.setattr(exceljob, "_target", lambda h, target, title, build_new=False: t)
    _job(_pre(), calculated)
    assert len(addin.asked) == 1 and wb.saved == 1
    addin.release()
    assert wb.saved == 2 and nb.mine.saved == 0


def test_a_new_workbooks_first_answer_whose_save_fails_leaves_no_file_and_no_session_workbook(
        new_book, calculated, monkeypatch, tmp_path):
    """The foreground save of a new workbook fails with the add-in loaded: today's clean-up - the workbook closed
    unsaved, its file deleted, the session not pointed at it."""
    monkeypatch.setattr(later, "_queued", {})
    monkeypatch.setattr(later, "_done", {})
    install(monkeypatch, new_book.app, tmp_path)

    def save(self):
        raise fakexl.ComError("the save was cancelled")
    monkeypatch.setattr(fakexl.Workbook, "Save", save)
    with pytest.raises(Exception, match="cancelled.*closed without saving and its file deleted"):
        _job(_pre(), calculated)
    assert not list(new_book.folder.iterdir()) and new_book.sessions == []
    assert list(new_book.app.Workbooks) == [new_book.mine]


def test_a_workbook_added_and_saved_as_for_the_question_is_saved_in_the_job_too(fake_target, calculated, addin):
    """The old route of a new workbook (Workbooks.Add of the template, then SaveAs: target.Target.new without made):
    its first answer is saved in the job as well."""
    fake_target.new = True
    _job(_pre(), calculated)
    assert addin.asked == [] and fake_target.wb.saved == 1 and fake_target.wb.AutoSaveOn is True


# ------------------------------------------------------------------ review fixes
def test_autosave_a_save_after_the_reply_could_not_switch_on_is_said_and_switched_on_by_the_next_answer(
        fake_target, calculated, addin):
    """Final review M6: the add-in's save went through but AutoSave could not go back on (the outcome's
    autosave_error): the next answer says so, and its own save switches it on in that workbook."""
    _job(_pre(), calculated)
    wb = fake_target.wb
    addin.autosave_refused = "AutoSave is not available now"
    addin.release()
    addin.autosave_refused = None
    assert wb.saved == 1 and wb.AutoSaveOn is False
    out, _ = _job(_pre(), calculated)
    assert [w for w in out["warnings"] if "AutoSave" in w] == [
        f"AutoSave could not be switched back on in {wb.Name} after the save after the last answer (AutoSave is not "
        "available now); this answer switches it on"]
    addin.release()
    assert wb.saved == 2 and wb.AutoSaveOn is True


def test_autosave_left_off_in_another_workbook_of_ours_is_switched_on_by_the_next_answer(fake_target, calculated,
                                                                                         addin):
    import fakexl
    other = fakexl.Workbook(["Q1 Answer"], app=fake_target.app, styles_ver=False)
    other.Name, other.FullName = "older answer.xlsx", r"C:\analyser\older answer.xlsx"
    other.__dict__["_autosave"] = False
    fake_target.app.ActiveWorkbook = fake_target.wb
    addin.pending = {"workbook": other.Name, "full_name": other.FullName, "autosave": True, "token": "6" * 32,
                     "queued": 0.0}
    addin.autosave_refused = "AutoSave is not available now"
    addin.release()
    addin.autosave_refused = None
    out, _ = _job(_pre(), calculated)
    assert other.AutoSaveOn is True and [w for w in out["warnings"] if "older answer" in w] == [
        "AutoSave could not be switched back on in older answer.xlsx after the save after the last answer (AutoSave "
        "is not available now); it is switched on now"]



def test_a_failed_outcome_without_the_files_path_never_saves_by_name(fake_target, calculated, addin):
    """Review M1: the user's own workbook can share a name with ours: no path, no save - a note only."""
    import fakexl
    other = fakexl.Workbook(["Mine"], app=fake_target.app, styles_ver=False)
    other.Name = "older answer.xlsx"
    fake_target.app.ActiveWorkbook = fake_target.wb
    addin.pending = {"workbook": other.Name, "autosave": False, "token": "3" * 32, "queued": 0.0}
    addin.save_error = "the network went away"
    addin.release()
    addin.save_error = None
    out, _ = _job(_pre(), calculated)
    assert other.saved == 0 and fake_target.wb.saved == 1
    assert [w for w in out["warnings"] if "older answer" in w] == [
        "the save of older answer.xlsx after the last answer failed (the network went away); which file it was is not "
        "known, so it was not saved: save it yourself"]


def test_a_failed_save_of_another_file_of_the_same_name_is_not_said_saved(fake_target, calculated, addin):
    """Review M4: the failed file had this workbook's name but another path: this job's save is not that file's."""
    addin.pending = {"workbook": fake_target.wb.Name, "full_name": r"C:\\elsewhere\\" + fake_target.wb.Name,
                     "autosave": False, "token": "4" * 32, "queued": 0.0}
    real = fake_target.wb.FullName
    fake_target.wb.FullName = r"C:\\elsewhere\\" + fake_target.wb.Name
    addin.save_error = "locked"
    addin.release()
    addin.save_error = None
    fake_target.wb.FullName = real
    out, _ = _job(_pre(), calculated)
    notes = [w for w in out["warnings"] if "after the last answer failed" in w]
    assert notes == [f"the save of {fake_target.wb.Name} after the last answer failed (locked); the workbook open by "
                     "that name now is another file, so it was not saved"]
    assert fake_target.wb.saved == 1                        # this job's own file, in the job


def test_another_smartass_servers_failed_save_is_said_as_its_own(fake_target, calculated, addin, monkeypatch):
    """Review (Task 1b): the outcome of another build's save (another smartass server in this Excel, told by this
    job's first contact) is said as that server's, not as this one's."""
    monkeypatch.setattr(later, "_unwritten", {id(fake_target.app): [
        {"token": "5" * 32, "workbook": "older answer.xlsx", "full_name": "", "ok": False, "error": "locked",
         "closed": False, "seconds": 0.2, "started": 1.0, "unwritten": "disk", "other_server": True}]})
    out, _ = _job(_pre(), calculated)
    assert [w for w in out["warnings"] if "older answer" in w] == [
        "the save of older answer.xlsx (asked by another smartass server in this Excel) after the last answer failed "
        "(locked); which file it was is not known, so it was not saved: save it yourself"]


def test_an_outcome_that_cannot_be_read_is_said_and_the_job_saves_in_the_foreground(fake_target, calculated, addin):
    """Review M3: a malformed outcome never fails a finished job."""
    from smartass.core import config as cfg
    _job(_pre(), calculated)
    addin.pending = None
    folder = cfg.INPROC_DIR / "saved"
    folder.mkdir(parents=True, exist_ok=True)
    (folder / f"{id(fake_target.app)}.json").write_text('{"token": "x", "ok": tru', encoding="utf-8")
    out, _ = _job(_pre(), calculated)
    assert out["checks"]["sheet_equals_model"] == "OK" and fake_target.wb.saved == 1 and len(addin.asked) == 1
    assert [w for w in out["warnings"] if "could not tell" in w] == [
        "smartass could not tell how its save after the last answer went (its outcome file could not be read): this "
        "answer is saved in the job"]



def test_the_workbooks_path_is_read_with_a_busy_excel_waited_for():
    """Live (the review-fix probe): right after a save after the reply Excel rejected a job's read of its workbook's
    FullName (_is_ours) - the read waits for a busy Excel as every other call does."""
    import pywintypes

    from smartass.xl.target import Target
    tries = []

    class Book:
        @property
        def FullName(self):
            tries.append(1)
            if len(tries) < 3:
                raise pywintypes.com_error(-2147418111, "Call was rejected by callee.", None, None)
            return r"C:\analyser\answer.xlsx"
    t = Target(app=None, wb=Book(), pid=0, hwnd=0)
    assert t.path.endswith("answer.xlsx") and len(tries) == 3



def test_finding_which_excel_holds_the_target_waits_for_a_busy_excel():
    """Live (the review-fix probe): a follow-up sent while the save after the reply ran had its scan of the open
    workbooks rejected - before its turn there - and failed; the scan is read again while Excel is busy."""
    import pywintypes

    from smartass.xl import target as xltarget
    tries = []

    class Book:
        Name, FullName = "answer.xlsx", r"C:\analyser\answer.xlsx"

    class App:
        @property
        def Workbooks(self):
            tries.append(1)
            if len(tries) < 2:
                raise pywintypes.com_error(-2147418111, "Call was rejected by callee.", None, None)
            return [Book()]
    from pathlib import Path
    t = xltarget._find_open([(7, 0, App())], Path(r"C:\analyser\answer.xlsx"))
    assert t is not None and t.pid == 7 and len(tries) == 2
