# engine/tests/v2/test_jobstate.py
"""Final review M3: the Excel settings a file job changes are recorded in a state file under the smartass home before
its first quiet batch and the file removed at its end, so what a job that did not finish left (a killed server, a
put-back that failed) is put back by the next job in that same Excel process - never in another - and said in its
notes. On the fake Excel through the COM executor; every figure is invented."""
import json

import pytest
from test_file_job import _pre
from test_v2_autosave import AutoSaveBook

import fakexl
from smartass import exceljob
from smartass.core import config
from smartass.xl import jobstate
from smartass.xl.inproc import runner
from smartass.xl.target import Target

KEY = (4242, 1_700_000_000.0)                   # the fake Excel's pid and process start time
QUIET = {"ScreenUpdating": False, "EnableEvents": False, "DisplayAlerts": False, "Calculation": -4135}
USER = {"ScreenUpdating": True, "EnableEvents": True, "DisplayAlerts": True, "Calculation": -4105}


def settings(app) -> dict:
    return {name: getattr(app, name) for name in USER}


@pytest.fixture
def excel(monkeypatch, tmp_path):
    """Our own analysis file (AutoSave on) in a fake Excel whose process is KEY; the jobs' targets are that file."""
    monkeypatch.setattr(config, "HOME", tmp_path)
    monkeypatch.setattr(exceljob, "remember", lambda pre, out: None)
    wb = AutoSaveBook(True)
    wb.Name, wb.FullName = "answer.xlsx", str(tmp_path / "analyser" / "answer.xlsx")
    monkeypatch.setattr(jobstate, "_start_time", lambda pid: KEY[1] if pid == KEY[0] else 0.0)
    monkeypatch.setattr(exceljob, "_target", lambda h, target, title, build_new=False: Target(
        app=wb.Application, wb=wb, pid=KEY[0], hwnd=0, created=True))
    monkeypatch.setattr(exceljob, "_is_ours", lambda t: t.created)
    return wb


def _left_by_a_dead_job(wb, started: float = KEY[1], autosave: bool = True) -> None:
    """A job that wrote its state and died in a quiet batch after its look switched AutoSave off: what it set is
    still so in Excel."""
    jobstate.begin(KEY[:1] + (started,), dict(USER), (wb.Name, wb.FullName) if autosave else None)
    for name, value in QUIET.items():
        setattr(wb.Application, name, value)
    wb._autosave = False


def test_a_job_puts_back_what_a_job_that_did_not_finish_left_and_says_so(excel):
    _left_by_a_dead_job(excel)
    out = exceljob.excel_job(None, _pre())
    assert settings(excel.Application) == USER and excel.AutoSaveOn is True
    (note,) = [w for w in out["warnings"] if "did not finish" in w]
    assert note == jobstate.NOTE.format("ScreenUpdating off, EnableEvents off, DisplayAlerts off, Calculation "
                                        "manual, AutoSave off in answer.xlsx")
    assert not list(jobstate.folder().glob("*.json"))


def test_the_state_of_another_excel_process_with_that_pid_is_removed_and_nothing_put_back(excel):
    _left_by_a_dead_job(excel, started=KEY[1] - 3600)
    excel.Application.Calculation = -4105                       # the user's own settings in this Excel
    out = exceljob.excel_job(None, _pre())
    assert not [w for w in out["warnings"] if "did not finish" in w]
    assert excel.Application.ScreenUpdating is False and excel.Application.EnableEvents is False   # left as found
    assert not list(jobstate.folder().glob("*.json"))


def test_a_setting_the_user_changed_since_is_left_as_it_is(excel):
    _left_by_a_dead_job(excel, autosave=False)
    excel.Application.Calculation = -4105                       # set back by the user meanwhile
    excel.Application.ScreenUpdating = True
    out = exceljob.excel_job(None, _pre())
    (note,) = [w for w in out["warnings"] if "did not finish" in w]
    assert note == jobstate.NOTE.format("EnableEvents off, DisplayAlerts off")
    assert settings(excel.Application) == USER


def test_a_job_records_its_state_before_its_first_quiet_batch_and_removes_it_at_its_end(excel, monkeypatch):
    seen = []
    real = exceljob._bring_in

    def bring_in(job, *a):
        (path,) = jobstate.folder().glob("*.json")
        seen.append(json.loads(path.read_text(encoding="utf-8")))
        return real(job, *a)
    monkeypatch.setattr(exceljob, "_bring_in", bring_in)
    exceljob.excel_job(None, _pre())
    assert seen == [{"pid": KEY[0], "started": KEY[1], "found": USER, "set": QUIET,
                     "autosave": {"name": "answer.xlsx", "full_name": excel.FullName}}]
    assert not list(jobstate.folder().glob("*.json"))


def test_a_failed_job_removes_its_state_too(excel, monkeypatch):
    def refuse(self, *a):
        raise fakexl.ComError("a table was refused")
    monkeypatch.setattr(fakexl._ListObjects, "Add", refuse)
    with pytest.raises(runner.BatchError):
        exceljob.excel_job(None, _pre())
    assert not list(jobstate.folder().glob("*.json")) and settings(excel.Application) == USER


def test_a_setting_a_batch_could_not_put_back_keeps_the_state_for_the_next_job(excel, monkeypatch):
    real = runner._restore

    def restore(app, changed):
        failed = real(app, changed)
        app.ScreenUpdating = False                              # as if Excel had refused this one
        return failed + ["ScreenUpdating=True: refused"]
    monkeypatch.setattr(runner, "_restore", restore)
    exceljob.excel_job(None, _pre())
    assert len(list(jobstate.folder().glob("*.json"))) == 1
    monkeypatch.setattr(runner, "_restore", real)
    out = exceljob.excel_job(None, _pre())
    assert excel.Application.ScreenUpdating is True
    assert jobstate.NOTE.format("ScreenUpdating off") in out["warnings"]
    assert not list(jobstate.folder().glob("*.json"))


def test_an_excel_that_cannot_be_told_gets_no_state_file(excel, monkeypatch):
    def unknown(pid):
        raise ProcessLookupError(pid)
    monkeypatch.setattr(jobstate, "_start_time", unknown)
    exceljob.excel_job(None, _pre())
    assert not jobstate.folder().exists() or not list(jobstate.folder().glob("*.json"))


class StuckAutoSave(AutoSaveBook):
    """Our own file whose AutoSave Excel refuses to switch back on while refuse is set."""
    refuse = True

    @property
    def AutoSaveOn(self):
        return self._autosave

    @AutoSaveOn.setter
    def AutoSaveOn(self, v):
        if v and self.refuse:
            raise fakexl.ComError("AutoSave cannot be switched on now")
        self._autosave = v


@pytest.mark.parametrize("fails", ["the job", "the save"])
def test_autosave_the_job_could_not_switch_back_on_is_said_and_kept_for_the_next_job(excel, monkeypatch, fails):
    """Task 3 review Minor 1: a failed job's clean-up, or the save batch, that cannot switch AutoSave back on in our
    own file says so (in the error, or the reply's warnings) and keeps the state, so the next job puts it right."""
    wb = StuckAutoSave(True)
    wb.Name, wb.FullName = excel.Name, excel.FullName
    monkeypatch.setattr(exceljob, "_target", lambda h, target, title, build_new=False: Target(
        app=wb.Application, wb=wb, pid=KEY[0], hwnd=0, created=True))
    monkeypatch.setattr(config, "SAVE", "sync")                 # the save batch switches it on
    if fails == "the job":
        real = fakexl._ListObjects.Add

        def refuse(self, *a):
            raise fakexl.ComError("a table was refused")
        monkeypatch.setattr(fakexl._ListObjects, "Add", refuse)
        with pytest.raises(runner.BatchError) as e:
            exceljob.excel_job(None, _pre())
        said = str(e.value)
        monkeypatch.setattr(fakexl._ListObjects, "Add", real)
    else:
        said = "; ".join(exceljob.excel_job(None, _pre())["warnings"])
    assert exceljob.AUTOSAVE_LEFT.format("answer.xlsx", "AutoSave cannot be switched on now").split(" (")[0] in said
    assert len(list(jobstate.folder().glob("*.json"))) == 1 and wb.AutoSaveOn is False
    wb.refuse = False
    out = exceljob.excel_job(None, _pre())
    assert wb.AutoSaveOn is True and jobstate.NOTE.format("AutoSave off in answer.xlsx") in out["warnings"]
    assert not list(jobstate.folder().glob("*.json"))
