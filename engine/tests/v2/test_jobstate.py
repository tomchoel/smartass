# engine/tests/v2/test_jobstate.py
"""Final review M3: the Excel settings a file job changes are recorded in a state file under the smartass home before
its first quiet batch and the file removed at its end, so what a job that did not finish left (a killed server, a
put-back that failed) is put back by the next job in that same Excel process - never in another - and said in its
notes. On the fake Excel through the COM executor; every figure is invented."""
import json

import psutil
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


def test_a_put_back_that_fails_keeps_the_state_for_the_next_job(excel, monkeypatch):
    """Task 3 review Minor 8: what could not be put back is not forgotten - the batch failed, or Excel refused one
    of its settings."""
    _left_by_a_dead_job(excel)

    def broken(app, found, sets, autosave):
        raise RuntimeError("the batch failed")
    real = jobstate._put_back
    monkeypatch.setattr(jobstate, "_put_back", broken)
    jobstate.put_back(excel.Application, KEY)
    assert len(list(jobstate.folder().glob("*.json"))) == 1
    calls = []
    real_run = runner.run

    def run(app, b, *a, **kw):                                  # the put-back's write batch: Excel refuses one
        res = real_run(app, b, *a, **kw)
        calls.append(1)
        if len(calls) == 2:
            res.skipped.append((0, "", "refused"))
        return res
    monkeypatch.setattr(jobstate, "_put_back", real)
    monkeypatch.setattr(runner, "run", run)
    jobstate.put_back(excel.Application, KEY)
    assert len(list(jobstate.folder().glob("*.json"))) == 1
    monkeypatch.setattr(runner, "run", real_run)
    jobstate.put_back(excel.Application, KEY)                       # (the fake did set it: nothing left to put back)
    assert not list(jobstate.folder().glob("*.json"))


def test_a_job_prunes_the_state_of_an_excel_that_is_gone(excel, monkeypatch):
    """Task 3 review Minor 4: a state whose Excel process is gone (no such pid, or one started at another time) is
    removed when a job starts; a live one of another Excel is left alone."""
    folder = jobstate.folder()
    folder.mkdir(parents=True, exist_ok=True)
    states = {4243: 5.0, 4244: 0.0, 4245: 0.0}                  # 4243 restarted since, 4244 alive, 4245 gone
    for pid, started in states.items():
        (folder / f"{pid}.json").write_text(json.dumps({"pid": pid, "started": started, "found": {}, "set": {},
                                                        "autosave": None}), encoding="utf-8")

    def start(pid):
        if pid == 4245:
            raise psutil.NoSuchProcess(pid)
        return KEY[1] if pid == KEY[0] else 0.0
    monkeypatch.setattr(jobstate, "_start_time", start)
    exceljob.excel_job(None, _pre())
    assert sorted(p.name for p in folder.glob("*.json")) == ["4244.json"]


def test_a_state_whose_excels_start_time_cannot_be_read_is_not_pruned(excel, monkeypatch):
    """Final review M12: a read error (access denied) is not "process gone": that Excel's state stays."""
    folder = jobstate.folder()
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "4246.json").write_text(json.dumps({"pid": 4246, "started": 0.0, "found": {}, "set": {},
                                                  "autosave": None}), encoding="utf-8")

    def start(pid):
        if pid == 4246:
            raise psutil.AccessDenied(pid)
        return KEY[1]
    monkeypatch.setattr(jobstate, "_start_time", start)
    exceljob.excel_job(None, _pre())
    assert sorted(p.name for p in folder.glob("*.json")) == ["4246.json"]


def _state() -> dict:
    (path,) = jobstate.folder().glob("*.json")
    return json.loads(path.read_text(encoding="utf-8"))


def test_what_a_failed_put_back_left_survives_the_next_whole_job_until_a_put_back_succeeds(excel, monkeypatch):
    """Final review M1: a dead job left only Calculation manual and the next job's put-back failed. That job's own
    record carries what could not be put back (with what the DEAD job found), so its end leaves it for the job after,
    which puts Calculation back to automatic - it is not taken for the user's own."""
    jobstate.begin(KEY, dict(USER), None)
    excel.Application.Calculation = QUIET["Calculation"]
    real = jobstate._put_back

    def broken(app, found, sets, autosave):
        raise RuntimeError("the batch failed")
    monkeypatch.setattr(jobstate, "_put_back", broken)
    exceljob.excel_job(None, _pre())
    state = _state()
    assert state["set"]["Calculation"] == QUIET["Calculation"] and state["found"]["Calculation"] == USER["Calculation"]
    assert excel.Application.Calculation == QUIET["Calculation"]
    monkeypatch.setattr(jobstate, "_put_back", real)
    out = exceljob.excel_job(None, _pre())
    assert excel.Application.Calculation == USER["Calculation"]
    assert jobstate.NOTE.format("Calculation manual") in out["warnings"]
    assert not list(jobstate.folder().glob("*.json"))


def test_a_put_back_excel_refused_in_part_keeps_only_what_was_not_put_back(excel, monkeypatch):
    _left_by_a_dead_job(excel, autosave=False)
    calls, real_run = [], runner.run

    def run(app, b, *a, **kw):                                  # the write batch: Excel refuses its first setting
        res = real_run(app, b, *a, **kw)
        calls.append(1)
        if len(calls) == 2:
            res.skipped.append((0, "", "refused"))
        return res
    monkeypatch.setattr(runner, "run", run)
    jobstate.put_back(excel.Application, KEY)
    assert _state() == {"pid": KEY[0], "started": KEY[1], "found": {"ScreenUpdating": True},
                        "set": {"ScreenUpdating": False}, "autosave": None}


def test_autosave_alone_kept_never_puts_back_settings_the_job_restored_itself(excel, monkeypatch):
    """Final review M2: a job finished but Excel refused AutoSave back on; what is kept for the next job is that
    AutoSave only - not the settings the job restored, which the user may change on purpose meanwhile."""
    wb = StuckAutoSave(True)
    wb.Name, wb.FullName = excel.Name, excel.FullName
    monkeypatch.setattr(exceljob, "_target", lambda h, target, title, build_new=False: Target(
        app=wb.Application, wb=wb, pid=KEY[0], hwnd=0, created=True))
    monkeypatch.setattr(config, "SAVE", "sync")
    exceljob.excel_job(None, _pre())
    assert _state()["set"] == {} and _state()["autosave"]["name"] == "answer.xlsx"
    wb.Application.Calculation = QUIET["Calculation"]          # the user sets manual on purpose
    wb.refuse = False
    out = exceljob.excel_job(None, _pre())
    assert wb.Application.Calculation == QUIET["Calculation"] and wb.AutoSaveOn is True
    assert [w for w in out["warnings"] if "did not finish" in w] == [jobstate.NOTE.format("AutoSave off in answer.xlsx")]
    assert not list(jobstate.folder().glob("*.json"))
