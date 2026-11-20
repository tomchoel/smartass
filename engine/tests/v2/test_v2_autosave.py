# engine/tests/v2/test_v2_autosave.py
"""AutoSave around a file job (phase 3b Task 2b: read and switched off as the look batch's last operations - left on,
it wrote our file between the batches, live 20.11.2026 - back on in the save batch): our own file is saved with
AutoSave off and gets it back afterwards, also after a failure; the user's workbook is never saved and keeps it off
with a note - also in the error text of a job that failed (final review M2). On the fake Excel through the COM
executor; every figure is invented."""
import pytest
from test_file_job import _pre

import fakexl
from smartass import pipeline
from smartass.xl.inproc.runner import BatchError
from smartass.xl.target import Target


class AutoSaveBook(fakexl.Workbook):
    """A workbook in OneDrive: AutoSaveOn and Save are events of the fake Excel, in order with the rest."""

    def __init__(self, on: bool):
        super().__init__(["Sources"], styles_ver=False)
        self._autosave = on

    @property
    def AutoSaveOn(self):
        return self._autosave

    @AutoSaveOn.setter
    def AutoSaveOn(self, v):
        self._autosave = v
        self.Application.events.append(("", "AutoSaveOn", v))

    def Save(self):
        self.Application.events.append(("", "Save", self._autosave))


class NoAutoSaveBook(fakexl.Workbook):
    """An Excel without AutoSave: reading the property raises."""

    @property
    def AutoSaveOn(self):
        raise fakexl.ComError("no autosave")

    def Save(self):
        self.Application.events.append(("", "Save", None))


@pytest.fixture
def job(monkeypatch, tmp_path):
    from smartass.core import config
    monkeypatch.setattr(config, "HOME", tmp_path)
    monkeypatch.setattr(pipeline, "remember", lambda pre, out: None)

    def run(wb, ours: bool):
        monkeypatch.setattr(pipeline, "_target", lambda h, target, title: Target(app=wb.Application, wb=wb, pid=0,
                                                                                hwnd=0, created=ours))
        monkeypatch.setattr(pipeline, "_is_ours", lambda t: t.created)
        return pipeline.excel_job(None, _pre())
    return run


def _steps(wb) -> list:
    return [(e[1], *e[2:]) for e in wb.Application.events if e[1] in ("AutoSaveOn", "Save", "Sheets.Add")]


def test_our_file_autosave_off_while_writing_then_saved_and_restored(job):
    """Task 8B tried leaving the save to AutoSave once it is back on: live (19.11.2026) the last answer was not on
    disk 30 s later, so our file is saved explicitly, AutoSave off, as before."""
    wb = AutoSaveBook(True)
    out = job(wb, ours=True)
    assert [s[:2] for s in _steps(wb)] == [("AutoSaveOn", False), ("Sheets.Add", ["Q1 Answer", "Q1 Data"]),
                                           ("Save", False), ("AutoSaveOn", True)]
    assert wb.AutoSaveOn is True and not [w for w in out["warnings"] if "AutoSave" in w]


def test_the_users_workbook_is_never_saved_and_keeps_autosave_off_with_a_note(job):
    wb = AutoSaveBook(True)
    out = job(wb, ours=False)
    assert [s for s in _steps(wb) if s[0] != "Sheets.Add"] == [("AutoSaveOn", False)] and wb.AutoSaveOn is False
    assert [w for w in out["warnings"] if "AutoSave was switched off in book.xlsx" in w]
    wb = AutoSaveBook(False)
    out = job(wb, ours=False)                           # already off: the look's off changes nothing, no note
    assert [s for s in _steps(wb) if s[0] != "Sheets.Add"] == [("AutoSaveOn", False)] and wb.AutoSaveOn is False
    assert not [w for w in out["warnings"] if "AutoSave" in w]


def test_autosave_already_off_or_unsupported_is_quiet(job):
    """AutoSave already off: the look's off changes nothing and it is not switched on afterwards; an Excel without
    AutoSave: nothing at all."""
    for wb, steps in ((AutoSaveBook(False), [("AutoSaveOn", False)]),
                      (NoAutoSaveBook(["Sources"], styles_ver=False), [])):
        out = job(wb, ours=True)
        assert [s[:2] for s in _steps(wb) if s[0] == "AutoSaveOn"] == steps
        assert not [w for w in out["warnings"] if "AutoSave" in w]


def test_a_failed_job_turns_autosave_back_on_in_our_file_after_the_rollback(job, monkeypatch):
    wb = AutoSaveBook(True)

    def refuse(self, *a):
        raise fakexl.ComError("a table was refused")
    monkeypatch.setattr(fakexl._ListObjects, "Add", refuse)
    with pytest.raises(BatchError, match="refused") as e:
        job(wb, ours=True)
    assert wb.AutoSaveOn is True and [s[0] for s in _steps(wb)] == ["AutoSaveOn", "Sheets.Add", "AutoSaveOn"]
    assert [ws.Name for ws in wb.order] == ["Sources"] and "AutoSave" not in str(e.value)


def test_the_notes_of_a_failed_job_reach_its_error(job, monkeypatch):
    """Final review M2: AutoSave left off in the user's workbook is said in the error of a job that failed."""
    wb = AutoSaveBook(True)

    def refuse(self, *a):
        raise fakexl.ComError("a table was refused")
    monkeypatch.setattr(fakexl._ListObjects, "Add", refuse)
    with pytest.raises(BatchError) as e:
        job(wb, ours=False)
    assert str(e.value).endswith("; " + pipeline.AUTOSAVE_NOTE.format("book.xlsx")) and "refused" in str(e.value)
    assert wb.AutoSaveOn is False and [ws.Name for ws in wb.order] == ["Sources"]


def test_autosave_is_not_said_off_when_the_job_failed_before_it_was_switched_off(job, monkeypatch):
    wb = AutoSaveBook(True)
    monkeypatch.setattr(pipeline, "_look", lambda t, name: (_ for _ in ()).throw(fakexl.ComError("Excel said no")))
    with pytest.raises(fakexl.ComError) as e:
        job(wb, ours=False)
    assert str(e.value) == "Excel said no" and wb.AutoSaveOn is True and _steps(wb) == []

def test_autosave_is_switched_off_in_the_look_batch_not_in_the_timed_move_in(job, monkeypatch):
    from smartass.xl.inproc import runner
    runs, real = [], runner.run
    monkeypatch.setattr(runner, "run", lambda app, b, *a, **k: runs.append(b) or real(app, b, *a, **k))
    job(AutoSaveBook(True), ours=True)
    sets = [(n, op["value"]) for n, b in enumerate(runs) for op in b.to_dict()["ops"]
            if op["k"] == "set" and op["name"] == "AutoSaveOn"]
    assert sets == [(0, False), (len(runs) - 1, True)]           # the look; the save
