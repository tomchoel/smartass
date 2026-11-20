# engine/tests/v2/test_v2_autosave.py
"""AutoSave around a file job (phase 3b Task 2b, controller ruling R3): AutoSave is never touched - not read, not
switched off, not switched back on - in our own file or in the user's workbook. Our own file is saved with one
Workbook.Save at the end of the job (Save costs the same with AutoSave on, and AutoSave does not save our edits between
the batches: measured by the controller 20.11.2026); the user's workbook is never saved. On the fake Excel through the
COM executor; every figure is invented."""
import pytest
from test_file_job import _pre

import fakexl
from smartass import pipeline
from smartass.xl.inproc import runner
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


@pytest.mark.parametrize("on", [True, False])
def test_our_file_is_saved_once_and_its_autosave_never_touched(job, on):
    wb = AutoSaveBook(on)
    out = job(wb, ours=True)
    assert [s[:2] for s in _steps(wb)] == [("Sheets.Add", ["Q1 Answer", "Q1 Data"]), ("Save", on)]
    assert wb.AutoSaveOn is on and not [w for w in out["warnings"] if "AutoSave" in w]


def test_the_users_workbook_is_never_saved_and_its_autosave_never_touched(job):
    wb = AutoSaveBook(True)
    out = job(wb, ours=False)
    assert [s[0] for s in _steps(wb)] == ["Sheets.Add"] and wb.AutoSaveOn is True
    assert not [w for w in out["warnings"] if "AutoSave" in w]


def test_an_excel_without_autosave_is_no_matter(job):
    wb = NoAutoSaveBook(["Sources"], styles_ver=False)
    out = job(wb, ours=True)
    assert [s[0] for s in _steps(wb)] == ["Sheets.Add", "Save"] and out["checks"]["sheet_equals_model"]


@pytest.mark.parametrize("ours", [True, False])
def test_a_failed_job_leaves_autosave_as_it_was_and_raises_excels_error(job, monkeypatch, ours):
    wb = AutoSaveBook(True)

    def refuse(self, *a):
        raise fakexl.ComError("a table was refused")
    monkeypatch.setattr(fakexl._ListObjects, "Add", refuse)
    with pytest.raises(BatchError) as e:
        job(wb, ours=ours)
    assert "AutoSave" not in str(e.value) and str(e.value).endswith("a table was refused")
    assert wb.AutoSaveOn is True and [s[0] for s in _steps(wb)] == ["Sheets.Add"]
    assert [ws.Name for ws in wb.order] == ["Sources"]


def test_no_batch_reads_or_sets_autosave(job, monkeypatch):
    runs, real = [], runner.run
    monkeypatch.setattr(runner, "run", lambda app, b, *a, **k: runs.append(b) or real(app, b, *a, **k))
    job(AutoSaveBook(True), ours=True)
    assert runs and not [op for b in runs for op in b.to_dict()["ops"] if op.get("name") == "AutoSaveOn"]
