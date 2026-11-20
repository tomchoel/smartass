"""Running a batch (xl/inproc/runner.py): the COM executor on the fake Excel, and the in-process executor's Python side
with a fake Application.Run that executes the job file through the COM executor, as the add-in does inside Excel."""
import json
import os
import threading
import time
from pathlib import Path

import pythoncom
import pytest
import pywintypes

from fakexl import Application, Workbook, count_com, snapshot
from smartass.core import config
from smartass.xl import jobs
from smartass.xl.host import HostError
from smartass.xl.inproc import build, runner
from smartass.xl.inproc.batch import MISSING, Batch

QUIET_OFF = {"ScreenUpdating": True, "EnableEvents": True, "DisplayAlerts": True, "Calculation": -4105}


def settings(app) -> dict:
    return {k: getattr(app, k) for k in QUIET_OFF}


@pytest.fixture
def book():
    app = Application()
    wb = Workbook(("S", "T"), app=app)
    return app, wb


# ------------------------------------------------------------------ the COM executor
def test_every_operation_kind_runs_and_values_come_back(book):
    app, wb = book
    wb.links = [r"C:\tmp\smartass_build_1.xlsx", r"C:\other\Budget.xlsx"]
    b = Batch()
    ws = b.workbook("book.xlsx").Worksheets("S")
    ws.Range("A1:C2").Value2 = ((1, "two", None), (2.5, True, "=A1"))
    ws.Name = "Q3 Data"
    b.ret(ws.Range("A1:C2").Value2, "grid")
    b.ret(ws.Range("B1").Value2, "one")
    b.ret(ws.Range("D9").Value2, "empty")
    b.ret(b.app.Workbooks.Count, "books")
    b.mark("written")
    b.ret(b.verb("break_links", b.workbook("book.xlsx"), "SMARTASS_BUILD_1.xlsx"), "broke")
    b.ret(b.workbook("book.xlsx").LinkSources(1), "links")
    res = runner.run(app, b)
    assert res.where == "com"
    assert res.values == {"grid": ((1, "two", None), (2.5, True, "=A1")), "one": "two", "empty": None, "books": 1,
                          "broke": 1, "links": (r"C:\other\Budget.xlsx",)}
    assert list(res.marks) == ["written"] and isinstance(res.marks["written"], float)
    assert res.skipped == [] and res.ms >= 0
    assert [s.Name for s in wb.order] == ["Q3 Data", "T"]


def test_missing_reaches_excel_as_an_omitted_argument(book):
    app, wb = book
    b = Batch()
    wbn = b.workbook("book.xlsx")
    wbn.Worksheets(["S"]).Copy(MISSING, wbn.Worksheets("T"))
    lo = wbn.Worksheets("T").ListObjects.Add(1, wbn.Worksheets("T").Range("A1:B3"), MISSING, 1)
    lo.Name = "Data_T"
    runner.run(app, b)
    assert [s.Name for s in wb.order] == ["S", "T", "S (2)"]
    assert wb.Worksheets("T").ListObjects("Data_T").Name == "Data_T"


def test_the_quiet_settings_are_captured_set_and_restored(book):
    app, wb = book
    app.DisplayAlerts = False                          # the user's own: left as it is
    b = Batch()
    b.ret(b.app.ScreenUpdating, "screen")
    b.ret(b.app.Calculation, "calc")
    b.ret(b.app.EnableEvents, "events")
    res = runner.run(app, b)
    assert res.values == {"screen": False, "calc": -4135, "events": False}
    assert settings(app) == {**QUIET_OFF, "DisplayAlerts": False}
    res = runner.run(app, _quiet_read(Batch(quiet=False)))
    assert res.values["screen"] is True


def _quiet_read(b):
    b.ret(b.app.ScreenUpdating, "screen")
    return b


def test_a_failing_operation_names_itself_after_finally_ran_and_the_settings_came_back(book):
    app, wb = book
    b = Batch()
    wbn = b.workbook("book.xlsx")
    wbn.Worksheets("S").Range("A1").Value2 = "before"
    wbn.Worksheets("Nope").Range("A1").Value2 = "never"
    wbn.Worksheets("S").Range("A2").Value2 = "never either"
    with b.always():
        wbn.Worksheets("T").Range("A1").Value2 = "finally"
    with pytest.raises(runner.BatchError) as e:
        runner.run(app, b)
    assert e.value.index == 4
    assert e.value.text == 'Workbooks("book.xlsx").Worksheets("Nope")'
    assert e.value.message.startswith("operation 4 (call Worksheets): ") and "Nope" in str(e.value)
    assert wb.Worksheets("S").cells[(1, 1)]["Value"] == "before"
    assert (2, 1) not in wb.Worksheets("S").cells
    assert wb.Worksheets("T").cells[(1, 1)]["Value"] == "finally"
    assert settings(app) == QUIET_OFF


class StubbornApp(Application):
    """An Excel that refuses to turn ScreenUpdating back on."""

    def __setattr__(self, k, v):
        if k == "ScreenUpdating" and v is True and getattr(self, "ScreenUpdating", None) is False:
            raise pywintypes.com_error(-2147352567, "Exception occurred.", (0, "Microsoft Excel", "No.", None, 0,
                                                                            -2146827284), None)
        super().__setattr__(k, v)


@pytest.mark.parametrize("on", [False, True])
def test_a_setting_that_cannot_be_put_back_is_a_warning_and_the_result_stands(monkeypatch, caplog, tmp_path, on):
    if on:
        _inproc_on(monkeypatch, tmp_path)
    app = StubbornInproc() if on else StubbornApp()
    Workbook(("S",), app=app)
    b = Batch()
    b.workbook("book.xlsx").Worksheets("S").Range("A1").Value2 = 1
    res = runner.run(app, b)
    assert res.where == ("inproc" if on else "com")
    warnings = [r.getMessage() for r in caplog.records if r.levelname == "WARNING"]
    assert len(warnings) == 1 and "ScreenUpdating=True" in warnings[0] and "No." in warnings[0]


def test_a_try_operation_that_fails_is_skipped_and_reported(book):
    app, wb = book
    b = Batch()
    ws = b.workbook("book.xlsx").Worksheets("S")
    old = b.attempt(ws.ListObjects("Old"))
    with b.attempt():
        old.Delete()
    ws.Range("A1").Value2 = 7
    res = runner.run(app, b)
    assert [(i, t) for i, t, _ in res.skipped] == [(2, 'Workbooks("book.xlsx").Worksheets("S").ListObjects("Old")'),
                                                  (3, 'Workbooks("book.xlsx").Worksheets("S").ListObjects("Old").Delete()')]
    assert "Old" in res.skipped[0][2]
    assert res.skipped[1][2] == "operation 3 (call Delete): " + runner.UNBOUND.format(3)
    assert wb.Worksheets("S").cells[(1, 1)]["Value"] == 7


def test_a_finally_operation_on_an_object_never_made_is_skipped_and_reported(book):
    app, wb = book
    b = Batch()
    src = b.app.Workbooks.Open(r"C:\nowhere\x.xlsx", 0, True)
    with b.always():
        src.Close(False)
    with pytest.raises(runner.BatchError) as e:
        runner.run(app, b)
    assert e.value.index == 1
    assert e.value.skipped == [(2, r'Application.Workbooks.Open("C:\nowhere\x.xlsx", 0, True).Close(False)',
                                "operation 2 (call Close): " + runner.UNBOUND.format(2))]


def test_a_com_object_asked_back_is_an_error(book):
    app, wb = book
    b = Batch()
    b.ret(b.workbook("book.xlsx").Worksheets("S").Range("A1"), "range")
    with pytest.raises(runner.BatchError, match="COM object") as e:
        runner.run(app, b)
    assert e.value.index == 3


def test_one_jobs_com_per_batch(book, monkeypatch):
    app, wb = book
    calls = count_com(monkeypatch)
    b = Batch()
    ws = b.workbook("book.xlsx").Worksheets("S")
    for r in range(1, 30):
        ws.Range(f"A{r}").Value2 = r
    runner.run(app, b)
    assert len(calls) == 1


def test_a_busy_excel_is_waited_for_on_the_rejected_operation_only(book, monkeypatch):
    app, wb = book
    monkeypatch.setattr("time.sleep", lambda s: None)
    ws = wb.Worksheets("S")
    seen = []
    real_range = ws.Range
    rejected = []

    def range_(a, b=None):
        if a == "A3" and not rejected:
            rejected.append(a)
            raise pywintypes.com_error(-2147418111, "Call was rejected by callee.", None, None)
        seen.append(a)
        return real_range(a, b)
    ws.__dict__["Range"] = range_
    b = Batch()
    wsn = b.workbook("book.xlsx").Worksheets("S")
    for a in ("A1", "A2", "A3", "A4"):
        wsn.Range(a).Value2 = a
    runner.run(app, b)
    assert rejected == ["A3"]
    assert seen == ["A1", "A2", "A3", "A4"]                    # each once: nothing before the rejection ran again


def test_an_abandoned_job_stops_after_the_finally_and_the_restore(book):
    app, wb = book
    token = jobs.JobToken()
    jobs.bind_job(token)
    try:
        b = Batch()
        ws = b.workbook("book.xlsx").Worksheets("S")
        ws.Range("A1").Value2 = 1
        ws.Range("A2").Value2 = 2
        with b.always():
            ws.Range("A3").Value2 = 3
        ws_obj = wb.Worksheets("S")
        real = ws_obj.Range

        def range_(a, b=None):
            if a == "A1":
                token.aborted = True
            return real(a, b)
        ws_obj.__dict__["Range"] = range_
        with pytest.raises(jobs.JobAborted):
            runner.run(app, b)
        assert (1, 1) not in ws_obj.cells and (2, 1) not in ws_obj.cells
        assert ws_obj.cells[(3, 1)]["Value"] == 3
        assert settings(app) == QUIET_OFF
    finally:
        jobs.bind_job()


def test_the_com_executor_returns_the_out_format_of_the_add_in(book):
    app, wb = book
    b = Batch(quiet=False)
    b.ret(b.workbook("book.xlsx").Worksheets("S").Name, "name")
    out = runner.execute(app, b.to_dict())
    assert out["ok"] is True and out["started"] == 4 and out["values"] == {"name": "S"}
    assert set(out) == {"ok", "started", "values", "marks", "ms", "skipped"}
    assert pythoncom.ArgNotFound is runner.OMITTED


# ------------------------------------------------------------------ the in-process executor's Python side
HASH = "abcd1234"
NOT_THERE = (-2147352567, "Exception occurred.", (0, "Microsoft Excel", "Cannot run the macro. The macro may not be "
                                                  "available in this workbook.", None, 0, -2146827284), None)


class InprocExcel(Application):
    """The fake Excel with the add-in's two entry points: RegisterXLL loads it; Run of SmartassRun_<hash> reads the
    job file, deletes it, runs it through the COM executor and returns the result record (a long one in the result
    file), as Runner.cs does."""

    def __init__(self):
        super().__init__()
        self.registered, self.runs = [], []
        self.loaded, self.register_answer, self.ping_answer = False, True, HASH
        self.fail_run = None                  # a function(job_path) that raises, instead of running

    def RegisterXLL(self, path):
        self.registered.append(str(path))
        self.loaded = self.loaded or self.register_answer is True
        return self.register_answer

    def Run(self, macro, *args):
        self.runs.append(macro)
        if not self.loaded or not macro.endswith("_" + HASH):
            raise pywintypes.com_error(*NOT_THERE)
        if macro.startswith("SmartassPing_"):
            return self.ping_answer
        path = Path(args[0])
        if self.fail_run is not None:
            return self.fail_run(path)
        job = json.loads(path.read_text(encoding="utf-8"))
        path.unlink()
        record = json.dumps(runner.execute(self, job))
        if len(record) <= 30000:
            return record
        Path(str(path) + ".out").write_text(record, encoding="utf-8")
        return "file"


def _inproc_on(monkeypatch, folder: Path) -> list:
    """SV_INPROC on, the add-in 'built' (no compiler), each fake Excel its own process, 64-bit."""
    monkeypatch.setattr(config, "INPROC", True)
    monkeypatch.setattr(config, "INPROC_DIR", folder / "inproc")
    monkeypatch.setattr(runner, "_EXCELS", {})
    monkeypatch.setattr(runner, "_identify", lambda app: (id(app), 0.0))
    monkeypatch.setattr(runner, "_is_64bit", lambda pid: True)
    built = []

    def ensure():
        built.append(1)
        return folder / f"smartass_run_{HASH}.xll", HASH
    monkeypatch.setattr(build, "ensure", ensure)
    return built


@pytest.fixture
def inproc(monkeypatch, tmp_path):
    return _inproc_on(monkeypatch, tmp_path)


class StubbornInproc(StubbornApp, InprocExcel):
    """The fake Excel with the add-in, refusing to turn ScreenUpdating back on."""


def excel_book():
    app = InprocExcel()
    return app, Workbook(("S", "T"), app=app)


AWKWARD = (0.1 + 0.2, 1 / 3, 1e22, 3.0, -0.0, 5e-324, 1.7976931348623157e308)


def value_batch():
    b = Batch()
    ws = b.workbook("book.xlsx").Worksheets("S")
    grid = ((None, True, False, 0, -7, 2 ** 40), AWKWARD[:6], ('say "hi"', "æøå ÆØÅ € 😀", "", "=A1", "x" * 30000, 1.5))
    ws.Range("A1:F3").Value2 = grid
    ws.Range("A5").Value2 = AWKWARD[6]
    ws.Range("A6").Value2 = -2146826281                       # a cell error, as pywin32 reads it
    b.ret(ws.Range("A1:F3").Value2, "grid")
    b.ret(ws.Range("A3").Value2, "one")
    b.ret(ws.Range("A5").Value2, "big")
    b.ret(ws.Range("A6").Value2, "error")
    b.ret(ws.Range("H9").Value2, "empty")
    wbn = b.workbook("book.xlsx")
    wbn.Worksheets(["S"]).Copy(MISSING, wbn.Worksheets("T"))
    b.ret(wbn.Worksheets.Count, "count")
    return b, grid


def typed(v):
    """v with every leaf's type, so 3 and 3.0 (or -0.0 and 0.0) differ."""
    if isinstance(v, tuple):
        return ("tuple", tuple(typed(x) for x in v))
    return (type(v).__name__, repr(v))


def test_every_value_kind_comes_back_from_the_add_in_as_the_com_executor_gives_it(inproc, monkeypatch):
    app, wb = excel_book()
    b, grid = value_batch()
    res = runner.run(app, b)
    assert res.where == "inproc"
    monkeypatch.setattr(config, "INPROC", False)
    app2, wb2 = excel_book()
    ref = runner.run(app2, value_batch()[0])
    assert ref.where == "com"
    assert {k: typed(v) for k, v in res.values.items()} == {k: typed(v) for k, v in ref.values.items()}
    assert typed(res.values["grid"]) == typed(tuple(tuple(None if c == "" else c for c in r) for r in grid))  # "": empty
    assert res.values["count"] == 3 and res.values["empty"] is None and res.values["error"] == -2146826281
    assert snapshot(wb) == snapshot(wb2)


def test_the_add_in_is_loaded_once_per_excel_and_each_batch_is_one_run(inproc):
    app, wb = excel_book()
    for _ in range(3):
        runner.run(app, Batch())
    assert app.registered == [str(Path(config.INPROC_DIR).parent / f"smartass_run_{HASH}.xll")]
    assert app.runs == [f"SmartassPing_{HASH}"] + [f"SmartassRun_{HASH}"] * 3
    assert list((config.INPROC_DIR / "jobs").iterdir()) == []          # job and result files are removed


def test_a_failing_operation_is_the_same_batch_error_in_both_executors(inproc, monkeypatch):
    def failing():
        b = Batch()
        ws = b.workbook("book.xlsx").Worksheets("S")
        with b.attempt():
            ws.ListObjects("Old").Delete()
        ws.Range("A1").Value2 = 1
        b.workbook("book.xlsx").Worksheets("Nope").Activate()
        with b.always():
            ws.Range("A2").Value2 = 2
        return b
    errors = []
    for on in (True, False):
        monkeypatch.setattr(config, "INPROC", on)
        app, wb = excel_book()
        with pytest.raises(runner.BatchError) as e:
            runner.run(app, failing())
        errors.append((e.value.index, e.value.text, e.value.message, e.value.skipped))
        assert wb.Worksheets("S").cells[(2, 1)]["Value"] == 2
        assert settings(app) == QUIET_OFF
    assert errors[0] == errors[1]
    assert errors[0][0] == 6 and errors[0][1] == 'Workbooks("book.xlsx").Worksheets("Nope")'


def test_an_add_in_gone_from_excel_is_registered_once_more(inproc):
    app, wb = excel_book()
    runner.run(app, Batch())
    app.loaded = False                                 # e.g. the add-in was removed from that Excel
    res = runner.run(app, Batch())
    assert res.where == "inproc" and len(app.registered) == 2


def test_an_add_in_that_cannot_be_reached_falls_back_before_any_operation_ran(inproc, caplog):
    app, wb = excel_book()
    runner.run(app, Batch())
    app.loaded, app.register_answer = False, False
    b = Batch()
    b.workbook("book.xlsx").Worksheets("S").Range("A1").Value2 = "once"
    res = runner.run(app, b)
    assert res.where == "com" and wb.Worksheets("S").cells[(1, 1)]["Value"] == "once"
    runs = len(app.runs)
    assert runner.run(app, Batch()).where == "com"
    assert len(app.runs) == runs                       # that Excel is not asked again
    assert len([r for r in caplog.records if r.levelname == "WARNING"]) == 1


def test_a_failure_of_the_add_in_before_its_first_operation_runs_the_batch_over_com(inproc):
    app, wb = excel_book()

    def before_first(path):
        path.unlink()
        return json.dumps({"ok": False, "started": 0, "index": -1, "message": "could not read the job", "values": {},
                           "marks": {}, "ms": 0.1, "skipped": []})
    app.fail_run = before_first
    b = Batch()
    b.workbook("book.xlsx").Worksheets("S").Range("A1").Value2 = "com"
    assert runner.run(app, b).where == "com"
    assert wb.Worksheets("S").cells[(1, 1)]["Value"] == "com"


def test_a_batch_the_add_in_took_is_never_run_again(inproc):
    app, wb = excel_book()
    ran = []

    def runs_then_fails(path):
        job = json.loads(path.read_text(encoding="utf-8"))
        path.unlink()
        runner.execute(app, job)
        ran.append(1)
        raise pywintypes.com_error(-2147352567, "Exception occurred.", None, None)
    runner.run(app, Batch())
    app.fail_run = runs_then_fails
    b = Batch()
    b.workbook("book.xlsx").Worksheets("S").Range("A1").Value2 = "x"
    with pytest.raises(runner.BatchError) as e:
        runner.run(app, b)
    assert e.value.index is None and ran == [1]
    assert [ev for ev in app.events if ev[1] == "Value2"] == [("S", "Value2", "A1")]


def test_sv_inproc_0_never_loads_the_add_in(inproc, monkeypatch):
    monkeypatch.setattr(config, "INPROC", False)
    app, wb = excel_book()
    assert runner.available(app) is False
    assert runner.run(app, Batch()).where == "com"
    assert app.registered == [] and app.runs == [] and inproc == []


@pytest.mark.parametrize("what", ["register false", "register raises", "ping mismatch", "32-bit", "no build"])
def test_each_load_failure_means_the_com_executor_for_that_excel_and_one_warning(inproc, monkeypatch, caplog, what):
    def broken():
        app, wb = excel_book()
        if what == "register false":
            app.register_answer = False
        elif what == "register raises":
            error = pywintypes.com_error(-2147352567, "Exception occurred.", None, None)
            monkeypatch.setattr(app, "RegisterXLL", lambda path: (_ for _ in ()).throw(error))
        elif what == "ping mismatch":
            app.ping_answer = "00000000"
        return app, wb
    if what == "32-bit":
        monkeypatch.setattr(runner, "_is_64bit", lambda pid: False)
    elif what == "no build":
        monkeypatch.setattr(build, "ensure", lambda: (_ for _ in ()).throw(build.BuildError("csc said no")))
    app, wb = broken()
    for _ in range(2):
        b = Batch()
        b.workbook("book.xlsx").Worksheets("S").Range("A1").Value2 = what
        assert runner.run(app, b).where == "com"
    assert wb.Worksheets("S").cells[(1, 1)]["Value"] == what
    warnings = [r.getMessage() for r in caplog.records if r.levelname == "WARNING"]
    assert len(warnings) == 1 and "COM" in warnings[0]
    other, _ = broken()
    assert runner.run(other, Batch()).where == "com"
    assert len([r for r in caplog.records if r.levelname == "WARNING"]) == 2     # once per Excel process


def test_a_busy_excel_is_an_error_and_not_a_fallback(inproc, monkeypatch):
    app, wb = excel_book()
    runner.run(app, Batch())
    monkeypatch.setattr("time.sleep", lambda s: None)
    app.fail_run = lambda path: (_ for _ in ()).throw(pywintypes.com_error(-2147418111, "rejected", None, None))
    b = Batch()
    b.workbook("book.xlsx").Worksheets("S").Range("A1").Value2 = "x"
    with pytest.raises(HostError, match="busy"):
        runner.run(app, b, limit=0.01)
    assert (1, 1) not in wb.Worksheets("S").cells


# ------------------------------------------------------------------ review fixes: outcomes, claims, the load
def ran_then(answer):
    """A fake add-in that claims and runs the job, then gives the answer given (a broken or missing result)."""
    def run(path, app):
        job = json.loads(path.read_text(encoding="utf-8"))
        path.unlink()
        runner.execute(app, job)
        return answer
    return run


@pytest.mark.parametrize("answer", ["{broken", '{"ok": true}', None, -2146826246, "file"])
def test_a_result_that_cannot_be_read_is_an_unknown_outcome_and_never_run_again(inproc, answer):
    app, wb = excel_book()
    runner.run(app, Batch())
    app.fail_run = lambda path: ran_then(answer)(path, app)        # "file": the result file is not there
    b = Batch()
    b.workbook("book.xlsx").Worksheets("S").Range("A1").Value2 = "once"
    with pytest.raises(runner.BatchError, match="outcome is unknown") as e:
        runner.run(app, b)
    assert e.value.index is None and e.value.where == "inproc"
    assert [ev for ev in app.events if ev[1] == "Value2"] == [("S", "Value2", "A1")]


def test_a_result_file_that_cannot_be_read_is_an_unknown_outcome(inproc, monkeypatch):
    app, wb = excel_book()
    runner.run(app, Batch())

    def unreadable(path):
        path.unlink()
        Path(f"{path}.out").write_text("{}", encoding="utf-8")
        return "file"
    app.fail_run = unreadable
    real = Path.read_text
    monkeypatch.setattr(Path, "read_text", lambda self, *a, **k: (_ for _ in ()).throw(PermissionError("locked"))
                        if self.name.endswith(".out") else real(self, *a, **k))
    with pytest.raises(runner.BatchError, match="outcome is unknown.*locked"):
        runner.run(app, Batch())


def test_a_job_the_add_in_could_not_claim_runs_once_over_com(inproc, caplog):
    app, wb = excel_book()
    runner.run(app, Batch())
    app.fail_run = lambda path: "not run: the job file could not be deleted: it is in use"
    b = Batch()
    b.workbook("book.xlsx").Worksheets("S").Range("A1").Value2 = "once"
    assert runner.run(app, b).where == "com"
    assert [ev for ev in app.events if ev[1] == "Value2"] == [("S", "Value2", "A1")]
    assert "could not be deleted" in caplog.text
    app.fail_run = None
    assert runner.run(app, Batch()).where == "inproc"                 # a claim that failed once is not for good


def test_a_refused_job_means_the_com_executor_for_that_excel(inproc):
    app, wb = excel_book()
    runner.run(app, Batch())
    app.fail_run = lambda path: "not run: refused: not in the jobs folder"
    assert runner.run(app, Batch()).where == "com"
    app.fail_run = None
    assert runner.run(app, Batch()).where == "com"


def test_a_run_that_fails_after_the_job_was_taken_is_an_unknown_outcome(inproc):
    app, wb = excel_book()
    runner.run(app, Batch())

    def taken_then_error(path):
        path.unlink()
        raise pywintypes.com_error(-2147352567, "Exception occurred.", None, None)
    app.fail_run = taken_then_error
    with pytest.raises(runner.BatchError, match="outcome is unknown"):
        runner.run(app, Batch())


def test_a_failure_outside_the_operations_is_not_called_before_the_first(inproc):
    app, wb = excel_book()
    runner.run(app, Batch())
    app.fail_run = lambda path: (path.unlink(), json.dumps(
        {"ok": False, "started": 2, "index": -1, "message": "x", "values": {}, "marks": {}, "ms": 1.0,
         "skipped": []}))[1]
    b = Batch()
    b.mark("a")
    b.mark("b")
    with pytest.raises(runner.BatchError) as e:
        runner.run(app, b)
    assert e.value.text == "outside the operations"


def test_job_and_result_files_are_removed_and_old_leftovers_swept(inproc):
    app, wb = excel_book()
    runner.run(app, Batch())
    folder = config.INPROC_DIR / "jobs"
    old, new = folder / ("a" * 32 + ".json"), folder / ("b" * 32 + ".json.out")
    for p in (old, new):
        p.write_text("{}", encoding="utf-8")
    os.utime(old, (time.time() - 3600, time.time() - 3600))
    b = Batch()
    b.ret(b.workbook("book.xlsx").Worksheets("S").Name, "s" * 31000)             # a long record: the result file
    runner.run(app, b)
    assert sorted(p.name for p in folder.iterdir()) == [new.name]


def test_a_batch_does_not_wait_long_for_another_jobs_load(inproc, monkeypatch, caplog):
    app, wb = excel_book()
    monkeypatch.setattr(runner, "LOCK_WAIT_S", 0.05)
    ex = runner._EXCELS.setdefault((id(app), 0.0), runner._Excel(id(app)))
    ex.lock.acquire()
    try:
        assert runner.run(app, Batch()).where == "com"
        assert runner.run(app, Batch()).where == "com"
    finally:
        ex.lock.release()
    assert len([r for r in caplog.records if r.levelname == "WARNING"]) == 1
    assert runner.run(app, Batch()).where == "inproc"


def test_a_register_xll_that_does_not_return_means_com_for_that_excel(inproc, monkeypatch, caplog):
    app, wb = excel_book()
    monkeypatch.setattr(runner, "REGISTER_LIMIT_S", 0.05)
    release = threading.Event()
    calls = []

    def blocked(path):
        calls.append(path)
        release.wait(5)
        return True
    monkeypatch.setattr(app, "RegisterXLL", blocked)
    try:
        assert runner.run(app, Batch()).where == "com"
        assert runner.run(app, Batch()).where == "com"
    finally:
        release.set()
    assert len(calls) == 1
    warnings = [r.getMessage() for r in caplog.records if r.levelname == "WARNING"]
    assert len(warnings) == 1 and "did not return" in warnings[0]


def test_a_busy_excel_during_the_load_is_tried_again_next_batch(inproc, monkeypatch):
    app, wb = excel_book()
    real, tries = app.RegisterXLL, []

    def busy_once(path):
        tries.append(path)
        if len(tries) == 1:
            raise pywintypes.com_error(-2147418111, "Call was rejected by callee.", None, None)
        return real(path)
    monkeypatch.setattr(app, "RegisterXLL", busy_once)
    assert runner.run(app, Batch()).where == "com"
    assert runner.run(app, Batch()).where == "inproc"


def test_an_unknown_excel_process_is_warned_about_once(inproc, monkeypatch, caplog):
    monkeypatch.setattr(runner, "_unknown_warned", False)
    monkeypatch.setattr(runner, "_identify", lambda app: (_ for _ in ()).throw(RuntimeError("no window")))
    app, wb = excel_book()
    for _ in range(3):
        assert runner.run(app, Batch()).where == "com"
    assert len([r for r in caplog.records if r.levelname == "WARNING"]) == 1



def test_a_result_file_that_cannot_be_deleted_keeps_the_result(inproc, monkeypatch):
    app, wb = excel_book()
    runner.run(app, Batch())
    real = Path.unlink
    held = []

    def scanner_holds(self, *a, **k):
        if self.name.endswith(".out"):
            held.append(self.name)
            raise PermissionError("in use by the virus scanner")
        return real(self, *a, **k)
    monkeypatch.setattr(Path, "unlink", scanner_holds)
    b = Batch()
    b.ret(b.workbook("book.xlsx").Worksheets("S").Name, "s" * 31000)             # a long record: the result file
    res = runner.run(app, b)
    assert res.where == "inproc" and res.values == {"s" * 31000: "S"} and held
