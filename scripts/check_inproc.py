r"""Live check of xl/inproc (light and fast phase 3b, Task 1): the add-in is built and loaded in a hidden Excel of this
script's own, and a battery of batches runs through BOTH executors - the COM executor and the in-process add-in - on
fresh workbooks. Each check compares what came back (values with their types, the error's operation, the skipped
operations) and the workbooks afterwards, read with plain COM (sheet order, cell values, formulas via Formula2,
number formats, fonts, fills, column widths, tables, defined names, links), and that the four quiet settings are as
they were before. Then: the ping, two builds loaded side by side, broken builds (a .dna pointing at a missing dll, a
dll with no functions: no window in Excel, the COM executor takes over), the add-in's guards (a formula calling it,
paths outside its jobs folder, a job file it cannot delete: nothing read, deleted or run), and the fixed cost of an
empty batch.

The Excel is one this script starts itself (DispatchEx, refused unless its pid is new; check_fastcom.start_excel) and
quits at the end; if Quit hangs, only that process is killed. The user's Excel is never touched. Nothing is saved: the
workbooks are new ones closed unsaved, the source files are written with openpyxl in a temporary folder, and
SMARTASS_HOME points at that folder too, so the builds, markers and job files never reach the real home.

Prints OK / DIFF per check with the milliseconds per executor; exit 0 only when all are OK.

Usage (from the repo root):
    engine\.venv\Scripts\python.exe scripts\check_inproc.py [--runs 30]
"""
import argparse
import datetime as dt
import json
import logging
import os
import shutil
import statistics
import sys
import tempfile
import threading
import time
import uuid
from pathlib import Path

TMP = Path(tempfile.mkdtemp(prefix="check_inproc_"))
os.environ["SMARTASS_HOME"] = str(TMP / "home")          # before smartass.core.config is imported

import pythoncom  # noqa: E402
import pywintypes  # noqa: E402

from check_fastcom import QUIT_WAIT_S, _wait_gone, kill_if_ours, start_excel  # noqa: E402
from smartass.core import config  # noqa: E402
from smartass.xl import fastcom, jobs  # noqa: E402
from smartass.xl.hostproc import windows_of  # noqa: E402
from smartass.xl.inproc import build, runner  # noqa: E402
from smartass.xl.inproc.batch import MISSING, Batch  # noqa: E402

AUTOMATIC, MANUAL = -4105, -4135
USER = {"ScreenUpdating": True, "EnableEvents": True, "DisplayAlerts": False, "Calculation": AUTOMATIC}
AWKWARD = (0.1 + 0.2, 1 / 3, 1e22, 3.0, -0.0, 2.5e-320)
LONG = "x" * 29_992 + "æøå\"'end"                         # 30 000 characters
REGISTER_WAIT_S = 60
WINDOW_WAIT_S = 3
results: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> bool:
    results.append((name, ok, detail))
    print(f"{'OK  ' if ok else 'DIFF'} {name}{'  ' + detail if detail else ''}", flush=True)
    return ok


# ------------------------------------------------------------------ the battery
def item_values(b, c):
    ws = b.workbook(c["book"]).Worksheets(1)
    ws.Range("A1:F3").Value2 = ((None, True, False, 0, -7, 2 ** 40), AWKWARD,
                                ('say "hi"', "æøå ÆØÅ € \U0001f600", "", "=1+1", "'=not a formula", 1.5))
    ws.Range("A5").Value2 = LONG
    ws.Range("A6").Formula2 = "=1/0"
    ws.Range("B6").Formula2 = "=NA()"
    ws.Range("A7").Value2 = 46053.5
    ws.Range("A7").NumberFormat = "dd.mm.yyyy hh:mm"
    ws.Range("A8").Value2 = 12.34
    ws.Range("A8").Style = "Currency"
    ws.Range("C8").Value2 = 7
    b.ret(ws.Range("A1:F3").Value2, "grid")
    b.ret(ws.Range("A1:F3").Formula2, "formulas")
    b.ret(ws.Range("B1").Value2, "one")
    b.ret(ws.Range("H20").Value2, "empty")
    b.ret(ws.Range("A5").Value2, "long")
    b.ret(ws.Range("A6:B6").Value2, "errors")
    b.ret(ws.Range("A7").Value2, "date as Value2")
    b.ret(ws.Range("A8").Value, "currency")
    b.ret(ws.Range("C8").Value2, "int written")
    b.ret(ws.Range("A1:A2").Value2, "column")


def item_long_inline(b, c):
    """A result record of about 20 000 characters: Run's return value, not the result file."""
    ws = b.workbook(c["book"]).Worksheets(1)
    ws.Range("A1").Value2 = "y" * 20_000
    b.ret(ws.Range("A1").Value2, "twenty thousand")


def item_optional_arguments(b, c):
    """Members whose arguments are all optional, called with arguments: the arguments go to the member itself (pywin32's
    attribute call would read the member first and call what it gave: Offset(2, 3) from B2 would be $D$3)."""
    ws = b.workbook(c["book"]).Worksheets(1)
    ws.Range("B2").Value2 = "hello"
    b.ret(ws.Range("B2").Offset(2, 3).Address, "Offset")
    b.ret(ws.Range("B2").Resize(2, 3).Address, "Resize")
    b.ret(ws.Range("A1:C3").Address(False, False), "Address")
    b.ret(b.app.International(5), "International(5)")
    b.ret(ws.Range("B2").Value(10), "Value(10)")
    b.ret(ws.Range("B2").Characters(1, 2).Text, "Characters")
    b.ret(ws.Cells(2, 3).Address, "Cells")
    b.ret(ws.Rows(4).Address, "Rows")
    b.ret(ws.Range("A1:C3").Item(2, 2).Address, "Item")
    b.ret(ws.Range("A1:B2").Areas(1).Address, "Areas")
    b.ret(ws.Range("B9").End(-4162).Row, "End")
    b.ret(ws.Range("A1:C9").Find("hello").Address, "Find")


OPTIONAL_EXPECTED = {"Offset": "$E$4", "Resize": "$B$2:$D$3", "Address": "A1:C3", "Value(10)": "hello",
                     "Characters": "he", "Cells": "$C$2", "Rows": "$4:$4", "Item": "$B$2", "Areas": "$A$1:$B$2",
                     "End": 2, "Find": "$B$2"}


def item_pywin32_alias(b, c):
    """pywin32's GetAddress is not an Excel name: it fails in both executors, at the same operation."""
    b.ret(b.workbook(c["book"]).Worksheets(1).Range("A1:C3").GetAddress(False, False), "address")


def item_text_to_cells(b, c):
    """Strings Excel parses on the way in (the locale of the call matters here)."""
    ws = b.workbook(c["book"]).Worksheets(1)
    ws.Range("A1:H1").Value2 = (("1.5", "1,5", "31.01.2026", "2026-01-31", "1/2/2026", "TRUE", "1 234", "12%"),)
    ws.Range("A2:B2").Formula2 = (("=SUM(1.5,2)", '=TEXT(0.5,"0.0")'),)
    ws.Range("A3").Value = "3,25"
    b.ret(ws.Range("A1:H3").Value2, "cells")


def item_number_formats(b, c):
    """A NumberFormat is read in the user's codes under the LCID both executors send (0): an en-US code that is not
    valid there fails, at the same operation."""
    ws = b.workbook(c["book"]).Worksheets(1)
    ws.Range("A1:A2").Value2 = ((1234.5,), (0.25,))
    ws.Range("A2").NumberFormat = "0%"
    b.ret(ws.Range("A2").NumberFormat, "percent")
    ws.Range("A1").NumberFormat = "#,##0.00"
    b.ret(ws.Range("A1").NumberFormat, "never")


def item_operations(b, c):
    wbn = b.workbook(c["book"])
    wss = wbn.Worksheets
    ws = wss(1)
    ws.Name = "Q1 Data"
    ws.Range("A1:C4").Value2 = (("Item", "Value", "Share"), ("a", 1.0, 0.25), ("b", 2.0, 0.5), ("c", 3.0, 0.25))
    lo = ws.ListObjects.Add(1, ws.Range("A1:C4"), MISSING, 1)
    lo.Name = "Data_Q1"
    lo.TableStyle = "TableStyleLight1"
    lo.ShowTableStyleRowStripes = False
    last = wss.Item(wss.Count)
    ans = wss.Add(MISSING, last)
    ans.Name = "Q1 Answer"
    ans.Range("B2").Formula2 = "=SUM(Data_Q1[Value])"
    ans.Range("B3:C3").Formula2 = (("=B2*2", "=AVERAGE(Data_Q1[Share])"),)
    ans.Range("B2:C3").NumberFormat = "0"
    font = ans.Range("B2").Font
    font.Bold = True
    font.Size = 14
    font.Color = 255
    font.Name = "Arial"
    ans.Range("B3").Interior.Color = 0xEEEEEE
    ans.Columns("B").ColumnWidth = 20
    ans.Range("C:C").ColumnWidth = 7.5
    wbn.Names.Add("Answer_Total", "='Q1 Answer'!$B$2")
    ans.Calculate()
    b.mark("built")
    ans.Move(wss("Q1 Data"))
    b.ret(ans.Range("B2").Value2, "total")
    b.ret(ans.Range("B3:C3").Value2, "row")
    b.ret(ans.UsedRange.Address, "used")
    b.ret(lo.Range.Address, "table")
    b.ret(wss.Count, "sheets")
    b.ret(wbn.ActiveSheet.Name, "active")


def item_copy(b, c):
    """Worksheets([...]).Copy between workbooks, with MISSING and with None for Before."""
    src = b.app.Workbooks.Open(c["src1"], 0, True)
    with b.always():
        src.Close(False)
    tgt = b.workbook(c["book"])
    src.Worksheets(["Src A", "Src B"]).Copy(MISSING, tgt.Worksheets(1))
    src.Worksheets("Other").Copy(None, tgt.Worksheets(tgt.Worksheets.Count))
    b.ret(tgt.Worksheets.Count, "count")
    b.ret(tgt.LinkSources(1), "links")


def item_fails_midway(b, c):
    ws = b.workbook(c["book"]).Worksheets(1)
    ws.Range("A1").Value2 = "before"
    b.ret(b.app.DisplayAlerts, "alerts inside")
    b.workbook(c["book"]).Worksheets("Nope").Range("A1").Value2 = "never"
    ws.Range("A2").Value2 = "never either"
    with b.always():
        ws.Range("A3").Value2 = "finally"
        ws.Range("ZZZ0").Value2 = "a finally operation that fails"
        ws.Range("A4").Value2 = "finally goes on"


def item_try(b, c):
    ws = b.workbook(c["book"]).Worksheets(1)
    with b.attempt():
        ws.ListObjects("Nope").Delete()
    ws.Range("A1").Value2 = "went on"
    b.ret(ws.Range("A1").Value2, "a1")


def item_com_object(b, c):
    b.ret(b.workbook(c["book"]).Worksheets(1).Range("A1"), "range")


def item_date(b, c):
    ws = b.workbook(c["book"]).Worksheets(1)
    ws.Range("A1").Value2 = 46053
    ws.Range("A1").NumberFormat = "dd.mm.yyyy"
    b.ret(ws.Range("A1").Value, "date")


def item_not_quiet(b, c):
    b.ret(b.app.ScreenUpdating, "screen")
    b.ret(b.app.Calculation, "calc")


BATTERY = [  # name, record, quiet, the user's settings before[, the values Excel gives]
    ("values in and out", item_values, True, USER),
    ("optional arguments go to the member", item_optional_arguments, True, USER, OPTIONAL_EXPECTED),
    ("a pywin32 alias is no Excel name", item_pywin32_alias, True, USER),
    ("a long result record inline", item_long_inline, True, USER),
    ("text parsed into cells", item_text_to_cells, True, USER),
    ("every operation kind", item_operations, True, USER),
    ("number formats", item_number_formats, True, USER),
    ("copy between workbooks (MISSING, None)", item_copy, True, USER),
    ("failing mid-batch", item_fails_midway, True,
     {"ScreenUpdating": True, "EnableEvents": False, "DisplayAlerts": True, "Calculation": MANUAL}),
    ("a try operation", item_try, True, USER),
    ("a COM object asked back", item_com_object, True, USER),
    ("a date asked back", item_date, True, USER),
    ("quiet=False", item_not_quiet, False, USER),
]


# ------------------------------------------------------------------ comparing
def typed(v):
    """v with each leaf's type, so 3 and 3.0 differ."""
    if isinstance(v, (tuple, list)):
        return (type(v).__name__, tuple(typed(x) for x in v))
    if isinstance(v, dict):
        return {k: typed(x) for k, x in v.items()}
    return (type(v).__name__, repr(v))


def settings(raw) -> dict:
    return {k: getattr(raw, k) for k in USER}


def put(raw, values: dict) -> None:
    for k, v in values.items():
        setattr(raw, k, v)


def _try(fn):
    try:
        return fn()
    except pywintypes.com_error as e:
        return ("com_error", e.hresult)


def sheet_picture(wb) -> dict:
    """The workbook as plain COM reads it (independent of both executors)."""
    sheets = []
    for ws in wb.Worksheets:
        used = ws.UsedRange
        cells = []
        for cell in used.Cells:
            font = cell.Font
            cells.append((cell.Address, typed(cell.Value2), _try(lambda c=cell: c.Formula2), cell.NumberFormat,
                          font.Name, font.Size, font.Bold, font.Color, cell.Interior.Color))
        last_col = used.Column + used.Columns.Count
        widths = [ws.Columns(i).ColumnWidth for i in range(1, last_col + 1)]
        tables = []
        for lo in ws.ListObjects:
            style = _try(lambda lo=lo: lo.TableStyle.Name)
            tables.append((lo.Name, lo.Range.Address, style, lo.ShowTableStyleRowStripes))
        sheets.append((ws.Name, ws.Visible, used.Address, cells, widths, tables))
    names = sorted((n.Name, n.RefersTo) for n in wb.Names)
    links = tuple(Path(str(x)).name for x in (wb.LinkSources(1) or ()))
    return {"sheets": sheets, "names": names, "links": links}


def run_once(app, raw, executor: str, record, quiet: bool, user: dict, ctx_of) -> dict:
    """One battery item through one executor on fresh workbooks: what came back, the workbooks, the settings."""
    ctx = ctx_of()
    b = Batch(quiet=quiet)
    record(b, ctx)
    config.INPROC = executor == "inproc"
    put(raw, user)
    t0 = time.perf_counter()
    try:
        res = runner.run(app, b)
        got = {"where": res.where, "values": typed(res.values), "marks": list(res.marks), "raw": res.values,
               "skipped": [(i, t.replace(ctx["book"], "<book>")) for i, t, _ in res.skipped]}
    except runner.BatchError as e:
        got = {"where": e.where, "error": (e.index, e.text.replace(ctx["book"], "<book>")),
               "skipped": [(i, t.replace(ctx["book"], "<book>")) for i, t, _ in e.skipped], "message": e.message}
    ms = (time.perf_counter() - t0) * 1000
    got["settings after"] = settings(raw)
    got["book"] = sheet_picture(raw.Workbooks(ctx["book"]))
    raw.Workbooks(ctx["book"]).Close(False)
    got["ms"] = ms
    return got


def battery(app, raw, sources) -> None:
    def fresh():
        return {"book": str(raw.Workbooks.Add().Name), **sources}
    print(f"\n{'battery item':<44} {'':4} {'COM ms':>9} {'in-process ms':>14}")
    for name, record, quiet, user, *expected in BATTERY:
        a = run_once(app, raw, "com", record, quiet, user, fresh)
        b = run_once(app, raw, "inproc", record, quiet, user, fresh)
        diffs = [k for k in ("values", "marks", "skipped", "error", "book") if a.get(k) != b.get(k)]
        for k, v in (expected[0] if expected else {}).items():
            if a.get("raw", {}).get(k) != v:
                diffs.append(f"{k}: Excel gives {v!r}, the COM executor {a.get('raw', {}).get(k)!r}")
        if a["where"] != "com" or b["where"] != "inproc":
            diffs.append(f"where {a['where']}/{b['where']}")
        for side in (a, b):
            if side["settings after"] != user:
                diffs.append(f"settings after ({side['where']}): {side['settings after']}")
        if "error" in a and "message" in a:
            for side in (a, b):
                if not side.get("message"):
                    diffs.append("no message")
        ok = not diffs
        detail = f"{a['ms']:9.1f} {b['ms']:14.1f}"
        if "error" in a:
            detail += f"   error at {a['error'][0]}: {a['error'][1]} | com: {a['message'][:60]} | inproc: {b.get('message', '')[:60]}"
        check(f"{name:<44}", ok, detail)
        for k in diffs:
            if k in a or k in b:
                print(f"     {k}: com {str(a.get(k))[:600]}\n     {' ' * len(k)}  inproc {str(b.get(k))[:600]}")
            else:
                print(f"     {k}")


# ------------------------------------------------------------------ the add-in itself
def register(raw, pid: int, ctime: float, path: Path):
    """RegisterXLL with a watchdog (a blocked call means a dialog: our Excel is killed and the check fails)."""
    dog = threading.Timer(REGISTER_WAIT_S, lambda: kill_if_ours(pid, ctime))
    dog.start()
    try:
        return raw.RegisterXLL(str(path))
    except pywintypes.com_error as e:
        return ("com_error", e.hresult)
    finally:
        dog.cancel()


def new_windows(pid: int, before: set) -> list:
    """Visible top-level windows of our Excel that were not there before, watched for WINDOW_WAIT_S."""
    t0, seen = time.monotonic(), []
    while time.monotonic() - t0 < WINDOW_WAIT_S:
        pythoncom.PumpWaitingMessages()
        seen = [w for w in windows_of(pid) if w not in before]
        if seen:
            break
        time.sleep(0.1)
    return seen


def broken_build(name: str, dll: str | None, config_file: bool) -> Path:
    folder = TMP / "broken" / name
    folder.mkdir(parents=True)
    shutil.copyfile(build.XLL, folder / f"{name}.xll")
    (folder / f"{name}.dna").write_text(build.DNA.replace("{name}.dll", dll or "missing.dll").format(name=name),
                                        encoding="utf-8")
    if config_file:
        (folder / f"{name}.xll.config").write_text(build.CONFIG, encoding="utf-8")
    return folder / f"{name}.xll"


def compile_empty(folder: Path) -> str:
    cs = folder / "Empty.cs"
    cs.write_text("public static class NothingHere { }\n", encoding="utf-8")
    code, text = build._compile([str(build.CSC), "-nologo", "-target:library", f"-out:{folder / 'empty.dll'}",
                                 str(cs)])
    if code:
        raise SystemExit(f"could not compile the empty dll: {text}")
    return "empty.dll"


class Warnings(logging.Handler):
    def __init__(self):
        super().__init__(logging.WARNING)
        self.seen = []

    def emit(self, record):
        self.seen.append(record.getMessage())


def addin_checks(app, raw, pid: int, ctime: float, ex) -> None:
    print()
    t0 = time.perf_counter()
    got = raw.Run(ex.ping_macro)
    check("ping", got == ex.hash, f"{got!r} in {(time.perf_counter() - t0) * 1000:.1f} ms")

    # a second build (another hash) beside the first: both answer; the first is not pruned while loaded
    first = Path(ex.xll)
    alt = TMP / "Runner_alt.cs"
    alt.write_text(build.RUNNER_CS.read_text(encoding="utf-8") + "\n// a second build\n", encoding="utf-8")
    real_cs = build.RUNNER_CS
    build.RUNNER_CS = alt
    try:
        xll2, h2 = build.ensure()
    finally:
        build.RUNNER_CS = real_cs
    ok2 = register(raw, pid, ctime, xll2)
    ex2 = runner._Excel(pid, ready=True, hash=h2, xll=str(xll2))
    b = Batch()
    b.ret(b.app.Workbooks.Count, "books")
    out = runner._run_inproc(app, ex2, b.to_dict(), 30)
    check("two builds side by side", ok2 is True and h2 != ex.hash and raw.Run(ex2.ping_macro) == h2
          and raw.Run(ex.ping_macro) == ex.hash and out is not None and out["ok"] and build._complete(first.parent),
          f"{ex.hash} and {h2}; the first build kept while loaded: {build._complete(first.parent)}")

    # broken builds: no window in Excel, and the runner falls back to the COM executor with one warning
    empty_dir = TMP / "empty"
    empty_dir.mkdir()
    empty = compile_empty(empty_dir)
    for name, dll in (("broken_missing_dll", None), ("broken_no_functions", empty)):
        xll = broken_build(name, dll, config_file=True)
        if dll:
            shutil.copyfile(empty_dir / dll, xll.parent / dll)
        before = set(windows_of(pid))
        loaded = register(raw, pid, ctime, xll)
        shown = new_windows(pid, before)
        handler = Warnings()
        logging.getLogger("smartview.inproc").addHandler(handler)
        real_ensure, real_state = build.ensure, dict(runner._EXCELS)
        build.ensure = lambda xll=xll: (xll, "0badc0de")
        runner._EXCELS.clear()
        try:
            b = Batch()
            b.ret(b.app.Workbooks.Count, "books")
            res = runner.run(app, b)
            res2 = runner.run(app, b)
        finally:
            build.ensure = real_ensure
            runner._EXCELS.clear()
            runner._EXCELS.update(real_state)
            logging.getLogger("smartview.inproc").removeHandler(handler)
        check(f"{name}: no window, COM executor", not shown and res.where == "com" and res2.where == "com"
              and len(handler.seen) == 1,
              f"RegisterXLL -> {loaded!r}; windows {shown}; warnings {handler.seen}")


def guard_checks(app, raw, ex) -> None:
    """The add-in runs nothing but a job file of its jobs folder, given by Application.Run: a formula calling it, a
    path outside the folder, a wrong name, ".." and other path tricks, a missing file and a job file it cannot delete
    are refused, and nothing is deleted or run."""
    print()
    wb = raw.Workbooks.Add()
    name = str(wb.Name)
    target = wb.Worksheets(1).Range("Z1")
    folder = build.jobs_folder()
    folder.mkdir(parents=True, exist_ok=True)

    def job_file(where: Path) -> Path:
        b = Batch()
        b.workbook(name).Worksheets(1).Range("Z1").Value2 = "ran"
        where.parent.mkdir(parents=True, exist_ok=True)
        where.write_text(json.dumps(runner._wire(b.to_dict())), encoding="utf-8")
        return where

    hexname = uuid.uuid4().hex + ".json"
    real = job_file(folder / hexname)
    cell = wb.Worksheets(1).Range("A1")
    cell.Formula = f'=SmartassRun_{ex.hash}("{real}")'
    shown = cell.Value2
    check("a formula calling the add-in runs nothing", real.exists() and target.Value2 is None
          and shown == -2146826246, f"the cell shows {shown!r} (#N/A is -2146826246); the job file is still there")
    cell.Clear()

    outside = job_file(TMP / "outside" / (uuid.uuid4().hex + ".json"))
    wrong = job_file(folder / "abc.json")
    tricks = {
        "outside the folder": str(outside),
        "a wrong name": str(wrong),
        "a .. path": str(folder / ".." / folder.name / hexname),
        "a ..\\ name": str(folder) + "\\..\\" + folder.name + "\\" + hexname,
        "forward slashes": str(real).replace("\\", "/"),
        "a network path": "\\\\localhost\\" + str(real).replace(":", "$", 1),
        "a stream": str(real) + ":x",
        "a relative path": hexname,
        "a missing file": str(folder / (uuid.uuid4().hex + ".json")),
    }
    for what, path in tricks.items():
        got = raw.Run(ex.run_macro, path)
        check(f"refused: {what}", isinstance(got, str) and got.startswith("not run:") and real.exists()
              and outside.exists() and wrong.exists() and target.Value2 is None, f"{path} -> {got!r}")

    held = open(real, "rb")                            # no FILE_SHARE_DELETE: the add-in cannot delete it
    try:
        got = raw.Run(ex.run_macro, str(real))
    finally:
        held.close()
    check("a job file the add-in cannot delete is not run", isinstance(got, str) and "could not be deleted" in got
          and real.exists() and target.Value2 is None, f"-> {got!r}")
    got = raw.Run(ex.run_macro, str(real))
    check("the same job file once it can be claimed runs", isinstance(got, str) and got.startswith("{")
          and not real.exists() and target.Value2 == "ran", f"-> {got[:60]!r}")
    for p in (outside, wrong):
        p.unlink(missing_ok=True)
    wb.Close(False)


def detector_check(raw, pid: int, ctime: float) -> None:
    """Without the .xll.config the same missing-dll build shows Excel-DNA's log window: the window watch can see one."""
    xll = broken_build("broken_without_config", None, config_file=False)
    before = set(windows_of(pid))
    loaded = register(raw, pid, ctime, xll)
    shown = new_windows(pid, before)
    print(f"INFO detector: without the config, RegisterXLL -> {loaded!r}, windows {shown or 'none'}"
          f" ({'the window watch sees it' if shown else 'no window to see'})")


def fixed_cost(app, raw, runs: int) -> None:
    print()
    wb = raw.Workbooks.Add()
    try:
        for quiet in (False, True):
            line = []
            for executor in ("com", "inproc"):
                config.INPROC = executor == "inproc"
                put(raw, USER)
                ts, inside = [], []
                for _ in range(runs):
                    t0 = time.perf_counter()
                    res = runner.run(app, Batch(quiet=quiet))
                    ts.append((time.perf_counter() - t0) * 1000)
                    inside.append(res.ms)
                line.append(f"{executor} {statistics.median(ts):6.2f} ms (inside {statistics.median(inside):5.2f})")
            print(f"empty batch, quiet={quiet!s:<5}: " + "   ".join(line))
    finally:
        wb.Close(False)


def write_sources() -> dict:
    import openpyxl
    out = {}
    for key in ("src1",):
        book = openpyxl.Workbook()
        a = book.active
        a.title = "Src A"
        a["A1"], a["B1"] = "label", "=Other!A1*2"
        book.create_sheet("Src B")["A1"] = "plain"
        book.create_sheet("Other")["A1"] = 21
        path = TMP / f"inproc_{key}.xlsx"
        book.save(path)
        out[key] = str(path)
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--runs", type=int, default=30, help="empty batches per executor for the fixed cost (median)")
    args = ap.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")
    logging.basicConfig(level=logging.WARNING, format="  log %(levelname)s %(message)s")
    sources = write_sources()
    pythoncom.CoInitialize()
    jobs.bind_job()
    raw, pid, ctime = start_excel()
    print(f"check_inproc {dt.datetime.now():%Y-%m-%d %H:%M}: hidden Excel pid {pid}, home {config.HOME}")
    try:
        app = fastcom.wrap(raw)
        config.INPROC = True
        t0 = time.perf_counter()
        ok = runner.available(app)
        ex = next(iter(runner._EXCELS.values()), None)
        check("build and load", ok and ex is not None and ex.ready,
              f"{(time.perf_counter() - t0) * 1000:.0f} ms (build + RegisterXLL + ping): {ex and ex.xll}")
        if not ok:
            return 1
        battery(app, raw, sources)
        fixed_cost(app, raw, args.runs)
        guard_checks(app, raw, ex)
        addin_checks(app, raw, pid, ctime, ex)
        detector_check(raw, pid, ctime)
    finally:
        dog = threading.Timer(QUIT_WAIT_S, lambda: kill_if_ours(pid, ctime))
        dog.start()
        try:
            for wb in list(raw.Workbooks):
                wb.Close(False)
            raw.Quit()
        except pywintypes.com_error:
            pass
        finally:
            dog.cancel()
        raw = app = None
        runner._EXCELS.clear()
        runner._last.__dict__.clear()
        fastcom._IFACE.clear()
        _wait_gone(pid, ctime)
        kill_if_ours(pid, ctime)
        shutil.rmtree(TMP, ignore_errors=True)
    failed = [n for n, ok, _ in results if not ok]
    print(f"\n{len(results) - len(failed)}/{len(results)} checks OK" + (f"; DIFF: {', '.join(failed)}" if failed else ""))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
