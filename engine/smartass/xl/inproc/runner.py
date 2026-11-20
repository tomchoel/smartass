"""Runs a Batch (batch.py) in the user's Excel: in Excel's own process through the add-in (one Application.Run), or
over COM from Python (the COM executor) - same operations, same values back, same workbook afterwards.

The in-process executor writes the batch as JSON (floats as their 64 bits in hex: exact both ways) to a job file and
calls the add-in's SmartassRun_<hash>(job path) with one Application.Run; the add-in deletes the job file once it has
read it, runs the operations inside Excel and returns the result record as Run's value, or writes it to
<job path>.out when it is too long for an Excel string (and returns "file"). The add-in is built on first use (build.py)
and loaded once per Excel process (RegisterXLL, then a ping that must answer the build's hash); a Run that finds the
macro gone registers once more. Anything that keeps it from loading or answering (SV_INPROC=0, a 32-bit Excel, no
compiler, a compile error, RegisterXLL False or failing, a wrong ping) means the COM executor for that Excel process,
with one warning. A batch falls back to the COM executor only when nothing of it ran: the job file is still there,
or the add-in's result says it failed before its first operation. Once the add-in has taken the job, a failure is a
BatchError and the batch is never run again.

The COM executor (execute) interprets the batch's plain data over the COM objects, inside one jobs.com. A busy Excel
is waited for on the rejected operation only (a rejected call did not run), never by running the batch again. A call
with arguments is made as the add-in makes it (IDispatch::Invoke, LCID 0, method or property get, the arguments on
the member itself): pywin32's obj.Name(args) reads a property whose arguments are all optional first and calls what
that gave, so Range.Offset(2, 3) would be the default member of the cell itself ($D$3 from B2, not $E$4) and
Address(False, False) would fail; pywin32's own aliases (GetAddress) are unknown names in both executors.

Both executors, around the operations: with quiet, ScreenUpdating, EnableEvents and DisplayAlerts off and Calculation
manual (each only when it is not so already), put back in reverse order afterwards, also when an operation failed.
An operation that fails stops the batch: the finally operations run, the settings come back, and run raises
BatchError(index, text, message). A failure of a try operation or of a finally operation is reported in
Result.skipped (or BatchError.skipped) and the batch goes on. An object an operation would act on that was never made
(its operation failed or did not run) is such a failure. Both produce the same result record (the add-in returns it
as JSON): ok, started (operations begun), index and message, values, marks, ms, skipped, and the settings that could
not be put back (logged as warnings)."""
from __future__ import annotations

import datetime
import json
import logging
import re
import struct
import threading
import time
import uuid
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

import pythoncom
import pywintypes
from win32com.client import dynamic

from ...core import config
from .. import fastcom, jobs
from ..host import HostError
from ..hostproc import is_busy, retry_busy
from . import build

log = logging.getLogger("smartview.inproc")

MANUAL = -4135                          # xlCalculationManual
QUIET = (("ScreenUpdating", False), ("EnableEvents", False), ("DisplayAlerts", False), ("Calculation", MANUAL))
LINKS = 1                               # xlExcelLinks (LinkSources, BreakLink)
RESTORE_LIMIT_S = 30
LOAD_LIMIT_S = 60                       # RegisterXLL and the ping: how long a busy Excel is waited for
OMITTED = pythoncom.ArgNotFound         # MISSING: sent as DISP_E_PARAMNOTFOUND, as C#'s Type.Missing
UNBOUND = "no object {}: the operation that makes it failed or did not run"
NOT_A_VALUE = "a COM object cannot be asked back"
NOT_A_DATE = "a date cannot be asked back (read Value2)"
METHOD_OR_GET = pythoncom.DISPATCH_METHOD | pythoncom.DISPATCH_PROPERTYGET


@dataclass
class Result:
    values: dict = field(default_factory=dict)            # ret name -> value
    marks: dict = field(default_factory=dict)             # mark name -> milliseconds since the batch started
    ms: float = 0.0                                       # the batch inside its executor
    skipped: list = field(default_factory=list)           # (index, text, message) of failed try / finally operations
    where: str = "com"                                    # "inproc" or "com"


class BatchError(Exception):
    """An operation of a batch failed: its index (None: unknown), its text as in code, Excel's message, the failed
    try / finally operations, and where it ran ("inproc" or "com"). The finally operations have run and the settings
    are back."""

    def __init__(self, index: int | None, text: str, message: str, skipped: list | None = None, where: str = "com"):
        super().__init__(f"Excel operation {index} failed: {text}: {message}")
        self.index, self.text, self.message, self.skipped = index, text, message, list(skipped or [])
        self.where = where


def run(app, batch, limit: float = 120) -> Result:
    """Run batch in app's Excel, in its process when the add-in is there, else over COM; Result, or BatchError when
    an operation failed. A busy Excel (beyond limit) or an abandoned job raises HostError, as jobs.com does."""
    job = batch.to_dict()
    ex = _ready(app)
    if ex is not None:
        out = _run_inproc(app, ex, job, limit)
        if out is not None:
            return _result(out, batch, "inproc")
    out = jobs.com(lambda: execute(app, job, limit), limit=limit)
    return _result(out, batch, "com")


def available(app) -> bool:
    """The add-in is loaded in app's Excel (loading it the first time this Excel process is asked)."""
    return _ready(app) is not None


def _result(out: dict, batch, where: str) -> Result:
    for text in out.get("unrestored", []):
        log.warning("could not restore Application.%s", text)
    skipped = [(i, batch.text(i), m) for i, m in out.get("skipped", [])]
    if not out["ok"]:
        i = out.get("index")
        text = batch.text(i) if i is not None and i >= 0 else "before the first operation"
        raise BatchError(i, text, out.get("message") or "", skipped, where)
    return Result({k: _tuples(v) for k, v in out["values"].items()}, dict(out["marks"]), float(out["ms"]), skipped,
                  where)


def _tuples(v):
    return tuple(_tuples(x) for x in v) if isinstance(v, (list, tuple)) else v


# ------------------------------------------------------------------ the COM executor
class _Failed(Exception):
    """An operation's failure, with the message both executors give."""


def execute(app, job: dict, limit: float = 120) -> dict:
    """The batch's operations over COM in this thread; the result record (as the add-in writes it). Raises only
    JobAborted (after the finally operations and the restore)."""
    t0 = time.perf_counter()
    refs: dict[int, object] = {0: app}
    values: dict = {}
    marks: dict = {}
    skipped: list = []
    out: dict = {"ok": True}
    started, changed, unrestored, aborted = 0, [], [], None
    ops, final = job["ops"], job.get("finally", [])
    try:
        if job.get("quiet"):
            _quiet(app, changed, limit)
        try:
            for i, op in enumerate(ops):
                started = i + 1
                try:
                    jobs._check_alive()
                    retry_busy(lambda op=op: _do(app, op, refs, values, marks, t0), limit)
                except jobs.JobAborted as e:
                    aborted = e
                    break
                except Exception as e:   # noqa: BLE001 - the failure is the result
                    if op.get("try"):
                        skipped.append([i, _message(e)])
                        continue
                    out = {"ok": False, "index": i, "message": _message(e)}
                    break
        finally:
            for j, op in enumerate(final):
                try:
                    retry_busy(lambda op=op: _do(app, op, refs, values, marks, t0), limit)
                except Exception as e:   # noqa: BLE001
                    skipped.append([len(ops) + j, _message(e)])
    except Exception as e:   # noqa: BLE001 - the settings could not be read or set: nothing ran
        out = {"ok": False, "index": -1, "message": _message(e)}
    finally:
        unrestored = _restore(app, changed)
    if aborted is not None:
        for text in unrestored:
            log.warning("could not restore Application.%s", text)
        raise aborted
    out.update(started=started, values=values, marks=marks, ms=(time.perf_counter() - t0) * 1000, skipped=skipped)
    if unrestored:
        out["unrestored"] = unrestored
    return out


def _quiet(app, changed: list, limit: float) -> None:
    """Each quiet setting that is not so already is set; changed gets (name, the user's value) for the restore.
    Calculation cannot be read or set with no workbook open: then it is left alone."""
    for name, value in QUIET:
        try:
            before = retry_busy(lambda name=name: getattr(app, name), limit)
            if before != value:
                retry_busy(lambda name=name, value=value: setattr(app, name, value), limit)
                changed.append((name, before))
        except Exception:   # noqa: BLE001
            if name != "Calculation":
                raise


def _restore(app, changed: list) -> list[str]:
    """Each changed setting put back, the last first; "name=value: message" of each that could not be."""
    failed = []
    for name, before in reversed(changed):
        try:
            retry_busy(lambda name=name, before=before: setattr(app, name, before), RESTORE_LIMIT_S)
        except Exception as e:   # noqa: BLE001
            failed.append(f"{name}={before}: {_message(e)}")
    return failed


def _do(app, op: dict, refs: dict, values: dict, marks: dict, t0: float) -> None:
    k = op["k"]
    if k == "mark":
        marks[op["name"]] = (time.perf_counter() - t0) * 1000
    elif k == "ret":
        values[op["name"]] = _plain(_ref(refs, op["on"]))
    elif k == "set":
        setattr(_ref(refs, op["on"]), op["name"], _arg(op["value"], refs))
    elif k == "get":
        refs[op["as"]] = getattr(_ref(refs, op["on"]), op["name"])
    elif k == "call":
        refs[op["as"]] = _call(_ref(refs, op["on"]), op["name"], [_arg(a, refs) for a in op["args"]])
    elif k == "verb":
        refs[op["as"]] = VERBS[op["name"]](*(_arg(a, refs) for a in op["args"]))
    else:
        raise _Failed(f"unknown operation {k!r}")


_DISPIDS: dict[tuple, int] = {}         # (fastcom's shared interface, name) -> dispid, learnt on first use


def _call(obj, name: str, args: list):
    """obj.name(*args) as the add-in calls it on a COM object (see the module docstring); a plain call on a fake.
    The dispid is asked once per interface and name (a fastcom object knows its interface; with SV_FASTCOM=0 it is
    asked every time)."""
    if not isinstance(obj, dynamic.CDispatch):
        return getattr(obj, name)(*args)
    ole = obj._oleobj_
    iface = obj.__dict__.get("_iface_")
    key = (iface, name) if iface is not None else None
    dispid = _DISPIDS.get(key) if key else None
    if dispid is None:
        dispid = ole.GetIDsOfNames(name)
        if key:
            _DISPIDS[key] = dispid
    return fastcom.wrap(ole.Invoke(dispid, 0, METHOD_OR_GET, True, *args))


def _ref(refs: dict, i: int):
    if i not in refs:
        raise _Failed(UNBOUND.format(i))
    return refs[i]


def _arg(v, refs: dict):
    """A value of the batch's plain data as pywin32 sends it (None: VT_NULL; a list: a 1-D array; a grid: 2-D)."""
    if isinstance(v, dict):
        if "r" in v:
            return _ref(refs, v["r"])
        if "d" in v:
            return _double(v["d"])
        if "grid" in v:
            return tuple(tuple(_arg(c, refs) for c in row) for row in v["grid"])
        if "list" in v:
            return [_arg(c, refs) for c in v["list"]]
        if "missing" in v:
            return OMITTED
        raise _Failed(f"unknown value {v!r}")
    return v


def _double(bits: str) -> float:
    """A float as it travels in the job file: its 64 bits in hex (exact both ways)."""
    return struct.unpack(">d", bytes.fromhex(bits))[0]


def _plain(v):
    """A value Excel gave, as both executors return it; a COM object or a date is an error."""
    if v is None or isinstance(v, (bool, int, float, str)):
        return v
    if isinstance(v, Decimal):                       # a currency (VT_CY)
        return float(v)
    if isinstance(v, (list, tuple)):
        return tuple(_plain(x) for x in v)
    if isinstance(v, datetime.datetime):
        raise _Failed(NOT_A_DATE)
    raise _Failed(NOT_A_VALUE)


def _message(e: BaseException) -> str:
    """Excel's own text for a failed call (the description and the error code), else the exception's."""
    if isinstance(e, pywintypes.com_error):
        info = e.excepinfo or ()
        text = (info[2] if len(info) > 2 and info[2] else None) or e.strerror or "COM error"
        code = info[5] if len(info) > 5 and info[5] else e.hresult
        if is_busy(e):
            text = jobs.BUSY
        return f"{text} (0x{code & 0xFFFFFFFF:08X})"
    return str(e) or type(e).__name__


# ------------------------------------------------------------------ the in-process executor
@dataclass
class _Excel:
    """One Excel process: whether the add-in is loaded there (None: not tried yet) and its macro names."""
    pid: int
    ready: bool | None = None
    hash: str = ""
    xll: str = ""
    lock: threading.Lock = field(default_factory=threading.Lock)

    @property
    def run_macro(self) -> str:
        return f"SmartassRun_{self.hash}"

    @property
    def ping_macro(self) -> str:
        return f"SmartassPing_{self.hash}"


_EXCELS: dict[tuple, _Excel] = {}       # (pid, process start time) -> _Excel
_EXCELS_LOCK = threading.Lock()
_last = threading.local()               # this thread's last (app, key): a job's batches share one app object


def _ready(app) -> _Excel | None:
    if not config.INPROC:
        return None
    try:
        key = _identify(app)
    except HostError:
        raise
    except Exception as e:   # noqa: BLE001 - which Excel process this is cannot be told: over COM
        log.warning("in-process executor not used: the Excel process is not known (%s)", e)
        return None
    with _EXCELS_LOCK:
        ex = _EXCELS.setdefault(key, _Excel(key[0]))
    with ex.lock:
        if ex.ready is None:
            _load(app, ex)
    return ex if ex.ready else None


def _identify(app) -> tuple:
    """(pid, start time) of app's Excel process, from its main window."""
    last = getattr(_last, "seen", None)
    if last is not None and last[0] is app:
        return last[1]
    import psutil
    import win32process
    hwnd = jobs.com(lambda: int(app.Hwnd))
    pid = win32process.GetWindowThreadProcessId(hwnd)[1]
    key = (pid, psutil.Process(pid).create_time())
    _last.seen = (app, key)
    return key


def _is_64bit(pid: int) -> bool:
    """The Excel process is 64-bit (a 32-bit Excel on 64-bit Windows runs under WOW64)."""
    import win32api
    import win32process
    h = win32api.OpenProcess(0x1000, False, pid)            # PROCESS_QUERY_LIMITED_INFORMATION
    try:
        return not win32process.IsWow64Process(h)
    finally:
        win32api.CloseHandle(h)


def _load(app, ex: _Excel) -> None:
    """Build (once per hash) and load the add-in in this Excel; on any failure the COM executor for this Excel,
    with one warning. An abandoned job leaves it undecided."""
    try:
        if not _is_64bit(ex.pid):
            raise build.BuildError("Excel is 32-bit; the add-in is built for 64-bit Excel")
        xll, ex.hash = build.ensure()
        ex.xll = str(xll)
        _register(app, ex)
        ex.ready = True
        log.info("in-process executor loaded in Excel pid %s (%s)", ex.pid, ex.xll)
    except jobs.JobAborted:
        raise
    except Exception as e:   # noqa: BLE001 - whatever it is, the COM executor does the work
        _give_up(ex, _message(e) if isinstance(e, pywintypes.com_error) else str(e))


def _register(app, ex: _Excel) -> None:
    if jobs.com(lambda: app.RegisterXLL(ex.xll), limit=LOAD_LIMIT_S) is not True:
        raise build.BuildError(f"RegisterXLL({ex.xll}) did not load it")
    got = jobs.com(lambda: app.Run(ex.ping_macro), limit=LOAD_LIMIT_S)
    if got != ex.hash:
        raise build.BuildError(f"the add-in answered {got!r} to its ping, not {ex.hash!r}")


def _give_up(ex: _Excel, why: str) -> None:
    ex.ready = False
    log.warning("in-process executor not used in Excel pid %s: %s; its Excel work runs over COM", ex.pid, why)


def _run_inproc(app, ex: _Excel, job: dict, limit: float) -> dict | None:
    """The add-in's result record, or None when nothing of the batch ran and the COM executor is to run it."""
    folder = config.INPROC_DIR / "jobs"
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f"{uuid.uuid4().hex}.json"
    out_path = Path(f"{path}.out")
    path.write_text(json.dumps(_wire(job)), encoding="utf-8")
    try:
        for again in (False, True):
            err, record = "Application.Run gave no result", None
            try:
                got = jobs.com(lambda: app.Run(ex.run_macro, str(path)), limit=limit)
                if isinstance(got, str) and got.startswith("{"):
                    record = got
                elif got == "file":
                    record = out_path.read_text(encoding="utf-8")
                else:
                    err = f"Application.Run gave {got!r}"
            except HostError:
                raise                                   # busy (the Run was rejected: nothing ran) or abandoned
            except Exception as e:   # noqa: BLE001 - what it means depends on the files below
                err = _message(e)
            if record is None and out_path.exists():
                record = out_path.read_text(encoding="utf-8")
            if record is not None:
                out = json.loads(record)
                if not out["ok"] and out.get("started", 0) == 0:
                    log.warning("the add-in failed before the batch's first operation (%s); it runs over COM",
                                out.get("message"))
                    return None
                return out
            if not path.exists():
                raise BatchError(None, "the batch in Excel", f"the add-in took the job and wrote no result ({err})",
                                 where="inproc")
            if again:
                break
            try:                                        # the macro is not there: load the add-in once more
                _register(app, ex)
            except jobs.JobAborted:
                raise
            except Exception as e:   # noqa: BLE001
                err = _message(e) if isinstance(e, pywintypes.com_error) else str(e)
                break
        _give_up(ex, err)
        return None
    finally:
        for p in (path, out_path):
            try:
                p.unlink(missing_ok=True)
            except OSError:
                pass


def _wire(v):
    """The job as it travels to the add-in: every float as {"d": its 64 bits in hex}."""
    if isinstance(v, float):
        return {"d": struct.pack(">d", v).hex()}
    if isinstance(v, dict):
        return {k: _wire(x) for k, x in v.items()}
    if isinstance(v, list):
        return [_wire(x) for x in v]
    return v


# ------------------------------------------------------------------ verbs
def _break_links(wb, file_name: str) -> int:
    """Break wb's links whose file name is file_name (any case); the user's own links to other workbooks stay.
    Reads LinkSources first: BreakLink needs the exact link text Excel holds."""
    broke = 0
    for link in _call(wb, "LinkSources", [LINKS]) or ():
        if re.split(r"[\\/]", str(link))[-1].casefold() == file_name.casefold():
            _call(wb, "BreakLink", [link, LINKS])
            broke += 1
    return broke


VERBS = {"break_links": _break_links}
