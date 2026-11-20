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
BatchError and the batch is never run again (see _run_inproc for how "nothing ran" is told).

The COM executor (execute) interprets the batch's plain data over the COM objects, inside one jobs.com. A busy Excel
is waited for on the rejected operation only (a rejected call did not run), never by running the batch again. A call
with arguments is made as the add-in makes it (IDispatch::Invoke, LCID 0, method or property get, the arguments on
the member itself): pywin32's obj.Name(args) reads a property whose arguments are all optional first and calls what
that gave, so Range.Offset(2, 3) would be the default member of the cell itself ($D$3 from B2, not $E$4) and
Address(False, False) would fail; pywin32's own aliases (GetAddress) are unknown names in both executors.

check (run's argument, default True): the job's abort token is checked before the batch and, over COM, before every
operation, so an abandoned job stops at its next batch; check=False runs the batch for an abandoned job too (the
clean-up after a failure: rollback, restore, the user's sheet landed on again).

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
import inspect
import json
import logging
import struct
import threading
import time
import uuid
import weakref
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

import pythoncom
import pywintypes
import winerror
from win32com.client import dynamic

from ...core import config, timing
from .. import fastcom, jobs
from ..host import HostError
from ..hostproc import is_busy, retry_busy
from . import build

log = logging.getLogger("smartview.inproc")

MANUAL = -4135                          # xlCalculationManual
QUIET = (("ScreenUpdating", False), ("EnableEvents", False), ("DisplayAlerts", False), ("Calculation", MANUAL))
RESTORE_LIMIT_S = 30
LOAD_LIMIT_S = 60                       # RegisterXLL and the ping: how long a busy Excel is waited for
REGISTER_LIMIT_S = 60                   # a RegisterXLL that has not returned by then: no add-in in that Excel
LOCK_WAIT_S = 10                        # a batch waits this long for another job's load of the add-in, then COM
LEFTOVER_S = 600                        # job / result files older than this are leftovers of a crash: removed
NOT_RUN = "not run: "                   # the add-in's answer when it read, deleted and ran nothing
OMITTED = pythoncom.ArgNotFound         # MISSING: sent as DISP_E_PARAMNOTFOUND, as C#'s Type.Missing
UNBOUND = "no object {}: the operation that makes it failed or did not run"
NOT_A_VALUE = "a COM object cannot be asked back"
NOT_A_DATE = "a date cannot be asked back (read Value2)"
TAKEN = "taken: "                        # how the verb exist's refusal ends: the names, ", " between
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


def run(app, batch, limit: float = 120, check: bool = True) -> Result:
    """Run batch in app's Excel, in its process when the add-in is there, else over COM; Result, or BatchError when
    an operation failed. A busy Excel (beyond limit) or an abandoned job raises HostError, as jobs.com does (check
    False: the batch runs for an abandoned job too, see the module docstring). Counted as one batch of the current
    stage (core/timing)."""
    timing.count_batch()
    job = batch.to_dict()
    ex = _ready(app, check)
    if ex is not None:
        out = _run_inproc(app, ex, job, limit, check)
        if out is not None:
            return _result(out, batch, "inproc")
    out = jobs.com(lambda: execute(app, job, limit, check), limit=limit, check=check)
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
        if i is not None and i >= 0:
            text = batch.text(i)
        else:
            text = "before the first operation" if out.get("started", 0) == 0 else "outside the operations"
        raise BatchError(i, text, out.get("message") or "", skipped, where)
    return Result({k: _tuples(v) for k, v in out["values"].items()}, dict(out["marks"]), float(out["ms"]), skipped,
                  where)


def _tuples(v):
    return tuple(_tuples(x) for x in v) if isinstance(v, (list, tuple)) else v


# ------------------------------------------------------------------ the COM executor
class _Failed(Exception):
    """An operation's failure, with the message both executors give."""


def execute(app, job: dict, limit: float = 120, check: bool = True) -> dict:
    """The batch's operations over COM in this thread; the result record (as the add-in writes it). Raises only
    JobAborted (after the finally operations and the restore; never with check False)."""
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
            _quiet(app, changed, limit, job.get("calculation", True))
        try:
            for i, op in enumerate(ops):
                started = i + 1
                try:
                    if check:
                        jobs._check_alive()
                    retry_busy(lambda op=op: _do(app, op, refs, values, marks, t0), limit)
                except jobs.JobAborted as e:
                    aborted = e
                    break
                except Exception as e:   # noqa: BLE001 - the failure is the result
                    if op.get("try"):
                        skipped.append([i, _failure(i, op, e)])
                        continue
                    out = {"ok": False, "index": i, "message": _failure(i, op, e)}
                    break
        finally:
            for j, op in enumerate(final):
                try:
                    retry_busy(lambda op=op: _do(app, op, refs, values, marks, t0), limit)
                except Exception as e:   # noqa: BLE001
                    skipped.append([len(ops) + j, _failure(len(ops) + j, op, e)])
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


def _quiet(app, changed: list, limit: float, calculation: bool = True) -> None:
    """Each quiet setting that is not so already is set; changed gets (name, the user's value) for the restore.
    Calculation cannot be read or set with no workbook open: then it is left alone; without calculation it is not
    touched at all."""
    for name, value in QUIET:
        if name == "Calculation" and not calculation:
            continue
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


def _failure(i: int, op: dict, e: BaseException) -> str:
    """"operation <index> (<kind> <name>): <message>": how a failing operation's text starts in both executors (the
    message after it is each executor's own: pywin32's and the add-in's texts for the same COM error can differ)."""
    return f"operation {i} ({op.get('k', '?')} {op.get('name', '?')}): {_message(e)}"


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
    """One Excel process: whether the add-in is loaded there (None: not decided yet) and its macro names."""
    pid: int
    ready: bool | None = None
    hash: str = ""
    xll: str = ""
    lock: threading.Lock = field(default_factory=threading.Lock)
    waited: bool = False                  # a batch gave up waiting for the lock (warned once)

    @property
    def run_macro(self) -> str:
        return f"SmartassRun_{self.hash}"

    @property
    def ping_macro(self) -> str:
        return f"SmartassPing_{self.hash}"


_EXCELS: dict[tuple, _Excel] = {}       # (pid, process start time) -> _Excel
_EXCELS_LOCK = threading.Lock()
_last = threading.local()               # this thread's last (weak app, key): a job's batches share one app object
_unknown_warned = False


def _ready(app, check: bool = True) -> _Excel | None:
    """app's Excel with its add-in loaded, or None (the COM executor). With check False an add-in not loaded yet is
    not loaded now (that needs a live job)."""
    global _unknown_warned
    if not config.INPROC:
        return None
    try:
        key = _identify(app, check)
    except HostError:
        raise
    except Exception as e:   # noqa: BLE001 - which Excel process this is cannot be told: over COM
        if not _unknown_warned:                   # without a process there is nothing to key "once per Excel" on
            _unknown_warned = True
            log.warning("in-process executor not used: the Excel process is not known (%s)", e)
        return None
    with _EXCELS_LOCK:
        ex = _EXCELS.setdefault(key, _Excel(key[0]))
    if not ex.lock.acquire(timeout=LOCK_WAIT_S):  # another job is still loading it (or stuck doing so)
        if not ex.waited:
            ex.waited = True
            log.warning("in-process executor not used for a batch in Excel pid %s: its add-in is still being "
                        "loaded after %s s; the batch runs over COM", ex.pid, LOCK_WAIT_S)
        return None
    try:
        if ex.ready is None and check:
            _load(app, ex)
    finally:
        ex.lock.release()
    return ex if ex.ready else None


def _identify(app, check: bool = True) -> tuple:
    """(pid, start time) of app's Excel process, from its main window."""
    last = getattr(_last, "seen", None)
    if last is not None and last[0]() is app:
        return last[1]
    import psutil
    import win32process
    hwnd = jobs.com(lambda: int(app.Hwnd), check=check)
    pid = win32process.GetWindowThreadProcessId(hwnd)[1]
    key = (pid, psutil.Process(pid).create_time())
    try:
        _last.seen = (weakref.ref(app), key)
    except TypeError:
        pass
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
    """Build (once per hash) and load the add-in in this Excel; on a failure the COM executor for this Excel, with
    one warning. A busy Excel or an abandoned job leaves it undecided (the next batch tries again)."""
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
        if isinstance(e, HostError) or (isinstance(e, pywintypes.com_error) and is_busy(e)):
            log.info("in-process executor not loaded yet in Excel pid %s: %s", ex.pid, e)
            return
        _give_up(ex, _message(e) if isinstance(e, pywintypes.com_error) else str(e))


class _Blocked(Exception):
    """RegisterXLL did not return in time."""


def _register(app, ex: _Excel) -> None:
    if _register_xll(app, ex.xll, REGISTER_LIMIT_S) is not True:
        raise build.BuildError(f"RegisterXLL({ex.xll}) did not load it")
    got = jobs.com(lambda: app.Run(ex.ping_macro), limit=LOAD_LIMIT_S)
    if got != ex.hash:
        raise build.BuildError(f"the add-in answered {got!r} to its ping, not {ex.hash!r}")


def _register_xll(app, path: str, limit: float):
    """app.RegisterXLL(path) on a thread of its own, waited for at most limit seconds (a dialog in Excel would hold
    it). On time out _Blocked: that thread stays blocked in Excel until Excel answers, and the caller gives up on
    the add-in for that Excel process."""
    jobs._check_alive()
    if isinstance(app, dynamic.CDispatch):           # a COM object: handed to the other thread through a stream
        stream = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, app._oleobj_)

        def call():
            pythoncom.CoInitialize()
            try:
                excel = dynamic.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(stream, pythoncom.IID_IDispatch))
                try:
                    return retry_busy(lambda: excel.RegisterXLL(path), LOAD_LIMIT_S)
                finally:
                    excel = None
            finally:
                pythoncom.CoUninitialize()
    else:                                              # the fakes of the tests
        def call():
            return app.RegisterXLL(path)
    box: dict = {}

    def worker():
        try:
            box["value"] = call()
        except BaseException as e:   # noqa: BLE001 - handed to the waiting thread
            box["error"] = e
    t = threading.Thread(target=worker, name="smartass RegisterXLL", daemon=True)
    t.start()
    t.join(limit)
    if t.is_alive():
        raise _Blocked(f"RegisterXLL did not return within {limit:.0f} s (a dialog in Excel?)")
    if "error" in box:
        raise box["error"]
    return box["value"]


def _give_up(ex: _Excel, why: str) -> None:
    ex.ready = False
    log.warning("in-process executor not used in Excel pid %s: %s; its Excel work runs over COM", ex.pid, why)


def _unknown(detail: str) -> BatchError:
    """The add-in may have run operations of the batch, and what it did cannot be told: never run it again."""
    return BatchError(None, "the batch in Excel", f"the batch's outcome is unknown: {detail}", where="inproc")


def _sweep(folder: Path) -> None:
    """Job and result files left by an earlier run (a crash) hold a batch's values: removed after LEFTOVER_S."""
    now = time.time()
    for p in folder.iterdir():
        try:
            if now - p.stat().st_mtime > LEFTOVER_S:
                p.unlink()
        except OSError:
            pass


def _run_inproc(app, ex: _Excel, job: dict, limit: float, check: bool = True) -> dict | None:
    """The add-in's result record, or None when nothing of the batch ran and the COM executor is to run it.

    Whether it started is told by what the add-in returned: a record (its "started"), "file" (a record in the result
    file), or "not run: ..." (nothing read, deleted or run). Anything else after the add-in was reached is an unknown
    outcome (BatchError), never run again. Only when Application.Run itself failed (the macro is not there) and the
    job file is untouched (the add-in deletes it before it runs anything) did nothing run: the add-in is registered
    once more and asked again."""
    folder = build.jobs_folder()
    folder.mkdir(parents=True, exist_ok=True)
    _sweep(folder)
    path = folder / f"{uuid.uuid4().hex}.json"
    out_path = Path(f"{path}.out")
    path.write_text(json.dumps(_wire(job)), encoding="utf-8")
    try:
        for again in (False, True):
            try:
                got = jobs.com(lambda: app.Run(ex.run_macro, str(path)), limit=limit, check=check)
            except HostError:
                raise                                   # busy (the Run was rejected: nothing ran) or abandoned
            except Exception as e:   # noqa: BLE001 - did the add-in get the job? Its file says
                if not path.exists():
                    raise _unknown(f"Application.Run failed after the add-in took the job ({_message(e)})") from None
                err = _message(e)
                if again:
                    break
                try:                                    # the macro is not there: load the add-in once more
                    _register(app, ex)
                    continue
                except jobs.JobAborted:
                    raise
                except Exception as e2:   # noqa: BLE001
                    err = _message(e2) if isinstance(e2, pywintypes.com_error) else str(e2)
                    break
            if isinstance(got, str) and got.startswith(NOT_RUN):
                log.warning("the add-in did not run the batch (%s); it runs over COM", got[len(NOT_RUN):])
                if got.startswith(NOT_RUN + "refused"):
                    _give_up(ex, got)                   # the add-in and this engine disagree on the job folder
                return None
            if got == "file":
                try:
                    record = out_path.read_text(encoding="utf-8")
                except OSError as e:
                    raise _unknown(f"the result file could not be read ({e})") from None
                finally:
                    try:
                        out_path.unlink(missing_ok=True)
                    except OSError:
                        pass                            # e.g. a scanner holds it: the sweep removes it later
            elif isinstance(got, str) and got.startswith("{"):
                record = got
            else:
                raise _unknown(f"the add-in answered {str(got)[:200]!r}")
            out = _parse(record)
            if not out["ok"] and out["started"] == 0:
                log.warning("the add-in failed before the batch's first operation (%s); it runs over COM",
                            out.get("message"))
                return None
            return out
        _give_up(ex, err)
        return None
    finally:
        for p in (path, out_path):
            try:
                p.unlink(missing_ok=True)
            except OSError:
                pass


def _parse(record: str) -> dict:
    """The add-in's result record, checked; an unreadable one is an unknown outcome."""
    try:
        out = json.loads(record)
        ok, started = out["ok"], out["started"]
        if not (isinstance(ok, bool) and isinstance(started, int) and isinstance(out["values"], dict)
                and isinstance(out["marks"], dict) and isinstance(out["ms"], (int, float))
                and isinstance(out["skipped"], list)):
            raise ValueError("a field of the wrong type")
        if not ok and not (isinstance(out["index"], int) and isinstance(out["message"], str)):
            raise ValueError("a failure without its operation")
    except (ValueError, KeyError, TypeError) as e:
        raise _unknown(f"the add-in's result could not be read ({e})") from None
    return out


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
def _each(collection, paths, prefixes):
    """The verb each (batch.py's module docstring): [[the value at each path] for every item, in order]."""
    keep = tuple(str(p).casefold() for p in prefixes) if prefixes else None
    rows = []
    for item in collection:
        row = [_path(item, str(p).split(".")) for p in paths]
        if keep is None or (isinstance(row[0], str) and row[0].casefold().startswith(keep)):
            rows.append(row)
    return rows


def _path(obj, parts: list[str]):
    """The value at a dotted path from obj ("*": the rest over every item of the collection reached); None when it
    cannot be read. Each step is a property read: pywin32 gives a method's name as a method object, never calling it
    (a COM object: no value, None). A busy Excel goes on up (the verb is read again as a whole)."""
    try:
        for i, part in enumerate(parts):
            if part == "*":
                return [_path(x, parts[i + 1:]) for x in obj]
            if not _is_property(obj, part):
                return None                          # a method's name: never called
            obj = getattr(obj, part)
        _plain(obj)                                   # a COM object or a date is not a value
        return obj
    except Exception as e:   # noqa: BLE001 - unreadable: None
        if is_busy(e):
            raise
        return None


_PROPERTIES: dict[tuple, bool] = {}     # (interface IID, member lower-cased) -> a property read without arguments


def _is_property(obj, name: str) -> bool:
    """name is a property of obj that reads without arguments, from the type information, as Runner.cs asks it:
    Excel runs a method even when it is asked for as a property. Cached per interface IID and member: a fastcom
    object knows its IID (its interface's shared data), so a cached answer costs no call to Excel; pywin32's own
    wrapper (SV_FASTCOM=0) asks the IID of every object (two calls), and an object without a real IID is never
    cached. A busy Excel goes up (the operation is waited for and run again), never cached; only a name the type
    information does not know, or a walk of its members that went through, is cached. A fake of the tests: any
    attribute that is not callable."""
    if not isinstance(obj, dynamic.CDispatch):
        try:
            return not inspect.isroutine(inspect.getattr_static(obj, name))
        except AttributeError:                       # given by __getattr__: a value
            return True
    ole, d = obj._oleobj_, obj.__dict__
    iface, ti, attr = d.get("_iface_"), d.get("_typeinfo_"), None
    try:
        if iface is None or ti is None:
            ti = ole.GetTypeInfo()
            attr = ti.GetTypeAttr()
            iid = attr[0]
        else:
            iid = iface.iid
    except pythoncom.com_error as e:
        if is_busy(e):
            raise
        return False                                 # no type information: nothing is read
    key = None if iid in fastcom.NOT_SHARED else (iid, name.lower())
    if key in _PROPERTIES:
        return _PROPERTIES[key]
    try:
        dispid = ole.GetIDsOfNames(name)
    except pythoncom.com_error as e:
        if is_busy(e):
            raise
        if key is not None and e.hresult == winerror.DISP_E_UNKNOWNNAME:
            _PROPERTIES[key] = False
        return False
    try:
        if attr is None:
            attr = ti.GetTypeAttr()
        prop = False
        for i in range(attr.cFuncs):
            fd = ti.GetFuncDesc(i)
            if fd.memid == dispid and fd.invkind & pythoncom.INVOKE_PROPERTYGET and len(fd.args) == fd.cParamsOpt:
                prop = True
                break
        prop = prop or any(ti.GetVarDesc(i).memid == dispid for i in range(attr.cVars))
    except pythoncom.com_error as e:
        if is_busy(e):
            raise
        return False                                 # not known this time: asked again next time
    if key is not None:
        _PROPERTIES[key] = prop
    return prop


def _nudge(obj, name: str, delta, above):
    """The verb nudge (batch.py's module docstring): the numeric property + delta when it is more than above (or
    above is None); the value it ends with."""
    if not _is_property(obj, name):
        raise _Failed(f"{name} is not a property")
    v = getattr(obj, name)
    if not isinstance(v, (int, float)) or isinstance(v, bool):
        raise _Failed(f"{name} is not a number")
    if above is not None and not v > above:
        return v
    new = float(v) + float(delta)
    setattr(obj, name, new)
    return new


def _exist(collection, keys, refuse):
    """The verb exist (batch.py's module docstring): the items' own names of the keys that name an item of the
    collection, in order; with refuse a key that is there fails the operation (TAKEN and those names). A busy Excel
    goes on up (the verb is run again)."""
    found = []
    for key in keys:
        try:
            name = getattr(_call(collection, "Item", [key]), "Name")
        except Exception as e:   # noqa: BLE001 - not an item
            if is_busy(e):
                raise
            continue
        if isinstance(name, str) and name.casefold() == str(key).casefold():
            found.append(name)
        elif (own := _named(collection, key)) is not None:
            found.append(own)
    if refuse and found:
        raise _Failed(TAKEN + ", ".join(found))
    return found



def _named(collection, key):
    """The Name of the collection's item named key (case ignored), by a walk over its items, or None: only when
    Item(key) gave an item of another name (with a name both the workbook and the active sheet have, Excel's
    Workbook.Names.Item gives the sheet's own, live 20.11.2026)."""
    want = str(key).casefold()
    for item in collection:
        try:
            name = getattr(item, "Name")
        except Exception as e:   # noqa: BLE001 - unreadable: not this one
            if is_busy(e):
                raise
            continue
        if isinstance(name, str) and name.casefold() == want:
            return name
    return None


VERBS: dict = {"each": _each, "nudge": _nudge, "exist": _exist}    # name -> function(*args); both executors have
                                                                   # each one (batch.VERBS, Runner.cs Verb)
