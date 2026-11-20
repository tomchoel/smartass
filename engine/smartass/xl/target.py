"""Resolve the workbook an analysis goes into: an open book, a file to open, or a new analysis file."""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import shutil
import subprocess
import time
import uuid
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import pywintypes

from . import instances, jobs, login
from ..core import config
from .host import ExcelHost, HostError, InputError

# ------------------------------------------------------------------ target workbook
@dataclass
class Target:
    """The workbook a job works in. made: the new workbook's file the job itself reserved and builds (wb is None
    until open_new opens it); a failed job closes it unsaved and deletes that file (discard_new). from_template: a
    new workbook added from the TN template (its look tells the base a new workbook has: remember_new_base)."""
    app: object
    wb: object
    pid: int
    hwnd: int
    bridge: object = None
    created: bool = False
    sheets_added: list[str] = field(default_factory=list)
    made: Path | None = None
    from_template: bool = False

    @property
    def name(self) -> str:
        return self.made.name if self.wb is None else self.wb.Name

    @property
    def path(self) -> str:
        return str(self.made) if self.wb is None else instances.local_path(self.wb.FullName)

    @property
    def full_name(self) -> str:
        return str(self.made) if self.wb is None else str(self.wb.FullName)

    def run(self, fn: str, *args):
        """Bridge call. Smart View acts on the ACTIVE workbook/sheet, so the target sheet is activated right before
        every call - if the user clicked into another workbook meanwhile, we act on ours, not theirs."""
        if args and isinstance(args[0], str) and fn not in ("B_Ping", "B_SetOption", "B_GetOption", "B_LastError"):
            name = args[0]
            try:
                ws = jobs.com(lambda: self.wb.Worksheets(name))
                jobs.com(self.wb.Activate)
                jobs.com(ws.Activate)
            except pywintypes.com_error:
                pass
        return jobs.com(lambda: self.app.Run(f"'{self.bridge.Name}'!modBridge.{fn}", *args))

    def status(self, text: str | None):
        try:
            jobs.com(lambda: setattr(self.app, "StatusBar", ("Claude: " + text) if text else False), limit=5)
        except Exception:   # noqa: BLE001 - cosmetic
            pass


def _wait_ready(app, limit: float = jobs.EDIT_WAIT_S, pid: int | None = None, name: str | None = None):
    t0 = time.monotonic()
    while time.monotonic() - t0 < limit:
        try:
            if app.Ready:
                return
        except pywintypes.com_error as e:
            if not jobs.is_busy(e):
                raise
        time.sleep(0.5)
    if pid:
        raise instances.busy_error(pid, name or "The workbook")
    raise HostError("Your Excel is in edit mode (a cell is being typed in) or showing a dialog. Press Enter or Esc, "
                    "then ask again.")


def _attach_bridge(h: ExcelHost, t: Target):
    ver = instances.ensure_bridge_file(h)
    path = instances.bridge_path(ver)
    br = None
    try:
        br = jobs.com(lambda: t.app.Workbooks(path.name), limit=20)    # this exact version already loaded
    except pywintypes.com_error as e:
        if e.hresult != -2147352567:                 # anything but "not in collection" is a real error
            raise
    if br is None:
        br = jobs.com(lambda: t.app.Workbooks.Open(str(path)), limit=120)
    t.bridge = br
    if t.run("B_Ping") != "pong":
        raise HostError("The Smart View bridge add-in did not load in your Excel.")


ARCHIVE_REFUSAL = ("'{name}' is in the Produktregnskapet archive (audit record) and is not written to. "
                   "Use target='new', or allow it explicitly for this one request (allow_archive=true).")


def resolve_target(h: ExcelHost, target: str = "session", allow_archive: bool = False, title: str | None = None,
                   build_new: bool = False) -> Target:
    """target: 'session' (default: this session's own analysis workbook - created on its first question, reopened
    if closed), 'active' (the workbook the user is working in), 'new' (a new file, which becomes the session's
    workbook), an absolute file path, or the name of an open workbook (e.g. an unsaved 'Book1'). build_new: the caller
    builds a new workbook itself (pipeline.excel_job: the first answer's built file is the workbook) - a target that
    resolves to a new workbook in a running Excel then comes back with its file reserved and not opened (Target.made,
    wb None; open_new opens it), made ready (bridge) all the same, and does not become the session's workbook here:
    the caller sets it once the answer is in it. A failure after the reservation deletes the file (discard_new)."""
    from ..core import session
    from ..epm import sso
    sso.ensure_global_watcher()
    if target == "session":
        wb_path = session.workbook()
        if wb_path:
            t = resolve_target(h, wb_path, allow_archive, title)
            t.created = True                          # the session's own file: saved after each answer
            return t
        target = "new"
    from ..core.timing import stage
    if target not in ("active", "new"):
        p = Path(target)
        if not p.is_absolute():                                 # the name of an open workbook, e.g. an unsaved 'Book1'
            with stage("find Excel windows"):
                insts = instances.excel_instances()
            t = _find_named([(i.pid, i.hwnd, i.app) for i in insts if i.app is not None], target)
            jobs.hold_turn(t.pid, what="this request")
            with stage("ready + bridge"):
                return _finish_target(h, t, allow_archive)      # the archive guard applies here too
        if instances.is_archive(p) and not allow_archive:       # refuse BEFORE opening anything
            raise InputError(ARCHIVE_REFUSAL.format(name=p.name))
    with stage("find Excel windows"):
        insts = instances.excel_instances()                     # waits for any open EPM login before touching an Excel
    excels = [(i.pid, i.hwnd, i.app) for i in insts if i.app is not None]
    if target == "active":
        for pid, hwnd, app in excels:
            try:
                wb = jobs.com(lambda: app.ActiveWorkbook, limit=20)
            except pywintypes.com_error:
                wb = None
            if wb is not None and not str(wb.Name).lower().endswith(".xlam"):
                jobs.hold_turn(pid, what="this request")   # same workbook: agents take turns
                return _finish_target(h, Target(app, wb, pid, hwnd), allow_archive)
        target = "new"                                # nothing open: fall back to a new analysis workbook
    if target == "new":
        with stage("new workbook"):
            t = _new_target(h, title or "Analysis", excels, {i.pid for i in insts}, build_new)
        if t.made is not None:                        # the caller builds it, opens it, and makes it the session's
            try:
                with stage("ready + bridge"):
                    return _finish_target(h, t, allow_archive)
            except BaseException as e:
                _noted(e, discard_new(t))
                raise
        with stage("ready + bridge"):
            t = _finish_target(h, t, allow_archive)
        session.set_workbook(t.path)                  # a new file becomes this session's workbook
        return t
    with stage("open workbook"):
        t = _open_path(h, Path(target), excels, insts)
    with stage("ready + bridge"):
        return _finish_target(h, t, allow_archive)


def _new_target(h: ExcelHost, title: str, excels, known: set[int] | None = None, build_new: bool = False) -> Target:
    """A new analysis workbook in a FREE Excel, so parallel agents do not wait for each other: the first Excel whose
    turn is free; else (below MAX_EXCELS) a new Excel window started on the TN template; else wait for the first.
    If it cannot be saved, the new workbook is closed unsaved (and an Excel started for it is closed again).
    build_new (the caller builds the file): in a running Excel only the file's name is reserved (_reserved_target);
    an Excel started for a parallel agent and no Excel at all make the workbook as before."""
    from . import styling
    from ..core.timing import stage
    tpl = styling.template_path()
    choice = next((e for e in excels if jobs.hold_turn(e[0], wait=False)), None)
    wb = None
    spawned = False
    if choice is None and excels and len(excels) < jobs.MAX_EXCELS and tpl.exists():
        with stage("start another Excel (parallel agent)"):
            choice, wb = _spawn_excel(tpl, {e[0] for e in excels} | (known or set()))
        spawned = True
    if choice is None and excels:
        jobs.hold_turn(excels[0][0], what="this request")
        choice = excels[0]
    if choice is None:                                # no Excel at all: the hidden one creates the file, /x opens it
        t = _open_path(h, _new_analysis_file(h, title), excels)
        t.created = True
        return t
    pid, hwnd, app = choice
    if wb is None and build_new:
        return _reserved_target(app, pid, hwnd, title)
    from_template = spawned
    try:
        if wb is None:
            with stage("add from template"):
                try:                                  # from the TN template: the styles come for free (saves ~11 s)
                    tpl_s = str(styling.ensure_template(app))
                    wb = jobs.com(lambda: app.Workbooks.Add(tpl_s))
                    from_template = True
                except Exception:   # noqa: BLE001 - template unavailable: plain workbook, styles installed later
                    wb = jobs.com(lambda: app.Workbooks.Add())
        with stage("save as"):
            _save_new(app, wb, title)
    except BaseException:
        _discard_new(app, wb, spawned)
        raise
    t = Target(app, wb, pid, hwnd, from_template=from_template)
    t.created = True
    return t


def _reserved_target(app, pid: int, hwnd: int, title: str) -> Target:
    """The Target of a new workbook the caller builds: a file name free on disk and not the name of a workbook open
    in that Excel (as _save_new's), reserved (_reserve). Nothing is done in Excel; its name is known (jobs.book_name
    reads it from the Target: the opened file's name)."""
    try:
        taken = {str(b.Name).lower() for b in jobs.com(lambda: app.Workbooks)}
    except Exception:   # noqa: BLE001
        taken = set()
    try:
        path = _reserve(title, taken)
    except OSError as e:
        raise HostError(f"Could not create the new workbook in the analysis folder ({e}). Nothing was left open; check "
                        "that the folder exists and is writable, then ask again.") from None
    t = Target(app, None, pid, hwnd, created=True, made=path)
    t.book_name = path.name
    return t


def open_new(t: Target) -> None:
    """The new workbook the job built (t.made) opened in its Excel with Workbooks.Open: t.wb from then on. Excel's
    error is raised as it is (the job's clean-up deletes the file). Opened read-only it is an error too (HostError):
    Excel opens a file whose sensitivity label it does not know read-only, and its Save then writes nothing, without
    an error (live, 20.11.2026) - as _save_new refuses a SaveAs that made no file."""
    def open_():
        wb = t.app.Workbooks.Open(str(t.made))
        return wb, wb.ReadOnly
    t.wb, read_only = jobs.com(open_, limit=180)
    if read_only:
        raise HostError(f"Excel opened the new workbook {t.made.name} read-only (sensitivity label?), so it could not "
                        "save the answer")


DELETE_TRIES, DELETE_PAUSE_S = 10, 0.2      # a file Excel has just closed can stay locked for a moment


def discard_new(t: Target) -> str:
    """A failed job's new workbook: closed without saving when it is open, and the file the job reserved deleted
    (only that one: t.made is the name the job itself created, exclusively). With no workbook in hand (the Open
    failed, or was abandoned and Excel may have opened the file all the same) it is looked for among that Excel's
    workbooks by its name first (_opened). Returns the note for the error's text, which names the file when it could
    not be closed or deleted. Also for an abandoned job (check False)."""
    name = t.made.name
    if t.wb is None:
        t.wb = _opened(t)
    if t.wb is not None:
        wb = t.wb
        try:
            jobs.com(lambda: wb.Close(False), limit=30, check=False)
        except Exception as e:   # noqa: BLE001 - the job's error is the one raised
            jobs.log.warning("could not close the new workbook %s: %s", name, e)
            return f"the new workbook {name} could not be closed: close it without saving and delete {t.made}"
        t.wb = None
        what, done = f"the new workbook {name} was closed without saving and its file", " deleted"
    else:
        what, done = f"the new workbook {name} was not opened; its file", " was deleted"
    for i in range(DELETE_TRIES):
        try:
            t.made.unlink(missing_ok=True)
            return what + done
        except OSError:
            if i < DELETE_TRIES - 1:
                time.sleep(DELETE_PAUSE_S)
    jobs.log.warning("could not delete the new workbook's file %s", t.made)
    return f"{what} could not be deleted ({t.made}): delete it yourself"


def _opened(t: Target):
    """The workbook t.made is open as in t's Excel, or None: found by its name, which no other workbook open there
    had when it was reserved (an Excel cannot open two workbooks of one name)."""
    name = t.made.name.lower()
    try:
        return jobs.com(lambda: next((b for b in t.app.Workbooks if str(b.Name).lower() == name), None), limit=30,
                        check=False)
    except Exception:   # noqa: BLE001 - nothing to close that can be found
        return None


def _noted(e: BaseException, note: str) -> None:
    """note added to the end of e's text: its first text argument (a COM error's is its second, after the hresult:
    the error keeps its shape), else a text argument at the end."""
    i = next((i for i, a in enumerate(e.args) if isinstance(a, str)), None)
    if i is None:
        e.args = (*e.args, note)
    else:
        e.args = (*e.args[:i], f"{e.args[i]}; {note}", *e.args[i + 1:])


def _save_new(app, wb, title: str) -> Path:
    """Name the first sheet 'Sources', label and SaveAs the new workbook under a file name that is free on disk and
    not the name of a workbook already open in that Excel (Excel refuses both with 0x800A03EC)."""
    try:
        taken = {str(b.Name).lower() for b in jobs.com(lambda: app.Workbooks)}
    except Exception:   # noqa: BLE001
        taken = set()
    path = _new_path(title, taken)
    jobs.com(lambda: setattr(wb.Worksheets(1), "Name", "Sources"))
    jobs.com(lambda: instances._label(wb))
    try:
        jobs.com(lambda: wb.SaveAs(str(path), 51))
    except pywintypes.com_error as e:
        raise HostError(f"Excel could not save the new workbook as {path} ({_com_text(e)}). Nothing was left open; "
                        "check that the folder exists and is writable, then ask again.") from None
    if not path.exists():
        raise HostError(f"Could not create {path} (sensitivity label?).")
    return path


def _com_text(e: pywintypes.com_error) -> str:
    info = e.excepinfo if e.excepinfo and len(e.excepinfo) > 2 else None
    return str(info[2]) if info and info[2] else str(e)


def _discard_new(app, wb, spawned: bool) -> None:
    """Close a new workbook that could not be saved, without saving; quit an Excel started only for it."""
    if wb is not None:
        try:
            jobs.com(lambda: wb.Close(False), limit=30, check=False)
        except Exception:   # noqa: BLE001
            jobs.log.warning("could not close the unsaved new workbook")
    if spawned:
        try:
            left = [b for b in jobs.com(lambda: app.Workbooks, limit=10, check=False)
                    if not str(b.Name).lower().endswith((".xlam", ".xla"))]
            if not left:
                jobs.com(lambda: app.Quit(), limit=10, check=False)
        except Exception:   # noqa: BLE001
            pass


def _spawn_excel(tpl: Path, known: set[int]):
    """Start a new interactive Excel on the TN template (opens as a new unsaved workbook) and take its turn."""
    proc = _launch_excel(tpl)
    t0 = time.monotonic()
    while time.monotonic() - t0 < 120:
        time.sleep(1)
        for pid, hwnd, app in instances.interactive_excels():
            if pid in known:
                continue
            for wb in app.Workbooks:
                if str(wb.Name).lower().startswith(tpl.stem.lower()):
                    jobs.hold_turn(pid, wait=False)
                    login.wait_login_cleared(pid, appear_grace=12)     # a fresh Excel may show the EPM login at start-up
                    jobs.log.info("started Excel pid %s for a parallel agent", pid)
                    return (pid, hwnd, app), wb
    raise HostError(f"A new Excel (pid {proc.pid}) did not open the TN template within 120 s.")


def _finish_target(h: ExcelHost, t: Target, allow_archive: bool) -> Target:
    login.wait_login_cleared(t.pid)                         # before any further COM into that Excel
    if (instances.is_archive(t.path) or instances.is_archive(t.full_name)) and not allow_archive:
        raise InputError(ARCHIVE_REFUSAL.format(name=t.name))
    _wait_ready(t.app, pid=t.pid, name=t.name)
    _attach_bridge(h, t)
    return t


def _find_named(excels, name: str) -> Target:
    """The open workbook (in any Excel that answers) whose Name is `name`, case-insensitive, with or without its
    extension; add-ins are skipped. Two such workbooks are ambiguous; none is refused naming the open workbooks."""
    want = name.strip().lower()
    hits: list[Target] = []
    names: list[str] = []
    for pid, hwnd, app in excels:
        for wb in app.Workbooks:
            n = str(wb.Name)
            if n.lower().endswith((".xlam", ".xla")):
                continue
            names.append(n)
            if want in (n.lower(), Path(n).stem.lower()):
                hits.append(Target(app, wb, pid, hwnd))
    if len(hits) > 1:
        raise InputError(f"target '{name}' is ambiguous: open as " + "; ".join(str(t.wb.FullName) for t in hits)
                         + ". Give the absolute path of the one you mean.")
    if not hits:
        raise InputError(f"Give an absolute path for target (got '{name}'), or use 'session' / 'active' / 'new', or "
                         f"the name of an open workbook ({', '.join(names) or 'none open'}).")
    return hits[0]


def _find_open(excels, path: Path) -> Target | None:
    for pid, hwnd, app in excels:
        for wb in app.Workbooks:
            if instances._norm(wb.FullName) == instances._norm(path):
                return Target(app, wb, pid, hwnd)
    return None


BUSY_WAIT_S = 10                    # a busy Excel that holds the target gets this long to answer before we say why


def _wait_busy_holder(pid: int, path: Path) -> Target | None:
    """The target seems open in Excel `pid` (a window caption shows its name), which does not answer COM (dialog, cell
    edit mode, busy). Take that Excel's turn (another agent may be working in it), give it BUSY_WAIT_S to answer, else
    name what blocks it. None when it answers but does not hold `path` (a same-named workbook from another folder)."""
    jobs.hold_turn(pid, what="this request")
    t0 = time.monotonic()
    while True:
        insts = [i for i in instances.excel_instances() if i.pid == pid]
        if not insts:
            raise HostError(f"The Excel (pid {pid}) that had '{path.name}' open has closed. Ask again.")
        i = insts[0]
        if i.app is not None:
            return _find_open([(i.pid, i.hwnd, i.app)], path)
        if time.monotonic() - t0 >= BUSY_WAIT_S:
            raise instances.busy_error(pid, path.name)
        time.sleep(0.5)


def _open_path(h: ExcelHost, path: Path, excels, insts: list[instances.XlInstance] | None = None) -> Target:
    """The workbook at `path`, wherever it is open: any Excel that answers, else a busy Excel showing its window
    (which is named if it stays busy); else opened in the user's running Excel; else in a new Excel."""
    if not path.exists():
        raise InputError(f"File not found: {path}")
    t = _find_open(excels, path)
    if t is not None:
        jobs.hold_turn(t.pid, what="this request")
        return t
    if insts is None:
        insts = instances.excel_instances()
    holder = instances.busy_holder(insts, path.name)
    if holder is not None:
        t = _wait_busy_holder(holder.pid, path)
        if t is not None:
            return t
    if instances._file_in_use(path):
        raise instances._locked_error(path)
    if excels:                                  # open it in the user's Excel that is already running
        pid, hwnd, app = excels[0]
        jobs.hold_turn(pid, what="this request")
        try:
            wb = jobs.com(lambda: app.Workbooks.Open(str(path)), limit=180)
        except pywintypes.com_error as e:
            if "cannot access" in str(e).lower():
                raise instances._locked_error(path) from None
            raise
        return Target(app, wb, pid, hwnd)
    proc = _launch_excel(path)
    t0 = time.monotonic()
    while time.monotonic() - t0 < OPEN_TIMEOUT_S:
        time.sleep(1)
        for pid, hwnd, app in instances.interactive_excels():
            for wb in app.Workbooks:
                if instances._norm(wb.FullName) == instances._norm(path):
                    jobs.hold_turn(pid, what="this request")
                    login.wait_login_cleared(pid, appear_grace=12)     # a fresh Excel shows the EPM login at start-up
                    return Target(app, wb, pid, hwnd)
        if proc.poll() is not None:
            raise HostError(f"Excel (pid {proc.pid}) closed before it opened {path.name}. "
                            "Next step: open Excel yourself (any workbook), then ask again.")
    raise HostError(f"Excel did not open {path} within {OPEN_TIMEOUT_S} s (pid {proc.pid}). "
                    "Next step: check that Excel window for a sign-in or OneDrive prompt, answer it, and ask again.")


# A cold Excel loads Smart View and the OneDrive file (sync, sensitivity label) before the workbook appears;
# on a busy PC that takes well over a minute.
OPEN_TIMEOUT_S = 240


def _launch_excel(path: Path) -> subprocess.Popen:
    """Start a new interactive Excel on path with a normal, visible window, so any start-up prompt can be seen
    (inherited start-up info otherwise left it minimised, hiding prompts from the user)."""
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 1                                # SW_SHOWNORMAL
    return subprocess.Popen([config.EXCEL_EXE, "/x", str(path)], startupinfo=si)


def _new_path(title: str, taken: set[str] | frozenset = frozenset()) -> Path:
    """analyser\\<project of this session>\\<date time> <title>.xlsx, with ' (2)', ' (3)' ... when that file exists
    (the same question again within the minute, e.g. a retry) or its name is in `taken` (open workbooks, lower case)."""
    from ..core import session
    safe = re.sub(r'[\\/:*?"<>|]+', "_", title).strip()[:60] or "Analysis"   # only characters Windows forbids
    stem = f"{dt.datetime.now():%Y-%m-%d_%H%M} {safe}"
    folder = session.project_dir()
    n = 1
    while True:
        p = folder / (f"{stem}.xlsx" if n == 1 else f"{stem} ({n}).xlsx")
        if not p.exists() and p.name.lower() not in taken:
            return p
        n += 1


def _reserve(title: str, taken: set[str] | frozenset = frozenset()) -> Path:
    """A new analysis file's name (_new_path: free on disk and not in taken) reserved by creating the file, empty and
    exclusively: a parallel job that chose the same name at the same moment takes the next one. OSError when the
    folder cannot be written."""
    path = _new_path(title, taken)
    for _ in range(50):
        try:
            os.close(os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY))
            return path
        except FileExistsError:
            path = _new_path(title, set(taken) | {path.name.lower()})
    raise FileExistsError(f"no free name for a new workbook next to {path}")


def _seed_has_label(seed: Path) -> bool:
    """The seed exists and carries OUR sensitivity label, so Excel will save copies: in docMetadata/LabelInfo.xml
    (Office's 2020 form, id="{<config.LABEL_ID>}": what Excel on this tenant writes, with no custom property - live,
    20.11.2026) or as MSIP_Label_<config.LABEL_ID>_ custom properties (the older form, label_properties); the id
    compared case-insensitively."""
    if not config.LABEL_ID:
        return False
    lid = config.LABEL_ID.lower().encode()
    try:
        with zipfile.ZipFile(seed) as z:
            parts = set(z.namelist())
            info, custom = (z.read(n).lower() if n in parts else b""
                            for n in ("docMetadata/LabelInfo.xml", "docProps/custom.xml"))
    except (OSError, zipfile.BadZipFile):
        return False
    return b'id="{' + lid + b'}"' in info or b"msip_label_" + lid + b"_" in custom


def label_properties(now: dt.datetime | None = None) -> dict[str, str]:
    """Our sensitivity label as the custom document properties of a built file (MSIP_Label_<id>_*): a file with them
    opens labelled with AutoSave on and its Save writes; one without opens read-only and its Save writes nothing
    (live, hidden Excel, 20.11.2026). Enabled, set now, assigned as instances._label assigns it (privileged), our
    site, name and content bits (config [label]: Excel records 2 for ours in its own files, the seed's), a new action
    id."""
    when = (now or dt.datetime.now(dt.timezone.utc)).astimezone(dt.timezone.utc)
    p = {"Enabled": "true", "SetDate": f"{when:%Y-%m-%dT%H:%M:%SZ}", "Method": "Privileged",
         "Name": str(config.LABEL_NAME or ""), "SiteId": str(config.LABEL_SITE or ""), "ActionId": str(uuid.uuid4()),
         "ContentBits": str(config.LABEL_CONTENT_BITS)}
    return {f"MSIP_Label_{config.LABEL_ID}_{k}": v for k, v in p.items()}


NEW_BASE_FILE = "new_workbook_base.json"


def new_base():
    """The base (render/xlsx.Base, its date format left at the default) a new analysis workbook has: the Normal style
    and standard sizes of a workbook added from the TN template of this style version, as the look of one found
    them (remember_new_base). None when none is stored for this version: a new workbook is then made from the
    template (its look stores the base)."""
    from ..render.xlsx import Base
    from . import styling
    try:
        d = json.loads((config.HOME / NEW_BASE_FILE).read_text(encoding="utf-8"))
        if d.get("template") != styling._styles_version():
            return None
        return Base(str(d["font_name"]), d["font_size"], int(d["row_px"]), int(d["col_px"]))
    except (OSError, ValueError, KeyError, TypeError):
        return None


def forget_new_base() -> None:
    """The stored base dropped (it did not describe a workbook built in it): the next new workbook is made from the
    template once, and its look stores a fresh one."""
    try:
        (config.HOME / NEW_BASE_FILE).unlink(missing_ok=True)
    except OSError as e:
        jobs.log.warning("could not drop the stored base of a new workbook: %s", e)


def remember_new_base(base) -> None:
    """The base a look found in a new workbook added from the TN template, stored for new_base (best effort)."""
    from . import styling
    d = {"template": styling._styles_version(), "font_name": base.font_name, "font_size": base.font_size,
         "row_px": base.row_px, "col_px": base.col_px}
    path = config.HOME / NEW_BASE_FILE
    try:
        if path.exists() and json.loads(path.read_text(encoding="utf-8")) == d:
            return
        path.write_text(json.dumps(d), encoding="utf-8")
    except (OSError, ValueError) as e:
        jobs.log.warning("could not store the base of a new workbook: %s", e)


def _new_analysis_file(h: ExcelHost, title: str) -> Path:
    """A copy of the labelled seed under the new name (no Excel started); without a usable seed the hidden Excel
    makes the file the old way and its result becomes the seed (best effort)."""
    path = _new_path(title)
    seed = config.NEW_WORKBOOK_SEED
    if _seed_has_label(seed):
        reserved = False
        try:
            path = _reserve(title)
            reserved = True
            shutil.copyfile(seed, path)
            return path
        except OSError:
            if reserved:
                try:
                    path.unlink(missing_ok=True)
                except OSError:
                    pass
    made = h.on_host(lambda: _new_analysis_file_host(h, path))
    try:
        seed.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(made, seed)
    except OSError:
        pass
    return made


def _new_analysis_file_host(h: ExcelHost, path: Path) -> Path:
    wb = h.com(h.app.Workbooks.Add)
    try:
        wb.Worksheets(1).Name = "Sources"
        instances._label(wb)
        h.com(lambda: wb.SaveAs(str(path), 51))
    finally:
        h.com(lambda: wb.Close(False))
    if not path.exists():
        raise HostError(f"Could not create {path} (sensitivity label?).")
    return path
