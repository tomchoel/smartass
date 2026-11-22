"""The user's Excel instances (attach via the main window), workbook paths and the sensitivity label of a new
workbook."""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from pathlib import Path

import psutil
import pythoncom
import pywintypes
import win32file
import win32gui
import win32process

from . import dialogs, fastcom
from ..core import config
from ..core.errors import HostError

WM_GETOBJECT, OBJID_NATIVEOM = 0x003D, -16


def _is_embedding(pid: int) -> bool:
    try:
        return "-embedding" in " ".join(psutil.Process(pid).cmdline()).lower()
    except psutil.Error:
        return True


@dataclass
class XlInstance:
    """One interactive Excel process. `app` is None when that Excel did not answer COM (a dialog is open, a cell is
    being edited, or it is busy); its workbook window captions (`captions`) are known either way."""
    pid: int
    hwnd: int                                        # front-most main window
    captions: list[str]
    app: object = None


def _excel_windows() -> list[tuple[int, int, list[tuple[int, str]]]]:
    """Thin Win32 part: [(pid, front-most XLMAIN hwnd, [(EXCEL7 hwnd, caption)])] of the visible Excel main windows,
    grouped by process, front-most process first (EnumWindows walks top-level windows in z-order)."""
    by_pid: dict[int, tuple[int, list[tuple[int, str]]]] = {}

    def cb(h, _):
        try:
            if win32gui.GetClassName(h) != "XLMAIN" or not win32gui.IsWindowVisible(h):
                return True
            pid = win32process.GetWindowThreadProcessId(h)[1]
            desk = win32gui.FindWindowEx(h, 0, "XLDESK", None)
            ex7 = win32gui.FindWindowEx(desk, 0, "EXCEL7", None) if desk else 0
            entry = by_pid.setdefault(pid, (h, []))
            if ex7:
                entry[1].append((ex7, win32gui.GetWindowText(ex7)))
        except pywintypes.error:
            pass
        return True
    win32gui.EnumWindows(cb, None)
    return [(pid, h, ex7s) for pid, (h, ex7s) in by_pid.items()]


def _app_from_window(ex7: int):
    """Excel.Application behind a workbook window (AccessibleObjectFromWindow), or None if that Excel refuses (busy
    too). Nothing in the scan waits for Excel's type information (fastcom.scan), as before Task D2: an Application got
    while Excel rejects that is late-bound by name, and its calls say busy as busy."""
    try:
        _, lres = win32gui.SendMessageTimeout(ex7, WM_GETOBJECT, 0, OBJID_NATIVEOM, 2, 3000)
        with fastcom.scan():
            win = fastcom.wrap(pythoncom.ObjectFromLresult(lres, pythoncom.IID_IDispatch, 0))
            return win.Application
    except Exception:   # noqa: BLE001 - modal dialog, cell edit mode, busy, or not a workbook window
        return None


def engine_excel_pids() -> set[int]:
    """The hidden Excels an older release's engine started (pids.json, while it exists: hostproc.sweep_leftovers):
    never a user's Excel."""
    from .hostproc import _PidRegistry
    return {int(i["pid"]) for i in _PidRegistry().load() if isinstance(i, dict) and i.get("pid")}


def excel_instances() -> list[XlInstance]:
    """Every visible, user-facing Excel process (not one an older release's engine started), most recently active
    first, including Excels that do not answer COM right now (app=None). An Excel showing Smart View's Login is never
    touched (COM into it blocks): the job is refused at once, naming that Excel by its windows' titles
    (dialogs.login_refusal) - nothing waits for or clicks the Login."""
    out = []
    engine = engine_excel_pids()
    for pid, h, ex7s in _excel_windows():
        if pid in engine or _is_embedding(pid):
            continue
        if dialogs.login_open(pid):               # window enumeration only: no COM into that Excel
            raise dialogs.login_refusal(pid, [c for _e, c in ex7s])
        if not ex7s:
            continue
        app = None
        for ex7, _cap in ex7s:                    # some windows (e.g. an add-in shown read-only) give no object
            app = _app_from_window(ex7)
            if app is not None:
                break
        out.append(XlInstance(pid, h, [c for _e, c in ex7s], app))
    return out


def interactive_excels() -> list[tuple[int, int, object]]:
    """[(pid, main hwnd, Application)] of the visible, user-started Excel instances that answer COM, most recently
    active first (excel_instances() also lists the ones that are busy)."""
    return [(i.pid, i.hwnd, i.app) for i in excel_instances() if i.app is not None]


def _caption_book(caption: str) -> str:
    """Workbook name in a window caption: 'Book  -  Read-Only  -  Compatibility Mode' / 'Book.xlsx:2' -> 'book'."""
    s = str(caption).split("  -  ")[0].strip()
    s = re.sub(r":\d+$", "", s)
    return re.sub(r"\.(xlsx|xlsm|xlsb|xls|xltx|xltm)$", "", s, flags=re.I).lower()


def busy_holder(instances: list[XlInstance], name: str) -> XlInstance | None:
    """The Excel that does not answer COM but shows a window for workbook `name` (a file name or path)."""
    want = _caption_book(Path(str(name)).name)
    return next((i for i in instances if i.app is None and any(_caption_book(c) == want for c in i.captions)), None)


def busy_error(pid: int, name: str) -> HostError:
    """Name what blocks the Excel holding `name`: a dialog (dialogs.scan) or else cell edit mode."""
    try:
        found = dialogs.scan(pid)
    except Exception:   # noqa: BLE001
        found = []
    if found:
        return HostError(f"'{name}' is open in an Excel (pid {pid}) that has a dialog open "
                         f"({dialogs.describe(found[0])}). Close that dialog in Excel, then ask again.")
    return HostError(f"'{name}' is open in an Excel (pid {pid}) that is in cell edit mode or busy (no dialog seen). "
                     "Press Enter or Esc in that Excel window, then ask again.")


SHARING_VIOLATION = {32, 33}        # ERROR_SHARING_VIOLATION, ERROR_LOCK_VIOLATION


def _file_in_use(path: Path) -> bool:
    """Another process holds the file open (Excel locks the workbooks it has open): only a sharing/lock violation
    counts. A read-only file (access denied, e.g. a SharePoint file with read rights only) is NOT in use - Excel
    opens it read-only."""
    try:
        h = win32file.CreateFile(str(path), win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                                 win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE,
                                 None, win32file.OPEN_EXISTING, 0, None)
    except pywintypes.error as e:
        return e.winerror in SHARING_VIOLATION
    h.Close()
    return False


def _locked_error(path: Path) -> HostError:
    return HostError(f"'{path.name}' is open in another program, or in an Excel this engine cannot reach (the file is "
                     "locked). Close it there (or save and close it in that Excel), then ask again.")


ARCHIVE_URL_MARK = "/sites/analysetelenornorgeas/"     # SharePoint site behind the Produktregnskapet archive


def _onedrive_mounts() -> list[tuple[str, str]]:
    """[(url namespace, local mount)] from OneDrive's sync registry, longest URL first."""
    import winreg
    out = []
    try:
        root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\SyncEngines\Providers\OneDrive")
    except OSError:
        return out
    i = 0
    while True:
        try:
            sub = winreg.EnumKey(root, i)
        except OSError:
            break
        i += 1
        try:
            k = winreg.OpenKey(root, sub)
            url = winreg.QueryValueEx(k, "UrlNamespace")[0]
            mount = winreg.QueryValueEx(k, "MountPoint")[0]
            if url and mount:
                out.append((url, mount))
        except OSError:
            continue
    return sorted(set(out), key=lambda x: -len(x[0]))


def local_path(full_name: str) -> str:
    """Workbook.FullName of a OneDrive/SharePoint file is an https URL; map it to the synced local path."""
    s = str(full_name)
    if not s.lower().startswith("http"):
        return s
    from urllib.parse import unquote
    for url, mount in _onedrive_mounts():
        if s.lower().startswith(url.lower()):
            rest = unquote(s[len(url):]).replace("/", os.sep)
            return os.path.join(mount, rest)
    return s


def _norm(p) -> str:
    return os.path.normcase(os.path.abspath(local_path(str(p))))


def is_archive(path) -> bool:
    s = str(path)
    if ARCHIVE_URL_MARK in s.replace("\\", "/").lower():
        return True
    local = local_path(s)
    if not os.path.isabs(local):                  # e.g. an unsaved 'Book1': not a file anywhere
        return False
    try:
        return _norm(local).startswith(_norm(config.ARCHIVE_ROOT) + os.sep)
    except Exception:   # noqa: BLE001
        return False


# ------------------------------------------------------------------ sensitivity label
def _label(wb):
    """Our sensitivity label on a workbook in the user's Excel, through Office's label API (target._save_new: a new
    workbook added from the template, before its SaveAs); a workbook already labelled is left as it is."""
    try:
        cur = wb.SensitivityLabel.GetLabel()
        if cur.LabelId:
            return
    except Exception:   # noqa: BLE001
        pass
    li = wb.SensitivityLabel.CreateLabelInfo()
    li.LabelId = config.LABEL_ID
    li.SiteId = config.LABEL_SITE
    li.AssignmentMethod = 1
    li.Justification = ""
    wb.SensitivityLabel.SetLabel(li, li)
