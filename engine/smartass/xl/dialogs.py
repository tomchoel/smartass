"""Dialogs in the user's Excel, found by window enumeration only (no COM): which dialog blocks an Excel that holds the
target (instances.busy_error names it in the reply), and Smart View's Login.

smartass uses no Smart View any more, but Smart View stays installed in the user's Excel and can show its own EPM
Login - at every Excel start when its start-up update check is on (doctor: smartview_startup). COM into an Excel whose
Login is open blocks, so a job never touches such an Excel: it is refused at once with a text that says which Excel
and what to do (refuse_login, login_refusal), and after smartass launches an Excel it watches LAUNCH_GRACE_S for a
Login before its first COM call (watch_after_start). Nothing here clicks, closes or waits for a sign-in.
classify is pure; the rest is thin Win32."""
from __future__ import annotations

import time
from dataclasses import dataclass

from ..core.errors import HostError

DIALOG_CLASSES = ("#32770", "NUIDialog")      # Win32 message boxes and Excel's own alerts
LOGIN_TITLES = ("login", "sign in")
LAUNCH_GRACE_S = 12            # a fresh Excel shows Smart View's Login at start-up (seen 19-22 s after; checked 12 s)


@dataclass(frozen=True)
class Dialog:
    hwnd: int
    cls: str
    title: str
    text: str
    kind: str              # 'login' | 'message'


def classify(cls: str, title: str, text: str) -> str | None:
    if (title or "").strip().lower() in LOGIN_TITLES:
        return "login"
    if cls in DIALOG_CLASSES:
        return "message"
    return None


def describe(d: Dialog) -> str:
    return f"'{d.title}: {d.text}'" if d.text else f"'{d.title}'"


def scan(pid: int) -> list[Dialog]:
    """Visible top-level dialogs of process `pid`."""
    out: list[Dialog] = []
    for hwnd, cls, title in windows(pid=pid):
        if classify(cls, title, "") is None:
            continue
        kids = [(c, t) for _h, c, t in windows(parent=hwnd)]
        text = " ".join(t.strip() for c, t in kids if c == "Static" and t.strip())[:300]
        kind = classify(cls, title, text)
        if kind:
            out.append(Dialog(hwnd, cls, title, text, kind))
    return out


def windows(pid: int | None = None, parent: int | None = None) -> list[tuple[int, str, str]]:
    """(hwnd, class, title) of the visible top-level windows of process `pid`, or of all child windows of `parent`."""
    import pywintypes
    import win32gui
    import win32process

    out: list[tuple[int, str, str]] = []

    def cb(h, _):
        try:
            if parent or (win32gui.IsWindowVisible(h) and win32process.GetWindowThreadProcessId(h)[1] == pid):
                out.append((h, win32gui.GetClassName(h), win32gui.GetWindowText(h)))
        except pywintypes.error:
            pass
        return True
    try:
        if parent:
            win32gui.EnumChildWindows(parent, cb, None)
        else:
            win32gui.EnumWindows(cb, None)
    except pywintypes.error:
        pass
    return out


def login_open(pid: int) -> list[int]:
    """Smart View's Login windows open in Excel `pid` (empty = none)."""
    return [d.hwnd for d in scan(pid) if d.kind == "login" and d.cls == "#32770" and d.title == "Login"]


STARTUP_HINT = ("If it shows at every Excel start, untick 'Check for updates when Microsoft Office starts' in Smart "
                "View > Options > Extensions (close Excel first).")


def titles(pid: int) -> list[str]:
    """The titles of Excel `pid`'s main windows ('Budget.xlsx - Excel'), read without COM."""
    return [t for _h, cls, t in windows(pid=pid) if cls == "XLMAIN" and t]


def login_refusal(pid: int, captions: list[str] | None = None) -> HostError:
    """The refusal for an Excel whose Smart View Login is open: which Excel (captions: its workbook windows' titles;
    else its main windows' titles - all read without COM) and what to do."""
    shown = [c for c in (titles(pid) if captions is None else captions) if c]
    which = f"the Excel showing {', '.join(repr(c) for c in shown[:3])} (pid {pid})" if shown else f"Excel (pid {pid})"
    return HostError(f"Smart View's Login is open in {which}; smartass does not touch an Excel while it is (COM "
                     "into it would block). Next step: sign in there or close that Login window, then ask again. "
                     + STARTUP_HINT)


def refuse_login(pid: int) -> None:
    """Refuse (login_refusal) when Smart View's Login is open in Excel `pid`: checked once, before the first COM."""
    if login_open(pid):
        raise login_refusal(pid)


def watch_after_start(pid: int, grace: float = LAUNCH_GRACE_S, alive=None, clock=None, sleep=None) -> None:
    """An Excel smartass just launched: watched for grace seconds (Smart View's Login may come right after its start)
    before the first COM call; a Login seen is refused at once (login_refusal). alive: the job's abort check, each
    round."""
    clock, sleep = clock or time.monotonic, sleep or time.sleep
    t0 = clock()
    while True:
        if alive is not None:
            alive()
        refuse_login(pid)
        if clock() - t0 >= grace:
            return
        sleep(0.5)
