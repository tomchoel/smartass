"""Smart View dialogs in our hidden Excel: find, classify, and decide whether the sign-in is stuck.

Smart View signs in through a 'Login' window. Normally it hosts WebView2 and the global SSO watcher clicks
'Telenor SSO' within 1-20 s. When WebView2 fails, Smart View says so in a message box and falls back to Internet
Explorer, whose Oracle IDCS page fails ('Script Error: Function expected'). In our hidden Excel that Excel is killed and
restarted. In a visible Excel the usual cause is an earlier, unfinished WebView2 login of the same Excel still holding
the WebView2 folder: ending that stale WebView2 and closing the IE windows lets Smart View reopen the login in WebView2
(userexcel.recover_ie_login; WM_CLOSE closes the IE Login and 'Script Error' windows - seen live 2026-10-23).
classify/blocking/ie_fallback/stale_sv_webview2_from are pure; the rest is thin Win32.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

DIALOG_CLASSES = ("#32770", "NUIDialog")      # Win32 message boxes (Smart View's) and Excel's own alerts
LOGIN_TITLES = ("login", "sign in")
IE_CLASS = "Internet Explorer_Server"
SV_WEBVIEW_DIR = r"oracle\smartview\webview"    # Smart View's fixed WebView2 user data folder under %APPDATA%


@dataclass(frozen=True)
class Dialog:
    hwnd: int
    cls: str
    title: str
    text: str
    kind: str              # 'login' | 'webview2_fallback' | 'message'
    ie: bool = False       # a Login window rendered by Internet Explorer (WebView2 already failed)


def classify(cls: str, title: str, text: str) -> str | None:
    t = (title or "").strip().lower()
    if "webview2" in f"{title} {text}".lower() or t == "script error":
        return "webview2_fallback"
    if t in LOGIN_TITLES:
        return "login"
    if cls in DIALOG_CLASSES:
        return "message"
    return None


def blocking(found: list[Dialog], first_seen: dict[int, float], now: float, grace: float) -> Dialog | None:
    """The dialog that means the sign-in is stuck, or None. A WebView2 Login gets `grace` s for the SSO watcher;
    an Internet Explorer Login or a WebView2-fallback box is stuck at once. first_seen ({hwnd: time}) is updated."""
    for d in found:
        first_seen.setdefault(d.hwnd, now)
    for d in found:
        if d.kind == "webview2_fallback" or (d.kind == "login" and (d.ie or now - first_seen[d.hwnd] >= grace)):
            return d
    return None


def describe(d: Dialog) -> str:
    return f"'{d.title}: {d.text}'" if d.text else f"'{d.title}'"


def scan(pid: int) -> list[Dialog]:
    """Visible top-level dialogs of process `pid` (EnumWindows + GetWindowThreadProcessId)."""
    import pywintypes
    import win32gui
    import win32process

    out: list[Dialog] = []

    def cb(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd) or win32process.GetWindowThreadProcessId(hwnd)[1] != pid:
                return True
            cls, title = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
            if classify(cls, title, "") is None:
                return True
            kids = _children(hwnd)
            text = " ".join(t.strip() for c, t in kids if c == "Static" and t.strip())[:300]
            kind = classify(cls, title, text)
            if kind:
                out.append(Dialog(hwnd, cls, title, text, kind, any(c == IE_CLASS for c, _t in kids)))
        except pywintypes.error:
            pass
        return True
    try:
        win32gui.EnumWindows(cb, None)
    except pywintypes.error:
        pass
    return out


def _children(hwnd: int) -> list[tuple[str, str]]:
    import pywintypes
    import win32gui

    out: list[tuple[str, str]] = []

    def cb(h, _):
        try:
            out.append((win32gui.GetClassName(h), win32gui.GetWindowText(h)))
        except pywintypes.error:
            pass
        return True
    try:
        win32gui.EnumChildWindows(hwnd, cb, None)
    except pywintypes.error:
        pass
    return out


def webview2_holders_from(rows: list[tuple[int, int, str]], exclude_parent: int | None) -> list[str]:
    """Apps (exe names) whose WebView2 browser process holds Smart View's WebView2 folder, from
    [(pid, parent pid, command line)] of msedgewebview2.exe; our own Excel (exclude_parent) is left out."""
    out: list[str] = []
    for _pid, ppid, cmd in rows:
        low = (cmd or "").lower()
        if "--type=" in low or SV_WEBVIEW_DIR not in low or ppid == exclude_parent:
            continue
        m = re.search(r"--webview-exe-name=(\S+)", cmd)
        name = m.group(1) if m else "?"
        if name not in out:
            out.append(name)
    return out


SV_BOX_TITLE = "Oracle Smart View for Office"


def sv_boxes(found: list[Dialog]) -> list[Dialog]:
    """Smart View's own modal message boxes (e.g. 'Empty response from provider.'). In our hidden Excel they are
    switched off (option 104); in a visible Excel they block the call that raised them until someone clicks OK."""
    return [d for d in found if d.kind == "message" and d.title == SV_BOX_TITLE]


def not_signed_in(attempts: list[list[str]]) -> bool:
    """EPM answered 'Empty response from provider' to two connect attempts (attempts = the box texts each failed
    attempt raised): this Excel's Smart View session is not valid (it shows no Login, so the SSO watcher cannot
    help). Retrying does not fix it; a manual sign-in does."""
    return sum(any("empty response from provider" in (t or "").lower() for t in a) for a in attempts) >= 2


def ok_only(buttons: list[str]) -> bool:
    """A box that only informs (its single button is OK) - safe to dismiss. Never OK/Cancel or Yes/No questions."""
    return [b.replace("&", "").strip().lower() for b in buttons] == ["ok"]


def buttons(hwnd: int) -> list[str]:
    return [t for c, t in _children(hwnd) if c == "Button"]


def press_ok(hwnd: int) -> bool:
    """Click OK on an information box; False (nothing done) for any box with other buttons."""
    import pywintypes
    import win32con
    import win32gui

    try:
        if not ok_only(buttons(hwnd)):
            return False
        ok = win32gui.FindWindowEx(hwnd, 0, "Button", None)
        win32gui.PostMessage(ok, win32con.BM_CLICK, 0, 0)
        return True
    except pywintypes.error:
        return False


def ie_fallback(found: list[Dialog]) -> list[int]:
    """Windows of a sign-in that already fell back to Internet Explorer (an IE-rendered Login, a 'Script Error' or a
    WebView2-failed box): broken for good, so they can be closed. Empty while a WebView2 Login is still in progress."""
    return [d.hwnd for d in found if d.kind == "webview2_fallback" or (d.kind == "login" and d.ie)]


def stale_sv_webview2_from(rows: list[tuple[int, int, str]], excel_pid: int, alive) -> list[int]:
    """Smart View WebView2 browser processes to end so the next login can start WebView2 again, from
    [(pid, parent pid, command line)]: the ones of this Excel (its login is already on the IE fallback, so they are
    a stuck earlier login) and orphans whose app is gone. Never one owned by another live app (e.g. PowerPoint)."""
    out = []
    for pid, ppid, cmd in rows:
        low = (cmd or "").lower()
        if "--type=" in low or SV_WEBVIEW_DIR not in low:
            continue
        if ppid == excel_pid or not alive(ppid):
            out.append(pid)
    return out


def end_stale_sv_webview2(excel_pid: int) -> list[int]:
    """End the stale Smart View WebView2 browser processes (with their child processes); returns their pids."""
    import psutil

    rows = []
    for p in psutil.process_iter(["name", "ppid", "cmdline"]):
        try:
            if (p.info["name"] or "").lower() == "msedgewebview2.exe":
                rows.append((p.pid, p.info["ppid"], " ".join(p.info["cmdline"] or [])))
        except psutil.Error:
            continue
    ended = []
    for pid in stale_sv_webview2_from(rows, excel_pid, psutil.pid_exists):
        try:
            p = psutil.Process(pid)
            for k in p.children(recursive=True) + [p]:
                k.kill()
            ended.append(pid)
        except psutil.Error:
            continue
    return ended


def close(hwnds: list[int]) -> None:
    import pywintypes
    import win32con
    import win32gui

    for h in hwnds:
        try:
            win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
        except pywintypes.error:
            pass


def webview2_holders(exclude_parent: int | None) -> list[str]:
    import psutil

    rows = []
    for p in psutil.process_iter(["name", "ppid", "cmdline"]):
        try:
            if (p.info["name"] or "").lower() == "msedgewebview2.exe":
                rows.append((p.pid, p.info["ppid"], " ".join(p.info["cmdline"] or [])))
        except psutil.Error:
            continue
    return webview2_holders_from(rows, exclude_parent)
