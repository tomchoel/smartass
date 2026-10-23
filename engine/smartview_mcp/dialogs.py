"""Smart View dialogs in our hidden Excel: find, classify, and decide whether the sign-in is stuck.

Smart View signs in through a 'Login' window. Normally it hosts WebView2 and the global SSO watcher clicks
'Telenor SSO' within 1-20 s. When WebView2 fails, Smart View says so in a message box and falls back to Internet
Explorer, whose Oracle IDCS page fails ('Script Error: Function expected'); that Login window is modal, ignores
WM_CLOSE, and only killing its Excel gets out. classify/blocking are pure; scan is the thin Win32 part.
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
