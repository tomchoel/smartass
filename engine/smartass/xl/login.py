"""Smart View login state in the user's Excel: the Login dialog, the SSO box watcher, IE-login recovery."""
from __future__ import annotations

import threading
import time

from . import jobs
from . import dialogs
from ..core import timing
from ..core.errors import HostError

LOGIN_NOTICE_S = 5            # a Login open this long: tell the user we wait for the SSO click
LOGIN_HUNG_S = 10             # a Login not responding this long: fail fast (nothing in the user's Excel is closed)
LOGIN_HUNG = ("The Smart View Login window in your Excel is not responding; nothing was closed. Close that Login "
              "window (or the Excel it belongs to), then ask again.")


def login_open(pid: int) -> list[int]:
    """The Smart View Login windows open in Excel `pid` (empty = none)."""
    return [d.hwnd for d in dialogs.scan(pid) if d.kind == "login" and d.cls == "#32770" and d.title == "Login"]


BOX_POLL_S = 0.5


class SmartViewBoxes:
    """While the engine drives Smart View in a visible Excel: a thread clicks OK on Smart View's modal message boxes
    in that Excel (they would block the call until the user clicks) and keeps their texts for the caller."""

    def __init__(self, pid: int):
        self.pid, self.texts = pid, []
        self._stop, self._seen = threading.Event(), set()
        self._thread = threading.Thread(target=self._run, name=f"sv-boxes-{pid}", daemon=True)

    def _run(self):
        while not self._stop.wait(BOX_POLL_S):
            try:
                boxes = dialogs.sv_boxes(dialogs.scan(self.pid))
                self._seen &= {d.hwnd for d in boxes}     # a reused window handle is a new box
                for d in boxes:
                    if d.hwnd not in self._seen:
                        self._seen.add(d.hwnd)
                        if dialogs.press_ok(d.hwnd):      # information boxes only, never a question
                            self.texts.append(d.text)
                            jobs.log.warning("Smart View message in Excel pid %s dismissed: %s", self.pid, d.text)
            except Exception:   # noqa: BLE001 - the guard must never break the pull
                jobs.log.debug("Smart View box guard failed", exc_info=True)

    def __enter__(self):
        self._thread.start()
        return self

    def __exit__(self, *exc):
        self._stop.set()
        self._thread.join(timeout=2)


NOT_SIGNED_IN = ("Smart View in this Excel is not signed in correctly: EPM answers 'Empty response from provider' to "
                 "every connect, and Smart View shows no Login to click. Ask the user to sign in once in that Excel "
                 "(Smart View ribbon -> Panel -> connect FinMax, or close that Excel so the next call starts a fresh "
                 "one), then ask again. If every Excel does this and no Login appears anywhere, Smart View's shared "
                 "sign-in is stale: close all Excels, rename the Cookies files in %APPDATA%\\Oracle\\SmartView\\WebView"
                 "\\EBWebView\\Default\\Network, then retry (docs\\troubleshooting.md, Known limits).")


RECOVER_SETTLE_S = 12         # after a login recovery: time for Smart View to reopen the login before we go on


def recover_ie_login(pid: int, hwnds: list[int]) -> dict:
    """The login in Excel `pid` is on the Internet Explorer fallback because an earlier, unfinished WebView2 login
    still holds Smart View's WebView2 folder. End that stale WebView2 (this Excel's own, or an orphan - never another
    app's) and close the broken IE windows; Smart View then opens a new login in WebView2. Excel itself is untouched."""
    ended = dialogs.end_stale_sv_webview2(pid)
    dialogs.close(hwnds)
    return {"ended": ended, "closed": hwnds}


def wait_login_cleared(pid: int, appear_grace: float = 0.0, limit: float = 90.0, clock=None, sleep=None) -> None:
    """Block until no Smart View Login dialog is open in `pid` (the global SSO watcher clicks it). A login that fell
    back to Internet Explorer ('Script Error') never completes: it is recovered (recover_ie_login), at most twice.
    The user is told once after LOGIN_NOTICE_S; a Login not responding for LOGIN_HUNG_S fails fast (LOGIN_HUNG)."""
    clock, sleep = clock or time.monotonic, sleep or time.sleep
    t0 = clock()
    seen = False
    recoveries, t_rec, opened, noticed, hung_since = 0, None, None, False, None
    while clock() - t0 < limit:
        jobs._check_alive()                       # an abandoned job stops waiting
        now = clock()
        if hwnds := login_open(pid):
            opened = now if opened is None else opened
            if not noticed and now - opened >= LOGIN_NOTICE_S:
                noticed = True
                timing.say("Smart View's Login is open in your Excel; waiting for the automatic SSO click")
            hung = any(dialogs.is_hung(w) for w in hwnds)
            hung_since = (now if hung_since is None else hung_since) if hung else None
            if hung_since is not None and now - hung_since >= LOGIN_HUNG_S:
                raise HostError(LOGIN_HUNG)
            broken = dialogs.ie_fallback(dialogs.scan(pid))
            if not broken:                        # a real (WebView2) login: the watcher clicks it
                seen = True
            # a second recovery only after RECOVER_SETTLE_S: a new WebView2 login of this Excel must not be ended
            elif recoveries < 2 and (t_rec is None or now - t_rec >= RECOVER_SETTLE_S):
                recoveries += 1
                res = recover_ie_login(pid, broken)
                jobs.log.warning("Smart View login in Excel pid %s fell back to Internet Explorer; ended stale WebView2 %s "
                            "and closed %s so it can sign in again", pid, res["ended"], res["closed"])
                seen, t_rec = False, clock()   # Smart View reopens the login (in WebView2): wait for it
                continue
        elif t_rec is not None:
            if seen or now - t_rec >= RECOVER_SETTLE_S:   # the new login came and went, or none came back
                return
        elif seen or now - t0 >= appear_grace:
            return
        sleep(0.5)
    raise HostError("The Oracle EPM login dialog is still open in your Excel and the automatic SSO click did not "
                    "complete. Ask the user to click 'Telenor SSO' in it, then ask again. If it keeps coming back, "
                    "Smart View's WebView2 may be held by another Office app (often an old PowerPoint): closing that "
                    "app usually fixes it.")
