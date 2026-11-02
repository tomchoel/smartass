"""Several drill pages at once in the hidden browser: each drill opens in its own tab of the signed-in context, so
the ~5 s each drill page needs to load and ask for its data overlap (measured 02.11.2026: a drill was ~5 s page
load + ~0.1 s per 500-line page; four drills of un.SY took 220 s one after the other).

All Playwright work stays on the browser's ONE worker thread: a pump job keeps up to config.DRILL_TABS tabs going,
takes new drills from a queue while it runs, and ends when none is left. A tab that does not capture cleanly (sign-in
redirect, slow page, expired session while paging, a crash) falls back to DrillBrowser._fetch - the one-tab route
with its re-login, reload and browser-restart rules. Read-only as before: each tab replays only its own page's
getData request (limit/offset changed)."""
from __future__ import annotations

import collections
import concurrent.futures as cf
import logging
import threading
import time

from . import config
from .drill_browser import (SIGNIN, CellTooWide, DrillBrowser, DrillError, TooManyLines, load_timeout_ms,
                            page_through, redact)

log = logging.getLogger("smartview.drill")
POLL_MS = 200


class _Slot:
    """One drill in its own tab: its captured request/view, and when it started."""

    def __init__(self, url: str, max_lines: int, fut: cf.Future):
        self.url, self.max_lines, self.fut = url, max_lines, fut
        self.tab = None
        self.got: dict = {}
        self.t0 = 0.0

    def on_request(self, req):
        if "/aif/ui/model/drilldown/getData?limit=" in req.url and "req" not in self.got:
            self.got["req"] = {"url": req.url, "body": req.post_data, "headers": dict(req.headers)}

    def on_response(self, resp):
        if "/aif/ui/model/view?page=DRILL_THROUGH" in resp.url and "view" not in self.got:
            try:
                self.got["view"] = resp.json()
            except Exception:   # noqa: BLE001 - the page asks again or the capture times out
                pass


class TabPump:
    def __init__(self, browser, tabs: int | None = None):
        self._browser = browser                   # callable -> the (current) shared DrillBrowser
        self.tabs = max(1, tabs or config.DRILL_TABS)
        self._mu = threading.Lock()
        self._pending: collections.deque = collections.deque()
        self._job: cf.Future | None = None
        self._cancel = threading.Event()

    def fetch(self, url: str, max_lines: int) -> cf.Future:
        """Queue one drill; the future gets {items, total, view, secs} or the drill's exception."""
        fut: cf.Future = cf.Future()
        with self._mu:
            self._pending.append((url, max_lines, fut))
            if self._job is not None:
                return fut
            b = self._browser()
            self._cancel.clear()
            job = self._job = b.submit(lambda s: self._run(b, s))
        job.add_done_callback(self._ended)
        return fut

    def cancel_all(self, why: str) -> None:
        """Hang guard: fail every queued and running drill (the pump stops at its next poll)."""
        self._cancel.set()
        self._fail_pending(DrillError(why))

    def _fail_pending(self, err: BaseException) -> None:
        with self._mu:
            left, self._pending = list(self._pending), collections.deque()
        for _, _, fut in left:
            if not fut.done() and (fut.running() or fut.set_running_or_notify_cancel()):
                fut.set_exception(err)

    def _ended(self, job: cf.Future) -> None:
        """The pump job finished: if it never ran (cancelled by a browser kill) or died, nothing may wait forever."""
        with self._mu:
            if self._job is not job:
                return
            self._job = None
        err = job.exception() if not job.cancelled() else None
        self._fail_pending(err or DrillError("the hidden browser was restarted; drill again"))

    # ------------------------------------------------------------ worker thread
    def _run(self, b: DrillBrowser, s) -> None:
        active: list[_Slot] = []
        try:
            while True:
                if self._cancel.is_set():
                    raise DrillError("the drill was stopped (hang guard)")
                with self._mu:
                    while len(active) < self.tabs and self._pending:
                        url, ml, fut = self._pending.popleft()
                        if fut.set_running_or_notify_cancel():
                            active.append(_Slot(url, ml, fut))
                    if not active:
                        self._job = None              # new drills from here on start a new pump job
                        return
                for slot in list(active):
                    if self._step(b, s, slot):
                        active.remove(slot)
                        self._close(slot)
                if active:
                    tab = next((x.tab for x in active if x.tab is not None and not x.tab.is_closed()), None)
                    (tab or s.page).wait_for_timeout(POLL_MS)    # lets every tab's events arrive
        except BaseException as e:
            for slot in active:
                self._close(slot)
                if not slot.fut.done():
                    slot.fut.set_exception(e if isinstance(e, Exception) else DrillError(str(e)))
            raise

    def _step(self, b: DrillBrowser, s, slot: _Slot) -> bool:
        """Advance one drill; True when it is finished (result or exception set)."""
        try:
            if slot.tab is None:
                return self._open(b, s, slot)
            if "req" in slot.got and "view" in slot.got:
                return self._page(b, s, slot)
            if SIGNIN in slot.tab.url:
                return self._fallback(b, s, slot, "signed out")
            epm = b._epm_error(slot.tab)
            if epm:
                raise CellTooWide(epm) if "EPMFDM-140373" in epm else DrillError(f"EPM refused the drill: {epm}")
            if time.monotonic() - slot.t0 >= config.DRILL_CAPTURE_TIMEOUT_S:
                return self._page(b, s, slot) if "req" in slot.got else self._fallback(b, s, slot, "no data request")
            return False
        except (TooManyLines, CellTooWide, DrillError) as e:
            slot.fut.set_exception(e)
            return True
        except Exception as e:   # noqa: BLE001 - a Playwright error in this tab: the one-tab route decides
            return self._fallback(b, s, slot, redact(str(e).splitlines()[0] if str(e) else type(e).__name__))

    def _open(self, b: DrillBrowser, s, slot: _Slot) -> bool:
        s.deadline = time.monotonic() + config.DRILL_SETUP_S
        b._ensure(s)                                  # start + sign in (once) before the first tab
        slot.tab = s.ctx.new_page()
        slot.tab.on("request", slot.on_request)
        slot.tab.on("response", slot.on_response)
        slot.tab.goto(b.base + "/" + slot.url.lstrip("./"), wait_until="commit",
                      timeout=load_timeout_ms(s.deadline, time.monotonic()))
        slot.t0 = time.monotonic()
        return False

    def _page(self, b: DrillBrowser, s, slot: _Slot) -> bool:
        t1 = time.monotonic()
        try:
            items, total = page_through(lambda off: DrillBrowser._post(slot.tab, slot.got["req"], off),
                                        slot.max_lines)
        except TooManyLines:
            raise
        except DrillError as e:                       # e.g. signed out while paging: the main tab re-logs in
            return self._fallback(b, s, slot, str(e))
        slot.fut.set_result({"items": items, "total": total, "view": slot.got.get("view"),
                             "secs": round(time.monotonic() - t1, 1)})
        return True

    def _fallback(self, b: DrillBrowser, s, slot: _Slot, why: str) -> bool:
        log.info("drill tab: %s; drilling this cell in the main tab", why)
        self._close(slot)
        try:
            slot.fut.set_result(b._fetch(s, slot.url, slot.max_lines, None))
        except Exception as e:   # noqa: BLE001 - forwarded to the drill
            slot.fut.set_exception(e)
        return True

    @staticmethod
    def _close(slot: _Slot) -> None:
        tab, slot.tab = slot.tab, None
        if tab is not None:
            try:
                tab.close()
            except Exception:   # noqa: BLE001 - a dead browser
                pass


_shared: TabPump | None = None
_shared_lock = threading.Lock()


def shared_pump() -> TabPump:
    """The process's one pump, on the shared drill browser (looked up whenever a pump job starts)."""
    global _shared
    with _shared_lock:
        if _shared is None:
            from .drill_browser import shared_browser
            _shared = TabPump(shared_browser)
        return _shared
