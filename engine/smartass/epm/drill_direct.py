"""Drill lines without loading a drill page per piece (spike 2026-11-02, docs/audit/2026-11-02-drill-speed-spike.md).

Per piece, from ONE warm tab (a tab that loaded one real drill page): GET the drill URL's JSON form
(&FORMAT=SUI) - EPM's own drill query - and page getData with it, field for field ({QUERY, PARAMS: PARAMSARRAY,
PARAMSTYPE: PARAMSTYPEARRAY}). Read-only: getData runs EPM's SELECT; only a split's siblings get a query built from
EPM's (drill_derive, tied to the cube). getData from a tab that never loaded a drill page gets the whole EPM session
signed out, so it is only ever sent from the warm tab."""
from __future__ import annotations

import asyncio
import json
import logging
import threading
import time
from typing import Callable

from ..core import config
from .drill_browser import shared_browser
from .drillcore import CellTooWide, DrillError, NoDrillData, Pager, SignedOut, SuiQuery, parse_sui, sui_url  # noqa: F401

log = logging.getLogger("smartview.drill")


_SUI_HEADERS = {"X-Requested-With": "XMLHttpRequest", "Accept": "*/*", "Content-Type": "application/json"}
_DROP = ("cookie", "content-length", "host")
_REQUEST_S = 120.0                 # one SUI / view / getData request (it aborts itself then)


class DirectDrill:
    """The warm tab, its getData headers and the view cache (worker-thread state), and the async drill on top.
    Every request is in flight in a tab (DrillBrowser.request), so many pieces run at once: Smart View-free SUI
    queries queue at EPM (one at a time per user, ~0.7 s each, spike 2026-11-03) while getData pages run together.

    INVARIANT: getData is POSTed only from self.tab, a tab that _warm() saw load a real drill page. getData from
    any other tab gets the whole EPM session signed out (spike section 2). The SUI GET is safe from any signed-in tab."""

    def __init__(self, capture_timeout_s: float | None = None, browser: Callable[[], object] | None = None):
        self.tab = None
        self.headers: dict[str, str] = {}
        self.views: dict[tuple, dict] = {}
        self.gen = 0                                   # warm tabs so far: a failure seen on an older one is stale
        self._ctx = None                               # the browser context the warm tab belongs to
        self._prewarming = False
        self._prewarm_lock = threading.Lock()
        self.capture_timeout_s = config.DRILL_CAPTURE_TIMEOUT_S if capture_timeout_s is None else capture_timeout_s
        self._browser = browser or shared_browser

    # ---------------------------------------------------------------- worker thread
    def _warm(self, b, s, url: str) -> None:
        """Load one real drill page in a new tab and keep it with the headers of its own getData request."""
        p = s.ctx.new_page()
        got: dict = {}

        def on_request(req):
            if "/aif/ui/model/drilldown/getData?limit=" in req.url and "h" not in got:
                got["h"] = {k: v for k, v in dict(req.headers).items() if k.lower() not in _DROP}
        p.on("request", on_request)
        try:
            p.goto(b.base + "/" + url.lstrip("./"), wait_until="domcontentloaded", timeout=60_000)
            t0 = time.monotonic()
            while "h" not in got and time.monotonic() - t0 < self.capture_timeout_s:
                p.wait_for_timeout(50)
                collect = getattr(b, "_collect", None)          # answers of requests in flight keep coming in
                if collect is not None and getattr(s, "inflight", None):
                    collect(s)
        finally:
            p.remove_listener("request", on_request)
        if "h" not in got:
            p.close()
            raise DrillError("The warm drill page never asked for its data; the page route takes over.")
        self.tab, self.headers, self._ctx = p, got["h"], s.ctx
        self.gen += 1
        log.info("direct drill: warm tab ready")

    @staticmethod
    def _signin_time(s) -> None:
        """The drill's own sign-in budget (an older job's deadline may have passed: _login would refuse at once)."""
        s.deadline = time.monotonic() + config.DRILL_SETUP_S

    def _ensure_warm(self, b, s, url: str) -> None:
        self._signin_time(s)
        b._ensure(s)
        if not self._alive(s):
            self._warm(b, s, url)

    def prewarm(self, url: str) -> None:
        """Warm the tab now (on the worker, not awaited) so it is ready when the first lines are asked for; the
        drill queries run meanwhile. Only with the link of a cell whose drill query EPM answered: a too-wide
        cell's page never asks for data and would hold the worker for the whole capture timeout. Once per tab: later
        calls do nothing while a warm is queued or the tab lives. A failure is only logged - lines() warms again and
        reports it."""
        with self._prewarm_lock:
            if self._prewarming or self._alive():
                return
            self._prewarming = True
        b = self._browser()

        def job(s):
            try:
                self._ensure_warm(b, s, url)
            except Exception as e:   # noqa: BLE001 - lines() tries again and reports
                log.info("direct drill: prewarm failed (%s)", e)

        def done(_f):                                  # also when the job was cancelled and never ran
            self._prewarming = False
        b.submit(job).add_done_callback(done)

    def _rewarm(self, b, s, url: str, seen_gen: int) -> None:
        """Signed out under a piece: sign in again and warm a new tab - unless another piece already did."""
        if self.gen != seen_gen and self._alive(s):
            return
        self._drop_tab()
        self._signin_time(s)
        b._login(s)
        self._warm(b, s, url)

    def _drop_tab(self) -> None:
        if self.tab is not None:
            try:
                self.tab.close()
            except Exception:   # noqa: BLE001 - already closed
                pass
        self.tab, self.headers, self.views = None, {}, {}

    def _alive(self, s=None) -> bool:
        """A live warm tab - of the current browser session when s is given (a killed or replaced browser leaves a
        page object behind that may not know it is gone)."""
        if self.tab is None or (s is not None and self._ctx is not s.ctx):
            return False
        try:
            return not self.tab.is_closed()
        except Exception:   # noqa: BLE001 - a page of a dead driver
            return False

    def _sui_headers(self) -> dict:
        extra = {k: v for k, v in self.headers.items() if k.lower() in ("page_id", "x-appbuilder-client-id")}
        return {**_SUI_HEADERS, **extra}

    def _warm_tab(self, s):
        """The warm tab for a getData or view request; never another tab (INVARIANT)."""
        if not self._alive(s):
            raise DrillError("the warm drill tab is gone")
        return self.tab

    # ---------------------------------------------------------------- any thread / event loop
    async def _ask(self, b, prep) -> tuple[int, str]:
        return await asyncio.wrap_future(b.request(prep, _REQUEST_S))

    async def query(self, url: str) -> SuiQuery:
        """EPM's own drill query for a drill link (SUI GET; safe from any signed-in tab). A request lost with its
        tab, or a sign-in page instead of JSON, is asked once more."""
        b = self._browser()

        def prep(s):
            self._signin_time(s)
            b._ensure(s)
            tab = self.tab if self._alive(s) else s.page
            return tab, sui_url(b.base, url, int(time.time() * 1000)), {"headers": self._sui_headers()}
        for attempt in (1, 2):
            try:
                _, text = await self._ask(b, prep)
                q = parse_sui(text)                    # NoDrillData / CellTooWide / SignedOut / DrillError
                self.prewarm(url)                      # a drillable cell: its page will ask for data
                return q
            except (NoDrillData, CellTooWide):
                raise
            except SignedOut as e:
                why = e
            except DrillError as e:
                if str(e).startswith("EPM refused the drill"):
                    raise
                why = e                                # the request was lost with its tab
            if attempt == 2:
                raise why
            log.info("direct drill: drill query lost (%s); asking once more", why)
        raise AssertionError("unreachable")

    async def lines(self, q: SuiQuery, warm_url: str, max_lines: int,
                    progress: Callable[[int, int | None], None] | None = None) -> dict:
        """The lines of a drill query from the warm tab (warmed with warm_url if needed); one re-sign-in."""
        b = self._browser()
        await asyncio.wrap_future(b.submit(lambda s: self._ensure_warm(b, s, warm_url)))
        key = (q.app_id, q.partition)
        if key not in self.views:
            _, v = await self._ask(b, lambda s: (self._warm_tab(s), f"{b.base}/aif/ui/model/view?page=DRILL_THROUGH"
                                                 f"&appId={q.app_id}&partitionkey={q.partition}"
                                                 f"&_={int(time.time() * 1000)}", {"headers": self._sui_headers()}))
            try:
                self.views[key] = json.loads(v)
            except ValueError:
                raise SignedOut("EPM answered the drill view with a page, not JSON (signed out?)") from None
        t1 = time.monotonic()
        pg = Pager(max_lines, progress, True)
        while True:
            gen, off = self.gen, pg.off
            try:
                status, text = await self._ask(b, lambda s: (
                    self._warm_tab(s), f"{b.base}/aif/ui/model/drilldown/getData?limit={config.DRILL_PAGE}"
                    f"&offset={off}", {"method": "POST", "body": q.body, "headers": self.headers}))
            except DrillError as e:                    # tab gone or navigated: like a signed-out answer
                status, text = 0, str(e)
            step = pg.take(status, text)
            if step == "recover":
                await asyncio.wrap_future(b.submit(lambda s, g=gen: self._rewarm(b, s, warm_url, g)))
            elif step == "done":
                break
        items, total = pg.result()
        return {"items": items, "total": total, "view": self.views.get(key), "secs": round(time.monotonic() - t1, 1),
                "query": q}

    async def fetch(self, url: str, max_lines: int,
                    progress: Callable[[int, int | None], None] | None = None) -> dict:
        """One drill piece: EPM's drill query, then its lines. The result carries the query ('query') so a split
        can build its siblings' queries from it (drill_derive)."""
        return await self.lines(await self.query(url), url, max_lines, progress)


_shared: DirectDrill | None = None
_shared_lock = threading.Lock()


def shared_direct() -> DirectDrill:
    global _shared
    with _shared_lock:
        if _shared is None:
            _shared = DirectDrill()
        return _shared
