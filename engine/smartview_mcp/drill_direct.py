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
from dataclasses import dataclass
from typing import Callable

from . import config
from .drill_browser import CellTooWide, DrillError, Pager, shared_browser

log = logging.getLogger("smartview.drill")


class NoDrillData(DrillError):
    """EPM's drill query has no lines here (data not loaded through Data Integration, or an empty cell)."""


class SignedOut(DrillError):
    """EPM answered with something that is not JSON (its sign-in / exit page)."""


@dataclass
class SuiQuery:
    body: str               # the getData POST body
    app_id: int | None
    partition: int | None


def sui_url(base: str, url: str, now_ms: int) -> str:
    """The drill URL's JSON form: DRILLURL dropped, FORMAT=SUI and a cache buster added (as the page asks it).

    Built by string operations, not parse/urlencode: every other parameter is passed through byte for byte, as the
    page sends Smart View's URL unchanged."""
    full = base.rstrip("/") + "/" + url.lstrip("./")
    full, _, frag = full.partition("#")
    path, _, query = full.partition("?")
    items = [p for p in query.split("&") if p and p.partition("=")[0] != "DRILLURL"]
    items += ["FORMAT=SUI", f"_={now_ms}"]
    return path + "?" + "&".join(items) + ("#" + frag if frag else "")


def parse_sui(text: str) -> SuiQuery:
    try:
        j = json.loads(text)
    except ValueError:
        raise SignedOut("EPM answered the drill query with a page, not JSON (signed out?)") from None
    if not isinstance(j, dict):
        raise SignedOut("EPM's drill query answer is not a JSON object")
    if "QUERY" not in j:
        why = str(j.get("details") or "no drill query in EPM's answer")
        if "EPMFDM-140373" in why:
            raise CellTooWide(why)
        if "does not return any data" in why:
            raise NoDrillData(why)
        raise DrillError(f"EPM refused the drill: {why}")
    body = json.dumps({"QUERY": j["QUERY"], "PARAMS": j.get("PARAMSARRAY", []),
                       "PARAMSTYPE": j.get("PARAMSTYPEARRAY", [])})
    return SuiQuery(body, j.get("APPLICATION_ID"), j.get("PARTITIONKEY"))


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
                p.wait_for_timeout(200)
        finally:
            p.remove_listener("request", on_request)
        if "h" not in got:
            p.close()
            raise DrillError("The warm drill page never asked for its data; the page route takes over.")
        self.tab, self.headers = p, got["h"]
        self.gen += 1
        log.info("direct drill: warm tab ready")

    def _ensure_warm(self, b, s, url: str) -> None:
        b._ensure(s)
        if not self._alive():
            self._warm(b, s, url)

    def _rewarm(self, b, s, url: str, seen_gen: int) -> None:
        """Signed out under a piece: sign in again and warm a new tab - unless another piece already did."""
        if self.gen != seen_gen and self._alive():
            return
        self._drop_tab()
        b._login(s)
        self._warm(b, s, url)

    def _drop_tab(self) -> None:
        if self.tab is not None:
            try:
                self.tab.close()
            except Exception:   # noqa: BLE001 - already closed
                pass
        self.tab, self.headers, self.views = None, {}, {}

    def _alive(self) -> bool:
        return self.tab is not None and not self.tab.is_closed()

    def _sui_headers(self) -> dict:
        extra = {k: v for k, v in self.headers.items() if k.lower() in ("page_id", "x-appbuilder-client-id")}
        return {**_SUI_HEADERS, **extra}

    def _warm_tab(self):
        """The warm tab for a getData or view request; never another tab (INVARIANT)."""
        if not self._alive():
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
            b._ensure(s)
            tab = self.tab if self._alive() else s.page
            return tab, sui_url(b.base, url, int(time.time() * 1000)), {"headers": self._sui_headers()}
        for attempt in (1, 2):
            try:
                _, text = await self._ask(b, prep)
                return parse_sui(text)                 # NoDrillData / CellTooWide / SignedOut / DrillError
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
            _, v = await self._ask(b, lambda s: (self._warm_tab(), f"{b.base}/aif/ui/model/view?page=DRILL_THROUGH"
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
                    self._warm_tab(), f"{b.base}/aif/ui/model/drilldown/getData?limit={config.DRILL_PAGE}"
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
