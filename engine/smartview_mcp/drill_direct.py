"""Drill lines without loading a drill page per piece (spike 2026-11-02, docs/audit/2026-11-02-drill-speed-spike.md).

Per piece, from ONE warm tab (a tab that loaded one real drill page): GET the drill URL's JSON form
(&FORMAT=SUI) - EPM's own drill query - and page getData with it, field for field ({QUERY, PARAMS: PARAMSARRAY,
PARAMSTYPE: PARAMSTYPEARRAY}). Read-only: nothing is built by us. getData from a tab that never loaded a drill page
gets the whole EPM session signed out, so it is only ever sent from the warm tab."""
from __future__ import annotations

import concurrent.futures as cf
import json
import logging
import threading
import time
from dataclasses import dataclass
from typing import Callable

from . import config
from .drill_browser import CellTooWide, DrillError, page_through, shared_browser

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


_GET = """async ([u, h]) => { try { const r = await fetch(u, {headers: h, credentials: 'include'});
                               return {s: r.status, t: await r.text()}; } catch (e) { return {s: 0, t: String(e)}; } }"""
_POST = """async ([u, b, h]) => { try { const r = await fetch(u, {method: 'POST', body: b, headers: h,
                                   credentials: 'include'}); return {s: r.status, t: await r.text()}; }
                                   catch (e) { return {s: 0, t: String(e)}; } }"""
_SUI_HEADERS = {"X-Requested-With": "XMLHttpRequest", "Accept": "*/*", "Content-Type": "application/json"}
_DROP = ("cookie", "content-length", "host")


class DirectDrill:
    """Worker-thread state: the warm tab, its getData headers and the view cache. One per drill browser session;
    use fetch() from any thread, fetch_on() on the worker thread (tests call it directly).

    INVARIANT: getData is POSTed only from self.tab, a tab that _warm() saw load a real drill page. getData from
    any other tab gets the whole EPM session signed out (spike section 2). The SUI GET is safe from any signed-in tab."""

    def __init__(self, capture_timeout_s: float | None = None):
        self.tab = None
        self.headers: dict[str, str] = {}
        self.views: dict[tuple, dict] = {}
        self.capture_timeout_s = config.DRILL_CAPTURE_TIMEOUT_S if capture_timeout_s is None else capture_timeout_s

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
        log.info("direct drill: warm tab ready")

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

    def fetch_on(self, b, s, url: str, max_lines: int,
                 progress: Callable[[int, int | None], None] | None = None) -> dict:
        b._ensure(s)
        tab = self.tab if self._alive() else s.page             # SUI only: safe from any signed-in tab
        r = tab.evaluate(_GET, [sui_url(b.base, url, int(time.time() * 1000)), self._sui_headers()])
        q = parse_sui(r["t"])                          # NoDrillData / CellTooWide / SignedOut / DrillError
        if not self._alive():
            self._warm(b, s, url)                      # from here on self.tab is a warmed tab: getData is allowed
        key = (q.app_id, q.partition)
        if key not in self.views:
            v = self.tab.evaluate(_GET, [f"{b.base}/aif/ui/model/view?page=DRILL_THROUGH&appId={q.app_id}"
                                         f"&partitionkey={q.partition}&_={int(time.time() * 1000)}",
                                         self._sui_headers()])
            try:
                self.views[key] = json.loads(v["t"])
            except ValueError:
                raise SignedOut("EPM answered the drill view with a page, not JSON (signed out?)") from None

        def post(off: int) -> tuple[int, str]:
            res = self.tab.evaluate(_POST, [f"{b.base}/aif/ui/model/drilldown/getData?limit={config.DRILL_PAGE}"
                                            f"&offset={off}", q.body, self.headers])
            return int(res["s"]), res["t"]

        def recover() -> None:
            self._drop_tab()
            b._login(s)
            self._warm(b, s, url)
        t1 = time.monotonic()
        items, total = page_through(post, max_lines, progress, recover)
        return {"items": items, "total": total, "view": self.views.get(key), "secs": round(time.monotonic() - t1, 1)}

    # ---------------------------------------------------------------- any thread
    def fetch(self, url: str, max_lines: int,
              progress: Callable[[int, int | None], None] | None = None) -> cf.Future:
        b = shared_browser()
        return b.submit(lambda s: self.fetch_on(b, s, url, max_lines, progress))


_shared: DirectDrill | None = None
_shared_lock = threading.Lock()


def shared_direct() -> DirectDrill:
    global _shared
    with _shared_lock:
        if _shared is None:
            _shared = DirectDrill()
        return _shared
