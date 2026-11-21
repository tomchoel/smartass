"""Drill lines over the user's OAuth token (epmhttp.EpmSession): no browser, no page per piece. The same protocol
the drill page runs (spike 2026-11-19), request for request:

1. the session key: GET /aif/ui/model/applications/servicetype with a page_id header; EPM's answer carries the key in
   its x-appbuilder-client-id header, and every later request carries the key and the page_id;
2. EPM's own drill query: the drill link's JSON form (drillcore.sui_url), read with drillcore.parse_sui;
3. the drill view, once per (application, partition) for the life of the key;
4. the lines: getData with that query, paged (drillcore.Pager).

A lost session (a redirect to the pod's exit page, 401/403, or a 200 that is not the JSON asked for) takes a new key
and the same request is sent once more (for getData through Pager's recover, at the same offset). Pieces that run at
once share the key: it is taken once, and a renewal several of them need happens once (the key a failed request
carried is compared with the current one). At most config.DRILL_PARALLEL getData calls are in flight.

A renewal drains first (live, 2026-11-21: a request with the old key that reaches EPM after the new key was taken
signs the new session out): from the moment a piece finds the session lost no request leaves with the old key, and
the key request waits until every request still out with it has come back - at most _DRAIN_S, the time one request
may take; after that it goes anyway and says so in the log.

The bound on retries: a piece retries once per key generation - never twice on the same key - and at most
MAX_RENEWALS times in all (its drill query, its view and its getData pages together). A session that is lost again
on a newer key is retried once more on the next key; a session lost after every new key ends the piece after
MAX_RENEWALS + 1 lost answers with the browser route's text (the drill query's / the view's "signed out?", Pager's
"getData returned HTTP n without data at offset m"). A piece therefore asks for at most MAX_RENEWALS new keys.

Read-only: the session refuses anything off its allowlist; getData runs EPM's SELECT. The key is a secret like the
token: it is never logged and never in an error. Nothing is cached beyond the key and the view: every drill is live."""
from __future__ import annotations

import asyncio
import json
import logging
import threading
import time
import weakref
from typing import Callable

from ..core import config
from . import epmhttp
from .drillcore import DrillError, Pager, SignedOut, SuiQuery, expired, parse_sui, redact, sui_url

log = logging.getLogger("smartview.drill")

PAGE_ID = "smartass"
SERVICE_TYPE = "/aif/ui/model/applications/servicetype"
VIEW = "/aif/ui/model/view"
GET_DATA = "/aif/ui/model/drilldown/getData"
_HEADERS = {"Accept": "application/json, text/javascript, */*; q=0.01", "Content-Type": "application/json",
            "X-Requested-With": "XMLHttpRequest", "page_id": PAGE_ID}
_REQUEST_S = 120.0                 # one key / SUI / view / getData request
_DRAIN_S = _REQUEST_S              # how long a new key waits for requests still out with the old one
MAX_RENEWALS = 3                   # new keys one piece may need (see the module docstring)


def _now_ms() -> int:
    return int(time.time() * 1000)


def _object(text: str) -> dict | None:
    """text as a JSON object, or None (a sign-in or exit page, an empty answer)."""
    try:
        j = json.loads(text)
    except ValueError:
        return None
    return j if isinstance(j, dict) else None


def _why(status: int) -> str:
    return "the request was lost" if status == 0 else "a page, not JSON" if status == 200 else f"HTTP {status}"


class _Loop:
    """One event loop's share of a TokenDrill: the lock (the key and the view cache), the getData gate and the
    requests out with a key (a new key waits until none is)."""

    def __init__(self):
        self.lock = asyncio.Lock()
        self.gate = asyncio.Semaphore(config.DRILL_PARALLEL)
        self.out = 0
        self.idle = asyncio.Event()
        self.idle.set()


class _Budget:
    """One piece's retries after a lost session: once per key generation, at most MAX_RENEWALS in all."""

    def __init__(self):
        self.spent_on: str | None = None              # the key whose loss the last retry answered
        self.renewals = 0

    def may_retry(self, key: str) -> bool:
        return key != self.spent_on and self.renewals < MAX_RENEWALS

    def spend(self, key: str) -> None:
        self.spent_on, self.renewals = key, self.renewals + 1


class TokenDrill:
    """The session key and the view cache of one pod (shared by every event loop), and the drill on top: the same
    calls and result as drill_direct.DirectDrill. http: an object with EpmSession's request()."""

    def __init__(self, base: str | None = None, http=None):
        self.base = (base or config.EPM_BASE or "").rstrip("/")
        self.http = http or epmhttp.session(self.base)
        self._key: str | None = None
        self._views: dict[tuple, dict] = {}
        self._taken = 0                                # key requests ended so far (with a key or not)
        self._dead: str | None = None                  # the last key found gone
        self._take_error: BaseException | None = None  # why the last one failed (None: it gave a key)
        self._per_loop: weakref.WeakKeyDictionary = weakref.WeakKeyDictionary()
        self._per_loop_lock = threading.Lock()

    # ---------------------------------------------------------------- the key and the view (under the loop's lock)
    def _sync(self) -> _Loop:
        """This event loop's lock, getData gate and count of requests out with a key."""
        loop = asyncio.get_running_loop()
        with self._per_loop_lock:
            got = self._per_loop.get(loop)
            if got is None:
                got = self._per_loop[loop] = _Loop()
            return got

    async def _call(self, method: str, url: str, key: str | None, body: str | None = None) -> epmhttp.Response:
        """One request on a worker thread; key None: the key request itself. A request lost on the way (OSError,
        epmhttp.RequestLost) is answered as status 0, like a page fetch that failed."""
        headers = {"page_id": PAGE_ID} if key is None else {**_HEADERS, "x-appbuilder-client-id": key}
        st = self._sync()
        if key is not None:
            st.out += 1
            st.idle.clear()
        try:
            return await asyncio.to_thread(self.http.request, method, url, headers=headers, body=body,
                                           timeout=_REQUEST_S)
        except OSError as e:
            log.info("token drill: %s %s was lost (%s)", method, redact(url), type(e).__name__)
            return epmhttp.Response(0, "", {})
        finally:
            if key is not None:
                st.out -= 1
                if st.out == 0:
                    st.idle.set()

    async def _drain(self) -> None:
        """Wait until no request is out with a key (the old one: no new request takes a key while there is none),
        at most _DRAIN_S. Only with the loop's lock held and the key dropped."""
        st = self._sync()
        if st.out == 0:
            return
        log.info("token drill: a new key waits for %d request(s) still out with the old one", st.out)
        try:
            await asyncio.wait_for(st.idle.wait(), _DRAIN_S)
        except TimeoutError:
            log.warning("token drill: %d request(s) with the old key still out after %g s; taking a new key anyway",
                        st.out, _DRAIN_S)

    async def _take_key(self) -> str:
        """A new key; the views of the old one go with it. Only with the loop's lock held."""
        self._key, self._views, self._take_error = None, {}, None
        try:
            await self._drain()
            r = await self._call("GET", SERVICE_TYPE, None)
            key = r.headers.get("x-appbuilder-client-id")
            if not key:
                log.info("token drill: no session key in EPM's answer (HTTP %s)", r.status)
                raise DrillError("EPM gave no session key for the drill")
        except BaseException as e:
            self._take_error = e
            raise
        finally:
            self._taken += 1
        self._key = key
        log.info("token drill: session key taken")
        return key

    async def _key_for(self, seen: str | None = None) -> str:
        """The current key, taking one if there is none. seen: the key the asking piece had (None: none yet). A key
        request that failed is the failure of every piece that waited for it or whose key it was to replace: one
        request, one error for all of them, as a renewal that succeeds is one request for all."""
        key = self._key
        if key is not None:
            return key
        waited_from = self._taken
        async with self._sync().lock:
            if self._key is not None:
                return self._key
            shared = self._taken != waited_from or (seen is not None and seen == self._dead)
            if self._take_error is not None and shared:
                raise self._take_error
            return await self._take_key()

    async def _renew(self, used: str) -> None:
        """A request with key used found the session gone: the key is dropped at once, so no request leaves with it
        from here on, and a new one is taken - unless another piece already did."""
        if self._key == used:
            self._key, self._dead = None, used
        await self._key_for(used)

    async def _view(self, q: SuiQuery, budget: _Budget) -> dict:
        """The drill view of q's application and partition, asked once per key; a lost session: a new key, within
        the piece's budget."""
        vk = (q.app_id, q.partition)
        async with self._sync().lock:
            while True:
                key = self._key if self._key is not None else await self._take_key()
                view = self._views.get(vk)
                if view is not None:
                    return view
                r = await self._call("GET", f"{VIEW}?page=DRILL_THROUGH&appId={q.app_id}&partitionkey={q.partition}"
                                            f"&_={_now_ms()}", key)
                view = _object(r.text)
                if r.status == 200 and view is not None:
                    self._views[vk] = view
                    return view
                if not expired(r.status, view):
                    raise DrillError(f"EPM answered the drill view with HTTP {r.status}")
                if not budget.may_retry(key):
                    log.info("token drill: the drill view found the session gone again (%s)", _why(r.status))
                    raise SignedOut("EPM answered the drill view with a page, not JSON (signed out?)")  # drill_direct's
                budget.spend(key)
                log.info("token drill: the drill view found the session gone (%s); a new key", _why(r.status))
                self._key, self._dead = None, key

    # ---------------------------------------------------------------- the drill
    async def query(self, url: str) -> SuiQuery:
        """EPM's own drill query for a drill link. EPM's refusals (CellTooWide, NoDrillData, DrillError) are raised
        at once; a lost session takes a new key and is asked once more (the retry bound: module docstring)."""
        return await self._query(url, _Budget())

    async def _query(self, url: str, budget: _Budget) -> SuiQuery:
        while True:
            key = await self._key_for()
            r = await self._call("GET", sui_url(self.base, url, _now_ms()), key)
            j = _object(r.text)
            if not expired(r.status, j):
                if j is None:
                    raise DrillError(f"EPM answered the drill query with HTTP {r.status}")
                return parse_sui(r.text)
            if not budget.may_retry(key):
                log.info("token drill: the drill query found the session gone again (%s)", _why(r.status))
                raise SignedOut("EPM answered the drill query with a page, not JSON (signed out?)")  # parse_sui's
            budget.spend(key)
            log.info("token drill: the drill query found the session gone (%s); a new key", _why(r.status))
            await self._renew(key)

    async def lines(self, q: SuiQuery, warm_url: str, max_lines: int,
                    progress: Callable[[int, int | None], None] | None = None) -> dict:
        """The lines of a drill query, paged; a lost session takes a new key (Pager's recover) and the same offset
        is asked again (the retry bound: module docstring). warm_url is DirectDrill's: there is no page to warm."""
        return await self._lines(q, max_lines, progress, _Budget())

    async def _lines(self, q: SuiQuery, max_lines: int, progress: Callable[[int, int | None], None] | None,
                     budget: _Budget) -> dict:
        view = await self._view(q, budget)
        gate = self._sync().gate
        t1 = time.monotonic()
        pg = Pager(max_lines, progress, True)
        key = self._key
        while True:
            async with gate:                         # the key is read in the gate: after a renewal began, a piece
                key = await self._key_for(key)       # that waited for its turn sends the new one, never the old
                r = await self._call("POST", f"{GET_DATA}?limit={config.DRILL_PAGE}&offset={pg.off}", key, q.body)
            pg.recovered = not budget.may_retry(key)  # Pager's one recover, given back for each newer key
            step = pg.take(r.status, r.text)
            if step == "recover":
                budget.spend(key)
                await self._renew(key)
            elif step == "done":
                break
        items, total = pg.result()
        return {"items": items, "total": total, "view": view, "secs": round(time.monotonic() - t1, 1), "query": q}

    async def fetch(self, url: str, max_lines: int,
                    progress: Callable[[int, int | None], None] | None = None) -> dict:
        """One drill piece: EPM's drill query, then its lines. The result carries the query ('query') so a split
        can build its siblings' queries from it (drill_derive)."""
        budget = _Budget()                          # one retry bound for the piece's query, view and pages
        return await self._lines(await self._query(url, budget), max_lines, progress, budget)


_shared: TokenDrill | None = None
_shared_lock = threading.Lock()


def shared() -> TokenDrill:
    """The process's one drill (config.EPM_BASE's pod)."""
    global _shared
    with _shared_lock:
        if _shared is None:
            _shared = TokenDrill()
        return _shared
