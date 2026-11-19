"""Grids over the Smart View provider's HTTP protocol - no Excel, no COM (Stage 0, 2026-10-25:
docs/specs/2026-10-25-stage0-epm-access-probe.md). retrieve() returns what ops._query_once returns,
(Layout, Parsed, seconds), so extract.route can use either route; the COM route stays the fallback.

Error contract for extract.route: InputError only for a query that is wrong whatever the route (a dimension
used twice, no members, no row/column axis); everything else - including an incomplete or stale atlas, which COM
does not need - is SvHttpError, so the caller falls back to Smart View in Excel.

Two transports. BearerTransport posts over plain HTTPS with the user's OAuth access token (epm/oauth.py; no browser)
for every pod that has a scope and a sign-in, and hands everything else to BrowserTransport, which posts from inside
the drill's hidden Edge (Telenor SSO), one tab per EPM pod, on the drill browser's worker thread. Tests use a fake
transport."""
from __future__ import annotations

import concurrent.futures as cf
import datetime as dt
import http.cookiejar
import logging
import re
import threading
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Protocol
from urllib.parse import urlsplit

from ..knowledge import atlas
from . import drill_browser, oauth, svxml
from ..core import config
from .drill_browser import DrillError
from .grid import INVALID, MISSING, NOACCESS, Axis, Layout, full_pov, parse, validate_axes, verify_labels
from ..xl.host import InputError

log = logging.getLogger(__name__)

# connection -> (Smart View data source = Server name, app, cube). By NAME: several servers share app FinMax
# (prod plan1, plan7 and the plan1-test pod); names are set in defaults.toml [svhttp].
CUBES = {"FinMax": (config.SVHTTP_FINMAX_SERVER, "FinMax", "FinMax"),
         "FinLite": (config.SVHTTP_FINMAX_SERVER, "FinMax", "FinLite"),
         "FinPlan": (config.SVHTTP_FINPLAN_SERVER, "FinPlan", "FinPlan")}


# FinPlan (live 10.11): a member name that clashes with an alias fails the whole grid, the same on every session:
# 'Found duplicate aliases or member name conflicting with alias under parent ...' and
# 'Duplicate name <m> found in grid; cannot resolve. Please replace it with qualified name.'
ALIAS_CLASH = ("duplicate aliases", "qualified name")
_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


class SvHttpError(Exception):
    """The HTTP route could not answer; extract.route then uses Smart View in Excel."""


class RequestLost(SvHttpError):
    """The request left but its answer did not come (tab navigated or closed, or the fetch timed out): EPM may still
    be working on it, so its session is dropped and the request goes again on another one."""


class Transport(Protocol):
    def post(self, base: str, body: str) -> str: ...


def _norm_dim(name: str) -> str:
    return name.removeprefix("All ").replace(" ", "").casefold()


def _first_line(e: BaseException) -> str:
    """One redacted line of an exception (Playwright messages carry call logs and can quote sso_token URLs)."""
    return drill_browser.redact(next((ln for ln in str(e).splitlines() if ln.strip()), type(e).__name__))[:200]


@dataclass
class Fetched:
    """A provider grid as fetched: Smart View's own shape from A1 (POV row, column headers, row labels, data)."""
    layout: Layout
    values: list[list]
    seconds: float


class SvHttp:
    def __init__(self, transport: Transport, home: str | None = None, db_path: Path | None = None):
        self.t = transport
        self.home = (home or config.EPM_BASE).rstrip("/")
        self.db_path = db_path or config.ATLAS_DB
        self._bases: dict[str, list[tuple[str, str]]] | None = None
        self._idle: dict[str, list[str]] = {}       # connection -> sIDs not in use (the session pool)
        self._slots: dict[str, threading.BoundedSemaphore] = {}
        self._lock = threading.Lock()                 # the idle pools; never held during a request
        self._bases_lock = threading.Lock()           # one data-source lookup at a time (a cold one signs in)

    def base_for(self, server: str, app: str) -> str:
        """The pod of the data source named `server`, which must serve `app` and must not be a test pod (unless
        config allows it). The list is cached once it is non-empty; a name missing from it is looked up once more."""
        entries = self._bases.get(server) if self._bases else None
        if not entries:
            bases = svxml.data_sources(self.t.post(self.home, svxml.provisioned_data_sources()))
            self._bases = bases or None
            entries = bases.get(server)
            if not entries:
                raise SvHttpError(f"EPM lists no Smart View data source named {server!r} "
                                  f"(has: {', '.join(sorted(bases)) or 'none'}); set svhttp server names in config")
        if len(entries) > 1:                 # only the name we need matters; other duplicates are not our problem
            raise SvHttpError(f"Smart View data source {server!r} is listed with different pods/apps "
                              f"({', '.join(f'{a} at {b}' for a, b in entries)}): ambiguous")
        got_app, base = entries[0]
        if got_app != app:
            raise SvHttpError(f"Smart View data source {server!r} serves app {got_app}, not {app}")
        if "-test" in (urlsplit(base).hostname or "") and not config.SVHTTP_ALLOW_TEST:
            raise SvHttpError(f"refused: data source {server!r} is on a test environment ({base}); "
                              "set svhttp.allow_test = true to use it")
        return base

    def _take(self, conn: str) -> tuple[str, str, str]:
        """(pod base, cube, sID) for ONE request: an idle session of the pool, else a new one. A session is never
        used by two requests at once: parallel grids on one sID come back wrong (2-3 of 8 corrupted, live
        2026-11-03), on their own sIDs all 8 are right."""
        server, app, cube = CUBES[conn]
        if not self._bases_lock.acquire(timeout=config.SVHTTP_TIMEOUT_S * 2):   # behind a cold start's sign-in
            raise SvHttpError("the HTTP route is busy looking up the EPM data sources")
        try:
            base = self.base_for(server, app)
        finally:
            self._bases_lock.release()
        if not self._lock.acquire(timeout=config.SVHTTP_TIMEOUT_S):
            raise SvHttpError("the HTTP route is busy with another query")
        try:
            idle = self._idle.get(conn)
            if idle:
                return base, cube, idle.pop()
        finally:
            self._lock.release()
        sid = svxml.sid_of(self.t.post(base, svxml.connect()))
        sid = svxml.sid_of(self.t.post(base, svxml.open_application(sid, app)))
        ans = self.t.post(base, svxml.open_cube(sid, app, cube))
        svxml.raise_for_exception(ans)
        return base, cube, svxml.sid_in(ans) or sid

    def _give(self, conn: str, sid: str) -> None:
        with self._lock:
            self._idle.setdefault(conn, []).append(sid)

    def _slot(self, conn: str) -> threading.BoundedSemaphore:
        return self._slots.setdefault(conn, threading.BoundedSemaphore(config.SVHTTP_SESSIONS))   # atomic

    def dimensions(self, conn: str) -> list[str]:
        """The connection's dimensions from the atlas; SvHttpError if it has none or any of them was not built
        (failed or empty; age does not matter) - the grid would then be built on a wrong dimension list, COM knows
        better."""
        db = atlas.connect(self.db_path)
        try:
            dims = atlas.known_dimensions(db, conn)              # incl. failed builds (their members are deleted)
            stale = [d for d in dims if not atlas.built(db, conn, d)]
        finally:
            db.close()
        if not dims:
            raise SvHttpError(f"the atlas has no dimensions for {conn}: build it with `smartass atlas build`")
        if stale:
            raise SvHttpError(f"the atlas is incomplete or stale for {conn} (dimensions {', '.join(stale)})")
        return dims

    def retrieve(self, conn: str, pov: dict | None, rows: list[Axis], cols: list[Axis],
                 suppress_missing: bool, suppress_zeros: bool):
        f = self.fetch(conn, pov, rows, cols)
        return f.layout, parse(f.layout, f.values, suppress_missing, suppress_zeros), f.seconds

    def fetch(self, conn: str, pov: dict | None, rows: list[Axis], cols: list[Axis]) -> "Fetched":
        """The checked provider grid, unparsed: what a source sheet is written from (core/source.py)."""
        if conn not in CUBES:
            raise SvHttpError(f"no HTTP route for connection {conn}")
        t0 = time.monotonic()
        dims = self.dimensions(conn)
        by_norm = {_norm_dim(d): d for d in dims}

        def canon(d: str) -> str:
            return by_norm.get(_norm_dim(d), d)
        pov = {canon(d): m for d, m in (pov or {}).items()}
        rows = [Axis(canon(a.dimension), a.members) for a in rows]
        cols = [Axis(canon(a.dimension), a.members) for a in cols]
        problems = validate_axes(dims, pov, rows, cols)
        unknown = [p for p in problems if p.startswith("unknown dimension")]
        if unknown:                       # the atlas may lack a dimension the cube has: let COM decide
            raise SvHttpError(f"the atlas is incomplete or stale for {conn} ({'; '.join(unknown)})")
        if problems:
            raise InputError("Invalid query: " + "; ".join(problems))
        lay = Layout(full_pov(dims, pov, rows, cols), rows, cols)
        values = self._execute(conn, lay)
        _check_grid(lay, values)
        full = parse(lay, values, False, False)
        if full.pov_problems:              # COM shows the same as a warning; here it means the layout is not the cube's
            raise SvHttpError("the provider appended POV members; the atlas may lack a dimension ("
                              + "; ".join(full.pov_problems) + ")")
        bad = verify_labels(lay, full)
        if bad:
            raise SvHttpError("Smart View provider returned a different layout than requested (rows/columns "
                              "moved): " + "; ".join(bad[:5]))
        return Fetched(lay, values, time.monotonic() - t0)

    def closed_month(self, conn: str = "FinMax") -> str:
        """The last closed month ('Current Reporting Month') as 'Sep 26': the one child of member CurrRepMonth in
        provider dimension TimePeriods, '[CurrRepMonth].[Y26M09]' (proven live 2026-10-25, FinMax only: FinLite is
        answered from the FinMax session). A month in the future or more than 3 months back is refused, so the
        caller's COM route gives a second opinion instead of a silently wrong month."""
        if conn not in ("FinMax", "FinLite"):
            raise SvHttpError(f"{conn} has no Current Reporting Month")
        names = self._enum("FinMax", "TimePeriods", "CurrRepMonth")
        if len(names) != 1:
            raise SvHttpError(f"CurrRepMonth should have exactly one child, got {len(names)}: {names[:5]}")
        m = re.search(r"Y(\d\d)M(\d\d)", names[0])
        if not m or not 1 <= int(m.group(2)) <= 12:
            raise SvHttpError(f"Current Reporting Month not understood: {names[:3]}")
        today = dt.date.today()
        back = (today.year * 12 + today.month) - ((2000 + int(m.group(1))) * 12 + int(m.group(2)))
        if not 0 <= back <= 3:
            raise SvHttpError(f"Current Reporting Month {names[0]} is not plausible today "
                              f"({'in the future' if back < 0 else f'{back} months back'})")
        return f"{_MONTHS[int(m.group(2)) - 1]} {m.group(1)}"

    def drill_link(self, conn: str, cell: dict[str, str]) -> dict:
        """ops.drill_link over HTTP: {cube_value, url (relative to the EPM base), report}. One 1x1 retrieve gives
        the value and the Default-alias labels Smart View sends; then GetDrillThroughReports (read-only)."""
        from .ops import DRILL_DEFAULTS, drill_cell_defaults
        cell, _ = drill_cell_defaults(cell)
        if cell.get("All Balance Types", "Periodic") != "Periodic":
            raise InputError("Drill-through works on Periodic data only; use a period member (quarter/year) "
                             "instead of Accumulated.")
        pov = {**DRILL_DEFAULTS, **{d: m for d, m in cell.items() if d not in ("All Accounts", "All Time Periods")}}
        f = self.fetch(conn, pov, [Axis("All Accounts", [cell["All Accounts"]])],
                       [Axis("All Time Periods", [cell["All Time Periods"]])])
        lay, v = f.layout, f.values
        cells = [("All Time Periods", "col", str(v[1][lay.nr])), ("All Accounts", "row", str(v[1 + lay.nc][0]))]
        cells += [(d, "pov", str(v[0][lay.nr + j])) for j, (d, _) in enumerate(lay.pov)]
        ords = (1 + lay.nc) * len(v[0]) + lay.nr
        val = v[1 + lay.nc][lay.nr]
        got = self._ask(conn, lambda sid, cube: svxml.drill_reports(sid, cube, ords, cells), svxml.drill_url,
                        lambda e: False)
        return {"cube_value": val if isinstance(val, float) else None,
                "url": got[1] if got else None, "report": got[0] if got else None}

    def child_values(self, conn: str, parts: list[dict]) -> list | None:
        """drill_split.child_values_on over HTTP: the parts differ in one dimension - one column retrieve,
        float or None per part (None for the whole list when they differ in nothing)."""
        from .ops import DRILL_DEFAULTS
        first = parts[0]
        dim = next((k for k in first if any(p.get(k) != first[k] for p in parts)), None)
        if dim is None:
            return None
        col = "All Time Periods" if dim != "All Time Periods" else "All Accounts"
        pov = {**DRILL_DEFAULTS, **{k: v for k, v in first.items() if k not in (dim, col)}}
        f = self.fetch(conn, pov, [Axis(dim, [p[dim] for p in parts])], [Axis(col, [first[col]])])
        lay = f.layout
        return [v if isinstance(v, float) else None
                for v in (f.values[1 + lay.nc + i][lay.nr] for i in range(len(parts)))]

    def _ask(self, conn: str, body, read, fatal) -> object:
        """POST body(sid, cube) on a session of its own and read(answer); at most SVHTTP_SESSIONS at once per
        connection. An error fatal(e) does not rule out (maybe an expired session) drops that session and tries
        once more on a new one. A session whose request failed in the browser is dropped too (EPM may still be
        working on it)."""
        slot = self._slot(conn)
        if not slot.acquire(timeout=config.SVHTTP_TIMEOUT_S):
            raise SvHttpError("the HTTP route is busy with other queries")
        try:
            for attempt in (1, 2):
                sid = None
                try:
                    base, cube, sid = self._take(conn)
                    ans = self.t.post(base, body(sid, cube))
                    sid = svxml.sid_in(ans) or sid            # each answer may carry a new sID: the latest one counts
                    out = read(ans)
                except RequestLost as e:              # the sID is dropped (not given back); a fresh one tries
                    if attempt == 2:
                        raise SvHttpError(str(e)) from None
                    log.info("svhttp %s: %s - once more on a new session", conn, e)
                    continue
                except svxml.SvXmlError as e:
                    if fatal(e) and sid is not None:
                        self._give(conn, sid)                 # a wrong query: the session itself is fine
                    elif not fatal(e):
                        with self._lock:
                            self._bases = None                # ... and re-read where the apps live
                    if fatal(e) or attempt == 2:
                        raise SvHttpError(f"Smart View provider: {e}") from None
                    log.info("svhttp %s: %s - reconnecting once", conn, e)
                    continue
                self._give(conn, sid)
                return out
        finally:
            slot.release()
        raise AssertionError("unreachable")

    def _enum(self, conn: str, dim: str, member: str) -> list[str]:
        return self._ask(conn, lambda sid, cube: svxml.enum_children(sid, dim, member), svxml.members,
                         lambda e: e.code in (4996, 4997))  # unknown member or dimension: a new session won't help

    def _execute(self, conn: str, lay: Layout) -> list[list]:
        return self._ask(conn, lambda sid, cube: svxml.execute_grid(sid, cube, lay), svxml.grid_values,
                         lambda e: "Grid is Invalid" in str(e)        # unknown member: deterministic
                         or any(m in str(e) for m in ALIAS_CLASH))   # alias clash: deterministic too


def _check_grid(lay: Layout, values: list[list]) -> None:
    """The returned slice must cover the requested layout (wider is fine: appended POV cells) and every data cell
    must be a number, blank or a Smart View label (#Missing, #NoAccess, #Invalid); anything else is not data."""
    m = lay.matrix()
    h, w = len(m), len(m[0])
    if len(values) < h:
        raise SvHttpError(f"the provider returned fewer rows than requested ({len(values)} < {h})")
    got_w = min((len(r) for r in values), default=0)
    if got_w < w:
        raise SvHttpError(f"the provider returned fewer columns than requested ({got_w} < {w})")
    labels = MISSING | NOACCESS | INVALID
    for r in range(1 + lay.nc, h):
        for c in range(lay.nr, lay.nr + len(lay.col_tuples)):
            v = values[r][c]
            if isinstance(v, str) and v.strip() not in labels:
                raise SvHttpError(f"non-numeric data in the provider grid (row {r + 1}, column {c + 1}: {v[:40]!r})")


def _origin(url: str) -> str:
    u = urlsplit(str(url or ""))
    return f"{u.scheme}://{u.netloc}".lower()


def _fetch_ms() -> int:
    return max(1, int(config.SVHTTP_TIMEOUT_S * 900))       # 90 % of the job's time limit, in ms


def _close(page) -> None:
    try:
        page.close()
    except Exception:   # noqa: BLE001 - the tab may already be gone
        pass


class BrowserTransport:
    """Posts provider XML from a signed-in tab of the drill browser: its own tab for its pod (plan1) and one extra
    tab per other pod (FinPlan: stratmodel), signed in once with the same SSO. Tab work runs on the browser's worker
    thread (_tabs is only touched there); the request itself is in flight in the tab (DrillBrowser._fire), so many
    posts run at once. Every failure to get an answer is an SvHttpError (so extract.route can fall back); a refused
    write stays a ValueError.

    `browser` is a callable returning the drill browser, resolved on every post: after
    drill_browser.shutdown_shared() the next post uses the new shared browser, not the shut one."""

    def __init__(self, browser: Callable[[], object]):
        self._browser = browser
        self._tabs: dict[str, tuple[object, object]] = {}       # pod base -> (browser context, page)

    def post(self, base: str, body: str) -> str:
        """Two tries: an answer that is not provider XML (signed out?) or a request lost with its tab signs in
        again - unless another request already did (a stale failure), then it is simply sent again."""
        svxml.check_read_only(body)
        base = base.rstrip("/")
        why, stale = "no answer", None
        for _ in (1, 2):
            status, text, stale = self._send(base, body, stale)
            if isinstance(text, str) and ("<res_" in text or "<exception" in text):
                return text                                       # provider XML, whatever the status: svxml judges
            if not status:                  # never again on this sID: EPM may still be working on the first one
                self._settle(base, stale)
                raise RequestLost(f"Smart View provider at {base}: lost the request "
                                  f"({_first_line(Exception(text))})")
            why = f"answered HTTP {status}"          # a sign-in page: EPM never ran it, so the same sID may go again
            log.warning("svhttp: %s from %s; trying again", why, base)
        raise SvHttpError(f"Smart View provider at {base}: {why} twice (sign-in failed?)")

    def _settle(self, base: str, stale) -> None:
        """After a lost request: let the next one sign in again if its tab is still the one that failed."""
        try:
            b = self._browser()
            b.submit(lambda s: self._drop_if_stale(b, s, base, stale)).result(timeout=30)
        except Exception as e:   # noqa: BLE001 - the next request's own checks still apply
            log.info("svhttp: could not reset the tab after a lost request: %s", _first_line(e))

    def _send(self, base: str, body: str, stale) -> tuple[int, str, tuple]:
        deadline = time.monotonic() + config.SVHTTP_TIMEOUT_S
        fut = None
        try:
            b = self._browser()
            fut = b.submit(lambda s: self._start(b, s, base, body, stale))
            inner, page, gen = fut.result(timeout=config.SVHTTP_TIMEOUT_S)
            try:
                status, text = inner.result(timeout=max(0.1, deadline - time.monotonic()))
            except DrillError as e:                               # the tab navigated or closed under the request
                return 0, str(e), (page, gen)
            return status, text, (page, gen)
        except cf.CancelledError:
            raise SvHttpError("the drill browser job was cancelled (browser restarting or shutting down)") from None
        except cf.TimeoutError:
            restarted = ""
            if fut is not None and not fut.done() and not fut.cancel():   # queued: cancelled, never runs
                try:                         # running and hung: restart OUR browser (like fin_drill's hang guard), so
                    b.kill()                 # later queries and drills do not queue behind a stuck worker
                    restarted = "; the hidden browser was restarted"
                except Exception as e:   # noqa: BLE001 - the fallback must still be offered
                    log.warning("svhttp: killing the hung drill browser failed: %s", _first_line(e))
            raise SvHttpError(f"the drill browser did not answer within {config.SVHTTP_TIMEOUT_S:.0f} s"
                              f"{restarted}") from None
        except SvHttpError:                       # (a refused write raised above, before the job)
            raise
        except DrillError as e:                                   # includes DrillBusy
            raise SvHttpError(str(e)) from None
        except Exception as e:   # noqa: BLE001 - e.g. a raw Playwright error: never hand it to the caller
            raise SvHttpError(f"browser error: {_first_line(e)}") from None

    @staticmethod
    def _own(b, base: str) -> bool:
        return base == b.base.rstrip("/")

    def _tab(self, b, s, base: str):
        if self._own(b, base):
            return s.page
        ctx, page = self._tabs.get(base, (None, None))
        if ctx is not s.ctx or page is None or page.is_closed():
            page = s.ctx.new_page()
            try:
                b._login_once(page, s.deadline, base=base)
            except DrillError as first:
                log.warning("svhttp: sign-in on %s failed (%s); retrying in a fresh tab", base, first)
                _close(page)
                page = s.ctx.new_page()
                try:
                    b._login_once(page, s.deadline, base=base)
                except DrillError as second:
                    _close(page)
                    raise SvHttpError(f"sign-in on {base} failed twice: {second}") from None
            self._tabs[base] = (s.ctx, page)
        return page

    def _drop(self, b, s, base: str) -> None:
        """Forget the tab that failed: the own pod's tab is the drill session (sign in again); another pod's extra
        tab is closed and opened afresh next time, leaving the drill session alone."""
        if self._own(b, base):
            s.ready = False
            return
        _, page = self._tabs.pop(base, (None, None))
        if page is not None:
            _close(page)

    def _drop_if_stale(self, b, s, base: str, stale) -> None:
        """Drop the tab a failed request used - unless it was renewed since (another request signed in again):
        many requests fail together when a tab goes, and only the first may sign in again."""
        page, gen = stale
        if self._own(b, base):
            if s.page is page and getattr(s, "gen", 0) == gen:
                self._drop(b, s, base)
        elif self._tabs.get(base, (None, None))[1] is page:
            self._drop(b, s, base)

    def _start(self, b, s, base: str, body: str, stale) -> tuple:
        """Worker thread: a signed-in tab on the pod, then the request in flight: (future, page, sign-in gen)."""
        s.deadline = time.monotonic() + config.SVHTTP_TIMEOUT_S
        if stale is not None:
            self._drop_if_stale(b, s, base, stale)
        why = "no answer"
        for attempt in (1, 2):
            try:
                b._ensure(s)
                page = self._tab(b, s, base)
                here = _origin(page.url)
                if here != _origin(base):       # never post the session id from another origin (e.g. the IdP page)
                    why = f"the tab is on {drill_browser.redact(here)}, not {base}"
                    log.warning("svhttp: %s; signing in again", why)
                    self._drop(b, s, base)
                    continue
                fut = b._fire(s, page, "/HyperionPlanning/SmartView",
                              {"method": "POST", "body": body, "headers": {"Content-Type": "text/xml"}},
                              _fetch_ms() / 1000)
                return fut, page, getattr(s, "gen", 0)
            except (SvHttpError, DrillError):                    # sign-in already retried inside
                raise
            except Exception as e:   # noqa: BLE001 - the tab navigated or the browser died: one fresh try
                if attempt == 2:
                    raise SvHttpError(f"browser error at {base}: {_first_line(e)}") from None
                self._drop(b, s, base)
        raise SvHttpError(f"Smart View provider at {base}: {why} twice (sign-in failed?)")


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None


AUTH_REFUSED = (301, 302, 303, 307, 401, 403)     # EPM's gate did not take the token: the request never ran


class BearerTransport:
    """Posts provider XML over plain HTTPS with the user's OAuth access token: no browser, no sign-in page (live
    2026-11-19: /HyperionPlanning/SmartView takes the token; a FinPlan grid is cell for cell the browser route's;
    on the pod alone 0.8 s cold against 17 s - but the data-source lookup on the home pod still signs in through
    the browser until that pod has a scope too). Only for a pod with a configured scope; a pod without one, a pod
    the user has not signed in to, and a token EPM refuses all go to `fallback` (the browser transport) - and stay there for
    SVHTTP_BEARER_RETRY_S, so a missing sign-in costs one log line, not one per request.

    One cookie jar per pod, as a browser has: EPM opens a server session for every request that brings no cookie.
    The cookie alone does not authenticate (302 without the token). Like every transport here it sends nothing
    that svxml.check_read_only refuses - the same token would be accepted by EPM's write endpoints."""

    def __init__(self, fallback: Transport, tokens=oauth, send: Callable[[str, str, str], tuple[int, str]] | None = None):
        self.fallback = fallback
        self._tokens = tokens
        self._send = send or self._https
        self._off: dict[str, float] = {}               # pod -> monotonic time until which the browser route is used
        self._openers: dict[str, urllib.request.OpenerDirector] = {}
        self._lock = threading.Lock()

    def post(self, base: str, body: str) -> str:
        svxml.check_read_only(body)
        base = base.rstrip("/")
        if self._tokens.scope_for(base) is None or time.monotonic() < self._off.get(base, 0.0):
            return self.fallback.post(base, body)
        try:
            token = self._tokens.access_token(base)
        except oauth.OAuthError as e:
            return self._fall_back(base, body, str(e))
        status = 0
        for attempt in (1, 2):
            try:
                status, text = self._send(base, body, token)
            except OSError as e:                  # it may have reached EPM: never again on this sID (as a lost tab)
                raise RequestLost(f"Smart View provider at {base}: lost the request ({type(e).__name__})") from None
            if "<res_" in text or "<exception" in text:
                return text                                       # provider XML, whatever the status: svxml judges
            if status not in AUTH_REFUSED:
                raise RequestLost(f"Smart View provider at {base}: answered HTTP {status} without provider XML")
            if attempt == 1:                                      # the hour may have ended early: one new token
                try:
                    token = self._tokens.access_token(base, reject=token)
                except oauth.OAuthError as e:
                    return self._fall_back(base, body, str(e))
        return self._fall_back(base, body, f"EPM refused the access token (HTTP {status})")

    def _fall_back(self, base: str, body: str, why: str) -> str:
        self._off[base] = time.monotonic() + config.SVHTTP_BEARER_RETRY_S
        log.warning("svhttp: no bearer route to %s (%s); using the browser route", base, why[:200])
        return self.fallback.post(base, body)

    def _opener(self, base: str) -> urllib.request.OpenerDirector:
        with self._lock:
            op = self._openers.get(base)
            if op is None:
                op = self._openers[base] = urllib.request.build_opener(
                    _NoRedirect, urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
            return op

    def _https(self, base: str, body: str, token: str) -> tuple[int, str]:
        req = urllib.request.Request(base + "/HyperionPlanning/SmartView", data=body.encode("utf-8"), method="POST",
                                     headers={"Authorization": "Bearer " + token, "Content-Type": "text/xml"})
        try:
            with self._opener(base).open(req, timeout=_fetch_ms() / 1000) as r:
                return r.status, r.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            return e.code, e.read().decode("utf-8", "replace")


_shared: SvHttp | None = None
_shared_lock = threading.Lock()


def shared() -> SvHttp:
    """The process's one HTTP route: the bearer transport in front of the shared drill browser (looked up on every
    post, and only started when a request needs it)."""
    global _shared
    with _shared_lock:
        if _shared is None:
            _shared = SvHttp(BearerTransport(BrowserTransport(lambda: drill_browser.shared_browser())))
        return _shared
