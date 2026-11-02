"""Grids over the Smart View provider's HTTP protocol - no Excel, no COM (Stage 0, 2026-10-25:
docs/specs/2026-10-25-stage0-epm-access-probe.md). retrieve() returns what ops._query_once returns,
(Layout, Parsed, seconds), so extract.route can use either route; the COM route stays the fallback.

Error contract for extract.route: InputError only for a query that is wrong whatever the route (a dimension
used twice, no members, no row/column axis); everything else - including an incomplete or stale atlas, which COM
does not need - is SvHttpError, so the caller falls back to Smart View in Excel.

The transport is BrowserTransport: it posts from inside the drill's hidden Edge (Telenor SSO), one tab per EPM pod,
on the drill browser's worker thread. Tests use a fake transport."""
from __future__ import annotations

import concurrent.futures as cf
import datetime as dt
import logging
import re
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Protocol
from urllib.parse import urlsplit

from . import atlas, config, drill_browser, svxml
from .drill_browser import DrillError
from .grid import INVALID, MISSING, NOACCESS, Axis, Layout, full_pov, parse, validate_axes, verify_labels
from .host import InputError

log = logging.getLogger(__name__)

# connection -> (Smart View data source = Server name, app, cube). By NAME: several servers share app FinMax
# (prod plan1, plan7 and the plan1-test pod); names are set in defaults.toml [svhttp].
CUBES = {"FinMax": (config.SVHTTP_FINMAX_SERVER, "FinMax", "FinMax"),
         "FinLite": (config.SVHTTP_FINMAX_SERVER, "FinMax", "FinLite"),
         "FinPlan": (config.SVHTTP_FINPLAN_SERVER, "FinPlan", "FinPlan")}


_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


class SvHttpError(Exception):
    """The HTTP route could not answer; extract.route then uses Smart View in Excel."""


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
        self._sids: dict[str, str] = {}
        self._lock = threading.Lock()

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

    def _note_sid(self, conn: str, text: str) -> None:
        """Each answer may carry a new sID; the latest one is the session."""
        sid = svxml.sid_in(text)
        if sid:
            self._sids[conn] = sid

    def _session(self, conn: str) -> tuple[str, str, str]:
        server, app, cube = CUBES[conn]
        base = self.base_for(server, app)
        if conn not in self._sids:
            sid = svxml.sid_of(self.t.post(base, svxml.connect()))
            sid = svxml.sid_of(self.t.post(base, svxml.open_application(sid, app)))
            ans = self.t.post(base, svxml.open_cube(sid, app, cube))
            svxml.raise_for_exception(ans)
            self._sids[conn] = sid
            self._note_sid(conn, ans)
        return base, cube, self._sids[conn]

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
            raise SvHttpError(f"the atlas has no dimensions for {conn}: build it with sv_atlas_build")
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
        if not self._lock.acquire(timeout=config.SVHTTP_TIMEOUT_S):
            raise SvHttpError("the HTTP route is busy with another query")
        try:
            values = self._execute(conn, lay)
        finally:
            self._lock.release()
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
        if not self._lock.acquire(timeout=config.SVHTTP_TIMEOUT_S):
            raise SvHttpError("the HTTP route is busy with another query")
        try:
            names = self._enum("FinMax", "TimePeriods", "CurrRepMonth")
        finally:
            self._lock.release()
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

    def _ask(self, conn: str, body, read, fatal) -> list:
        """POST body(sid, cube) to conn's provider and read(answer). An error fatal(e) does not rule out (maybe an
        expired session) reconnects once."""
        for attempt in (1, 2):
            try:
                base, cube, sid = self._session(conn)
                ans = self.t.post(base, body(sid, cube))
                self._note_sid(conn, ans)
                return read(ans)
            except svxml.SvXmlError as e:
                if not fatal(e):
                    self._sids.pop(conn, None)
                    self._bases = None                      # ... and re-read where the apps live
                if fatal(e) or attempt == 2:
                    raise SvHttpError(f"Smart View provider: {e}") from None
                log.info("svhttp %s: %s - reconnecting once", conn, e)
        raise AssertionError("unreachable")

    def _enum(self, conn: str, dim: str, member: str) -> list[str]:
        return self._ask(conn, lambda sid, cube: svxml.enum_children(sid, dim, member), svxml.members,
                         lambda e: e.code in (4996, 4997))  # unknown member or dimension: a new session won't help

    def _execute(self, conn: str, lay: Layout) -> list[list]:
        return self._ask(conn, lambda sid, cube: svxml.execute_grid(sid, cube, lay), svxml.grid_values,
                         lambda e: "Grid is Invalid" in str(e))   # unknown member: deterministic


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


# The fetch aborts itself (AbortSignal) a little before the job's own time limit, so a hung request frees the worker.
_JS = """async ([b, ms]) => { const r = await fetch('/HyperionPlanning/SmartView', {method: 'POST',
    credentials: 'include', body: b, headers: {'Content-Type': 'text/xml'}, signal: AbortSignal.timeout(ms)});
    return [r.status, await r.text()]; }"""


def _fetch_ms() -> int:
    return max(1, int(config.SVHTTP_TIMEOUT_S * 900))       # 90 % of the job's time limit, in ms


def _close(page) -> None:
    try:
        page.close()
    except Exception:   # noqa: BLE001 - the tab may already be gone
        pass


class BrowserTransport:
    """Posts provider XML from a signed-in tab of the drill browser: its own tab for its pod (plan1) and one extra
    tab per other pod (FinPlan: stratmodel), signed in once with the same SSO. All page work runs on the browser's
    worker thread via submit(); _tabs is only touched there. Every failure to get an answer is an SvHttpError
    (so extract.route can fall back); a refused write stays a ValueError.

    `browser` is a callable returning the drill browser, resolved on every post: after
    drill_browser.shutdown_shared() the next post uses the new shared browser, not the shut one."""

    def __init__(self, browser: Callable[[], object]):
        self._browser = browser
        self._tabs: dict[str, tuple[object, object]] = {}       # pod base -> (browser context, page)

    def post(self, base: str, body: str) -> str:
        svxml.check_read_only(body)
        fut = None
        try:
            b = self._browser()
            fut = b.submit(lambda s: self._post(b, s, base.rstrip("/"), body))
            return fut.result(timeout=config.SVHTTP_TIMEOUT_S)
        except cf.CancelledError:
            raise SvHttpError("the drill browser job was cancelled (browser restarting or shutting down)") from None
        except cf.TimeoutError:
            restarted = ""
            if fut is not None and not fut.cancel() and not fut.done():   # still queued: cancelled, never runs
                try:                         # running and hung: restart OUR browser (like sv_drill's hang guard), so
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

    def _post(self, b, s, base: str, body: str) -> str:
        s.deadline = time.monotonic() + config.SVHTTP_TIMEOUT_S
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
                status, text = page.evaluate(_JS, [body, _fetch_ms()])
            except (SvHttpError, DrillError):                    # sign-in already retried inside
                raise
            except Exception as e:   # noqa: BLE001 - the tab navigated or the browser died: one fresh try
                if attempt == 2:
                    raise SvHttpError(f"browser error at {base}: {_first_line(e)}") from None
                self._drop(b, s, base)
                continue
            if isinstance(text, str) and ("<res_" in text or "<exception" in text):
                return text                                       # provider XML, whatever the status: svxml judges
            why = f"answered HTTP {status}"
            log.warning("svhttp: HTTP %s from %s; signing in again", status, base)
            self._drop(b, s, base)
        raise SvHttpError(f"Smart View provider at {base}: {why} twice (sign-in failed?)")


_shared: SvHttp | None = None
_shared_lock = threading.Lock()


def shared() -> SvHttp:
    """The process's one HTTP route (on the shared drill browser, looked up on every post)."""
    global _shared
    with _shared_lock:
        if _shared is None:
            _shared = SvHttp(BrowserTransport(lambda: drill_browser.shared_browser()))
        return _shared
