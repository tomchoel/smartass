"""Grids over the Smart View provider's HTTP protocol - no Excel, no COM (Stage 0, 2026-10-25:
docs/specs/2026-10-25-stage0-epm-access-probe.md). retrieve() returns what ops._query_once returns,
(Layout, Parsed, seconds), so sv_query can use either route; the COM route stays the fallback.

Error contract for sv_query(source='auto'): InputError only for a query that is wrong whatever the route (a dimension
used twice, no members, no row/column axis); everything else - including an incomplete or stale atlas, which COM
does not need - is SvHttpError, so the caller falls back to Smart View in Excel.

The transport is BrowserTransport: it posts from inside the drill's hidden Edge (Telenor SSO), one tab per EPM pod,
on the drill browser's worker thread. Tests use a fake transport."""
from __future__ import annotations

import concurrent.futures as cf
import logging
import threading
import time
from pathlib import Path
from typing import Callable, Protocol

from . import atlas, config, drill_browser, svxml
from .drill_browser import DrillError
from .grid import Axis, Layout, full_pov, parse, validate_axes, verify_labels
from .host import InputError

log = logging.getLogger(__name__)

CUBES = {"FinMax": ("FinMax", "FinMax"), "FinLite": ("FinMax", "FinLite"), "FinPlan": ("FinPlan", "FinPlan")}


class SvHttpError(Exception):
    """The HTTP route could not answer; sv_query(source='auto') then uses Smart View in Excel."""


class Transport(Protocol):
    def post(self, base: str, body: str) -> str: ...


def _norm_dim(name: str) -> str:
    return name.removeprefix("All ").replace(" ", "").casefold()


def _first_line(e: BaseException) -> str:
    """One redacted line of an exception (Playwright messages carry call logs and can quote sso_token URLs)."""
    return drill_browser.redact(str(e).splitlines()[0] if str(e).strip() else type(e).__name__)[:200]


class SvHttp:
    def __init__(self, transport: Transport, home: str | None = None, db_path: Path | None = None):
        self.t = transport
        self.home = (home or config.EPM_BASE).rstrip("/")
        self.db_path = db_path or config.ATLAS_DB
        self._bases: dict[str, str] | None = None
        self._sids: dict[str, str] = {}
        self._lock = threading.Lock()

    def base_for(self, app: str) -> str:
        """The app's pod. The list is cached once it is non-empty; an app missing from it is looked up once more."""
        if self._bases and app in self._bases:
            return self._bases[app]
        bases = svxml.data_sources(self.t.post(self.home, svxml.provisioned_data_sources()))
        self._bases = bases or None
        if app not in bases:
            raise SvHttpError(f"EPM lists no Smart View data source for app {app} "
                              f"(has: {', '.join(sorted(bases)) or 'none'})")
        return bases[app]

    def _note_sid(self, conn: str, text: str) -> None:
        """Each answer may carry a new sID; the latest one is the session."""
        sid = svxml.sid_in(text)
        if sid:
            self._sids[conn] = sid

    def _session(self, conn: str) -> tuple[str, str, str]:
        app, cube = CUBES[conn]
        base = self.base_for(app)
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
            dims = atlas.dimensions(db, conn)
            stale = [d for d in dims if not atlas.built(db, conn, d)]
        finally:
            db.close()
        if not dims:
            raise SvHttpError(f"the atlas has no dimensions for {conn}: build it with sv_atlas_build")
        if stale:
            raise SvHttpError(f"the atlas is incomplete or stale for {conn} (dimensions {', '.join(stale)}); "
                              "Smart View in Excel will be used")
        return dims

    def retrieve(self, conn: str, pov: dict | None, rows: list[Axis], cols: list[Axis],
                 suppress_missing: bool, suppress_zeros: bool):
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
            raise SvHttpError(f"the atlas is incomplete or stale for {conn} ({'; '.join(unknown)}); "
                              "Smart View in Excel will be used")
        if problems:
            raise InputError("Invalid query: " + "; ".join(problems))
        lay = Layout(full_pov(dims, pov, rows, cols), rows, cols)
        if not self._lock.acquire(timeout=config.SVHTTP_TIMEOUT_S):
            raise SvHttpError("the HTTP route is busy with another query")
        try:
            values = self._execute(conn, lay)
        finally:
            self._lock.release()
        bad = verify_labels(lay, parse(lay, values, False, False))
        if bad:
            raise SvHttpError("Smart View provider returned a different layout than requested (rows/columns "
                              "moved): " + "; ".join(bad[:5]))
        return lay, parse(lay, values, suppress_missing, suppress_zeros), time.monotonic() - t0

    def _execute(self, conn: str, lay: Layout) -> list[list]:
        for attempt in (1, 2):
            try:
                base, cube, sid = self._session(conn)
                ans = self.t.post(base, svxml.execute_grid(sid, cube, lay))
                self._note_sid(conn, ans)
                return svxml.grid_values(ans)
            except svxml.SvXmlError as e:
                if "Grid is Invalid" in str(e):           # unknown member: deterministic, a new session won't help
                    raise SvHttpError(f"Smart View provider: {e}") from None
                self._sids.pop(conn, None)                  # maybe an expired session: reconnect once
                self._bases = None                          # ... and re-read where the apps live
                if attempt == 2:
                    raise SvHttpError(f"Smart View provider: {e}") from None
                log.info("svhttp %s: %s - reconnecting once", conn, e)
        raise AssertionError("unreachable")


_JS = """async (b) => { const r = await fetch('/HyperionPlanning/SmartView', {method: 'POST', credentials: 'include',
    body: b, headers: {'Content-Type': 'text/xml'}}); return [r.status, await r.text()]; }"""


def _close(page) -> None:
    try:
        page.close()
    except Exception:   # noqa: BLE001 - the tab may already be gone
        pass


class BrowserTransport:
    """Posts provider XML from a signed-in tab of the drill browser: its own tab for its pod (plan1) and one extra
    tab per other pod (FinPlan: stratmodel), signed in once with the same SSO. All page work runs on the browser's
    worker thread via submit(); _tabs is only touched there. Every failure to get an answer is an SvHttpError
    (so sv_query can fall back); a refused write stays a ValueError.

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
            if fut is not None:
                fut.cancel()                                      # still queued: do not run it for nobody
            raise SvHttpError(f"the drill browser did not answer within {config.SVHTTP_TIMEOUT_S:.0f} s") from None
        except (SvHttpError, ValueError):
            raise
        except DrillError as e:                                   # includes DrillBusy
            raise SvHttpError(str(e)) from None
        except Exception as e:   # noqa: BLE001 - e.g. a raw Playwright error: never hand it to sv_query
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
        status = None
        for attempt in (1, 2):
            try:
                b._ensure(s)
                status, text = self._tab(b, s, base).evaluate(_JS, body)
            except (SvHttpError, DrillError):                    # sign-in already retried inside
                raise
            except Exception as e:   # noqa: BLE001 - the tab navigated or the browser died: one fresh try
                if attempt == 2:
                    raise SvHttpError(f"browser error at {base}: {_first_line(e)}") from None
                self._drop(b, s, base)
                continue
            if isinstance(text, str) and ("<res_" in text or "<exception" in text):
                return text                                       # provider XML, whatever the status: svxml judges
            log.warning("svhttp: HTTP %s from %s; signing in again", status, base)
            self._drop(b, s, base)
        raise SvHttpError(f"Smart View provider at {base} answered HTTP {status} twice (sign-in failed?)")


_shared: SvHttp | None = None
_shared_lock = threading.Lock()


def shared() -> SvHttp:
    """The process's one HTTP route (on the shared drill browser, looked up on every post)."""
    global _shared
    with _shared_lock:
        if _shared is None:
            _shared = SvHttp(BrowserTransport(lambda: drill_browser.shared_browser()))
        return _shared
