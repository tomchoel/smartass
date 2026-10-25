"""Grids over the Smart View provider's HTTP protocol - no Excel, no COM (Stage 0, 2026-10-25:
docs/specs/2026-10-25-stage0-epm-access-probe.md). retrieve() returns what ops._query_once returns,
(Layout, Parsed, seconds), so sv_query can use either route; the COM route stays the fallback.

The transport is BrowserTransport: it posts from inside the drill's hidden Edge (Telenor SSO), one tab per EPM pod,
on the drill browser's worker thread. Tests use a fake transport."""
from __future__ import annotations

import concurrent.futures as cf
import logging
import threading
import time
from pathlib import Path
from typing import Protocol

from . import atlas, config, svxml
from .grid import Axis, Layout, full_pov, parse, validate_axes
from .drill_browser import DrillError
from .host import InputError

log = logging.getLogger(__name__)

CUBES = {"FinMax": ("FinMax", "FinMax"), "FinLite": ("FinMax", "FinLite"), "FinPlan": ("FinPlan", "FinPlan")}


class SvHttpError(Exception):
    """The HTTP route could not answer; sv_query(source='auto') then uses Smart View in Excel."""


class Transport(Protocol):
    def post(self, base: str, body: str) -> str: ...


def _norm_dim(name: str) -> str:
    return name.removeprefix("All ").replace(" ", "").casefold()


class SvHttp:
    def __init__(self, transport: Transport, home: str | None = None, db_path: Path | None = None):
        self.t = transport
        self.home = (home or config.EPM_BASE).rstrip("/")
        self.db_path = db_path or config.ATLAS_DB
        self._bases: dict[str, str] | None = None
        self._sids: dict[str, str] = {}
        self._lock = threading.Lock()

    def base_for(self, app: str) -> str:
        if self._bases is None:
            self._bases = svxml.data_sources(self.t.post(self.home, svxml.provisioned_data_sources()))
        if app not in self._bases:
            raise SvHttpError(f"EPM lists no Smart View data source for app {app} "
                              f"(has: {', '.join(sorted(self._bases))})")
        return self._bases[app]

    def _session(self, conn: str) -> tuple[str, str, str]:
        app, cube = CUBES[conn]
        base = self.base_for(app)
        if conn not in self._sids:
            sid = svxml.sid_of(self.t.post(base, svxml.connect()))
            sid = svxml.sid_of(self.t.post(base, svxml.open_application(sid, app)))
            svxml.raise_for_exception(self.t.post(base, svxml.open_cube(sid, app, cube)))
            self._sids[conn] = sid
        return base, cube, self._sids[conn]

    def dimensions(self, conn: str) -> list[str]:
        dims = atlas.dimensions(atlas.connect(self.db_path), conn)
        if not dims:
            raise SvHttpError(f"the atlas has no dimensions for {conn}: build it with sv_atlas_build")
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
        if problems:
            raise InputError("Invalid query: " + "; ".join(problems))
        lay = Layout(full_pov(dims, pov, rows, cols), rows, cols)
        with self._lock:
            values = self._execute(conn, lay)
        return lay, parse(lay, values, suppress_missing, suppress_zeros), time.monotonic() - t0

    def _execute(self, conn: str, lay: Layout) -> list[list]:
        for attempt in (1, 2):
            try:
                base, cube, sid = self._session(conn)
                return svxml.grid_values(self.t.post(base, svxml.execute_grid(sid, cube, lay)))
            except svxml.SvXmlError as e:
                if str(e).startswith("Grid is Invalid"):    # unknown member: deterministic, a new session won't help
                    raise SvHttpError(f"Smart View provider: {e}") from None
                self._sids.pop(conn, None)                   # maybe an expired session: reconnect once
                if attempt == 2:
                    raise SvHttpError(f"Smart View provider: {e}") from None
                log.info("svhttp %s: %s - reconnecting once", conn, e)
        raise AssertionError("unreachable")


_JS = """async (b) => { const r = await fetch('/HyperionPlanning/SmartView', {method: 'POST', credentials: 'include',
    body: b, headers: {'Content-Type': 'text/xml'}}); return [r.status, await r.text()]; }"""


class BrowserTransport:
    """Posts provider XML from a signed-in tab of the drill browser: its own tab for its pod (plan1) and one extra
    tab per other pod (FinPlan: stratmodel), signed in once with the same SSO. All page work runs on the browser's
    worker thread via submit(); _tabs is only touched there. Every failure to get an answer is an SvHttpError
    (so sv_query can fall back); a refused write stays a ValueError."""

    def __init__(self, browser):
        self.b = browser
        self._tabs: dict[str, tuple[object, object]] = {}       # pod base -> (browser context, page)

    def post(self, base: str, body: str) -> str:
        svxml.check_read_only(body)
        try:
            fut = self.b.submit(lambda s: self._post(s, base.rstrip("/"), body))
            return fut.result(timeout=config.SVHTTP_TIMEOUT_S)
        except cf.CancelledError:
            raise SvHttpError("the drill browser job was cancelled (browser restarting or shutting down)") from None
        except cf.TimeoutError:
            raise SvHttpError(f"the drill browser did not answer within {config.SVHTTP_TIMEOUT_S:.0f} s") from None
        except DrillError as e:                                   # includes DrillBusy
            raise SvHttpError(str(e)) from None

    def _tab(self, s, base: str):
        if base == self.b.base.rstrip("/"):
            return s.page
        ctx, page = self._tabs.get(base, (None, None))
        if ctx is not s.ctx or page is None or page.is_closed():
            page = s.ctx.new_page()
            try:
                self.b._login_once(page, s.deadline, base=base)
            except DrillError as first:
                log.warning("svhttp: sign-in on %s failed (%s); retrying in a fresh tab", base, first)
                try:
                    page.close()
                except Exception:   # noqa: BLE001 - the tab may already be gone
                    pass
                page = s.ctx.new_page()
                try:
                    self.b._login_once(page, s.deadline, base=base)
                except DrillError as second:
                    raise SvHttpError(f"sign-in on {base} failed twice: {second}") from None
            self._tabs[base] = (s.ctx, page)
        return page

    def _post(self, s, base: str, body: str) -> str:
        s.deadline = time.monotonic() + config.SVHTTP_TIMEOUT_S
        status = None
        for attempt in (1, 2):
            self.b._ensure(s)
            try:
                status, text = self._tab(s, base).evaluate(_JS, body)
            except SvHttpError:
                raise
            except Exception as e:   # noqa: BLE001 - the tab navigated or the browser died: one fresh try
                if attempt == 2:
                    raise SvHttpError(f"browser error at {base}: {str(e).splitlines()[0][:200]}") from None
                s.ready = False
                self._tabs.pop(base, None)
                continue
            if status == 200 and ("<res_" in text or "<exception" in text):
                return text
            log.warning("svhttp: HTTP %s from %s; signing in again", status, base)
            s.ready = False
            self._tabs.pop(base, None)
        raise SvHttpError(f"Smart View provider at {base} answered HTTP {status} twice (sign-in failed?)")


_shared: SvHttp | None = None
_shared_lock = threading.Lock()


def shared() -> SvHttp:
    """The process's one HTTP route (on the shared drill browser)."""
    global _shared
    with _shared_lock:
        if _shared is None:
            from .drill_browser import shared_browser
            _shared = SvHttp(BrowserTransport(shared_browser()))
        return _shared
