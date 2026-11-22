"""Grids over the Smart View provider's HTTP protocol - no Excel, no COM (Stage 0, 2026-10-25:
docs/specs/2026-10-25-stage0-epm-access-probe.md). retrieve() returns (Layout, Parsed, seconds); fetch() the checked
grid unparsed (Fetched). The second engine (epm/rest.py: Planning REST, fetch_rest) is the fallback for a grid this
route could not fetch (data/provider.py).

Error contract: InputError only for a query that is wrong whatever the route (a dimension used twice, no members, no
row/column axis); everything else is SvHttpError - NoFallback when the second engine would meet it as well (the same
token, scope or atlas, or an alias clash), so only the others are fetched again over REST. A member the cube does not
have where the grid puts it (WrongMember) is final on this route (no reconnect) but goes on to REST, whose HTTP 400
names it (data/provider.py leaves it out). A missing sign-in (oauth.NotSignedIn) is raised as it is and never falls
back.

BearerTransport posts over plain HTTPS with the user's OAuth access token (epm/oauth.py; no browser) to every pod that
has a scope, each request through the pod's epmhttp.EpmSession (its allowlist, its origin, its one new token on a
refusal); without a sign-in it raises oauth.NotSignedIn, which the tool call answers with the sign-in page and code.
Tests use a fake transport."""
from __future__ import annotations

import datetime as dt
import logging
import re
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol
from urllib.parse import urlsplit

from ..knowledge import atlas
from . import epmhttp, oauth, svxml
from ..core import config
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
# A member the cube does not have where the grid puts it (live 2026-11-22, stored answers with members renamed since):
# deterministic, and REST's HTTP 400 names it ('The member ... does not exist in the ... dimension'). Matched from the
# start of the provider's message (its desc), the whole sentence, so a transient error that only quotes a phrase is
# not taken for one.
WRONG_MEMBER = re.compile(r"(?:Error: )?(?:Grid is Invalid\.(?: .*)?|Member name and Alias placed on row \d+ do not "
                          r"match\.?|Member \S.* is out of place\.?)", re.S)
_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


class SvHttpError(Exception):
    """The HTTP route could not answer; the second engine (epm/rest.py) is tried unless it is NoFallback."""


class NoFallback(SvHttpError):
    """An error the second engine would meet as well - the same token, scope or atlas, or a grid EPM refuses for its
    members - so the grid is not fetched again over REST: this error is the answer."""


class NoScope(NoFallback):
    """The pod has no scope under [oauth.scopes]: no token can be asked for it (serverkit.NEXT_STEP says which scope
    to add and where it comes from). The second engine sends to the same pod with the same scope."""


class TokenRefused(NoFallback):
    """EPM's gate refused the access token twice (the session's one new token did not help either)."""


class AtlasStale(NoFallback):
    """The atlas has no dimensions for the connection, or one of them was not built, or the grid names a dimension
    it does not have: build it (`smartass atlas build`). Both engines lay the grid out from the atlas."""


class GridRefused(NoFallback):
    """The provider refuses the grid for an alias clash whatever the session (a member name that is another member's
    alias: ambiguous - REST would resolve it without saying which)."""


class WrongMember(SvHttpError):
    """The provider refuses the grid for a member the cube does not have where the grid puts it (WRONG_MEMBER: 'Grid
    is Invalid', a name and alias that do not match, a member out of place), whatever the session - so no reconnect;
    but not NoFallback: Planning REST's HTTP 400 names the member, so the second engine can leave it out."""


class RequestLost(SvHttpError):
    """The request left but its answer did not come (the connection dropped, or the request timed out): EPM may still
    be working on it, so its session is dropped and the request goes again on another one."""


class Transport(Protocol):
    def post(self, base: str, body: str) -> str: ...


def _norm_dim(name: str) -> str:
    return name.removeprefix("All ").replace(" ", "").casefold()


@dataclass
class Fetched:
    """A provider grid as fetched: Smart View's own shape from A1 (POV row, column headers, row labels, data). route:
    which engine fetched it ("http": the provider; "rest": epm/rest.py); warnings: what that engine wants said."""
    layout: Layout
    values: list[list]
    seconds: float
    route: str = "http"
    warnings: list[str] = field(default_factory=list)


def dimensions_in(db, conn: str) -> list[str]:
    """The connection's dimensions in the atlas (an open atlas.db); SvHttpError if it has none or any of them was
    not built (failed or empty; age does not matter) - a grid would then be built on a wrong dimension list."""
    dims = atlas.known_dimensions(db, conn)              # incl. failed builds (their members are deleted)
    stale = [d for d in dims if not atlas.built(db, conn, d)]
    if not dims:
        raise AtlasStale(f"the atlas has no dimensions for {conn}: build it with `smartass atlas build`")
    if stale:
        raise AtlasStale(f"the atlas is incomplete or stale for {conn} (dimensions {', '.join(stale)})")
    return dims


def layout(conn: str, dims: list[str], pov: dict | None, rows: list[Axis], cols: list[Axis]) -> Layout:
    """The grid's layout on the cube's dimensions (any spelling of a dimension, 'Accounts' or 'All Accounts'): the
    axes as asked and a POV member for every other dimension. InputError for a query wrong on every route;
    SvHttpError for a dimension the atlas does not have."""
    by_norm = {_norm_dim(d): d for d in dims}

    def canon(d: str) -> str:
        return by_norm.get(_norm_dim(d), d)
    pov = {canon(d): m for d, m in (pov or {}).items()}
    rows = [Axis(canon(a.dimension), a.members) for a in rows]
    cols = [Axis(canon(a.dimension), a.members) for a in cols]
    problems = validate_axes(dims, pov, rows, cols)
    unknown = [p for p in problems if p.startswith("unknown dimension")]
    if unknown:                       # the atlas may lack a dimension the cube has
        raise AtlasStale(f"the atlas is incomplete or stale for {conn} ({'; '.join(unknown)})")
    if problems:
        raise InputError("Invalid query: " + "; ".join(problems))
    return Layout(full_pov(dims, pov, rows, cols), rows, cols)


def check_layout(lay: Layout, values: list[list], who: str = "the provider") -> None:
    """The layout check a grid must pass before anything is built on it (both engines): it covers the layout, its
    data cells are numbers or Smart View's labels, no POV member was appended, and every label that names its member
    is the member asked for at that place. SvHttpError otherwise; who: the engine, for the text."""
    _check_grid(lay, values, who)
    full = parse(lay, values, False, False)
    if full.pov_problems:              # COM shows the same as a warning; here it means the layout is not the cube's
        raise SvHttpError(f"{who} appended POV members; the atlas may lack a dimension ("
                          + "; ".join(full.pov_problems) + ")")
    bad = verify_labels(lay, full)
    if bad:
        raise SvHttpError(f"{who} returned a different layout than requested (rows/columns moved): "
                          + "; ".join(bad[:5]))


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

    def _base(self, conn: str) -> str:
        """The pod of the connection's data source (base_for), one lookup at a time."""
        server, app, _ = CUBES[conn]
        if not self._bases_lock.acquire(timeout=config.SVHTTP_TIMEOUT_S * 2):   # behind a cold start's sign-in
            raise SvHttpError("the HTTP route is busy looking up the EPM data sources")
        try:
            return self.base_for(server, app)
        finally:
            self._bases_lock.release()

    def rest_session(self, conn: str) -> epmhttp.EpmSession:
        """The EpmSession of the connection's pod that the grid route sends through (one cookie jar per pod): the
        second engine (epm/rest.py) sends through it too. The pod is the data source's (base_for: a cached list
        once it was read), checked by pod_origin; SvHttpError when it cannot be found or the transport has no
        sessions (only the bearer route has them)."""
        if conn not in CUBES:
            raise SvHttpError(f"no HTTP route for connection {conn}")
        session = getattr(self.t, "session", None)
        if session is None:
            raise SvHttpError("the second engine needs the bearer route's sessions")
        return session(pod_origin(self._base(conn)))

    def _take(self, conn: str) -> tuple[str, str, str]:
        """(pod base, cube, sID) for ONE request: an idle session of the pool, else a new one. A session is never
        used by two requests at once: parallel grids on one sID come back wrong (2-3 of 8 corrupted, live
        2026-11-03), on their own sIDs all 8 are right."""
        _, app, cube = CUBES[conn]
        base = self._base(conn)
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
            return dimensions_in(db, conn)
        finally:
            db.close()

    def fetch_rest(self, conn: str, pov: dict | None, rows: list[Axis], cols: list[Axis]) -> "Fetched":
        """The same grid over the second engine (epm/rest.py: Planning REST) on the pod's session. The atlas is opened
        here, on the calling thread (a sqlite connection belongs to the thread that opened it), and closed."""
        from . import rest
        session = self.rest_session(conn)
        db = atlas.connect(self.db_path)
        try:
            return rest.fetch_grid(session, conn, pov, rows, cols, db)
        finally:
            db.close()

    def retrieve(self, conn: str, pov: dict | None, rows: list[Axis], cols: list[Axis],
                 suppress_missing: bool, suppress_zeros: bool):
        f = self.fetch(conn, pov, rows, cols)
        return f.layout, parse(f.layout, f.values, suppress_missing, suppress_zeros), f.seconds

    def fetch(self, conn: str, pov: dict | None, rows: list[Axis], cols: list[Axis]) -> "Fetched":
        """The checked provider grid, unparsed: what a source sheet is written from (core/source.py)."""
        if conn not in CUBES:
            raise SvHttpError(f"no HTTP route for connection {conn}")
        t0 = time.monotonic()
        lay = layout(conn, self.dimensions(conn), pov, rows, cols)
        values = self._execute(conn, lay)
        check_layout(lay, values)
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

    def _ask(self, conn: str, body, read, fatal, refused=lambda e: SvHttpError) -> object:
        """POST body(sid, cube) on a session of its own and read(answer); at most SVHTTP_SESSIONS at once per
        connection. An error fatal(e) does not rule out (maybe an expired session) drops that session and tries
        once more on a new one; a fatal one is raised at once as the class refused(e) gives. A session whose request
        was lost on the way is dropped too (EPM may still be working on it)."""
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
                        raise (refused(e) if fatal(e) else SvHttpError)(f"Smart View provider: {e}") from None
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
        def wrong(e) -> bool:                             # a member the cube lacks there
            return bool(WRONG_MEMBER.fullmatch(str(e).strip()))

        def clash(e) -> bool:                             # alias clash: deterministic too
            return any(m in str(e) for m in ALIAS_CLASH)
        return self._ask(conn, lambda sid, cube: svxml.execute_grid(sid, cube, lay), svxml.grid_values,
                         lambda e: wrong(e) or clash(e), lambda e: WrongMember if wrong(e) else GridRefused)


def _check_grid(lay: Layout, values: list[list], who: str = "the provider") -> None:
    """The returned slice must cover the requested layout (wider is fine: appended POV cells) and every data cell
    must be a number, blank or a Smart View label (#Missing, #NoAccess, #Invalid); anything else is not data."""
    m = lay.matrix()
    h, w = len(m), len(m[0])
    if len(values) < h:
        raise SvHttpError(f"{who} returned fewer rows than requested ({len(values)} < {h})")
    got_w = min((len(r) for r in values), default=0)
    if got_w < w:
        raise SvHttpError(f"{who} returned fewer columns than requested ({got_w} < {w})")
    labels = MISSING | NOACCESS | INVALID
    for r in range(1 + lay.nc, h):
        for c in range(lay.nr, lay.nr + len(lay.col_tuples)):
            v = values[r][c]
            if isinstance(v, str) and v.strip() not in labels:
                raise SvHttpError(f"non-numeric data in the grid {who} returned (row {r + 1}, column {c + 1}: "
                                  f"{v[:40]!r})")


_URL = re.compile(r"[\x21-\x7e]+")                      # printable ASCII: urlsplit drops tab and line breaks silently


def pod_origin(base: str) -> str:
    """'https://host' of a pod base URL - the origin its scope is configured for (oauth.origin) - when base is https
    on that host's own port (443) with no user, path, query or fragment; anything else (EPM's data-source list keeps
    the scheme and the port it was given) is SvHttpError before a token is asked for. The message shows the scheme,
    the host and the port only."""
    base = str(base or "").rstrip("/")
    try:
        parts = urlsplit(base)
        port = parts.port
    except ValueError:
        parts, port = urlsplit(base), -1
    shown = f"{parts.scheme}://{parts.hostname or ''}" + (f":{port}" if port not in (None, -1) else "")
    if (not _URL.fullmatch(base) or parts.scheme.lower() != "https" or not parts.hostname or "@" in parts.netloc
            or port not in (None, 443) or parts.path or "?" in base or "#" in base):
        raise SvHttpError(f"refused: {shown} is not https on an EPM pod's own host and port; the access token is not "
                          "sent there")
    return oauth.origin(base)


SLOW_READ_BYTES_S = 256 << 10      # a slow but working line: the largest answer (64 MB) in 256 s


def timeout_s() -> float:
    """One request's deadline for its whole answer (EpmSession reads to it). config.SVHTTP_TIMEOUT_S was urllib's
    limit per receive, so a big grid could take longer in all: the deadline is its old 90 % (the server's time to the
    first byte) plus the read of the largest answer the session takes (MAX_ANSWER_BYTES) at SLOW_READ_BYTES_S -
    364 s by default - within the most an EpmSession takes (600 s). A request that never answers waits that long
    before it is RequestLost (and sent again on a new sID, see SvHttp._ask)."""
    first_byte = max(1, int(config.SVHTTP_TIMEOUT_S * 900)) / 1000
    return min(first_byte + epmhttp.MAX_ANSWER_BYTES / SLOW_READ_BYTES_S, epmhttp.MAX_TIMEOUT_S)


def _provider_xml(text: str) -> bool:
    """The provider answered (its XML): the request reached it, whatever the status - never sent again on its sID."""
    return "<res_" in text or "<exception" in text


class BearerTransport:
    """Posts provider XML over plain HTTPS with the user's OAuth access token: no browser, no sign-in page (live
    2026-11-19 on plan1 and stratmodel: /HyperionPlanning/SmartView takes the token; grids, the closed month and
    drill links are cell for cell the browser route's; a cold first grid 0.7 s against about 20 s). It builds no
    request of its own: each goes through the pod's epmhttp.EpmSession, whose allowlist takes exactly this POST
    (Content-Type text/xml, a body svxml.check_read_only admits) and which sends the bearer only over https to the
    pod's origin. The base is checked first (pod_origin: never the scheme or port EPM's data-source list gives - the
    session would quietly make any of them its pod's origin). oauth.NotSignedIn reaches the caller as it is (the tool
    call answers with the sign-in page and code), and so does oauth.NotSetUp when this PC has no [oauth] set-up at
    all; a pod without a scope, any other oauth.OAuthError and a token EPM refuses twice (the session's one new token
    on a gate answer, epmhttp.GATE_STATUSES: TokenRefused) are SvHttpError; a request that left without a whole
    answer is RequestLost.

    One session (one cookie jar) per pod, as a browser has: EPM opens a server session for every request that brings
    no cookie. The cookie alone does not authenticate (302 without the token). The sessions are the grid route's
    own, not the drill's (epmhttp.session): the drill's key belongs to the EPM session of its jar, and a refused
    token clears a session's jar - a grid's 401 would sign the drill's key out mid-drill, and the drill's renewal
    waits only for its own requests. Like every transport here it sends nothing that svxml.check_read_only refuses -
    the same token would be accepted by EPM's write endpoints. opener: tests put a fake wire under the sessions."""

    def __init__(self, tokens=oauth, opener=None):
        self._tokens = tokens
        self._opener = opener
        self._sessions: dict[str, epmhttp.EpmSession] = {}
        self._lock = threading.Lock()

    def session(self, origin: str) -> epmhttp.EpmSession:
        """The grid route's EpmSession for a pod origin (pod_origin's answer)."""
        with self._lock:
            s = self._sessions.get(origin)
            if s is None:
                s = self._sessions[origin] = epmhttp.EpmSession(origin, tokens=self._tokens, opener=self._opener)
            return s

    def post(self, base: str, body: str) -> str:
        svxml.check_read_only(body)
        if not self._tokens.configured():
            raise oauth.NotSetUp()                       # the reply says what config.toml needs, not "retry"
        base = pod_origin(base)                          # the only origin the bearer goes to, whatever EPM listed
        try:
            r = self.session(base).request("POST", epmhttp.SMART_VIEW, headers={"Content-Type": "text/xml"},
                                           body=body, timeout=timeout_s(), renew_on=epmhttp.GATE_STATUSES,
                                           renew_if=lambda r: not _provider_xml(r.text))
        except oauth.NotSignedIn:
            raise
        except oauth.OAuthError as e:
            raise SvHttpError(f"no EPM access token for {base}: {e}") from None
        except epmhttp.NoScope:
            raise NoScope(f"no OAuth scope is configured for {base} ([oauth.scopes] in config.toml)") from None
        except (epmhttp.Refused, epmhttp.TooLarge) as e:   # their texts quote no body, token or query
            raise SvHttpError(f"Smart View provider at {base}: {e}") from None
        except OSError as e:                             # it may have reached EPM: never again on this sID
            raise RequestLost(f"Smart View provider at {base}: lost the request ({type(e).__name__})") from None
        if _provider_xml(r.text):
            return r.text                                # provider XML, whatever the status: svxml judges
        if r.status in epmhttp.GATE_STATUSES:            # the session took one new token for it already
            raise TokenRefused(f"Smart View provider at {base}: EPM refused the access token twice (HTTP {r.status})")
        raise RequestLost(f"Smart View provider at {base}: answered HTTP {r.status} without provider XML")


_shared: SvHttp | None = None
_shared_lock = threading.Lock()


def shared() -> SvHttp:
    """The process's one HTTP route: the bearer transport, with no browser behind it."""
    global _shared
    with _shared_lock:
        if _shared is None:
            _shared = SvHttp(BearerTransport())
        return _shared
