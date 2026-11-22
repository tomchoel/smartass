"""The second engine: one grid over Oracle's documented Planning REST API - POST
/HyperionPlanning/rest/v3/applications/{app}/plantypes/{cube}/exportdataslice (a read: it exports a slice) through the
pod's epmhttp.EpmSession - returned as the provider route's Fetched (the same layout, labels and cells), so everything
after the fetch is unchanged. It is the fallback for a grid the provider route could not fetch
(svhttp.SvHttp.fetch_rest, data/provider.py); live 2026-11-22 it equalled the provider route cell for cell on 20
grids of the run log (docs/audit/2026-11-21-rest-vs-provider.md).

What differs from the provider, and how it is made the same (live 2026-11-21, read-only, one-cell grids on plan1 and
stratmodel, compared with the provider route by key paths, types and booleans only):
- dimension names: REST takes the cube's own names ('Accounts', 'TimePeriods'; FinPlan's 'Account', 'BalanceTypes');
  the atlas has Smart View's ('All Accounts', 'Account', 'All Balance Types'). The plan type's own list (GET
  .../plantypes/{cube}/dimensions?aliasTableName=Default, Oracle 26.04) gives each name with its Default alias, and
  every atlas dimension is its alias or its name there (map_dimensions); one it does not list is an error naming it;
- members: REST takes a member by name or alias and answers its NAME. The atlas keeps Smart View's labels and, where
  a label has no 'code alias (prefix)' form ('Sep 26', 'Periodic', FinPlan's members), the label as the name - REST
  then answers a name the atlas does not know ('Y26M09'). Each label is the atlas's: of the member asked for at that
  place (REST resolved the spelling the atlas keeps - its answered name may even be one the atlas holds for another
  member, live), else of the member REST answered; else the name is kept and one warning says so. A top member is
  sent and answered under its dimension's REST name. Where the atlas was built over REST (epm/members.py) its side
  table holds the real names behind each key (two members can give one key), and REST must answer one of them at its
  place - the key itself only when it is one of them (else EPM resolved the key to another member whose name it is);
  a key without side rows there passes only as itself, and not when the side table records that name behind another
  key. In an atlas without a trusted side table (the hidden Excel's, or one an older release rebuilt) the check is
  the one before the side table: a member kept under its label has its place as the only link between label and
  number, so the order matters there: REST answers rows and columns in the order asked (live 2026-11-22: three such
  members asked against outline order on rows and on columns, per application, each read alone first for its name,
  cell for cell the provider route's answer);
- cells: numbers come as text and are read only in their plain form (digits, one '.', an exponent; anything else is
  an error naming the cell's shape, never a guess); an empty cell ("" live, " " in Oracle's documentation) is None,
  as the provider route's Fetched holds a blank (live 2026-11-22: every missing cell of 20 grids is "" over REST and
  None from the provider);
- no suppression: every row and column asked comes back (the provider route suppresses nothing either - grid.parse
  does). The answer must be exactly the grid asked for (_answer, _matrix), then the provider route's own layout check
  runs (svhttp.check_layout);
- a grid over MAX_CELLS (cells as EPM counts them, epm_cells: its rows and two header rows, times a row's member
  headers and data cells) is split along its rows into requests of at most MAX_CELLS cells (and bodies within the
  rule's size cap, epmhttp.MAX_BODY_BYTES), stitched in order; EPM itself refuses a request over 500 000 (live);
- no access: Oracle documents that exportdataslice exports only cells the user has read-write access to, and no
  marker for the others (#NoAccess was never seen on either route, live 2026-11-22): a cell the user may not read
  cannot be told apart from an empty one, which is read as None - the reply's warning for REST grids says so
  (pipeline.prepare). A "#NoAccess" text is still read as Smart View's label.

Errors are svhttp.SvHttpError (RestError): EPM refusing the grid (HTTP 400, Oracle's message scrubbed), a 404 on the
plan type (NotConfigured), an answer that is not the grid asked for, a dimension the plan type lists and the atlas
lacks (AtlasStale). A 400 whose message is Oracle's sentence for a member the plan type does not have there (or the
user may not read: _NO_MEMBER, live 2026-11-22) and names a member asked - exactly as it was sent, in that dimension -
is MemberRefused: data/provider.py leaves a member of an axis of several out of the request and fetches it again; a
POV member or an axis's only member cannot be left out. A 400 naming a member that was not asked is never acted on
(an error that says to rebuild the atlas). A gate's answer (401, 3xx, 403: epmhttp.GATE_STATUSES, as on the provider
route) is the session's one new token, then TokenRefused; a lost request is svhttp.RequestLost; oauth.NotSignedIn
reaches the caller as it is; a query wrong on every route is InputError and a stale atlas svhttp.AtlasStale, as on
the provider route."""
from __future__ import annotations

import json
import math
import re
import threading
import time
from collections import Counter

from ..knowledge import atlas
from . import epmhttp, oauth, svhttp, svxml
from .drillcore import redact
from .grid import Axis, Layout, name_from_label
from .svhttp import Fetched, SvHttpError

ROUTE = "rest"
# Cells per request as EPM counts them (epm_cells). EPM refuses a grid over its threshold with HTTP 400 ('Unable to
# load the form as the number of cells exceeded the threshold ... Error Threshold: 500000', live 2026-11-22; 466 000
# counted cells answered in 5.7 s, 2.8 MB): a request takes half of it.
EPM_CELL_THRESHOLD = 500_000
MAX_CELLS = EPM_CELL_THRESHOLD // 2
HEADER_ROWS = 2                                       # the POV row and the one column dimension's row
LIST_TTL_S = 3600.0                                   # a plan type's dimension list is read again after an hour
MAX_MESSAGE = 300                                     # characters of Oracle's message in an error text
_JWT = re.compile(r"[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}")
# Oracle's HTTP 400 for a member the plan type does not have there (live 2026-11-22, invented members on rows,
# columns and the POV of both applications: the same text under 'detail', 'message' and 'localizedMessage'; the
# member as sent - in double quotes when it has a space -, the dimension's REST name; only the first such member of
# the grid is named). The whole sentence must match: anything else is a refused grid like any other 400.
_NO_MEMBER = re.compile(r'The member (?:"(?P<quoted>.+)"|(?P<bare>\S+)) does not exist in the (?P<dim>.+?) dimension '
                        r'for the specified cube or you do not have access to it\.')
ATLAS_BUILD = "check the member or rebuild the atlas (`smartass atlas build`)"


class RestError(SvHttpError):
    """The REST engine could not fetch the grid: EPM refused it, or its answer is not the grid asked for."""


class NotConfigured(RestError):
    """HTTP 404 on the plan type: the pod has no such application or plan type (svhttp.CUBES, the data source's pod)."""


class RequestLost(RestError, svhttp.RequestLost):
    """The request left and no whole answer came (it is a read: a caller may send it again)."""


class TokenRefused(RestError, svhttp.TokenRefused):
    """EPM's gate refused the access token twice (the session's one new token did not help)."""


class AtlasStale(RestError, svhttp.AtlasStale):
    """The plan type lists a dimension the atlas lacks: build the atlas (`smartass atlas build`)."""


class MemberRefused(RestError):
    """EPM refused the grid for one member asked (HTTP 400: it does not exist there, or the user may not read it).
    dimension: the atlas's; member: as asked; droppable: it is on an axis of several members, so the grid can be
    fetched without it - a POV member or the only member of its axis cannot be left out (every remaining cell would
    mean something else, or there would be none)."""

    def __init__(self, text: str, dimension: str, member: str, droppable: bool):
        super().__init__(text)
        self.dimension, self.member, self.droppable = dimension, member, droppable


class _NamesMember(RestError):
    """A 400 with Oracle's wrong-member sentence, before fetch_grid has matched its member against the grid."""

    def __init__(self, text: str, rest_dimension: str, member: str):
        super().__init__(text)
        self.rest_dimension, self.member = rest_dimension, member


# (pod, app, cube) -> (when read, the plan type's (name, alias, type), the atlas's dimensions it did not map when it
# was read, else None): read again after LIST_TTL_S, or at once (once) when it no longer maps the atlas's dimensions
# (a dimension renamed in EPM needs no restart)
_listed: dict[tuple[str, str, str], tuple[float, list[tuple], tuple[str, ...] | None]] = {}
_listed_lock = threading.Lock()
_clock = time.monotonic


def fetch_grid(session: epmhttp.EpmSession, conn: str, pov: dict | None, rows: list[Axis], cols: list[Axis],
               atlas_db) -> Fetched:
    """One grid over REST - the arguments and the answer of svhttp.SvHttp.fetch - on the session of the
    connection's pod. atlas_db: an open atlas.db (knowledge.atlas.connect) of the calling thread."""
    if conn not in svhttp.CUBES:
        raise RestError(f"no REST route for connection {conn}")
    t0 = time.monotonic()
    _, app, cube = svhttp.CUBES[conn]
    dims = svhttp.dimensions_in(atlas_db, conn)
    lay = svhttp.layout(conn, dims, pov, rows, cols)
    if lay.nc != 1:                                   # every grid smartass asks has one; Oracle's shape for more unseen
        raise RestError(f"a REST grid takes one column dimension, not {lay.nc}")
    for a in lay.rows + lay.cols:                     # live: REST answers a member asked twice once
        twice = [m for m, n in Counter(m.casefold() for m in a.members).items() if n > 1]
        if twice:
            raise RestError(f"{a.dimension} member {twice[0]} is asked twice: Planning REST answers it once")
    names = _rest_dimensions(session, conn, app, cube, dims)
    per_row = lay.nr + len(lay.col_tuples)            # as EPM counts a row: its member headers and its cells
    most = MAX_CELLS // per_row - HEADER_ROWS         # data rows per request
    if most < 1:
        raise RestError(f"one row of {per_row} cells and the header rows ({epm_cells(1, per_row)}) are over the cell "
                        f"budget of a REST request ({MAX_CELLS})")
    got_pov = got_cols = None
    heads: list[list[str]] = []
    data: list[list] = []
    try:
        for part, body in _requests(lay, names, most):
            p, c, h, d = _answer(_post(session, app, cube, body), lay, part)
            if got_pov is not None and (p, c) != (got_pov, got_cols):
                raise RestError("the REST requests of one grid answered different POV or column members")
            got_pov, got_cols = p, c
            heads += h
            data += d
    except _NamesMember as e:
        raise _refused(conn, lay, names, e) from None
    values, warnings = _matrix(atlas_db, conn, lay, names, got_pov, got_cols, heads, data)
    svhttp.check_layout(lay, values, "Planning REST")
    return Fetched(lay, values, time.monotonic() - t0, ROUTE, warnings)


def epm_cells(rows: int, per_row: int) -> int:
    """Cells of a request as EPM counts them against its threshold: the grid's data rows and its HEADER_ROWS, times a
    row's width (its member headers and its data cells). Live 2026-11-22: 'Current Value: 588028' for 42 000 rows of
    2 headers and 12 cells, (42 000 + 2) x 14."""
    return (rows + HEADER_ROWS) * per_row


# ---------------------------------------------------------------- dimension names
def map_dimensions(conn: str, dims: list[str], listed: list[tuple]) -> dict[str, str]:
    """{atlas dimension: REST name}: the one listed dimension whose Default alias or name is the atlas's name
    (exactly, else in any case, as Essbase compares names). listed: (name, alias[, dimType]) of the plan type. A
    dimension with no single match is an error naming it; a listed dimension the atlas lacks (attribute dimensions
    aside) means the atlas is incomplete or stale - the grid would leave it out."""
    base = [(n, a) for n, a, *t in listed if not (t and str(t[0]).startswith("Attribute"))]
    out, unmapped = {}, []
    for d in dims:
        hits = {n for n, a in base if d in (n, a)} or {n for n, a in base if d.casefold() in (n.casefold(),
                                                                                           (a or "").casefold())}
        if len(hits) == 1:
            out[d] = hits.pop()
        else:
            unmapped.append(d)
    if unmapped:
        raise RestError(f"Planning REST lists no single dimension for {conn}'s {', '.join(map(repr, unmapped))} "
                        f"(it lists {', '.join(n + (f' ({a})' if a else '') for n, a in base) or 'none'})")
    extra = [n for n, _ in base if n not in out.values()]
    if extra:
        raise AtlasStale(f"the atlas is incomplete or stale for {conn} (Planning REST also lists {', '.join(extra)}): "
                         "build it with `smartass atlas build`")
    return out


def _rest_dimensions(session, conn: str, app: str, cube: str, dims: list[str]) -> dict[str, str]:
    """The atlas's dimensions mapped to the plan type's REST names, from its list - cached, read again once it is
    LIST_TTL_S old, and at once (once) when the cached one no longer maps. A list read for these very dimensions that
    did not map them is not read again before it expires: the same error (a stale atlas) from the cached list."""
    key = (session.origin, app, cube)
    with _listed_lock:
        read_at, listed, failed = _listed.get(key, (None, None, None))
    if listed is not None and _clock() - read_at < LIST_TTL_S:
        if failed == tuple(dims):
            return map_dimensions(conn, dims, listed)     # raises as it did when it was read
        try:
            return map_dimensions(conn, dims, listed)
        except RestError:                             # renamed in EPM since it was read: read it again
            pass
    listed = read_dimensions(session, app, cube)
    try:
        out = map_dimensions(conn, dims, listed)
    except RestError:
        with _listed_lock:
            _listed[key] = (_clock(), listed, tuple(dims))
        raise
    with _listed_lock:
        _listed[key] = (_clock(), listed, None)
    return out


def read_dimensions(session, app: str, cube: str) -> list[tuple]:
    """The plan type's dimensions as (name, Default alias, type), one page - uncached (the atlas build reads them
    once per build, epm/members.Source; fetch_grid through _rest_dimensions' cache)."""
    r = _send(session, "GET", epmhttp.dimension_list(app, cube) + "?aliasTableName=Default", None, app, cube)
    j = _json(r.text)
    items = j.get("items") if isinstance(j, dict) else None
    if not isinstance(items, list) or not all(isinstance(it, dict) and isinstance(it.get("name"), str)
                                              for it in items):
        raise RestError("not a Planning REST dimension list")
    if j.get("hasMore"):
        raise RestError(f"Planning REST lists more than {len(items)} dimensions for {app}/{cube} (one page)")
    return [(it["name"], it.get("alias") if isinstance(it.get("alias"), str) else None, it.get("dimType"))
            for it in items]


# ---------------------------------------------------------------- the request
def _requests(lay: Layout, names: dict[str, str], most: int) -> list[tuple[list[Axis], str]]:
    """(row axes, JSON body) of each request, in the grid's row order: at most 'most' rows each (_split), and a body
    within the rule's size cap (epmhttp.MAX_BODY_BYTES) - a part whose body is over it is cut in two along its rows.
    All are built before the first is sent; one row whose body is over the cap is an error."""
    out: list[tuple[list[Axis], str]] = []
    todo = _split(lay.rows, most)[::-1]
    while todo:
        part = todo.pop()
        body = json.dumps(_body(lay, part, names))
        if len(body.encode("utf-8")) <= epmhttp.MAX_BODY_BYTES:
            out.append((part, body))
            continue
        n = math.prod(len(a.members) for a in part)
        if n == 1:
            raise RestError(f"one row's request is over the {epmhttp.MAX_BODY_BYTES} bytes a request takes")
        todo += _split(part, (n + 1) // 2)[::-1]
    return out


def _split(rows: list[Axis], most: int) -> list[list[Axis]]:
    """The row axes as cross products of at most 'most' rows each, in the grid's row order: the outermost axis is
    cut into runs; where one of its members is still too many rows, its inner axes are cut the same way."""
    n = math.prod(len(a.members) for a in rows)
    if n <= most:
        return [rows]
    first, inner = rows[0], rows[1:]
    per = math.prod(len(a.members) for a in inner)
    if per <= most:
        k = most // per
        return [[Axis(first.dimension, first.members[i:i + k]), *inner] for i in range(0, len(first.members), k)]
    return [[Axis(first.dimension, [m]), *part] for m in first.members for part in _split(inner, most)]


def _sent(dim: str, member: str, names: dict[str, str]) -> str:
    """A member as REST takes it: a dimension's top member under the dimension's REST name, any other as asked."""
    return names[dim] if member.casefold() == dim.casefold() else member


def _body(lay: Layout, rows: list[Axis], names: dict[str, str]) -> dict:
    def segment(axes: list[Axis]) -> dict:
        return {"dimensions": [names[a.dimension] for a in axes],
                "members": [[_sent(a.dimension, m, names) for m in a.members] for a in axes]}
    return {"exportPlanningData": False, "gridDefinition": {
        "suppressMissingBlocks": False, "suppressMissingRows": False, "suppressMissingColumns": False,
        "pov": {"dimensions": [names[d] for d, _ in lay.pov], "members": [[_sent(d, m, names)] for d, m in lay.pov]},
        "columns": [segment(lay.cols)], "rows": [segment(rows)]}}


def _post(session, app: str, cube: str, body: str):
    return _json(_send(session, "POST", epmhttp.export_slice(app, cube), body, app, cube).text)


def _send(session, method: str, url: str, body: str | None, app: str, cube: str):
    """One request through the session; its answer when HTTP 200, else the error for its status. A gate's answer
    takes the session's one new token, as on the provider route."""
    where = f"Planning REST at {session.origin}"
    try:
        r = session.request(method, url, headers={"Content-Type": "application/json"} if body is not None else None,
                            body=body, timeout=svhttp.timeout_s(), renew_on=epmhttp.GATE_STATUSES)
    except oauth.NotSignedIn:
        raise
    except oauth.OAuthError as e:
        raise RestError(f"no EPM access token for {session.origin}: {e}") from None
    except epmhttp.NoScope:
        raise svhttp.NoScope(f"no OAuth scope is configured for {session.origin} ([oauth.scopes] in config.toml)") \
            from None
    except (epmhttp.Refused, epmhttp.TooLarge) as e:   # their texts quote no body, token or query
        raise RestError(f"{where}: {e}") from None
    except OSError as e:
        raise RequestLost(f"{where}: lost the request ({type(e).__name__})") from None
    if r.status == 200:
        return r
    if r.status == 400:
        text = f"{where} refused the grid (HTTP 400): {_message(r.text)}"
        named = _named(r.text)
        raise _NamesMember(text, *named) if named else RestError(text)
    if r.status in epmhttp.GATE_STATUSES:             # the session took one new token for it already
        raise TokenRefused(f"{where}: EPM refused the access token twice (HTTP {r.status})")
    if r.status == 404:
        raise NotConfigured(f"{where} has no plan type {cube} in application {app} (HTTP 404): check the "
                            "application and plan type names (svhttp.CUBES) and the data source's pod")
    raise RestError(f"{where} answered HTTP {r.status}")


def _message(text: str) -> str:
    """Oracle's message from an error answer, without a URL query, a token-like string or a line break, cut short."""
    try:
        j = json.loads(text)
    except ValueError:
        return "not JSON"
    msg = next((j[k] for k in ("detail", "message", "title", "error")
                if isinstance(j, dict) and isinstance(j.get(k), str) and j[k].strip()), None)
    if msg is None:
        return "no message"
    msg = redact(_JWT.sub("...", re.sub(r"\s+", " ", msg))).strip()
    return msg[:MAX_MESSAGE] + (" ..." if len(msg) > MAX_MESSAGE else "")


def _named(text: str) -> tuple[str, str] | None:
    """(REST dimension, member as sent) from Oracle's wrong-member sentence (_NO_MEMBER, the whole message), else
    None."""
    try:
        j = json.loads(text)
    except ValueError:
        return None
    for k in ("detail", "message", "localizedMessage"):
        m = _NO_MEMBER.fullmatch(j.get(k) or "") if isinstance(j, dict) and isinstance(j.get(k), str) else None
        if m:
            return m["dim"], m["quoted"] if m["quoted"] is not None else m["bare"]
    return None


def _refused(conn: str, lay: Layout, names: dict[str, str], e: _NamesMember) -> RestError:
    """The member a 400 names, matched against the grid asked: exactly as it was sent (_sent), in the dimension it
    names. MemberRefused when it was asked there; else an error that acts on nothing."""
    grid = f"the {conn} grid ({' / '.join(m for _, m in lay.pov)})"
    dim = next((d for d, n in names.items() if n == e.rest_dimension), None)
    pov = [m for d, m in lay.pov if d == dim and _sent(d, m, names) == e.member]
    axis = next((a for a in lay.rows + lay.cols if a.dimension == dim
                 and any(_sent(dim, m, names) == e.member for m in a.members)), None)
    if dim is None or not (pov or axis):
        return RestError(f"{e} - Planning REST refused {grid} for a member that was not asked: {ATLAS_BUILD}")
    if pov:
        return MemberRefused(f"Planning REST refused {grid}: {dim} member {pov[0]} does not exist there or you have no "
                             f"access to it; as a POV member it cannot be left out - {ATLAS_BUILD}", dim, pov[0], False)
    member = next(m for m in axis.members if _sent(dim, m, names) == e.member)
    if len(axis.members) == 1:
        return MemberRefused(f"Planning REST refused {grid}: {dim} member {member} does not exist there or you have no "
                             f"access to it; as the only {dim} member asked it cannot be left out - {ATLAS_BUILD}",
                             dim, member, False)
    return MemberRefused(f"Planning REST refused {grid}: {dim} member {member} does not exist there or you have no "
                         "access to it", dim, member, True)


def _json(text: str):
    try:
        return json.loads(text)
    except ValueError:
        raise RestError("not a Planning REST answer (not JSON)") from None


# ---------------------------------------------------------------- the answer
def _answer(j, lay: Layout, rows: list[Axis]) -> tuple[list[str], list[str], list[list[str]], list[list]]:
    """(POV names, column names, row names, row cells) of one answer, in Oracle's shape: {"pov": [names], "columns":
    [[a name per column]], "rows": [{"headers": [names], "data": [cells]}]} - exactly as many of each as asked."""
    def names(xs, n: int) -> bool:
        return isinstance(xs, list) and len(xs) == n and all(isinstance(x, str) for x in xs)
    if (not isinstance(j, dict) or not {"pov", "columns", "rows"} <= set(j) or not isinstance(j["rows"], list)
            or not all(isinstance(r, dict) for r in j["rows"])):
        raise RestError("not a Planning REST grid (no pov, columns and rows)")
    ncols, nrows = len(lay.col_tuples), math.prod(len(a.members) for a in rows)
    if not names(j["pov"], len(lay.pov)):
        raise RestError(f"the Planning REST grid's POV is not the {len(lay.pov)} members asked")
    cols = j["columns"]
    if not (isinstance(cols, list) and len(cols) == 1 and names(cols[0], ncols)):
        raise RestError(f"the Planning REST grid's columns are not the {ncols} column members asked")
    if len(j["rows"]) != nrows:
        raise RestError(f"the Planning REST grid has {nrows} rows asked, {len(j['rows'])} answered")
    heads, data = [], []
    for r in j["rows"]:
        if not names(r.get("headers"), lay.nr):
            raise RestError(f"a Planning REST row has not {lay.nr} member headers")
        if not isinstance(r.get("data"), list) or len(r["data"]) != ncols:
            raise RestError(f"a Planning REST row has not {ncols} cells")
        heads.append(r["headers"])
        data.append([_cell(v) for v in r["data"]])
    return j["pov"], cols[0], heads, data


# A plain number as REST writes one: digits, one '.' and digits, an exponent - nothing a number format adds (a
# thousands separator, a decimal comma, a sign '+', an underscore): such a cell is an error, never a guess.
_PLAIN = re.compile(r"-?[0-9]+(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?")


def _shape(s: str) -> str:
    """What a cell looks like, never its value: each run of digits as 9, of letters as a, the rest as it is, cut."""
    out = re.sub(r"[^\W\d_]+", "a", re.sub(r"\d+", "9", s))
    return out if len(out) <= 20 else f"{out[:20]}..."


def _cell(v):
    """A REST cell as the provider route's Fetched holds it: a float, None (an empty cell, svxml's blank) or Smart
    View's label (#Missing, #NoAccess, #Invalid). Anything else fails the grid (RestError, 'non-numeric data', naming
    the cell's shape or JSON type, never its value): a text that is not a plain number (_PLAIN) is never guessed - under
    an EPM number format of another user '1.234' could mean 1 234 (final review A)."""
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            f = float(v)
        except OverflowError:                         # a JSON integer past a float's range (10**400)
            raise RestError("non-numeric data in the Planning REST grid: a JSON integer too large for a number"
                            ) from None
        if not math.isfinite(f):
            raise RestError("non-numeric data in the Planning REST grid: a number that is not finite")
        return f
    if not isinstance(v, str):
        kind = {True: "true", False: "false"}.get(v) if isinstance(v, bool) else (
            "list" if isinstance(v, list) else "object" if isinstance(v, dict) else type(v).__name__)
        raise RestError(f"non-numeric data in the Planning REST grid: a JSON {kind}, not a number")
    s = v.strip()
    if not s:
        return None
    special = svxml.SPECIAL.get(s.casefold())
    if special:
        return special
    if not _PLAIN.fullmatch(s):
        raise RestError(f"non-numeric data in the Planning REST grid: a cell shaped '{_shape(s)}' is not a plain "
                        "number (digits, one '.', an exponent), and it is never guessed")
    f = float(s)
    if not math.isfinite(f):                          # '1e999': plain in form, too large for a number
        raise RestError("non-numeric data in the Planning REST grid: a number that is not finite")
    return f


def _labels(db, conn: str, dim: str, spellings: set[str]) -> tuple[dict, dict]:
    """The atlas's members of one dimension among these spellings: ({folded name: (name, label, real names)}, {folded
    label: (name, label, real names)}) - by name exactly (one indexed query), else by name in any case or by label (one
    pass). The real names are EPM's from the atlas's side table where it is trusted (atlas.real_names: an atlas the
    build over REST wrote, not rebuilt since by an older release); else empty."""
    by_name: dict[str, tuple[str, str]] = {}
    by_label: dict[str, tuple[str, str]] = {}
    want = {s.casefold() for s in spellings}
    ordered = sorted(spellings)
    for i in range(0, len(ordered), 500):              # SQLite's limit on query parameters
        part = ordered[i:i + 500]
        for n, lab in db.execute(f"SELECT name, label FROM member WHERE connection=? AND dimension=? AND name IN "
                                 f"({','.join('?' * len(part))})", (conn, dim, *part)):
            by_name[n.casefold()] = (n, lab or n)
    if want - set(by_name):
        for n, lab in db.execute("SELECT name, label FROM member WHERE connection=? AND dimension=?", (conn, dim)):
            if n.casefold() in want:
                by_name.setdefault(n.casefold(), (n, lab or n))
            if lab and lab.casefold() in want:
                by_label.setdefault(lab.casefold(), (n, lab))
    reals = atlas.real_names(db, conn, dim, {n for n, _ in [*by_name.values(), *by_label.values()]})

    def with_reals(d: dict) -> dict:
        return {k: (n, lab, frozenset(reals.get(n, ()))) for k, (n, lab) in d.items()}
    return with_reals(by_name), with_reals(by_label)


def _matrix(db, conn: str, lay: Layout, names: dict[str, str], pov: list[str], cols: list[str],
            heads: list[list[str]], data: list[list]) -> tuple[list[list], list[str]]:
    """The provider's matrix from A1 with the atlas's labels for the names REST answered, and the warnings. Each
    place's name must be the member asked there: one of its real names where the atlas's side table is trusted and
    holds the key (an atlas built over REST; the key itself only as one of them), exactly the key where the atlas keeps
    the member under a 'code alias (prefix)' label, the dimension's REST name for a top member; elsewhere (an older
    atlas that keeps the label as the name) one member asked is always one name answered, and two members two names -
    and the place itself ties label and number, so REST must answer in the order asked (live: it does). A key without
    side rows in a trusted side table passes only as itself, and not when the side table records that name behind
    another key (atlas.real_owners). Otherwise RestError, naming the key by its code: an error for the grid, never
    another member's numbers under the key's label."""
    places: dict[str, list[tuple[str, str]]] = {}     # dimension -> [(asked, answered), ...] in grid order
    for (d, m), got in zip(lay.pov, pov):
        places.setdefault(d, []).append((m, got))
    for k, combo in enumerate(lay.col_tuples):
        places.setdefault(lay.cols[0].dimension, []).append((combo[0], cols[k]))
    for i, combo in enumerate(lay.row_tuples):
        for d, a in enumerate(lay.rows):
            places.setdefault(a.dimension, []).append((combo[d], heads[i][d]))
    label: dict[tuple[str, str, str], str] = {}       # (dimension, asked, answered) -> label
    warnings: list[str] = []
    for d, pairs in places.items():
        _one_to_one(d, pairs)
        looked = {a for a, _ in pairs} | {g for a, g in pairs if a.casefold() not in (g.casefold(), d.casefold())}
        by_name, by_label = _labels(db, conn, d, looked)   # not a top member's REST name: no pass over the dimension
        owners: list = []                             # atlas.real_owners of d, read once at the first key it needs
        for asked, got in dict.fromkeys(pairs):
            top = asked.casefold() == d.casefold()
            known = by_name.get(asked.casefold()) or by_label.get(asked.casefold())
            if known and not top and not known[2]:    # no side row: only the key itself, and not another key's member
                if not owners:
                    owners.append(atlas.real_owners(db, conn, d))
                if owners[0] is not None:             # the side table is trusted: a build over REST, every member in it
                    other = sorted(owners[0].get(got.casefold(), set()) - {known[0]})
                    if other or got.casefold() != known[0].casefold():
                        why = (f"EPM's name of the member behind {other[0]} in the atlas" if other else
                               "a name the atlas's side table does not record for it")
                        raise RestError(f"Planning REST returned a different layout than requested: {d} {known[0]} "
                                        f"answered as {got}, {why}")
            if top or (known and (known[2] or name_from_label(known[1]))):   # REST answers that very name
                if got.casefold() not in ({names[d].casefold()} if top else
                                          {x.casefold() for x in known[2] or (known[0],)}):
                    raise RestError(f"Planning REST returned a different layout than requested: {d} "
                                    f"{asked if top else known[0]} answered as {got}")
                label[d, asked, got] = known[1] if known else d
            elif known:                               # REST resolved the label the atlas keeps as the name
                label[d, asked, got] = known[1]
            elif hit := _answered(db, conn, d, got, by_name, owners):   # asked in a spelling the atlas does not keep
                label[d, asked, got] = hit[1]
            else:
                label[d, asked, got] = got
                warnings.append(f"{conn}: the atlas has no label for {d} member {got} (kept as its name)")
    m = lay.matrix()
    for j, ((d, asked), got) in enumerate(zip(lay.pov, pov)):
        m[0][lay.nr + j] = label[d, asked, got]
    for k, combo in enumerate(lay.col_tuples):
        m[1][lay.nr + k] = label[lay.cols[0].dimension, combo[0], cols[k]]
    for i, combo in enumerate(lay.row_tuples):
        r = 1 + lay.nc + i
        for d, a in enumerate(lay.rows):
            m[r][d] = label[a.dimension, combo[d], heads[i][d]]
        m[r][lay.nr:lay.nr + len(data[i])] = data[i]
    return m, warnings


def _answered(db, conn: str, dim: str, got: str, by_name: dict, owners: list) -> tuple | None:
    """The atlas's member EPM answered as got, for a member asked in a spelling the atlas does not keep: with a trusted
    side table the key whose side rows record got as EPM's name (atlas.real_owners) - never a key that is merely
    spelled like got, which may be another member's (Task 4c review, Minor 3); none or two such keys: None. Else (no
    trusted side table) the key spelled got. owners: real_owners of dim, read once per grid and dimension."""
    if not owners:
        owners.append(atlas.real_owners(db, conn, dim))
    if owners[0] is None:
        return by_name.get(got.casefold())
    keys = owners[0].get(got.casefold(), set())
    if len(keys) != 1:
        return None
    key = next(iter(keys))
    return _labels(db, conn, dim, {key})[0].get(key.casefold())


def _one_to_one(dim: str, pairs: list[tuple[str, str]]) -> None:
    """One member asked is always one name answered, and two members asked are two names."""
    seen: dict[str, str] = {}
    back: dict[str, str] = {}
    for asked, got in pairs:
        a, g = asked.casefold(), got.casefold()
        if seen.setdefault(a, g) != g or back.setdefault(g, a) != a:
            raise RestError(f"Planning REST returned a different layout than requested: {dim} {asked} answered as "
                            f"{got}")
