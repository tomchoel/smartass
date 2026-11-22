"""The token route's one way to EPM: HTTPS requests to one pod with the user's OAuth access token (epm/oauth.py) and
a cookie jar, as a browser tab of that pod sends them - no browser, no Excel. The grids (svhttp.BearerTransport, the
Smart View provider XML), the second grid engine (rest: Planning REST exportdataslice) and the drill over the token
(drill_token) are built on it.

What it holds to, whatever a caller asks for:
- read-only towards EPM: a request that is not exactly on ALLOWED is refused before a token is asked for - its
  method, its path, the query names that path takes, the header names it takes and, for the POST targets, the body
  (Smart View provider XML that svxml.check_read_only admits; the drill's getData query object; one exportdataslice
  grid definition). The same token would be accepted by EPM's write endpoints (importdataslice, jobs): never here;
- the bearer goes only to the session's own pod, and only when config.OAUTH_SCOPES has a scope for it: the URL that
  is sent is built from the pod's origin and the allowed path, never from the host a caller wrote; a redirect is
  never followed (a 3xx is the answer, for the caller to judge);
- a token, the x-appbuilder-client-id key, a cookie or a URL query is never in a log line or an exception text
  (drillcore.redact on every URL shown), nor in a Response's repr (Set-Cookie stays in the jar).

A request that left and got no whole answer raises OSError (RequestLost for http.client's own errors): EPM may have
run it."""
from __future__ import annotations

import http.client
import http.cookiejar
import json
import logging
import re
import socket
import threading
import time
import urllib.error
import urllib.request
from typing import NamedTuple
from urllib.parse import unquote_plus, urlsplit

from . import oauth, svxml
from .drillcore import redact

log = logging.getLogger(__name__)

SMART_VIEW = "/HyperionPlanning/SmartView"
GET_DATA = "/aif/ui/model/drilldown/getData"
PLANNING_REST = "/HyperionPlanning/rest/v3/applications"
# Planning REST (Oracle's documented API): the (application, plan type) pairs the grid route reads - svhttp.CUBES's,
# pinned equal by tests/test_epmhttp.py (svhttp imports this module, not the other way round). Each rule below is a
# whole path built from them: no pattern, so an encoded, dotted, longer or other-case path is none of them.
REST_CUBES = (("FinMax", "FinMax"), ("FinMax", "FinLite"), ("FinPlan", "FinPlan"))


def export_slice(app: str, cube: str) -> str:
    """POST: one grid's cells (exportdataslice exports: a read; importdataslice is never on ALLOWED)."""
    return f"{PLANNING_REST}/{app}/plantypes/{cube}/exportdataslice"


def dimension_list(app: str, cube: str) -> str:
    """GET (Oracle's "Get Dimensions for a Plan Type", 26.04): the plan type's dimensions under the names REST takes,
    with their Default aliases (?aliasTableName=Default) - rest.py maps the atlas's names onto them."""
    return f"{PLANNING_REST}/{app}/plantypes/{cube}/dimensions"

# Oracle's "Get Member" (GET, the member and by default its whole visible subtree; it creates nothing - Add Member is a
# POST to .../members, never here): the atlas reads a dimension's tree from its top member, which is the dimension's
# own name. Only these (application, dimension) pairs, as whole paths, no query.
MEMBER_TREES = (("FinMax", "TimePeriods"), ("FinPlan", "Period"))


def member_tree(app: str, dim: str) -> str:
    """GET: dimension dim of application app as one tree, read from its top member (Oracle's Get Member)."""
    return f"{PLANNING_REST}/{app}/dimensions/{dim}/members/{dim}"


ANY = re.compile(r".*")
DIGITS = re.compile(r"\d{1,12}")
MILLIS = re.compile(r"\d{1,13}")                       # the cache buster: int(time.time() * 1000), 13 digits now
VERSION = re.compile(r"\d{1,6}(?:\.\d{1,6}){0,3}")    # Smart View's rcp_version on its drill link (live 2026-11-21)


class _Rule(NamedTuple):
    method: str
    query: dict[str, tuple[int, int | None, re.Pattern]]   # name -> (least, most times (None: any), its value)
    headers: frozenset[str]                                  # the caller's header names it takes (lower case)
    content_type: str | None                                 # a POST's body type (the session sets it if missing)


_DRILL_HEADERS = frozenset({"accept", "content-type", "x-requested-with", "page_id", "x-appbuilder-client-id"})
# The brief's table, with what each request may carry besides: the query names are those the drill sends today
# (sui_url keeps Smart View's link minus DRILLURL; drill_token's view and getData URLs), nothing else.
ALLOWED: dict[str, _Rule] = {
    # The grid route's one request (svhttp.BearerTransport: no query, Content-Type text/xml only); tests/test_epmhttp.py
    # pins that nothing else sends it.
    SMART_VIEW: _Rule("POST", {}, frozenset({"content-type"}), "text/xml"),          # + svxml.check_read_only(body)
    "/aif/ui/model/applications/servicetype": _Rule("GET", {}, _DRILL_HEADERS, None),
    "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp": _Rule("GET", {
        "FORMAT": (1, 1, re.compile("SUI")), "module": (1, 1, re.compile(r"aif\.drilldown")),
        "attribute": (0, None, ANY), "sso_token": (0, 1, ANY), "rcp_version": (0, 1, VERSION), "_": (0, 1, MILLIS)},
        _DRILL_HEADERS, None),
    "/aif/ui/model/view": _Rule("GET", {
        "page": (1, 1, re.compile("DRILL_THROUGH")), "appId": (0, 1, ANY), "partitionkey": (0, 1, ANY),
        "_": (0, 1, MILLIS)}, _DRILL_HEADERS, None),
    GET_DATA: _Rule("POST", {"limit": (1, 1, re.compile(r"\d{1,5}")), "offset": (1, 1, DIGITS)}, _DRILL_HEADERS,
                    "application/json"),
    # The REST grid (rest.py): no query, Content-Type application/json only, a body _slice_grid admits.
    **{export_slice(app, cube): _Rule("POST", {}, frozenset({"content-type"}), "application/json")
       for app, cube in REST_CUBES},
    **{dimension_list(app, cube): _Rule("GET", {"aliasTableName": (1, 1, re.compile("Default"))}, frozenset(), None)
       for app, cube in REST_CUBES},
    # A dimension's member tree (the atlas): no query, no header, no body.
    **{member_tree(app, dim): _Rule("GET", {}, frozenset(), None) for app, dim in MEMBER_TREES},
}
SLICE_PATHS = frozenset(export_slice(app, cube) for app, cube in REST_CUBES)
GET_DATA_KEYS = {"QUERY": str, "PARAMS": list, "PARAMSTYPE": list}   # drillcore.parse_sui / drill_derive build these
# exportdataslice's body (Oracle's documented request; rest.py builds it): exactly these keys at every level
SLICE_KEYS = frozenset({"exportPlanningData", "gridDefinition"})
SLICE_GRID_KEYS = frozenset({"suppressMissingBlocks", "suppressMissingRows", "suppressMissingColumns", "pov", "columns",
                             "rows"})
SLICE_AXIS_KEYS = frozenset({"dimensions", "members"})
_SLICE_NAME = re.compile(r"[^\x00-\x1f\x7f]{1,255}")   # a dimension or member name: no control character
# Oracle's member functions (Children(x), ILvl0Descendants(x) ...) and Essbase's @-functions: a grid smartass builds
# names its members (GridRequest holds names; no atlas member has either form, live 2026-11-22), so a member that
# reads as a function - '@...', a word right before a bracket, or a function's name before one - is refused.
_SLICE_FUNCTION = re.compile(r"\s*(?:@|[A-Za-z_]\w*\(|(?:I?L?R?Siblings|I?Children|I?Descendants|I?Ancestors|I?Parents?"
                             r"|I?Lvl0Descendants|Level0Descendants|Relative|UDA|Attribute|Generation|Level|Match"
                             r"|Range)\s*\()", re.I)
# Statement words that write or administer: never a word of a drill query (outside strings, quoted names, comments)
SQL_WRITES = frozenset({"INSERT", "UPDATE", "DELETE", "MERGE", "DROP", "ALTER", "TRUNCATE", "CREATE", "GRANT", "REVOKE",
                        "EXEC", "EXECUTE", "CALL", "BEGIN", "COMMIT", "DECLARE"})
_SQL_WORD = re.compile(r"[A-Za-z_][A-Za-z0-9_$#]*")
HEADER_VALUES = {"x-requested-with": re.compile("XMLHttpRequest")}
SECRET_HEADERS = ("x-appbuilder-client-id",)            # values scrubbed from an error text and a repr, like the token
HIDDEN_HEADERS = ("set-cookie", "set-cookie2")          # kept by the jar, never handed to a caller
MAX_BODY_BYTES = 1 << 20                                # a getData query is far smaller
# A provider request is about 4 bytes a grid cell (svxml.execute_grid, measured 2026-11-21): 1 MB would stop a grid of
# some 260 000 cells, which the provider route always sent; 16 MB is about 4 million, about as many as an answer under
# MAX_ANSWER_BYTES holds (12 to 18 bytes a cell in the provider's answer shape: 3.7 to 5.6 million).
MAX_SMART_VIEW_BODY_BYTES = 16 << 20
MAX_ANSWER_BYTES = 64 << 20                             # a 5000-line getData page is about 7 MB (1.3 kB a line)
MAX_TIMEOUT_S = 600.0
# A gate's answers: EPM did not take the token, so the request never ran and may go once more with a new one. A 200
# or a 500 did run (sending it again would run it twice): request() refuses a renew_on outside these.
GATE_STATUSES = (301, 302, 303, 307, 401, 403)        # a tuple: renew_on takes one (svhttp, rest)
_URL = re.compile(r"[\x21-\x7e]+")                      # printable ASCII: urlsplit drops tab and line breaks silently
_NAME = re.compile(r"[!#$%&'*+.^_`|~0-9A-Za-z-]+")      # an HTTP header name (RFC 9110 token)
_VALUE = re.compile(r"[\x20-\x7e]*")                    # no line break: http.client would quote the value in its error


class Response(NamedTuple):
    status: int
    text: str
    headers: dict[str, str]                             # names lower-cased; no Set-Cookie

    def __repr__(self) -> str:
        shown = {k: "..." if k in SECRET_HEADERS else redact(v) for k, v in self.headers.items()}
        return f"Response(status={self.status}, text=<{len(self.text)} chars>, headers={shown})"


class Refused(ValueError):
    """The request is not on the allowlist or leaves the pod: nothing was sent and no token was asked for."""


class NoScope(Refused):
    """The pod has no scope under [oauth.scopes]: nothing was sent (serverkit.NEXT_STEP says which scope to add)."""


class RequestLost(ConnectionError):
    """http.client's own errors (a cut-off answer, a bad status line) as an OSError: the request left, its answer
    did not come whole."""


class TooLarge(ValueError):
    """EPM's answer is over MAX_ANSWER_BYTES: reading stopped there."""


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None                                     # urllib then raises HTTPError: the 3xx is the answer


def _opener(*handlers) -> urllib.request.OpenerDirector:
    """urllib's opener that never follows a redirect (handlers: tests put a fake wire under it). Cookies are the
    session's own jar, not a handler, so a fake opener sees them too."""
    return urllib.request.build_opener(_NoRedirect, *handlers)


class EpmSession:
    """Allowlisted requests to one pod with the user's bearer and one cookie jar. tokens: epm/oauth.py (scope_for,
    access_token); opener: anything with open(request, timeout)."""

    def __init__(self, base: str, tokens=oauth, opener=None):
        self.base = base
        self.origin = oauth.origin(base)                # 'https://host', lower case; '' when base has no host
        self._host = urlsplit(self.origin).hostname or ""
        self._tokens = tokens
        self._opener = opener or _opener()
        self._jar = http.cookiejar.CookieJar()          # thread-safe; only ever sent to self.origin

    def request(self, method: str, url: str, headers: dict | None = None, body: str | None = None,
                timeout: float = 120.0, renew_on: tuple[int, ...] = (401,), renew_if=None) -> Response:
        """Send one allowlisted request with the bearer; a 401 (or another status of renew_on, all of them gate
        answers: GATE_STATUSES) takes one new token (reject=the refused one), drops the session's cookies (they came
        with the refused token) and is sent once more - unless renew_if(the answer) says no, which keeps that answer
        as it is (it can only narrow renew_on: the grid route keeps a gate status that carries provider XML, which
        reached the provider). Any other status is an answer (a 3xx too: never followed). Refused before anything is
        sent; oauth.OAuthError / NotSignedIn as the token source raises them; OSError when the request left and no
        whole answer came (its text scrubbed of secrets); TooLarge for an answer over MAX_ANSWER_BYTES."""
        try:
            if (not isinstance(renew_on, tuple)
                    or not all(type(s) is int and s in GATE_STATUSES for s in renew_on)):   # a bool is no status
                raise Refused("refused: only a gate's answer (3xx, 401, 403) may take a new token and go again")
            target, rule = self._target(method, url, body)
            sent = self._headers(headers, rule)
            if not isinstance(timeout, (int, float)) or isinstance(timeout, bool) or not 0 < timeout <= MAX_TIMEOUT_S:
                raise Refused(f"refused: a timeout must be a number of seconds over 0 and at most {MAX_TIMEOUT_S:g}")
            if self._tokens.scope_for(self.origin) is None:
                raise NoScope(f"refused: no OAuth scope is configured for {self.origin or redact(self.base)!r}")
        except Refused as e:
            log.warning("epmhttp: %s", e)
            raise
        data = None if body is None else body.encode("utf-8")
        t0 = time.monotonic()
        token = self._tokens.access_token(self.origin)
        r = self._send(method, target, sent, data, token, timeout)
        if r.status in renew_on and (renew_if is None or renew_if(r)):   # the hour may have ended early: one more
            self._jar.clear()                           # first: the cookies go even when no new token comes
            token = self._tokens.access_token(self.origin, reject=token)
            r = self._send(method, target, sent, data, token, timeout)
        log.debug("epmhttp: %s %s -> HTTP %s (%.2f s)", method, redact(target), r.status, time.monotonic() - t0)
        return r

    def _target(self, method: str, url: str, body: str | None) -> tuple[str, _Rule]:
        """The URL to send - the pod's origin, the allowed path and the caller's query - and its rule, or Refused."""
        if not self._host:
            raise Refused(f"refused: {redact(self.base)!r} is not an EPM pod URL")
        if not isinstance(url, str) or not _URL.fullmatch(url):
            raise Refused("refused: a URL with a space, a control or a non-ASCII character")
        shown = redact(url)
        try:
            parts = urlsplit(url)
            port = parts.port
        except ValueError:
            raise Refused(f"refused: {shown} is not a URL") from None
        if parts.scheme or parts.netloc:
            if (parts.scheme.lower() != "https" or "@" in parts.netloc or parts.hostname != self._host
                    or port not in (None, 443)):
                raise Refused(f"refused: {shown} is not on {self.origin}")
        rule = ALLOWED.get(parts.path)
        if rule is None or rule.method != method:
            raise Refused(f"refused: {method!r} {shown} is not on the allowlist")
        _check_query(method, shown, parts.query, rule)
        if method == "GET" and body is not None:
            raise Refused(f"refused: GET {shown} with a body")
        if method == "POST":
            _check_body(method, shown, parts.path, body)
        return self.origin + parts.path + (f"?{parts.query}" if parts.query else ""), rule

    @staticmethod
    def _headers(headers: dict | None, rule: _Rule) -> dict[str, str]:
        """The caller's headers if each name is one the request takes (the session sets Authorization, Cookie,
        Host, Content-Length and Accept-Encoding itself), plus the body's Content-Type when it was left out."""
        out = dict(headers or {})
        for name, value in out.items():
            if not isinstance(name, str) or not _NAME.fullmatch(name):
                raise Refused("refused: a header name that is not a token")
            low = name.lower()
            if low not in rule.headers:
                raise Refused(f"refused: the {name} header is not one this request takes")
            _check_value(name, value)
            if low in HEADER_VALUES and not HEADER_VALUES[low].fullmatch(value):
                raise Refused(f"refused: the {name} header has a value this request does not take")
            kind = value.split(";")[0].strip().lower()
            if low == "content-type" and kind != (rule.content_type or "application/json"):
                raise Refused(f"refused: the content type {kind[:40]!r} for this request")
        if [k.lower() for k in out].count("content-type") > 1:
            raise Refused("refused: two Content-Type headers")
        if rule.content_type and not any(k.lower() == "content-type" for k in out):
            out["Content-Type"] = rule.content_type
        return out

    def _send(self, method: str, url: str, headers: dict[str, str], data: bytes | None, token: str,
              timeout: float) -> Response:
        bearer = "Bearer " + str(token)
        _check_value("Authorization", bearer)
        req = urllib.request.Request(url, data=data, method=method,
                                     headers={**headers, "Accept-Encoding": "identity", "Authorization": bearer})
        self._jar.add_cookie_header(req)
        deadline = time.monotonic() + timeout
        secrets = [token, *(v for k, v in headers.items() if k.lower() in SECRET_HEADERS)]
        err = None
        try:
            try:
                r = self._opener.open(req, timeout=timeout)
                status = r.status
            except urllib.error.HTTPError as e:         # outside 2xx, a 3xx too (_NoRedirect): an answer
                r, status = e, e.code
            with r:
                self._jar.extract_cookies(r, req)
                text = _read(r, deadline, method, url)
                info = r.headers
        except OSError as e:
            err = self._scrubbed(e, secrets)
        except http.client.HTTPException as e:          # IncompleteRead, BadStatusLine, LineTooLong ...
            err = self._scrubbed(RequestLost(f"{type(e).__name__}: {e}"), secrets)
        if err is not None:                             # raised out here: a scrubbed error has no __context__
            raise err from None
        return Response(status, text, {k.lower(): v for k, v in info.items() if k.lower() not in HIDDEN_HEADERS})

    def _scrubbed(self, e: OSError, secrets: list[str]) -> OSError:
        """e itself when its text quotes nothing secret; else the same kind of error with the token, the key, the
        cookies and any URL query cut out."""
        with self._jar._cookies_lock:                   # CookieJar.__iter__ does not lock: another request may write
            cookies = [c.value for c in self._jar]
        secrets = [s for s in (*secrets, *cookies) if s]

        def scrub(text: str) -> str:
            for s in secrets:
                text = text.replace(s, "...")
            return redact(text)
        text = str(e)
        if scrub(text) == text:
            return e
        try:
            new = (type(e)(scrub(str(e.reason))) if isinstance(e, urllib.error.URLError) else type(e)(scrub(text)))
        except Exception:   # noqa: BLE001 - an OSError kind that wants other arguments
            new = OSError(scrub(text))
        return new if scrub(str(new)) == str(new) else OSError(scrub(text))


def _check_query(method: str, shown: str, query: str, rule: _Rule) -> None:
    """Every query name is one the path takes (compared percent-decoded, as the servlet decodes it), as many times
    as it may appear, with the value it must have (compared as sent)."""
    seen: dict[str, list[str]] = {}
    for item in (p for p in query.split("&") if p):
        k, _, v = item.partition("=")
        seen.setdefault(unquote_plus(k), []).append(v)
    for name, values in seen.items():
        if name not in rule.query:
            raise Refused(f"refused: {method} {shown} with a query item this request does not take")
        least, most, value = rule.query[name]
        if (most is not None and len(values) > most) or not all(value.fullmatch(v) for v in values):
            raise Refused(f"refused: {method} {shown}: {name} has a value or a count this request does not take")
    for name, (least, _, value) in rule.query.items():
        if len(seen.get(name, ())) < least:
            want = f"{name}={value.pattern}" if value is not ANY and value is not DIGITS else name
            raise Refused(f"refused: {method} {shown} needs {want} in its query")


def _check_body(method: str, shown: str, path: str, body) -> None:
    if not isinstance(body, str):
        raise Refused(f"refused: POST {shown} without a body")
    cap = MAX_SMART_VIEW_BODY_BYTES if path == SMART_VIEW else MAX_BODY_BYTES
    if len(body.encode("utf-8")) > cap:
        raise Refused(f"refused: POST {shown} with a body over {cap} bytes")
    if path == SMART_VIEW:
        try:
            svxml.check_read_only(body)
        except ValueError:                              # its text quotes the body's start (a session id): not here
            raise Refused("refused: not a single read-only Smart View request") from None
        return
    j = _json_object(shown, body)
    if path in SLICE_PATHS:
        if not _slice_grid(j):                          # its text is never quoted (member names)
            raise Refused(f"refused: POST {shown} with a body that is not one exportdataslice grid")
        return
    if not isinstance(j, dict) or any(k not in GET_DATA_KEYS or not isinstance(v, GET_DATA_KEYS[k])
                                      for k, v in j.items()):
        raise Refused(f"refused: POST {shown} with a body that is not a drill query ({', '.join(GET_DATA_KEYS)})")
    if "QUERY" in j and not _read_only_sql(j["QUERY"]):
        raise Refused("refused: the drill query is not one read-only SELECT")      # its text is never quoted


def _json_object(shown: str, body: str):
    """The body as JSON with no key twice (a second key would be read by one parser and not by another)."""

    def no_twice(pairs):
        keys = [k for k, _ in pairs]
        if len(keys) != len(set(keys)):
            raise ValueError("a key twice")
        return dict(pairs)
    try:
        return json.loads(body, object_pairs_hook=no_twice)
    except ValueError:
        raise Refused(f"refused: POST {shown} with a body that is not one JSON object") from None


def _slice_grid(j) -> bool:
    """One exportdataslice grid: exportPlanningData false (no supporting detail or comments), the three suppress
    flags as booleans, a POV with one member per dimension, and at least one column and one row segment - each a
    list of dimension names with one non-empty member list per dimension; no name is a member function. Nothing
    else, at any level."""
    def names(xs) -> bool:
        return isinstance(xs, list) and all(isinstance(x, str) and _SLICE_NAME.fullmatch(x)
                                            and not _SLICE_FUNCTION.match(x) for x in xs)

    def segment(s, pov: bool = False) -> bool:
        if not isinstance(s, dict) or set(s) != SLICE_AXIS_KEYS:
            return False
        dims, members = s["dimensions"], s["members"]
        return (names(dims) and (pov or bool(dims)) and isinstance(members, list) and len(members) == len(dims)
                and all(names(m) and (len(m) == 1 if pov else bool(m)) for m in members))
    if not isinstance(j, dict) or set(j) != SLICE_KEYS or j["exportPlanningData"] is not False:
        return False
    g = j["gridDefinition"]
    return (isinstance(g, dict) and set(g) == SLICE_GRID_KEYS
            and all(type(g[k]) is bool for k in ("suppressMissingBlocks", "suppressMissingRows",
                                                 "suppressMissingColumns"))
            and segment(g["pov"], pov=True)
            and all(isinstance(g[k], list) and g[k] and all(segment(s) for s in g[k]) for k in ("columns", "rows")))


_Q_CLOSE = {"[": "]", "{": "}", "(": ")", "<": ">"}   # Oracle's q'[...]': a bracket closes with its pair
_SPACE = " \t\r\n\f\v"                                  # ASCII whitespace only: anything else is not certain


def _q_end(query: str, k: int) -> int | None:
    """The index after Oracle's alternative quoting q'<d>...<d>' whose delimiter <d> is at k (a bracket closes with
    its pair, any other character with itself, followed by a quote); None when there is no delimiter, it is
    whitespace, a quote or not printable ASCII, or the literal does not end."""
    d = query[k] if k < len(query) else ""
    if not d or d in _SPACE or d == "'" or not (d.isascii() and d.isprintable()):
        return None
    j = query.find(_Q_CLOSE.get(d, d) + "'", k + 1)
    return None if j < 0 else j + 2


def _sql_tokens(query: str) -> list[str] | None:
    """The words (upper case) and other characters of an SQL text outside its strings ('...', and Oracle's q'[...]'
    and nq'[...]'), quoted names ("...") and comments (-- and /* */); None when one of those does not end, or when
    anything outside them cannot be read with certainty (a character that is not printable ASCII or ASCII
    whitespace, a q-quote glued to what comes before it). A tokeniser, not a parser: it only has to know what EPM
    would read as a statement word or a statement end - and refuse what it cannot tell."""
    out, i, n = [], 0, len(query)
    while i < n:
        c = query[i]
        if c in "'\"":
            j = i + 1
            j = query.find(c, i + 1)                    # 'it''s' reads as two strings: the same here
            if j < 0:
                return None
            i = j + 1
        elif query.startswith("--", i):
            ends = [k for k in (query.find("\n", i), query.find("\r", i)) if k >= 0]
            i = min(ends) if ends else n
        elif query.startswith("/*", i):
            j = query.find("*/", i + 2)
            if j < 0:
                return None
            i = j + 2
        elif c in _SPACE:
            i += 1
        elif m := _SQL_WORD.match(query, i):
            word = m.group().upper()
            if word in ("Q", "NQ") and query.startswith("'", m.end()):
                if i > 0 and (query[i - 1].isalnum() or query[i - 1] in "_$#.@'\""):
                    return None                         # 1q'[...]', 'a'q'[...]', t.q'[...]': not certain
                i = _q_end(query, m.end() + 1)
                if i is None:
                    return None
                continue
            out.append(word)
            i = m.end()
        elif c.isascii() and c.isprintable():
            out.append(c)
            i += 1
        else:
            return None                                 # not ASCII, or a control character: not certain
    return out


def _read_only_sql(query: str) -> bool:
    """EPM's drill query (it hands it out; the client posts it back): one SELECT (or WITH ... SELECT) - it starts
    with that word, has no ';' and no word that writes or administers, outside strings and comments."""
    tokens = _sql_tokens(query)
    return (bool(tokens) and tokens[0] in ("SELECT", "WITH") and ";" not in tokens
            and not SQL_WRITES.intersection(tokens))


def _read(r, deadline: float, method: str, url: str) -> str:
    """The answer's text, at most MAX_ANSWER_BYTES and no later than the deadline: one receive at a time (read1),
    each with the socket's timeout set to the time left, so a slow drip cannot outlast it."""
    sock, read = _socket(r), getattr(r, "read1", None) or r.read
    chunks, n = [], 0
    while True:
        left = deadline - time.monotonic()
        if left <= 0:
            raise TimeoutError(f"EPM's answer to {method} {redact(url)} did not end in time")
        if sock is not None and sock.fileno() != -1:   # closed for good once the body is read (urllib let go of it)
            sock.settimeout(left)
        chunk = read(1 << 16)
        if not chunk:
            return b"".join(chunks).decode("utf-8", "replace")
        n += len(chunk)
        if n > MAX_ANSWER_BYTES:
            raise TooLarge(f"EPM's answer to {method} {redact(url)} is over {MAX_ANSWER_BYTES >> 20} MB")
        chunks.append(chunk)


def _socket(r) -> socket.socket | None:
    """The socket under an answer (http.client.HTTPResponse, or the one an HTTPError wraps); None for a fake."""
    fp = r
    for _ in range(4):
        sock = getattr(getattr(fp, "raw", None), "_sock", None)
        if isinstance(sock, socket.socket):
            return sock
        fp = getattr(fp, "fp", None)
    return None


def _check_value(name: str, value) -> None:
    if not isinstance(value, str) or not _VALUE.fullmatch(value):
        raise Refused(f"refused: the {name} header has a line break or a character HTTP cannot carry")


_sessions: dict[str, EpmSession] = {}
_sessions_lock = threading.Lock()


def session(base: str) -> EpmSession:
    """The process's one session (one cookie jar) per pod origin."""
    pod = oauth.origin(base)
    with _sessions_lock:
        s = _sessions.get(pod)
        if s is None:
            s = _sessions[pod] = EpmSession(pod)
        return s
