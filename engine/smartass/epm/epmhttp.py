"""The token route's one way to EPM: HTTPS requests to one pod with the user's OAuth access token (epm/oauth.py) and
a cookie jar, as a browser tab of that pod sends them - no browser, no Excel. The drill over the token (drill_token)
is built on it.

What it holds to, whatever a caller asks for:
- read-only towards EPM: a request that is not exactly on ALLOWED is refused before a token is asked for - its
  method, its path, the query names that path takes, the header names it takes and, for the two POST targets, the
  body (Smart View provider XML that svxml.check_read_only admits; the drill's getData query object). The same token
  would be accepted by EPM's write endpoints;
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
ANY = re.compile(r".*")
DIGITS = re.compile(r"\d{1,12}")
MILLIS = re.compile(r"\d{1,13}")                       # the cache buster: int(time.time() * 1000), 13 digits now


class _Rule(NamedTuple):
    method: str
    query: dict[str, tuple[int, int | None, re.Pattern]]   # name -> (least, most times (None: any), its value)
    headers: frozenset[str]                                  # the caller's header names it takes (lower case)
    content_type: str | None                                 # a POST's body type (the session sets it if missing)


_DRILL_HEADERS = frozenset({"accept", "content-type", "x-requested-with", "page_id", "x-appbuilder-client-id"})
# The brief's table, with what each request may carry besides: the query names are those the drill sends today
# (sui_url keeps Smart View's link minus DRILLURL; drill_direct's view and getData URLs), nothing else.
ALLOWED: dict[str, _Rule] = {
    SMART_VIEW: _Rule("POST", {}, frozenset({"content-type"}), "text/xml"),          # + svxml.check_read_only(body)
    "/aif/ui/model/applications/servicetype": _Rule("GET", {}, _DRILL_HEADERS, None),
    "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp": _Rule("GET", {
        "FORMAT": (1, 1, re.compile("SUI")), "module": (1, 1, re.compile(r"aif\.drilldown")),
        "attribute": (0, None, ANY), "sso_token": (0, 1, ANY), "_": (0, 1, MILLIS)}, _DRILL_HEADERS, None),
    "/aif/ui/model/view": _Rule("GET", {
        "page": (1, 1, re.compile("DRILL_THROUGH")), "appId": (0, 1, ANY), "partitionkey": (0, 1, ANY),
        "_": (0, 1, MILLIS)}, _DRILL_HEADERS, None),
    GET_DATA: _Rule("POST", {"limit": (1, 1, re.compile(r"\d{1,5}")), "offset": (1, 1, DIGITS)}, _DRILL_HEADERS,
                    "application/json"),
}
GET_DATA_KEYS = {"QUERY": str, "PARAMS": list, "PARAMSTYPE": list}   # drillcore.parse_sui / drill_derive build these
HEADER_VALUES = {"x-requested-with": re.compile("XMLHttpRequest")}
SECRET_HEADERS = ("x-appbuilder-client-id",)            # values scrubbed from an error text and a repr, like the token
HIDDEN_HEADERS = ("set-cookie", "set-cookie2")          # kept by the jar, never handed to a caller
MAX_BODY_BYTES = 1 << 20                                # a getData query or a provider request is far smaller
MAX_ANSWER_BYTES = 64 << 20                             # a 5000-line getData page is about 7 MB (1.3 kB a line)
MAX_TIMEOUT_S = 600.0
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
                timeout: float = 120.0) -> Response:
        """Send one allowlisted request with the bearer; a 401 takes one new token (reject=the refused one), drops
        the session's cookies (they came with the refused token) and is sent once more. Any status is an answer
        (a 3xx too: never followed). Refused before anything is sent; oauth.OAuthError / NotSignedIn as the token
        source raises them; OSError when the request left and no whole answer came (its text scrubbed of
        secrets); TooLarge for an answer over MAX_ANSWER_BYTES."""
        try:
            target, rule = self._target(method, url, body)
            sent = self._headers(headers, rule)
            if not isinstance(timeout, (int, float)) or isinstance(timeout, bool) or not 0 < timeout <= MAX_TIMEOUT_S:
                raise Refused(f"refused: a timeout must be a number of seconds over 0 and at most {MAX_TIMEOUT_S:g}")
            if self._tokens.scope_for(self.origin) is None:
                raise Refused(f"refused: no OAuth scope is configured for {self.origin or redact(self.base)!r}")
        except Refused as e:
            log.warning("epmhttp: %s", e)
            raise
        data = None if body is None else body.encode("utf-8")
        t0 = time.monotonic()
        token = self._tokens.access_token(self.origin)
        r = self._send(method, target, sent, data, token, timeout)
        if r.status == 401:                             # the hour may have ended early: one new token, one more try
            token = self._tokens.access_token(self.origin, reject=token)
            self._jar.clear()
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
            raise self._scrubbed(e, secrets) from None
        except http.client.HTTPException as e:          # IncompleteRead, BadStatusLine, LineTooLong ...
            raise self._scrubbed(RequestLost(f"{type(e).__name__}: {e}"), secrets) from None
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
    if len(body.encode("utf-8")) > MAX_BODY_BYTES:
        raise Refused(f"refused: POST {shown} with a body over {MAX_BODY_BYTES} bytes")
    if path == SMART_VIEW:
        try:
            svxml.check_read_only(body)
        except ValueError:                              # its text quotes the body's start (a session id): not here
            raise Refused("refused: not a single read-only Smart View request") from None
        return

    def no_twice(pairs):
        keys = [k for k, _ in pairs]
        if len(keys) != len(set(keys)):
            raise ValueError("a key twice")
        return dict(pairs)
    try:
        j = json.loads(body, object_pairs_hook=no_twice)
    except ValueError:
        raise Refused(f"refused: POST {shown} with a body that is not one JSON object") from None
    if not isinstance(j, dict) or any(k not in GET_DATA_KEYS or not isinstance(v, GET_DATA_KEYS[k])
                                      for k, v in j.items()):
        raise Refused(f"refused: POST {shown} with a body that is not a drill query ({', '.join(GET_DATA_KEYS)})")


def _read(r, deadline: float, method: str, url: str) -> str:
    """The answer's text, at most MAX_ANSWER_BYTES and no later than the deadline."""
    chunks, n = [], 0
    while chunk := r.read(1 << 20):
        n += len(chunk)
        if n > MAX_ANSWER_BYTES:
            raise TooLarge(f"EPM's answer to {method} {redact(url)} is over {MAX_ANSWER_BYTES >> 20} MB")
        if time.monotonic() > deadline:
            raise TimeoutError(f"EPM's answer to {method} {redact(url)} did not end in time")
        chunks.append(chunk)
    return b"".join(chunks).decode("utf-8", "replace")


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
