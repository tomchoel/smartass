"""The token route's one way to EPM: HTTPS requests to one pod with the user's OAuth access token (epm/oauth.py) and
a cookie jar, as a browser tab of that pod sends them - no browser, no Excel. The drill over the token (drill_token)
is built on it.

What it holds to, whatever a caller asks for:
- read-only towards EPM: a request that is not exactly on ALLOWED is refused before a token is asked for. The only
  POST targets are the Smart View provider (bodies svxml.check_read_only admits) and the drill's getData - the same
  token would be accepted by EPM's write endpoints;
- the bearer goes only to the session's own pod, and only when config.OAUTH_SCOPES has a scope for it: the URL that
  is sent is built from the pod's origin and the allowed path, never from the host a caller wrote; a redirect is
  never followed (a 3xx is the answer, for the caller to judge);
- a token, the x-appbuilder-client-id key, a cookie or a URL query is never in a log line or an exception text
  (drillcore.redact on every URL shown)."""
from __future__ import annotations

import http.cookiejar
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
# path -> (the one method it is sent with, a query item it must have exactly once: (name, value) or None)
ALLOWED: dict[str, tuple[str, tuple[str, str] | None]] = {
    SMART_VIEW: ("POST", None),                                                   # + svxml.check_read_only(body)
    "/aif/ui/model/applications/servicetype": ("GET", None),
    "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp": ("GET", ("FORMAT", "SUI")),
    "/aif/ui/model/view": ("GET", ("page", "DRILL_THROUGH")),
    "/aif/ui/model/drilldown/getData": ("POST", None),
}
OWN_HEADERS = ("authorization", "cookie", "host")       # the session sets these; a caller's would bypass it
SECRET_HEADERS = ("x-appbuilder-client-id",)            # values scrubbed from an error text, like the token
_URL = re.compile(r"[\x21-\x7e]+")                      # printable ASCII: urlsplit drops tab and line breaks silently
_NAME = re.compile(r"[!#$%&'*+.^_`|~0-9A-Za-z-]+")      # an HTTP header name (RFC 9110 token)
_VALUE = re.compile(r"[\x20-\x7e]*")                    # no line break: http.client would quote the value in its error


class Response(NamedTuple):
    status: int
    text: str
    headers: dict[str, str]                             # names lower-cased


class Refused(ValueError):
    """The request is not on the allowlist or leaves the pod: nothing was sent and no token was asked for."""


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
        """Send one allowlisted request with the bearer; a 401 takes one new token (reject=the refused one) and is
        sent once more. Any status is an answer (a 3xx too: never followed). Refused before anything is sent;
        oauth.OAuthError / NotSignedIn as the token source raises them; OSError from the network as it is (a text
        that would quote a secret is scrubbed first)."""
        try:
            target = self._target(method, url, body)
            sent = self._headers(headers)
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
            r = self._send(method, target, sent, data, token, timeout)
        log.debug("epmhttp: %s %s -> HTTP %s (%.2f s)", method, redact(target), r.status, time.monotonic() - t0)
        return r

    def _target(self, method: str, url: str, body: str | None) -> str:
        """The URL to send - the pod's origin, the allowed path and the caller's query - or Refused."""
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
        if rule is None or rule[0] != method:
            raise Refused(f"refused: {method!r} {shown} is not on the allowlist")
        if rule[1]:
            name, value = rule[1]
            got = [v for k, _, v in (p.partition("=") for p in parts.query.split("&")) if unquote_plus(k) == name]
            if got != [value]:
                raise Refused(f"refused: {method} {shown} needs {name}={value} once in its query")
        if method == "GET" and body is not None:
            raise Refused(f"refused: GET {shown} with a body")
        if method == "POST" and not isinstance(body, str):
            raise Refused(f"refused: POST {shown} without a body")
        if parts.path == SMART_VIEW:
            try:
                svxml.check_read_only(body)
            except ValueError as e:
                raise Refused(str(e)) from None
        return self.origin + parts.path + (f"?{parts.query}" if parts.query else "")

    @staticmethod
    def _headers(headers: dict | None) -> dict[str, str]:
        out = dict(headers or {})
        for name, value in out.items():
            if not isinstance(name, str) or not _NAME.fullmatch(name):
                raise Refused("refused: a header name that is not a token")
            if name.lower() in OWN_HEADERS:
                raise Refused(f"refused: the {name} header is the session's own")
            _check_value(name, value)
        return out

    def _send(self, method: str, url: str, headers: dict[str, str], data: bytes | None, token: str,
              timeout: float) -> Response:
        bearer = "Bearer " + str(token)
        _check_value("Authorization", bearer)
        req = urllib.request.Request(url, data=data, headers={**headers, "Authorization": bearer}, method=method)
        self._jar.add_cookie_header(req)
        try:
            try:
                r = self._opener.open(req, timeout=timeout)
                status = r.status
            except urllib.error.HTTPError as e:         # outside 2xx, a 3xx too (_NoRedirect): an answer
                r, status = e, e.code
            with r:
                self._jar.extract_cookies(r, req)
                text = r.read().decode("utf-8", "replace")
                info = r.headers
        except OSError as e:
            raise self._scrubbed(e, [token, *(v for k, v in headers.items() if k.lower() in SECRET_HEADERS)]) from None
        return Response(status, text, {k.lower(): v for k, v in info.items()})

    def _scrubbed(self, e: OSError, secrets: list[str]) -> OSError:
        """e itself when its text quotes nothing secret; else the same kind of error with the token, the key, the
        cookies and any URL query cut out."""
        secrets = [s for s in (*secrets, *(c.value for c in self._jar)) if s]

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
