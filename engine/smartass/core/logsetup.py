"""Engine logging: one rotating file per process kind, every line redacted (redact: session keys, tokens, cookies,
credential headers, URL query strings and the HTTP secret)."""
from __future__ import annotations

import logging
import os
import re
import time
from logging.handlers import RotatingFileHandler
from pathlib import Path

_TAG = r"(?:<|&(?:amp;)*lt;)"                               # '<', also HTML-escaped once or more
_END = r"(?:>|&(?:amp;)*gt;)"
_EQ = r"[ \t]*(?:=|%3D|%3A)[ \t]*"                             # name = value, URL-encoded name=value
_Q = r"(?:[\"']|&(?:amp;)*(?:quot|apos|#0?39|#34|#[xX]0?2[27]);)"   # a quote, also HTML-escaped (once or more)
_NAMES = (r"sID|JSESSIONID|ORA_[\w-]*|OAM[\w-]*|_WL_[\w-]*|X-Smartass-Secret|device_code|"
          r"[\w-]*(?:session|token|secret|password|passwd)[\w-]*")
# The secret keys by their whole names (no prefix: 'Oamaru: 3' is text): after one of them "name: value" is scrubbed
# whatever the value's length. Never a generic name: value rule: the scrub runs on texts whose sentences have colons -
# every log line, every tool error (core/serverkit), the run log, the provider's own error text (epm/svxml) and the two
# reply warnings that quote an engine's error (pipeline.prepare: the grids not fetched, the grids fetched over REST);
# the other reply warnings are the engine's own words and do not pass it. sID and session_id only in their own case
# ('SID: 4', 'Session_id: unknown' are text), the others in any; client secret also with a space or a hyphen.
# id_token is not here: 'id_token: missing' is text, and an ID token is a JWT (the JWT rule) of more than 16
# characters (the rule of the 16 key-like characters, which also takes 'session ID: <value>').
_KEYS = (r"sID|session_id|(?i:sessionId|JSESSIONID|ssoToken|sso_token|access_token|refresh_token|client[ _-]secret|"
         r"device_code|password|passwd|X-Smartass-Secret)")
_KEYLIKE = r"[A-Za-z0-9._~+/=%-]"                              # a key's characters, URL-encoded ones too
# A credential's shape: 16 key-like characters, or 12 with a digit, one of + / = % or a capital after a small letter
# (base64: 'dXNlcjpwYXNz'); 'Basic subscriptions' is text
_SHAPED = rf"(?:(?={_KEYLIKE}*(?:[0-9+/=%]|(?-i:[a-z][A-Z]))){_KEYLIKE}{{12}}|{_KEYLIKE}{{16}})"
_EQUOTE = r"&(?:amp;)*(?:quot|apos|#0?39|#34|#[xX]0?2[27]);"   # a quote HTML-escaped (once or more)
# A query after a space or at the start is scrubbed only when it names a secret key ('Use ?period=Sep' is text).
_QUERY_KEYS = rf"code|{_NAMES}"
_PATTERNS = [
    (re.compile(rf"({_TAG}\s*sID\s*{_END}).*?({_TAG}\s*/\s*sID\s*{_END})", re.S | re.I), r"\1***\2"),
    (re.compile(r"(?<=[\w/.%-])\?(?=[\w%])[^\s'\"<>]+"), "?..."),          # a URL's query string, whole (not 'label?)')
    (re.compile(rf"(?:^|(?<=\s))\?(?=(?:[^\s'\"<>&]*&)*(?:{_QUERY_KEYS})=)[^\s'\"<>]+", re.I),
     "?..."),                                                      # ... after a space or at the start: a secret key's only
    (re.compile(r"%3F(?=[\w.~-]+%3D)[^\s'\"<>]+", re.I), "%3F..."),           # a URL-encoded query (%3Fname%3D...)
    (re.compile(r"\b((?:proxy-)?authorization[\"']?[ \t]*[:=][ \t]*[\"']?)(?:(?:Bearer|Basic|Digest|Negotiate|NTLM)"
                rf"[ \t]+[^\s\"',;<>&}}]+|(?:[A-Za-z][\w-]*[ \t]+)?(?=[^\s\"',;<>&}}]*?{_SHAPED})"
                r"[^\s\"',;<>&}]+)", re.I),
     r"\1***"),                    # an Authorization header: a scheme's value, else a credential's shape ('approved')
    (re.compile(r"\b(Bearer(?:\s+|%20|\+))(?!\*\*\*)[A-Za-z0-9._~+/=-]{16,}", re.I),
     r"\1***"),                                                   # a token's length: not 'the cost bearer of'
    (re.compile(rf"\b(Basic\s+)(?!\*\*\*)(?={_SHAPED})[A-Za-z0-9._~+/=-]+"), r"\1***"),   # basic credentials
    (re.compile(r"\beyJ[\w-]*\.[\w-]+(?:\.[\w-]*)?"), "***"),                          # a JWT (header.payload.sig)
    (re.compile(r"\b((?:set-)?cookie[ \t]*:[ \t]*)[^\n<]*", re.I), r"\1***"),            # a cookie header, whole
    (re.compile(r"([\"'](?:set-)?cookie[\"'][ \t]*[:,][ \t]*[\"'])[^\"']+", re.I), r"\1***"),  # ... dict key, pair
    (re.compile(r"([\"'](?:set-)?cookie[\"'][ \t]*:[ \t]*\[)[^\]]*", re.I), r"\1***"),     # ... a list of values
    (re.compile(r"\b(x-[\w-]+[ \t]*:[ \t]*)(?!\*\*\*)(?=[A-Za-z0-9._~+/=-]{12})[^\s,;<>\"']+", re.I),
     r"\1***"),                                                         # an X-... header's key-like value
    # a session key, token, secret or cookie: name=value (any value), "name": "value" (JSON), name: value (a value of
    # 16 key-like characters or more - not the words after a class name in a logged error, 'NoToken: no EPM ...')
    (re.compile(rf"({_Q}?(?<![\w-])(?:{_NAMES}){_Q}?{_EQ}{_Q}?)(?!\*\*\*)[^\s;<>\"'&,}}]+", re.I), r"\1***"),
    (re.compile(rf"([\"'](?:{_NAMES})[\"'][ \t]*:[ \t]*[\"'])(?!\*\*\*)[^\"']+", re.I), r"\1***"),
    (re.compile(rf"({_EQUOTE}(?:{_NAMES}){_EQUOTE}[ \t]*:[ \t]*{_EQUOTE})(?!\*\*\*)[^&\"'<>]+", re.I),
     r"\1***"),                                                   # ... the same pair with its quotes HTML-escaped
    (re.compile(rf"((?<![\w-])(?:{_NAMES}|session[ \t]+id)[ \t]*:[ \t]*{_Q}?)(?!\*\*\*)(?={_KEYLIKE}{{16}})"
                r"[^\s;<>\"'&,}]+", re.I), r"\1***"),
    (re.compile(rf"((?<![\w-])(?:{_KEYS})[ \t]*:[ \t]*{_Q}?)(?!\*\*\*)[^\s;<>\"'&,}}]+"), r"\1***"),
]
_secrets: list[str] = []


def add_secret(value: str | None) -> None:
    if value and len(value) >= 8 and value not in _secrets:
        _secrets.append(value)


def redact(text) -> str:
    """text as it may go into a log line, a tool reply, the run log or an exception text: every shape of a session
    key, token, cookie, credential header or URL query string replaced (_PATTERNS), and the secrets this process
    holds (add_secret). The one scrub for all of them (also svxml's: the provider's own error text)."""
    out = str(text)
    for p, repl in _PATTERNS:
        out = p.sub(repl, out)
    for s in _secrets:
        out = out.replace(s, "***")
    return out


class _RedactingFormatter(logging.Formatter):
    def format(self, record):
        return redact(super().format(record))       # covers tracebacks too


class _SharedRotatingFileHandler(RotatingFileHandler):
    """RotatingFileHandler for a file several processes append to (one server per MCP client).

    On Windows the stock rollover shifts/deletes backups first and then fails to rename the live file (another
    process has it open), so the record is lost, and every later record fails and rotates the backups away. Here the
    live file is moved aside first; backups are only shifted once that worked. If it is in use, the record is
    appended anyway and rollover is paused for a while."""

    RETRY_AFTER = 600.0                              # s to wait before trying to rotate again after a failed rename

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._paused_until = 0.0

    def shouldRollover(self, record):
        if time.monotonic() < self._paused_until:
            return False
        return super().shouldRollover(record)

    def doRollover(self):
        if self.stream:
            self.stream.close()
            self.stream = None
        aside = self.baseFilename + ".rotating"
        try:
            if self.backupCount <= 0:
                raise OSError("no backups configured")
            os.rename(self.baseFilename, aside)
        except OSError:
            self._paused_until = time.monotonic() + self.RETRY_AFTER
        else:
            for i in range(self.backupCount - 1, 0, -1):
                src, dst = f"{self.baseFilename}.{i}", f"{self.baseFilename}.{i + 1}"
                if os.path.exists(src):
                    try:
                        os.replace(src, dst)
                    except OSError:
                        pass
            try:
                os.replace(aside, self.baseFilename + ".1")
            except OSError:
                pass
        if not self.delay:
            self.stream = self._open()               # with delay=True emit() reopens in append mode


def configure(path: Path, level: int = logging.INFO, max_bytes: int = 5_000_000, backups: int = 5) -> logging.Handler:
    """Attach one rotating handler to the root logger (idempotent per path)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    root = logging.getLogger()
    for h in root.handlers:
        if getattr(h, "_smartass_path", None) == str(path):
            return h
    logging.raiseExceptions = False                  # two clients' servers may rotate the same file; never crash
    h = _SharedRotatingFileHandler(path, maxBytes=max_bytes, backupCount=backups, encoding="utf-8", delay=True)
    h._smartass_path = str(path)
    h.setFormatter(_RedactingFormatter("%(asctime)s %(levelname)s %(name)s %(message)s"))
    root.addHandler(h)
    root.setLevel(level)
    return h
