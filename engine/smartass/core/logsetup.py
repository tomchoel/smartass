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
_SEP = r"[ \t]*(?:[=:]|%3D|%3A)[ \t]*"                         # name = value, name: value, URL-encoded '='
_NAMES = (r"sID|JSESSIONID|ORA_[\w-]*|OAM[\w-]*|_WL_[\w-]*|X-Smartass-Secret|"
          r"[\w-]*(?:session|token|secret|password|passwd)[\w-]*")
_PATTERNS = [
    (re.compile(rf"({_TAG}\s*sID\s*{_END}).*?({_TAG}\s*/\s*sID\s*{_END})", re.S | re.I), r"\1***\2"),
    (re.compile(r"(?<=[\w/.%-])\?[^\s'\"<>]+"), "?..."),                       # a URL's query string, whole
    (re.compile(r"\b((?:proxy-)?authorization[\"']?[ \t]*[:=][ \t]*[\"']?)(?:[A-Za-z]+[ \t]+)?[^\s\"',;<>&}]+",
                re.I), r"\1***"),                                                  # an Authorization header
    (re.compile(r"\b(Bearer\s+)(?!\*\*\*)[A-Za-z0-9._~+/=-]+", re.I), r"\1***"),
    (re.compile(r"\b(Basic\s+)(?!\*\*\*)[A-Za-z0-9._~+/=-]{8,}"), r"\1***"),        # basic credentials
    (re.compile(r"\beyJ[\w-]*\.[\w-]+(?:\.[\w-]*)?"), "***"),                          # a JWT (header.payload.sig)
    (re.compile(r"\b((?:set-)?cookie[ \t]*:[ \t]*)[^\n<]*", re.I), r"\1***"),            # a cookie header, whole
    (re.compile(r"\b(x-[\w-]+[ \t]*:[ \t]*)(?!\*\*\*)[^\s,;<>\"']+", re.I), r"\1***"),   # an X-... header's value
    (re.compile(rf"([\"']?(?<![\w-])(?:{_NAMES})[\"']?{_SEP}[\"']?)(?!\*\*\*)[^\s;<>\"'&,}}]+", re.I), r"\1***"),
]                                                       # a session key, token, secret or cookie pair (also JSON)
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
