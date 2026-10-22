"""Engine logging: one rotating file per process kind, with sso tokens, bearer tokens and the HTTP secret redacted."""
from __future__ import annotations

import logging
import os
import re
import time
from logging.handlers import RotatingFileHandler
from pathlib import Path

_PATTERNS = [re.compile(r"(sso_token=)[^&\s'\"]+"),
             re.compile(r"(Bearer\s+)[A-Za-z0-9._~+/=-]+"),
             re.compile(r"((?:X-Smartass-Secret|secret)\s*[:=]\s*['\"]?)[^\s'\",}]+", re.I)]
_secrets: list[str] = []


def add_secret(value: str | None) -> None:
    if value and len(value) >= 8 and value not in _secrets:
        _secrets.append(value)


def redact(text) -> str:
    out = str(text)
    for p in _PATTERNS:
        out = p.sub(r"\1***", out)
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
