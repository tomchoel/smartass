"""Engine logging: one rotating file per process kind, with sso tokens, bearer tokens and the HTTP secret redacted."""
from __future__ import annotations

import logging
import re
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


def configure(path: Path, level: int = logging.INFO, max_bytes: int = 5_000_000, backups: int = 5) -> logging.Handler:
    """Attach one rotating handler to the root logger (idempotent per path)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    root = logging.getLogger()
    for h in root.handlers:
        if getattr(h, "_smartass_path", None) == str(path):
            return h
    logging.raiseExceptions = False                  # two clients' servers may rotate the same file; never crash
    h = RotatingFileHandler(path, maxBytes=max_bytes, backupCount=backups, encoding="utf-8", delay=True)
    h._smartass_path = str(path)
    h.setFormatter(_RedactingFormatter("%(asctime)s %(levelname)s %(name)s %(message)s"))
    root.addHandler(h)
    root.setLevel(level)
    return h
