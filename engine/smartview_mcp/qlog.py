"""Append-only query log (JSONL) — an audit trail of what was pulled from EPM, when."""
from __future__ import annotations

import getpass
import json
import time

from . import config


def write(entry: dict) -> None:
    rec = {"ts": time.strftime("%Y-%m-%dT%H:%M:%S"), "user": getpass.getuser(), **entry}
    try:
        config.QUERY_LOG.parent.mkdir(parents=True, exist_ok=True)
        with open(config.QUERY_LOG, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")
    except OSError:
        pass   # logging must never break a query (OneDrive lock etc.)
