"""Ledger lines of CLOSED periods kept on disk: a closed month (and a past year) never changes in the ledger, so a
drill of it is read back from svkit\\drill_cache instead of EPM (measured 02.11.2026: ~110 s for un.SY Yr 25).

Safety: only a period that ends on or before the closed month is kept; an entry lives DRILL_CACHE_DAYS; and the
caller re-checks the cell's cube value in Smart View before using it (a different value -> drill again). Only drills
that reconcile (gap 0) are saved. The sheet's source line says the lines came from the cache."""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import logging
import re
import time

from .. import config

log = logging.getLogger("smartview")
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _last_month(period: str) -> int | None:
    """'Yr 25' / 'Q3 26' / 'Oct 26' -> an index of the period's last month; None for anything else (YTD, totals)."""
    p = str(period or "").strip()
    m = re.fullmatch(r"Yr (\d\d)", p)
    if m:
        return int(m[1]) * 12 + 12
    m = re.fullmatch(r"Q([1-4]) (\d\d)", p)
    if m:
        return int(m[2]) * 12 + int(m[1]) * 3
    m = re.fullmatch(r"([A-Z][a-z]{2}) (\d\d)", p)
    if m and m[1] in MONTHS:
        return int(m[2]) * 12 + MONTHS.index(m[1]) + 1
    return None


def closed_period(cell: dict, closed: str | None) -> bool:
    last, c = _last_month(cell.get("All Time Periods", "")), _last_month(closed or "")
    return last is not None and c is not None and last <= c


def _path(connection: str, cell: dict):
    key = json.dumps([connection, sorted(cell.items())], ensure_ascii=False)
    return config.SVKIT / "drill_cache" / (hashlib.sha1(key.encode("utf-8")).hexdigest() + ".json")


def load(connection: str, cell: dict, closed: str | None) -> dict | None:
    """{pieces, parent, saved} of a closed cell drilled before (within DRILL_CACHE_DAYS), else None."""
    if not closed_period(cell, closed):
        return None
    f = _path(connection, cell)
    try:
        if time.time() - f.stat().st_mtime >= config.DRILL_CACHE_DAYS * 86400:
            return None
        hit = json.loads(f.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return hit if hit.get("cell") == cell and hit.get("connection") == connection else None


def save(connection: str, cell: dict, closed: str | None, pieces: list[dict], parent: float | None) -> None:
    if not closed_period(cell, closed) or parent is None:
        return
    f = _path(connection, cell)
    try:
        f.parent.mkdir(parents=True, exist_ok=True)
        tmp = f.with_suffix(".tmp")
        tmp.write_text(json.dumps({"connection": connection, "cell": cell, "parent": parent, "pieces": pieces,
                                   "saved": dt.date.today().strftime("%d.%m.%Y")}, ensure_ascii=False),
                       encoding="utf-8")
        tmp.replace(f)
    except (OSError, TypeError, ValueError) as e:
        log.info("drill cache not saved: %s", e)


def note(saved: str) -> str:
    return f" Cached {saved} (closed period; cube value re-checked)."
