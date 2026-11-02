"""In-memory facts for one pnl recipe call: rows read live by extract.fill, read back by fin_pnl.compute.

One row per cube cell (cube, scenario, version, entity, account, period, balance); a put replaces it. FinMax rows
have version ''. FinPlan periods are stored like FinMax's ('Sep 26' = Period Sep of Years FY26). closed_month is the
closed month when the row was read, or OPEN for a row read as the open month or the latest forecast: such rows are
never taken for closed history."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Iterable

WhKey = tuple[str, str, str, str, str, str, str]      # cube, scenario, version, entity, account, period, balance
_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
OPEN = "open"                                          # closed_month of a row read as the open month / latest forecast


def _month_index(label: str) -> int | None:
    """'Sep 26' -> 2026*12 + 9 (None for anything else, e.g. OPEN)."""
    try:
        return (2000 + int(label[4:6])) * 12 + _MONTHS.index(label[:3]) + 1
    except (ValueError, IndexError):
        return None


def protects(old_period: str, old_closed: str) -> bool:
    """A stored row is closed history - never replaced by an OPEN read - when it was read as closed and its period
    is on or before the month that was closed then. Plan months after it (forecast, target) stay replaceable."""
    if old_closed == OPEN:
        return False
    p, c = _month_index(old_period), _month_index(old_closed)
    return p is not None and c is not None and p <= c


@dataclass(frozen=True)
class Row:
    key: WhKey
    value: float | None
    source: str                 # 'http' | 'com'
    fetched_at: float
    closed_month: str           # the closed month at read time, or OPEN


@dataclass(frozen=True)
class Fill:
    job: str                    # live | top_up
    what: str
    started_at: float
    seconds: float
    cells: int
    result: str                 # ok | warn | error
    detail: dict

    def as_dict(self) -> dict:
        return asdict(self)


def _match(key: WhKey, cube, entities, accounts, periods, scenario, version, balance) -> bool:
    c, s, v, e, a, p, b = key
    return (c == cube and s == scenario and v == version and b == balance and e in entities and a in accounts
            and p in periods)


class FactStore:
    """Cube cells in dicts, plus the record of each fill."""

    def __init__(self):
        self.rows: dict[WhKey, Row] = {}
        self.fills: list[Fill] = []

    def put(self, rows: Iterable[Row]) -> int:
        written = set()
        for r in rows:
            old = self.rows.get(r.key)
            if r.closed_month == OPEN and old is not None and protects(old.key[5], old.closed_month):
                continue                                   # an open read never replaces closed history
            self.rows[r.key] = r
            written.add(r.key)
        return len(written)                                # distinct keys written

    def get(self, cube: str, entities, accounts, periods, scenario: str, version: str, balance: str,
            closed_only: bool = False) -> dict[WhKey, float | None]:
        ents, accs, pers = set(entities), set(accounts), set(periods)
        return {k: r.value for k, r in self.rows.items()
                if _match(k, cube, ents, accs, pers, scenario, version, balance)
                and not (closed_only and r.closed_month == OPEN)}

    def freshness(self, cube: str | None = None, entities=None) -> dict:
        rows = [r for k, r in self.rows.items() if (cube is None or k[0] == cube)
                and (entities is None or k[3] in set(entities))]
        stamps = [r.fetched_at for r in rows]
        return {"rows": len(rows), "oldest": min(stamps) if stamps else None, "newest": max(stamps) if stamps else None,
                "open_rows": sum(r.closed_month == OPEN for r in rows),
                "closed_months": sorted({r.closed_month for r in rows if r.closed_month != OPEN})}

    def record_fill(self, fill: Fill) -> None:
        self.fills.append(fill)
