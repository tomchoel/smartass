"""The local warehouse (Stage 2): cube cells in DuckDB at config.WAREHOUSE_DB, normally
%LOCALAPPDATA%\\smartass\\warehouse.duckdb (outside OneDrive, never in git). Storage only - no business logic.

- fact: one row per cube cell (cube, scenario, version, entity, account, period, balance); an upsert replaces it.
  FinMax rows have version ''. FinPlan periods are stored like FinMax's ('Sep 26' = Period Sep of Years FY26).
  closed_month is the closed month when the row was read, or 'open' for a row read live at question time (the open
  month, the latest forecast): such rows are never taken for closed history.
- fill: one row per extract job (what, when, cells, seconds, result, detail) - freshness and the doctor read it.
- state: small key/value facts of the refresh logic (last closed month filled, atlas snapshot, forecast name).

One writer per machine: writes hold an exclusive lock file (drill_browser.ProfileLock, reused here as a plain file
lock, freed when the process dies); put() writes in chunks of PUT_CHUNK rows - atomic per chunk, not per job; every call opens and closes its own DuckDB connection, so other processes can read in between. A missing or locked
warehouse raises WarehouseUnavailable; callers then answer from live pulls and say so. MemoryWarehouse has the
same API in plain Python (tests, and the live-pull fallback)."""
from __future__ import annotations

import json
import threading
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

from . import config

WhKey = tuple[str, str, str, str, str, str, str]      # cube, scenario, version, entity, account, period, balance
PUT_CHUNK = 10_000                                     # rows per write: readers in other processes get in between
TYPES = ("VARCHAR",) * 7 + ("DOUBLE", "VARCHAR", "DOUBLE", "VARCHAR", "VARCHAR")
COLUMNS = ("cube", "scenario", "version", "entity", "account", "period", "balance", "value", "source", "fetched_at",
           "closed_month", "atlas_build")
SCHEMA = """
CREATE TABLE IF NOT EXISTS fact (
  cube VARCHAR, scenario VARCHAR, version VARCHAR, entity VARCHAR, account VARCHAR, period VARCHAR, balance VARCHAR,
  value DOUBLE, source VARCHAR, fetched_at DOUBLE, closed_month VARCHAR NOT NULL, atlas_build VARCHAR,
  PRIMARY KEY (cube, scenario, version, entity, account, period, balance));
CREATE SEQUENCE IF NOT EXISTS fill_id;
CREATE TABLE IF NOT EXISTS fill (
  id BIGINT DEFAULT nextval('fill_id') PRIMARY KEY, job VARCHAR, what VARCHAR, started_at DOUBLE, seconds DOUBLE,
  cells BIGINT, result VARCHAR, detail VARCHAR);
CREATE TABLE IF NOT EXISTS state (key VARCHAR PRIMARY KEY, value VARCHAR);
"""
OPEN = "open"                                          # closed_month of a row read live at question time


class WarehouseUnavailable(Exception):
    """The warehouse file is missing or another process holds it; answer from live pulls instead."""


@dataclass(frozen=True)
class Row:
    key: WhKey
    value: float | None
    source: str                 # 'http' | 'com'
    fetched_at: float
    closed_month: str           # the closed month at read time, or OPEN
    atlas_build: str = ""       # the atlas snapshot the members were checked against

    def record(self) -> tuple:
        return (*self.key, self.value, self.source, self.fetched_at, self.closed_month, self.atlas_build)


@dataclass(frozen=True)
class Fill:
    job: str                    # base | new_month | restatement | forecast | top_up
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


class MemoryWarehouse:
    """The warehouse API over dicts: for tests and for answering from live pulls when the file is unavailable."""

    def __init__(self):
        self.rows: dict[WhKey, Row] = {}
        self.fills: list[Fill] = []
        self._state: dict[str, str] = {}

    def put(self, rows: Iterable[Row]) -> int:
        n = 0
        for r in rows:
            n += 1
            old = self.rows.get(r.key)
            if r.closed_month == OPEN and old is not None and old.closed_month != OPEN:
                continue                                   # a live read never replaces closed history
            self.rows[r.key] = r
        return n

    def get(self, cube: str, entities, accounts, periods, scenario: str, version: str, balance: str,
            closed_only: bool = False) -> dict[WhKey, float | None]:
        ents, accs, pers = set(entities), set(accounts), set(periods)
        return {k: r.value for k, r in self.rows.items()
                if _match(k, cube, ents, accs, pers, scenario, version, balance)
                and not (closed_only and r.closed_month == OPEN)}

    def freshness(self, cube: str | None = None, entities=None) -> dict:
        rows = [r for k, r in self.rows.items() if (cube is None or k[0] == cube)
                and (entities is None or k[3] in set(entities))]
        return _freshness(len(rows), [r.fetched_at for r in rows], sum(r.closed_month == OPEN for r in rows),
                          {r.closed_month for r in rows if r.closed_month != OPEN})

    def record_fill(self, fill: Fill) -> None:
        self.fills.append(fill)

    def last_fills(self, n: int = 10) -> list[dict]:
        return [f.as_dict() for f in reversed(self.fills[-n:])]

    def state(self) -> dict[str, str]:
        return dict(self._state)

    def set_state(self, **values: str) -> None:
        self._state.update({k: str(v) for k, v in values.items()})

    def status(self) -> dict:
        return {"path": None, "exists": True, "rows": len(self.rows), "fills": self.last_fills(5),
                "state": self.state(), **self.freshness()}


def _freshness(n: int, stamps: list[float], open_rows: int, closed: set[str]) -> dict:
    return {"rows": n, "oldest": min(stamps) if stamps else None, "newest": max(stamps) if stamps else None,
            "open_rows": open_rows, "closed_months": sorted(closed)}


def _lock_conflict(e: Exception) -> bool:
    """Another connection or process holds the file (DuckDB 2.0, Windows: 'being used by another process' /
    'File is already open in'; in-process: ConnectionException). Anything else is a bad file."""
    import duckdb
    if isinstance(e, duckdb.ConnectionException):
        return True
    m = str(e).lower()
    return "used by another process" in m or "already open" in m or "conflicting lock" in m


_mu = threading.RLock()                                # one DuckDB connection at a time in this process


class Warehouse:
    """The DuckDB file. Each method opens, works and closes; writes hold the machine-wide lock file."""

    def __init__(self, path: Path | None = None):
        from .drill_browser import ProfileLock
        self.path = Path(path or config.WAREHOUSE_DB)
        self._lock = ProfileLock(self.path)            # <path>.lock; used as a plain file lock

    def _connect(self, write: bool, deadline: float):
        import duckdb
        if not write and not self.path.exists():
            raise WarehouseUnavailable(f"no warehouse yet ({self.path})")
        while True:
            try:
                if write:
                    self.path.parent.mkdir(parents=True, exist_ok=True)
                    db = duckdb.connect(str(self.path))
                    db.execute(SCHEMA)
                    return db
                return duckdb.connect(str(self.path), read_only=True)
            except (duckdb.IOException, duckdb.ConnectionException) as e:
                first = str(e).splitlines()[0] if str(e) else type(e).__name__
                if not _lock_conflict(e):                  # a 0-byte or foreign file: retrying cannot help
                    raise WarehouseUnavailable(f"cannot open {self.path}: {first}") from None
                if time.monotonic() > deadline:
                    raise WarehouseUnavailable(f"warehouse {self.path} is in use: {first}") from None
                time.sleep(0.1)

    def _write(self, fn):
        with _mu:
            deadline = time.monotonic() + config.WAREHOUSE_LOCK_WAIT_S     # one budget for lock file and connect
            while not self._lock.acquire():
                if time.monotonic() > deadline:
                    raise WarehouseUnavailable(f"another smartass process is writing {self.path}")
                time.sleep(0.1)
            try:
                db = self._connect(True, deadline)
                try:
                    return fn(db)
                finally:
                    db.close()
            finally:
                self._lock.release()

    def _read(self, fn):
        with _mu:
            db = self._connect(False, time.monotonic() + config.WAREHOUSE_LOCK_WAIT_S)
            try:
                return fn(db)
            except Exception as e:                       # a damaged or foreign file: say so, never a silent gap
                raise WarehouseUnavailable(f"cannot read {self.path}: {e}") from None
            finally:
                db.close()

    def put(self, rows: Iterable[Row]) -> int:
        """Upsert; a key repeated in `rows` keeps its last row; an OPEN row never replaces a closed one. Returns the number of distinct keys written."""
        recs = [r.record() for r in {r.key: r for r in rows}.values()]
        sel = ", ".join(f"unnest(${i + 1}::{t}[]) AS {c}" for i, (t, c) in enumerate(zip(TYPES, COLUMNS)))
        pk = " AND ".join(f"f.{c} = n.{c}" for c in COLUMNS[:7])
        # one statement: ~0.3 s/6500 rows; an OPEN row never replaces an existing closed one
        sql = (f"INSERT OR REPLACE INTO fact SELECT * FROM (SELECT {sel}) n WHERE NOT (n.closed_month = '{OPEN}' "
               f"AND EXISTS (SELECT 1 FROM fact f WHERE {pk} AND f.closed_month <> '{OPEN}'))")
        for i in range(0, len(recs), PUT_CHUNK):
            cols = [list(c) for c in zip(*recs[i:i + PUT_CHUNK])]
            self._write(lambda db, cols=cols: db.execute(sql, cols))
        return len(recs)

    def get(self, cube: str, entities, accounts, periods, scenario: str, version: str, balance: str,
            closed_only: bool = False) -> dict[WhKey, float | None]:
        sql = ("SELECT cube, scenario, version, entity, account, period, balance, value FROM fact WHERE cube=? AND "
               "scenario=? AND version=? AND balance=? AND list_contains(?, entity) AND list_contains(?, account) "
               "AND list_contains(?, period)" + (" AND closed_month <> ?" if closed_only else ""))
        args = [cube, scenario, version, balance, list(entities), list(accounts), list(periods)]
        if closed_only:
            args.append(OPEN)
        return self._read(lambda db: {tuple(r[:7]): r[7] for r in db.execute(sql, args).fetchall()})

    def freshness(self, cube: str | None = None, entities=None) -> dict:
        where, args = ["1=1"], []
        if cube is not None:
            where.append("cube=?")
            args.append(cube)
        if entities is not None:
            where.append("list_contains(?, entity)")
            args.append(list(entities))
        sql = (f"SELECT count(*), min(fetched_at), max(fetched_at), count(*) FILTER (WHERE closed_month=?), "
               f"list(DISTINCT closed_month) FILTER (WHERE closed_month<>?) FROM fact WHERE {' AND '.join(where)}")
        args = [OPEN, OPEN, *args]
        n, lo, hi, op, closed = self._read(lambda db: db.execute(sql, args).fetchone())
        return {"rows": n, "oldest": lo, "newest": hi, "open_rows": op or 0, "closed_months": sorted(closed or [])}

    def record_fill(self, fill: Fill) -> None:
        self._write(lambda db: db.execute(
            "INSERT INTO fill (job, what, started_at, seconds, cells, result, detail) VALUES (?,?,?,?,?,?,?)",
            [fill.job, fill.what, fill.started_at, fill.seconds, fill.cells, fill.result,
             json.dumps(fill.detail, ensure_ascii=False)]))

    def last_fills(self, n: int = 10) -> list[dict]:
        rows = self._read(lambda db: db.execute(
            "SELECT job, what, started_at, seconds, cells, result, detail FROM fill ORDER BY id DESC LIMIT ?",
            [n]).fetchall())
        return [{"job": j, "what": w, "started_at": s, "seconds": sec, "cells": c, "result": r,
                 "detail": json.loads(d or "{}")} for j, w, s, sec, c, r, d in rows]

    def state(self) -> dict[str, str]:
        try:
            return dict(self._read(lambda db: db.execute("SELECT key, value FROM state").fetchall()))
        except WarehouseUnavailable:
            if self.path.exists():
                raise
            return {}                                   # no file yet: nothing filled

    def set_state(self, **values: str) -> None:
        self._write(lambda db: db.execute(
            "INSERT OR REPLACE INTO state SELECT unnest($1::VARCHAR[]), unnest($2::VARCHAR[])",
            [list(values), [str(v) for v in values.values()]]))

    def status(self) -> dict:
        if not self.path.exists():
            return {"path": str(self.path), "exists": False, "rows": 0, "fills": [], "state": {}}
        try:
            return {"path": str(self.path), "exists": True, "fills": self.last_fills(5), "state": self.state(),
                    **self.freshness()}
        except WarehouseUnavailable as e:               # feeds the doctor: never throws
            return {"path": str(self.path), "exists": True, "error": str(e), "rows": 0, "fills": [], "state": {}}
