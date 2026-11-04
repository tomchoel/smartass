"""Local cache of drilled lines (config.HOME\\drill.db, normally %LOCALAPPDATA%\\smartass, outside OneDrive). Follow-up
reports and key points read the lines from here; drills older than DRILL_KEEP_DAYS are pruned."""
from __future__ import annotations

import datetime as dt
import json
import sqlite3
import uuid
from pathlib import Path

from smartass.core import config

_SCHEMA = """
CREATE TABLE IF NOT EXISTS drill (id TEXT PRIMARY KEY, created TEXT, connection TEXT, cell TEXT, rec TEXT,
                                  columns TEXT, workbook TEXT, q INT, alias TEXT);
CREATE TABLE IF NOT EXISTS line (drill_id TEXT, n INT, data TEXT, PRIMARY KEY (drill_id, n));
CREATE INDEX IF NOT EXISTS drill_wq ON drill (workbook, q);
"""


def connect(path: Path = config.DRILL_DB) -> sqlite3.Connection:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(path)
    db.executescript(_SCHEMA)
    return db


def _enc(v):
    return {"__dt": v.isoformat()} if isinstance(v, dt.datetime) else v


def _dec(d: dict) -> dict:
    return {k: dt.datetime.fromisoformat(v["__dt"]) if isinstance(v, dict) and "__dt" in v else v
            for k, v in d.items()}


def save(db, *, connection: str, cell: dict, rec: dict, columns: list[str], rows: list[dict], workbook: str | None,
         q: int | None, alias: str, created: dt.datetime | None = None) -> str:
    did = uuid.uuid4().hex[:12]
    with db:
        if workbook and q is not None:            # a re-drill of the same question/alias replaces the old one
            for (old,) in db.execute("SELECT id FROM drill WHERE workbook=? AND q=? AND alias=?",
                                     (workbook.lower(), q, alias)).fetchall():
                db.execute("DELETE FROM line WHERE drill_id=?", (old,))
                db.execute("DELETE FROM drill WHERE id=?", (old,))
        db.execute("INSERT INTO drill VALUES (?,?,?,?,?,?,?,?,?)",
                   (did, (created or dt.datetime.now()).isoformat(timespec="seconds"), connection,
                    json.dumps(cell), json.dumps(rec, default=str), json.dumps(columns),
                    (workbook or "").lower(), q, alias))
        db.executemany("INSERT INTO line VALUES (?,?,?)",
                       ((did, i, json.dumps({k: _enc(v) for k, v in r.items()}, ensure_ascii=False))
                        for i, r in enumerate(rows)))
    return did


def load(db, drill_id: str) -> tuple[dict, list[dict]]:
    m = db.execute("SELECT created, connection, cell, rec, columns, workbook, q, alias FROM drill WHERE id=?",
                   (drill_id,)).fetchone()
    if m is None:
        raise KeyError(drill_id)
    meta = {"id": drill_id, "created": m[0], "connection": m[1], "cell": json.loads(m[2]), "rec": json.loads(m[3]),
            "columns": json.loads(m[4]), "workbook": m[5], "q": m[6], "alias": m[7]}
    rows = [_dec(json.loads(d)) for (d,) in db.execute("SELECT data FROM line WHERE drill_id=? ORDER BY n",
                                                        (drill_id,))]
    return meta, rows


def find(db, workbook: str, q: int) -> dict[str, str]:
    """alias -> drill id for one question of one workbook."""
    return {a: i for i, a in db.execute("SELECT id, alias FROM drill WHERE workbook=? AND q=?",
                                        (workbook.lower(), q))}


def prune(db, keep_days: int = config.DRILL_KEEP_DAYS) -> int:
    cut = (dt.datetime.now() - dt.timedelta(days=keep_days)).isoformat(timespec="seconds")
    old = [i for (i,) in db.execute("SELECT id FROM drill WHERE created < ?", (cut,))]
    with db:
        for i in old:
            db.execute("DELETE FROM line WHERE drill_id=?", (i,))
            db.execute("DELETE FROM drill WHERE id=?", (i,))
    return len(old)


def meta(db, drill_id: str) -> dict | None:
    """The drill's metadata without its lines (cheap, also for 300 000-line drills)."""
    m = db.execute("SELECT created, connection, cell, rec, columns, workbook, q, alias FROM drill WHERE id=?",
                   (drill_id,)).fetchone()
    if m is None:
        return None
    return {"id": drill_id, "created": m[0], "connection": m[1], "cell": json.loads(m[2]), "rec": json.loads(m[3]),
            "columns": json.loads(m[4]), "workbook": m[5], "q": m[6], "alias": m[7]}
