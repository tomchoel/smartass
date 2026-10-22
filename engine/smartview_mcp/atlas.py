"""The cube atlas: every dimension tree of FinMax / FinLite / FinPlan in a local SQLite catalogue."""
from __future__ import annotations

import sqlite3
import time
from pathlib import Path

from . import config

SCHEMA = """
CREATE TABLE IF NOT EXISTS member (
  connection TEXT, dimension TEXT, name TEXT, label TEXT, parent TEXT, depth INT, ord INT,
  PRIMARY KEY (connection, dimension, name, parent)
);
CREATE INDEX IF NOT EXISTS ix_member_name ON member(name);
CREATE INDEX IF NOT EXISTS ix_member_parent ON member(connection, dimension, parent);
-- descendants() joins on (connection, parent) without dimension; without this index every recursion step scanned
-- the whole atlas (cost_base plan: 17.8 s)
CREATE INDEX IF NOT EXISTS ix_member_conn_parent ON member(connection, parent);
CREATE TABLE IF NOT EXISTS build (
  connection TEXT, dimension TEXT, built_at REAL, n INT, seconds REAL, error TEXT,
  PRIMARY KEY (connection, dimension)
);
"""


def connect(path: Path = config.ATLAS_DB) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(path)
    db.executescript(SCHEMA)
    return db


def store(db, conn: str, dim: str, members: list[dict], seconds: float, error: str | None = None):
    with db:
        db.execute("DELETE FROM member WHERE connection=? AND dimension=?", (conn, dim))
        db.executemany(
            "INSERT OR IGNORE INTO member VALUES (?,?,?,?,?,?,?)",
            [(conn, dim, m["name"], m["label"], m["parent"], m["depth"], m["order"]) for m in members])
        db.execute("INSERT OR REPLACE INTO build VALUES (?,?,?,?,?,?)",
                   (conn, dim, time.time(), len(members), seconds, error))


def purge(db, conn: str) -> None:
    with db:
        db.execute("DELETE FROM member WHERE connection=?", (conn,))
        db.execute("DELETE FROM build WHERE connection=?", (conn,))


def status(db) -> dict:
    rows = db.execute("SELECT connection, dimension, built_at, n, seconds, error FROM build ORDER BY 1, 2").fetchall()
    oldest = min((r[2] for r in rows), default=None)
    age_days = (time.time() - oldest) / 86400 if oldest else None
    return {
        "built": [{"connection": c, "dimension": d, "members": n, "seconds": round(s or 0, 1), "error": e,
                   "built_at": time.strftime("%Y-%m-%d %H:%M", time.localtime(b))} for c, d, b, n, s, e in rows],
        "oldest_age_days": round(age_days, 1) if age_days is not None else None,
        "stale": age_days is None or age_days > 30,
    }


def search(db, text: str, connection: str | None = None, dimension: str | None = None, limit: int = 25) -> list[dict]:
    """Case-insensitive search on member name and label; exact matches first."""
    q = f"%{text.lower()}%"
    sql = ("SELECT connection, dimension, name, label, group_concat(DISTINCT parent), min(depth) FROM member "
           "WHERE (lower(name) LIKE ? OR lower(label) LIKE ?)")
    args: list = [q, q]
    if connection:
        sql += " AND connection=?"
        args.append(connection)
    if dimension:
        sql += " AND dimension=?"
        args.append(dimension)
    sql += (" GROUP BY connection, dimension, name"
            " ORDER BY (lower(name)=?) DESC, (lower(label) LIKE ?) DESC, min(depth), length(name) LIMIT ?")
    args += [text.lower(), f"{text.lower()}%", limit]
    out = []
    for c, d, n, lab, parents, depth in db.execute(sql, args).fetchall():
        out.append({"connection": c, "dimension": d, "name": n, "label": lab, "depth": depth,
                    "parents": [p for p in (parents or "").split(",") if p]})
    return out


def children(db, connection: str, name: str, limit: int = 500) -> list[dict]:
    rows = db.execute("SELECT dimension, name, label, depth FROM member WHERE connection=? AND parent=? "
                      "ORDER BY ord LIMIT ?", (connection, name, limit)).fetchall()
    return [{"dimension": d, "name": n, "label": lab, "depth": dep} for d, n, lab, dep in rows]


def descendants(db, connection: str, name: str, leaves_only: bool = False, limit: int = 5000) -> list[dict]:
    sql = """
    WITH RECURSIVE t(name, label, depth, dimension) AS (
      SELECT name, label, depth, dimension FROM member WHERE connection=?1 AND parent=?2
      UNION
      SELECT m.name, m.label, m.depth, m.dimension FROM t CROSS JOIN member m ON m.connection=?1 AND m.parent=t.name
    ) SELECT DISTINCT name, label, depth, dimension FROM t LIMIT ?3"""
    # CROSS JOIN fixes the join order: drive from the rows found so far and look children up by the
    # (connection, parent) index. The planner otherwise scanned all members of the cube per step (11-45 s).
    rows = db.execute(sql, (connection, name, limit)).fetchall()
    out = [{"name": n, "label": lab, "depth": dep, "dimension": d} for n, lab, dep, d in rows]
    if leaves_only:
        parents = {r[0] for r in db.execute("SELECT DISTINCT parent FROM member WHERE connection=?", (connection,))}
        out = [o for o in out if o["name"] not in parents]
    return out


def ancestors(db, connection: str, name: str) -> list[str]:
    chain, cur, seen = [], name, set()
    while cur and cur not in seen:
        seen.add(cur)
        r = db.execute("SELECT parent FROM member WHERE connection=? AND name=? LIMIT 1", (connection, cur)).fetchone()
        cur = r[0] if r else None
        if cur:
            chain.append(cur)
    return chain
