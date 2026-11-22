"""The cube atlas: every dimension tree of FinMax / FinLite / FinPlan in a local SQLite catalogue."""
from __future__ import annotations

import sqlite3
import time
from pathlib import Path

from ..core import config

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
-- behaviour cards derived from data (knowledge/cards_fetch.py): one row per (member, parent) occurrence
CREATE TABLE IF NOT EXISTS card (
  connection TEXT, dimension TEXT, name TEXT, parent TEXT, operator TEXT, time_balance TEXT, account_type TEXT,
  shared INT, derived_at REAL, PRIMARY KEY (connection, dimension, name, parent)
);
"""


def connect(path: Path = config.ATLAS_DB) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(path)
    db.executescript(SCHEMA)
    return db


# What the build over REST adds to a member (epm/members.py EXTRA): its real name (the key stays the one derived from
# its label, as the zoom gave it), its Default alias, data storage, account type and time balance. Added to the
# member table by add_member_columns - only by a build, on its new file: reading the atlas never adds them, and the
# old readers name their columns, so they do not notice.
MEMBER_EXTRA = ("real_name", "alias", "storage", "account_type", "time_balance")
_BASE = ("connection", "dimension", "name", "label", "parent", "depth", "ord")


def add_member_columns(db) -> None:
    """The schema migration of the build over REST: the MEMBER_EXTRA columns, where the member table lacks them."""
    have = {r[1] for r in db.execute("PRAGMA table_info(member)")}
    with db:
        for c in MEMBER_EXTRA:
            if c not in have:
                db.execute(f"ALTER TABLE member ADD COLUMN {c} TEXT")


def store(db, conn: str, dim: str, members: list[dict], seconds: float, error: str | None = None):
    """Replace a dimension's members (the first of a (name, parent) given twice is kept) and record the build. A
    member's MEMBER_EXTRA values are stored where the table has those columns (add_member_columns)."""
    extra = [c for c in MEMBER_EXTRA if c in {r[1] for r in db.execute("PRAGMA table_info(member)")}]
    cols = (*_BASE, *extra)
    with db:
        db.execute("DELETE FROM member WHERE connection=? AND dimension=?", (conn, dim))
        db.executemany(
            f"INSERT OR IGNORE INTO member ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
            [(conn, dim, m["name"], m["label"], m["parent"], m["depth"], m["order"], *(m.get(c) for c in extra))
             for m in members])
        db.execute("INSERT OR REPLACE INTO build VALUES (?,?,?,?,?,?)",
                   (conn, dim, time.time(), len(members), seconds, error))


def store_cards(db, conn: str, dim: str, rows: list[tuple]) -> None:
    """rows: (name, parent, operator, time_balance, account_type, shared); replaces the dimension's cards. A row
    without a time balance keeps the one its (name, parent) had (the time balance is derived in a later step, which
    may fail). A purge of the members keeps the cards (the last good derivation stays until the next one succeeds)."""
    now = time.time()
    with db:
        old = dict(((n, p), tb) for n, p, tb in db.execute(
            "SELECT name, parent, time_balance FROM card WHERE connection=? AND dimension=?", (conn, dim)))
        db.execute("DELETE FROM card WHERE connection=? AND dimension=?", (conn, dim))
        db.executemany("INSERT OR REPLACE INTO card VALUES (?,?,?,?,?,?,?,?,?)",
                       [(conn, dim, n, p, op, tb if tb is not None else old.get((n, p)), at, sh, now)
                        for n, p, op, tb, at, sh in rows])


def cards_status(db) -> list[dict]:
    """Per connection and dimension: rows, operator counts, time balance counts, derivation time."""
    out = []
    for c, d, n, at in db.execute("SELECT connection, dimension, count(*), max(derived_at) FROM card "
                                  "GROUP BY 1, 2 ORDER BY 1, 2").fetchall():
        def counts(col: str) -> dict:
            return {(k or "unknown"): v for k, v in db.execute(
                f"SELECT {col}, count(*) FROM card WHERE connection=? AND dimension=? GROUP BY 1", (c, d))}
        out.append({"connection": c, "dimension": d, "rows": n, "operator": counts("operator"),
                    "time_balance": counts("time_balance"), "account_type": counts("account_type"),
                    "derived_at": time.strftime("%Y-%m-%d %H:%M", time.localtime(at))})
    return out


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
        "stale": age_days is None or age_days > MAX_AGE_DAYS,
    }


MAX_AGE_DAYS = 30


def usable(db, connection: str, dimension: str) -> bool:
    """The dimension was built without error, has members and is at most MAX_AGE_DAYS old."""
    r = db.execute("SELECT built_at, n, error FROM build WHERE connection=? AND dimension=?",
                   (connection, dimension)).fetchone()
    return bool(r and r[1] and not r[2] and (time.time() - r[0]) / 86400 <= MAX_AGE_DAYS)


def built(db, connection: str, dimension: str) -> bool:
    """The dimension was built without error and has members, whatever its age (dimension names do not go stale;
    the HTTP route needs only those)."""
    r = db.execute("SELECT n, error FROM build WHERE connection=? AND dimension=?", (connection, dimension)).fetchone()
    return bool(r and r[0] and r[1] is None)


def known_dimensions(db, connection: str) -> list[str]:
    """Every dimension the atlas knows for the connection, sorted: those with members AND those with only a build
    row (a failed rebuild deletes the members but records the build), so a failed dimension is not silently lost."""
    return [r[0] for r in db.execute("SELECT dimension FROM member WHERE connection=? UNION "
                                     "SELECT dimension FROM build WHERE connection=? ORDER BY 1",
                                     (connection, connection))]


def canonical(db, connection: str, dimension: str, name: str) -> str | None:
    """The atlas spelling of a member given by name (or, failing that, alias label), case-insensitive like Essbase;
    None if the dimension does not have it. The dimension's own name is its top member."""
    if name.casefold() == dimension.casefold():
        return dimension
    db.create_function("fold", 1, lambda s: s.casefold() if isinstance(s, str) else s, deterministic=True)
    sql = "SELECT name FROM member WHERE connection=? AND dimension=? AND fold({}) = ? LIMIT 1"   # æøå too
    for col in ("name", "label"):
        r = db.execute(sql.format(col), (connection, dimension, name.casefold())).fetchone()
        if r:
            return r[0]
    return None


def has_member(db, connection: str, dimension: str, name: str) -> bool:
    return canonical(db, connection, dimension, name) is not None


def search(db, text: str, connection: str | None = None, dimension: str | None = None, limit: int = 25) -> list[dict]:
    """Case-insensitive search on member name and label; exact matches first."""
    q = f"%{text.lower()}%"
    sql = ("SELECT connection, dimension, name, label, group_concat(DISTINCT parent), min(depth) FROM member "
           "WHERE (lower(name) LIKE ? OR lower(label) LIKE ?) "
           "AND coalesce(?, connection)=connection AND coalesce(?, dimension)=dimension")
    args: list = [q, q, connection or None, dimension or None]
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
