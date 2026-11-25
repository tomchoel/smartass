"""The cube atlas: every dimension tree of FinMax / FinLite / FinPlan in a local SQLite catalogue."""
from __future__ import annotations

import re
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


def empty() -> sqlite3.Connection:
    """An atlas with no members, in memory: what a home with no atlas reads as (no file is made)."""
    db = sqlite3.connect(":memory:")
    db.executescript(SCHEMA)
    return db


def connect_read() -> sqlite3.Connection:
    """The home's atlas for a read (fin_status, `smartass atlas status`): connect() when the file is there (its CREATE
    IF NOT EXISTS gives an older atlas the tables a status reads), else empty() - a status never makes an atlas.db on
    a fresh home."""
    path = Path(config.ATLAS_DB)                       # checked and opened: the same file
    return connect(path) if path.exists() else empty()


# What the build over REST knows of a member beyond today's seven columns: its real name (the key stays the one derived
# from its label, as the zoom gave it), its Default alias, data storage, account type and time balance. They live in
# a SIDE table, never in member: an installed older release (0.6.0) purges and rewrites member with a positional
# 7-value INSERT, which must keep working on an atlas this release built. The side table holds every real member
# behind a key - two members that give one key under one parent included (member keeps one row for them) - and the
# stamp of the build that wrote it: real_names trusts it only while the dimension's build row carries that stamp (an
# older release that rebuilt the members left it behind: then it is ignored). Only a store of real names creates it.
REAL_FIELDS = ("real_name", "alias", "storage", "account_type", "time_balance")
_REAL_SCHEMA = """CREATE TABLE IF NOT EXISTS member_real (
  connection TEXT, dimension TEXT, name TEXT, parent TEXT, real_name TEXT, alias TEXT, storage TEXT,
  account_type TEXT, time_balance TEXT, built_at REAL, PRIMARY KEY (connection, dimension, name, parent, real_name)
)"""


def _has_real(db) -> bool:
    return db.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='member_real'").fetchone() is not None


def store(db, conn: str, dim: str, members: list[dict], seconds: float, error: str | None = None):
    """Replace a dimension's members (the first of a (name, parent) given twice is kept) and record the build. The
    members' real names and properties (REAL_FIELDS, from the build over REST) replace the dimension's side rows,
    stamped with this build; a store without them only clears its side rows."""
    now = time.time()
    real = [m for m in members if m.get("real_name")]
    with db:
        db.execute("DELETE FROM member WHERE connection=? AND dimension=?", (conn, dim))
        db.executemany(
            "INSERT OR IGNORE INTO member (connection, dimension, name, label, parent, depth, ord) "
            "VALUES (?,?,?,?,?,?,?)",
            [(conn, dim, m["name"], m["label"], m["parent"], m["depth"], m["order"]) for m in members])
        if real:
            db.execute(_REAL_SCHEMA)
        if real or _has_real(db):
            db.execute("DELETE FROM member_real WHERE connection=? AND dimension=?", (conn, dim))
            db.executemany("INSERT OR IGNORE INTO member_real VALUES (?,?,?,?,?,?,?,?,?,?)",
                           [(conn, dim, m["name"], m["parent"], *(m.get(f) for f in REAL_FIELDS), now) for m in real])
        db.execute("INSERT OR REPLACE INTO build VALUES (?,?,?,?,?,?)", (conn, dim, now, len(members), seconds, error))


def real_names(db, conn: str, dim: str, names: set[str]) -> dict[str, set[str]]:
    """{key: EPM's real names behind it} for these keys of one dimension, from a side table whose stamp is the
    dimension's build's; {} for a key it does not hold, and for all when there is none or it is stale."""
    if not names or not _has_real(db):
        return {}
    out: dict[str, set[str]] = {}
    ordered = sorted(names)
    for i in range(0, len(ordered), 500):              # SQLite's limit on query parameters
        part = ordered[i:i + 500]
        for n, rn in db.execute(
                "SELECT r.name, r.real_name FROM member_real r JOIN build b ON b.connection = r.connection AND "
                "b.dimension = r.dimension AND b.built_at = r.built_at WHERE r.connection=? AND r.dimension=? AND "
                f"r.name IN ({','.join('?' * len(part))})", (conn, dim, *part)):
            out.setdefault(n, set()).add(rn)
    return out


def real_owners(db, conn: str, dim: str) -> dict[str, set[str]] | None:
    """{casefolded real name: the keys whose trusted side rows hold it} over one dimension; None where the dimension has
    no trusted side rows (none stored, or an older release rebuilt it since). A store gives all of a dimension's side
    rows one stamp, so one row tells (an untrusted dimension costs no pass); a trusted one is read in one pass - for
    the rare key that has no side row of its own."""
    if not _has_real(db):
        return None
    stamp = db.execute("SELECT built_at FROM member_real WHERE connection=? AND dimension=? LIMIT 1",
                       (conn, dim)).fetchone()
    built = db.execute("SELECT built_at FROM build WHERE connection=? AND dimension=?", (conn, dim)).fetchone()
    if stamp is None or built is None or stamp[0] != built[0]:
        return None
    out: dict[str, set[str]] = {}
    for n, rn in db.execute("SELECT name, real_name FROM member_real WHERE connection=? AND dimension=? AND built_at=?",
                            (conn, dim, built[0])):
        out.setdefault(rn.casefold(), set()).add(n)
    return out


def _read_only_uri(path: Path) -> str:
    """SQLite's read-only URI of a file: absolute() keeps a mapped drive a drive (resolve() would make it a UNC path),
    and a UNC path's host goes into the path ('file:////host/share/...'): SQLite takes no URI authority."""
    uri = Path(path).absolute().as_uri()
    if not uri.startswith("file:///"):
        uri = "file:////" + uri[len("file://"):]
    return uri + "?mode=ro"


def present(path: Path | None = None) -> bool:
    """True when the atlas file (default: the home's) holds members - read only, never created: a missing file, or one
    with no member rows (an older release's atlas.connect() made such a file on a fresh home), is no atlas. A file that
    cannot be read, or that another process holds locked, is there (its own error says why; nothing builds over it).
    Never waits: the tool calls ask it on the event loop (fin_drill)."""
    p = Path(path or config.ATLAS_DB)
    if not p.exists():
        return False
    try:                                # timeout 0: a lock (a build's swap) counts as there, never waited on
        db = sqlite3.connect(_read_only_uri(p), uri=True, timeout=0)
    except sqlite3.Error:
        return True
    try:
        return db.execute("SELECT 1 FROM member LIMIT 1").fetchone() is not None
    except sqlite3.OperationalError as e:
        return "no such table" not in str(e)
    except sqlite3.Error:
        return True
    finally:
        db.close()


# A PC with no atlas (the controller's ruling of 22.11): one command - `smartass login` while a sign-in is missing (it
# signs in, then builds the atlas), else `smartass atlas build`. Here, beside present(), so that the tool calls and the
# knowledge layer need not import the CLI's modules.
def no_atlas_command() -> str:
    """The one command for a PC with no atlas: `smartass login` while a sign-in is missing (it signs in, then builds
    the atlas), else `smartass atlas build`. From the stored sign-ins only (no network)."""
    from ..epm import oauth
    try:
        ready = oauth.configured() and all(s["state"] == "signed_in" for s in oauth.status())
    except Exception:   # noqa: BLE001 - sign-ins that cannot be read: the sign-in comes first
        ready = False
    return "smartass atlas build" if ready else "smartass login"


def need_atlas() -> None:
    """The entries that need the atlas (fin_lookup's members, a drill, every answer's grids and knowledge layer) refuse
    on a PC with none, naming one command (no_atlas_step); "none" is read without creating the file (present)."""
    from ..model.errors import GuardError
    if not present():
        raise GuardError(f"no atlas on this PC yet ({config.ATLAS_DB}): every answer is built on it",
                         f"{no_atlas_step()}, then ask again")


def no_atlas_warning() -> str:
    """fin_lookup's one warning on a PC with no atlas that still answers from its knowledge.db: what is missing and the
    one command (no_atlas_step)."""
    return (f"no atlas on this PC yet ({config.ATLAS_DB}): no cube member is found and no answer can be built - "
            f"{no_atlas_step()}")


def no_atlas_step() -> str:
    """What a PC with no atlas does next, one command (no_atlas_command)."""
    cmd = no_atlas_command()
    return f"run `{cmd}` in a terminal" + (" (it signs you in to EPM, then builds the atlas)" if cmd == "smartass login"
                                            else " (read-only, about half a minute, then the behaviour cards)")


def trust(path: Path) -> tuple[int, int] | None:
    """(dimensions whose side rows are trusted, dimensions built with members) of the atlas file at path - read only,
    never written; None when it cannot be read.
    Where a dimension has no trusted side rows (an atlas built before 0.7.0, or rebuilt since by one) the REST
    engine's strict member check is off (epm/rest.py: the check before the side table)."""
    try:
        db = sqlite3.connect(_read_only_uri(path), uri=True)
    except sqlite3.Error:
        return None
    try:
        built = db.execute("SELECT connection, dimension, built_at FROM build WHERE n > 0 AND error IS NULL").fetchall()
        if not _has_real(db):
            return 0, len(built)
        trusted = sum(db.execute("SELECT 1 FROM member_real WHERE connection=? AND dimension=? AND built_at=? LIMIT 1",
                                 row).fetchone() is not None for row in built)
        return trusted, len(built)
    except sqlite3.Error:
        return None
    finally:
        db.close()


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


def purge(db, conn: str, keep: set[str] | frozenset[str] = frozenset()) -> None:
    """Every dimension of the connection goes (members, build rows, side rows), except those in keep."""
    marks = ",".join("?" * len(keep))
    tables = ["member", "build", *(["member_real"] if _has_real(db) else [])]
    with db:
        for t in tables:
            db.execute(f"DELETE FROM {t} WHERE connection=? AND dimension NOT IN ({marks})", (conn, *sorted(keep)))


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


def dimensions_info(db, connection: str | None = None) -> list[dict]:
    """Summary of built dimensions in atlas.db, optionally narrowed to one cube/connection."""
    sql = ("SELECT connection, dimension, built_at, n, error FROM build "
           "WHERE (? IS NULL OR connection = ?) ORDER BY connection, dimension")
    out = []
    for c, d, b, n, err in db.execute(sql, (connection, connection)).fetchall():
        out.append({
            "connection": c,
            "dimension": d,
            "members": n,
            "built_at": time.strftime("%Y-%m-%d %H:%M", time.localtime(b)) if b else None,
            **({"error": err} if err else {}),
        })
    return out


def _clean_label(name: str, label: str | None) -> str:
    if not label:
        return name
    return re.sub(r"\s*\([^()]*\)\s*$", "", label).strip() or name


def tree(db, connection: str, dimension: str | None = None, parent: str | None = None, depth: int = 1,
         leaves_only: bool = False, limit: int = 200) -> list[dict]:
    """Walk hierarchy nodes under `parent` (or top levels of `dimension`) up to `depth` hops, enriched with card & alias."""
    max_depth = max(1, min(int(depth), 6))
    lim = max(1, min(int(limit), 1000))
    root = parent or dimension
    if not root:
        return []
    dim_filter = dimension
    if not dim_filter and root:
        row = db.execute("SELECT dimension FROM member WHERE connection=? AND name=? COLLATE NOCASE LIMIT 1",
                         (connection, root)).fetchone()
        if row:
            dim_filter = row[0]

    has_real = _has_real(db)
    parents_set = {r[0] for r in db.execute(
        "SELECT DISTINCT parent FROM member WHERE connection=? AND parent IS NOT NULL", (connection,))}
    cards = {(n, p): (op, tb, at, sh) for n, p, op, tb, at, sh in db.execute(
        "SELECT name, parent, operator, time_balance, account_type, shared FROM card WHERE connection=?",
        (connection,))}
    real_map: dict[tuple[str, str], tuple[str | None, str | None]] = {}
    if has_real:
        for n, p, alias, storage in db.execute(
                "SELECT r.name, r.parent, r.alias, r.storage FROM member_real r JOIN build b "
                "ON b.connection = r.connection AND b.dimension = r.dimension AND b.built_at = r.built_at "
                "WHERE r.connection=?", (connection,)):
            real_map.setdefault((n, p or ""), (alias, storage))

    queue: list[tuple[str, int]] = [(root, 0)]
    seen_edges: set[tuple[str, str]] = set()
    out: list[dict] = []
    while queue and len(out) < lim:
        cur_parent, lvl = queue.pop(0)
        next_lvl = lvl + 1
        sql = ("SELECT dimension, name, label, parent, depth FROM member WHERE connection=? AND parent=? "
               + ("AND dimension=? " if dim_filter else "") + "ORDER BY ord")
        params = (connection, cur_parent, dim_filter) if dim_filter else (connection, cur_parent)
        rows = db.execute(sql, params).fetchall()
        if not rows and lvl == 0 and dim_filter and cur_parent.casefold() == dim_filter.casefold():
            sql_top = ("SELECT dimension, name, label, parent, depth FROM member WHERE connection=? AND dimension=? "
                       "AND depth <= 1 ORDER BY depth, ord")
            rows = db.execute(sql_top, (connection, dim_filter)).fetchall()
        for d, name, label, par, mem_depth in rows:
            if (par or "", name) in seen_edges:
                continue
            seen_edges.add((par or "", name))
            is_parent = name in parents_set
            if not leaves_only or not is_parent:
                item: dict = {
                    "dimension": d,
                    "name": name,
                    "label": _clean_label(name, label),
                    "parent": par,
                    "level": next_lvl,
                    "depth": mem_depth,
                    "has_children": is_parent,
                }
                c = cards.get((name, par or ""))
                if c:
                    op, tb, at, sh = c
                    item["card"] = {k: v for k, v in (
                        ("operator", op), ("time_balance", tb), ("account_type", at),
                        ("shared", bool(sh) if sh is not None else None)) if v not in (None, "")}
                r_info = real_map.get((name, par or ""))
                if r_info:
                    alias, storage = r_info
                    if alias:
                        item["alias"] = alias
                    if storage:
                        item["storage"] = storage
                out.append(item)
                if len(out) >= lim:
                    break
            if next_lvl < max_depth and is_parent:
                queue.append((name, next_lvl))
    return out
