"""Hierarchy change log for the atlas: snapshots of `member` and diffs between them."""
from __future__ import annotations

import time
from datetime import datetime

SCHEMA = """
CREATE TABLE IF NOT EXISTS snapshot (
  id INTEGER PRIMARY KEY, connection TEXT, dimension TEXT, taken_at REAL, n INT,
  UNIQUE(connection, dimension, taken_at)
);
CREATE TABLE IF NOT EXISTS snapshot_member (
  snapshot_id INTEGER, name TEXT, label TEXT, parent TEXT
);
CREATE INDEX IF NOT EXISTS ix_snapmem_snap ON snapshot_member(snapshot_id, name);
CREATE INDEX IF NOT EXISTS ix_snap_cd ON snapshot(connection, dimension, taken_at);
"""


def ensure_schema(db) -> None:
    db.executescript(SCHEMA)


def _iso(ts: float) -> str:
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")


def _rows(db, snap_id: int) -> set[tuple]:
    return {(n, lab or "", p or "") for n, lab, p in
            db.execute("SELECT name, label, parent FROM snapshot_member WHERE snapshot_id=?", (snap_id,))}


def take_snapshot(db, connection: str, dimension: str) -> int | None:
    ensure_schema(db)
    cur = {(n, lab or "", p or "") for n, lab, p in db.execute(
        "SELECT name, label, parent FROM member WHERE connection=? AND dimension=?", (connection, dimension))}
    if not cur:
        return None
    last = db.execute("SELECT id FROM snapshot WHERE connection=? AND dimension=? "
                      "ORDER BY taken_at DESC, id DESC LIMIT 1", (connection, dimension)).fetchone()
    if last and _rows(db, last[0]) == cur:
        return last[0]
    now = time.time()
    with db:
        c = db.execute("INSERT INTO snapshot(connection, dimension, taken_at, n) VALUES (?,?,?,?)",
                       (connection, dimension, now, len(cur)))
        sid = c.lastrowid
        db.executemany("INSERT INTO snapshot_member VALUES (?,?,?,?)", [(sid, n, lab, p) for n, lab, p in cur])
    return sid


def snapshot_all(db) -> list[dict]:
    ensure_schema(db)
    out = []
    for conn, dim in db.execute("SELECT connection, dimension FROM build WHERE error IS NULL OR error=''"
                                " ORDER BY 1, 2").fetchall():
        before = db.execute("SELECT max(id) FROM snapshot WHERE connection=? AND dimension=?",
                            (conn, dim)).fetchone()[0]
        sid = take_snapshot(db, conn, dim)
        out.append({"connection": conn, "dimension": dim, "snapshot_id": sid,
                    "new": sid is not None and sid != before})
    return out


def list_snapshots(db, connection=None, dimension=None) -> list[dict]:
    ensure_schema(db)
    sql = ("SELECT id, connection, dimension, taken_at, n FROM snapshot WHERE coalesce(?, connection)=connection "
           "AND coalesce(?, dimension)=dimension ORDER BY connection, dimension, taken_at")
    return [{"id": i, "connection": c, "dimension": d, "taken_at": _iso(t), "ts": t, "members": n}
            for i, c, d, t, n in db.execute(sql, [connection or None, dimension or None])]


def _to_ts(v, end: bool) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    try:
        return float(s)
    except ValueError:
        pass
    dt = datetime.fromisoformat(s.replace("Z", ""))
    if len(s) <= 10 and end:  # date only: include the whole day
        dt = dt.replace(hour=23, minute=59, second=59)
    return dt.timestamp()


def _group(rows: set[tuple]) -> dict[str, dict]:
    g: dict[str, dict] = {}
    for n, lab, p in rows:
        e = g.setdefault(n, {"label": lab, "parents": set()})
        if lab and not e["label"]:
            e["label"] = lab
        if p:
            e["parents"].add(p)
    return g


def _diff(old_rows: set[tuple], new_rows: set[tuple]) -> dict:
    a, b = _group(old_rows), _group(new_rows)
    added = [{"name": n, "label": b[n]["label"], "parents": sorted(b[n]["parents"])} for n in sorted(b.keys() - a.keys())]
    removed = [{"name": n, "label": a[n]["label"], "parents": sorted(a[n]["parents"])} for n in sorted(a.keys() - b.keys())]
    moved, relabelled = [], []
    for n in sorted(a.keys() & b.keys()):
        if a[n]["parents"] != b[n]["parents"]:
            moved.append({"name": n, "label": b[n]["label"], "from": sorted(a[n]["parents"]),
                          "to": sorted(b[n]["parents"])})
        if a[n]["label"] != b[n]["label"]:
            relabelled.append({"name": n, "old": a[n]["label"], "new": b[n]["label"]})
    return {"added": added, "removed": removed, "moved": moved, "relabelled": relabelled}


def changes(db, connection: str, dimension: str | None = None, since=None, until=None, limit: int = 500) -> dict:
    ensure_schema(db)
    t_since, t_until = _to_ts(since, False), _to_ts(until, True)
    if dimension:
        dims = [dimension]
    else:
        dims = [r[0] for r in db.execute("SELECT DISTINCT dimension FROM snapshot WHERE connection=? ORDER BY 1",
                                         (connection,))]
    if not dims:
        return {"connection": connection, "dimensions": {}, "message":
                f"No snapshots for {connection}. Run snapshot_all after an atlas build to take a baseline."}
    res: dict = {}
    summaries = []
    msgs = []
    for dim in dims:
        snaps = db.execute("SELECT id, taken_at FROM snapshot WHERE connection=? AND dimension=? "
                           "ORDER BY taken_at, id", (connection, dim)).fetchall()
        if not snaps:
            msgs.append(f"{dim}: no snapshots.")
            continue
        if len(snaps) == 1:
            msgs.append(f"{dim}: only one snapshot ({_iso(snaps[0][1])}); a baseline was just taken - "
                        "changes will appear after the next atlas build.")
            res[dim] = {"snapshots": 1, "baseline": _iso(snaps[0][1]), "counts": {}}
            continue
        ui = len(snaps) - 1
        if t_until is not None:
            cand = [i for i, s in enumerate(snaps) if s[1] <= t_until]
            ui = cand[-1] if cand else 0
        if t_since is not None:
            cand = [i for i, s in enumerate(snaps) if s[1] <= t_since]
            bi = cand[-1] if cand else 0
        else:
            bi = max(ui - 1, 0)
        if bi >= ui:
            res[dim] = {"snapshots": len(snaps), "counts": {},
                        "from": _iso(snaps[bi][1]), "to": _iso(snaps[ui][1])}
            msgs.append(f"{dim}: no distinct snapshots in the requested range.")
            continue
        d = _diff(_rows(db, snaps[bi][0]), _rows(db, snaps[ui][0]))
        counts = {k: len(v) for k, v in d.items()}
        entry = {"from": _iso(snaps[bi][1]), "to": _iso(snaps[ui][1]), "snapshots": len(snaps),
                 "counts": counts, "truncated": any(c > limit for c in counts.values())}
        for k, v in d.items():
            entry[k] = v[:limit]
        res[dim] = entry
        summaries.append(f"{dim}: " + (", ".join(f"{c} {k}" for k, c in counts.items() if c) or "no changes"))
    summary = "; ".join(summaries + msgs) or "No changes."
    return {"connection": connection, "dimensions": res, "summary": summary,
            "message": " ".join(msgs) if msgs and not summaries else None}


def _lab(label) -> str:
    return f" {label}" if label else ""


def changes_markdown(result: dict, max_lines: int = 30) -> str:
    out = [f"## Hierarchy changes - {result.get('connection', '')}", ""]
    dims = result.get("dimensions") or {}
    if result.get("message"):
        out += [result["message"], ""]
    for dim, e in dims.items():
        c = e.get("counts") or {}
        if not c:
            out += [f"### {dim}", "No comparison available (see message).", ""]
            continue
        out.append(f"### {dim} ({e['from']} -> {e['to']})")
        out.append(" | ".join(f"{k}: {v}" for k, v in c.items()))
        cats = [
            ("added", lambda x: f"added `{x['name']}`{_lab(x['label'])} under {', '.join(x['parents']) or '-'}"),
            ("removed", lambda x: f"removed `{x['name']}`{_lab(x['label'])} (was under {', '.join(x['parents']) or '-'})"),
            ("moved", lambda x: f"moved `{x['name']}`{_lab(x['label'])}: "
                                f"{', '.join(x['from']) or '-'} → {', '.join(x['to']) or '-'}"),
            ("relabelled", lambda x: f"relabelled `{x['name']}`: {x['old']} → {x['new']}"),
        ]
        for key, fmt in cats:
            items = e.get(key) or []
            if not items:
                continue
            out.append("")
            out += [f"- {fmt(x)}" for x in items[:max_lines]]
            if c.get(key, 0) > max_lines:
                out.append(f"- ... and {c[key] - max_lines} more {key}")
        out.append("")
    return "\n".join(out).rstrip() + "\n"
