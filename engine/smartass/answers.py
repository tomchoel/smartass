# engine/smartass/answers.py
"""Answer memory (sub-spec 2 §7): every fin_analyze / follow-up / refresh run is stored in HOME/answers.db (per user,
never in git, never shared) with its question, plan purpose, spec, key values, metrics, entities, periods, critic
brief and critic score. Searchable (FTS5 over question, purpose, metrics, entities) for fin_lookup(kind='answer');
rows older than 24 months are purged. A run of the same workbook and Q adds a row; get and search show the latest."""
from __future__ import annotations

import contextlib
import datetime as dt
import json
import os
import re
import sqlite3
from pathlib import Path

KEEP_DAYS = 730                         # 24 months

SCHEMA = """
CREATE TABLE IF NOT EXISTS answer (
  id INTEGER PRIMARY KEY, at TEXT NOT NULL, workbook TEXT NOT NULL, wb_key TEXT NOT NULL, q INTEGER NOT NULL,
  question TEXT NOT NULL, purpose TEXT NOT NULL DEFAULT '', spec TEXT NOT NULL DEFAULT '{}',
  key_values TEXT NOT NULL DEFAULT '{}', metrics TEXT NOT NULL DEFAULT '[]', entities TEXT NOT NULL DEFAULT '[]',
  periods TEXT NOT NULL DEFAULT '[]', critic_brief TEXT, score REAL, notes TEXT NOT NULL DEFAULT '');
CREATE INDEX IF NOT EXISTS answer_wb_q ON answer (wb_key, q, id);
CREATE VIRTUAL TABLE IF NOT EXISTS answer_fts USING fts5(question, purpose, metrics, entities);
"""
JSON_COLS = ("spec", "key_values", "metrics", "entities", "periods", "critic_brief")


def db_path() -> Path:
    from .core import config
    return Path(config.HOME) / "answers.db"


def _key(workbook: str) -> str:
    """One file, whatever its spelling: 'C:/a/X.xlsx' and 'c:\\a\\x.xlsx' are the same answer workbook."""
    return os.path.normcase(os.path.normpath(workbook or ""))


def _now() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def _purge(db: sqlite3.Connection) -> None:
    cut = (dt.datetime.now() - dt.timedelta(days=KEEP_DAYS)).isoformat(timespec="seconds")
    ids = [r[0] for r in db.execute("SELECT id FROM answer WHERE at < ?", (cut,))]
    db.executemany("DELETE FROM answer_fts WHERE rowid = ?", [(i,) for i in ids])
    db.executemany("DELETE FROM answer WHERE id = ?", [(i,) for i in ids])


@contextlib.contextmanager
def _connect():
    path = db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(path)
    try:
        db.row_factory = sqlite3.Row
        db.executescript(SCHEMA)
        with db:
            yield db
    finally:
        db.close()


def _row(r: sqlite3.Row | None) -> dict | None:
    if r is None:
        return None
    out = dict(r)
    for c in JSON_COLS:
        out[c] = json.loads(out[c]) if out.get(c) else None
    out.pop("wb_key", None)
    return out


def store(run: dict, at: str | None = None) -> int:
    """run: workbook, q, question, purpose, spec, key_values, metrics, entities, periods, critic_brief (optional).
    Returns the run id."""
    with _connect() as db:
        _purge(db)
        vals = {c: json.dumps(run.get(c), ensure_ascii=False, default=str) if run.get(c) is not None else None
                for c in JSON_COLS}
        for c in ("spec", "key_values"):
            vals[c] = vals[c] or "{}"
        for c in ("metrics", "entities", "periods"):
            vals[c] = vals[c] or "[]"
        cur = db.execute(
            "INSERT INTO answer (at, workbook, wb_key, q, question, purpose, spec, key_values, metrics, entities, "
            "periods, critic_brief) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (at or _now(), run["workbook"], _key(run["workbook"]), int(run["q"]), run.get("question") or "",
             run.get("purpose") or "", vals["spec"], vals["key_values"], vals["metrics"], vals["entities"],
             vals["periods"], vals["critic_brief"]))
        rid = cur.lastrowid
        db.execute("INSERT INTO answer_fts (rowid, question, purpose, metrics, entities) VALUES (?, ?, ?, ?, ?)",
                   (rid, run.get("question") or "", run.get("purpose") or "", " ".join(run.get("metrics") or []),
                    " ".join(run.get("entities") or [])))
        return rid


def _latest(db: sqlite3.Connection, wb_key: str, q: int) -> sqlite3.Row | None:
    return db.execute("SELECT * FROM answer WHERE wb_key = ? AND q = ? ORDER BY id DESC LIMIT 1",
                      (wb_key, int(q))).fetchone()


def get(workbook: str, q: int) -> dict | None:
    """The latest stored run of Q<q> in this workbook, or None."""
    with _connect() as db:
        return _row(_latest(db, _key(workbook), q))


def latest(q: int | None = None, workbook: str | None = None) -> dict | None:
    """The most recent stored run (of Q<q> when given) in any workbook, or in `workbook` (its path, or the name of an
    unsaved book, as store keys it): fin_critic without an Excel session, the default target of follow-ups."""
    where, args = [], []
    if q is not None:
        where.append("q = ?")
        args.append(int(q))
    if workbook is not None:
        where.append("wb_key = ?")
        args.append(_key(workbook))
    with _connect() as db:
        sql = ("SELECT * FROM answer" + (" WHERE " + " AND ".join(where) if where else "")
               + " ORDER BY id DESC LIMIT 1")
        return _row(db.execute(sql, args).fetchone())


def set_score(workbook: str, q: int, score: float, notes: str = "") -> bool:
    """The critic score (1-5) on the latest run of Q<q>; False when no run is stored."""
    with _connect() as db:
        r = _latest(db, _key(workbook), q)
        if r is None:
            return False
        db.execute("UPDATE answer SET score = ?, notes = ? WHERE id = ?", (float(score), notes or "", r["id"]))
        return True


def _match(text: str, op: str) -> str:
    return f" {op} ".join(f'"{t}"*' for t in re.findall(r"\w+", text))


def search(text: str, limit: int = 10) -> list[dict]:
    """Past answers matching every word of text (prefix; question, purpose, metrics, entities), else any word.
    One hit per workbook and Q (its latest run), best match first."""
    with _connect() as db:
        _purge(db)
        ranked: list = []
        for op in ("AND", "OR"):
            m = _match(text, op)
            if not m:
                break
            ranked = db.execute("SELECT rowid, bm25(answer_fts) FROM answer_fts WHERE answer_fts MATCH ? "
                                "ORDER BY bm25(answer_fts), rowid DESC", (m,)).fetchall()
            if ranked:
                break
        keys: dict[tuple, None] = {}                    # (wb_key, q) in best-match order
        for rid, _ in ranked:
            r = db.execute("SELECT wb_key, q FROM answer WHERE id = ?", (rid,)).fetchone()
            if r is not None:
                keys.setdefault((r["wb_key"], r["q"]), None)
        return [_row(_latest(db, k, q)) for k, q in list(keys)[:int(limit)]]
