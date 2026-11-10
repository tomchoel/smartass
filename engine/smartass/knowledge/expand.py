"""A fact's member scope expanded over the atlas tree (spec 2026-11-10 §4). Pure reads of atlas.db; the build
stores the result in fact_scope with `via` = the member the fact names."""
from __future__ import annotations

import sqlite3

from .semantic import Fact


def _rows(db: sqlite3.Connection, cube: str, member: str) -> list[tuple[str, str, str]]:
    """(name, dimension, parent) of every placement of the member (a shared member sits under several parents)."""
    return db.execute("SELECT name, dimension, coalesce(parent, '') FROM member WHERE connection=? AND name=? "
                      "COLLATE NOCASE", (cube, member)).fetchall()


def descendants(db: sqlite3.Connection, cube: str, member: str) -> set[str]:
    out: set[str] = set()
    todo = [(r[0], r[1]) for r in _rows(db, cube, member)]
    while todo:
        name, dim = todo.pop()
        for (kid,) in db.execute("SELECT DISTINCT name FROM member WHERE connection=? AND dimension=? AND parent=?",
                                 (cube, dim, name)):
            if kid not in out:
                out.add(kid)
                todo.append((kid, dim))
    return out


def ancestors(db: sqlite3.Connection, cube: str, member: str, stops: set[str]) -> set[str]:
    """Up every parent chain to the nearest ancestor in `stops` (included); a chain without one gives its parent."""
    out: set[str] = set()
    folded = {s.casefold() for s in stops}
    for _name, _dim, parent in _rows(db, cube, member):
        chain, p = [], parent
        while p:
            chain.append(p)
            if p.casefold() in folded:
                break
            row = _rows(db, cube, p)
            p = row[0][2] if row else ""
        else:
            chain = chain[:1]
        out.update(chain)
    return out


def expand(fact: Fact, db: sqlite3.Connection, stops: dict[str, set[str]]) -> dict[tuple[str, str], str]:
    """{(cube, member): via} for the fact's member scope; `via` is the member the fact itself names."""
    out: dict[tuple[str, str], str] = {}
    for cube, members in fact.scope_members.items():
        for m in members:
            out.setdefault((cube, m), m)
            reach: set[str] = set()
            if fact.inherit in ("down", "both"):
                reach |= descendants(db, cube, m)
            if fact.inherit in ("up", "both"):
                reach |= ancestors(db, cube, m, stops.get(cube, set()))
            for x in reach:
                out.setdefault((cube, x), m)
    return out
