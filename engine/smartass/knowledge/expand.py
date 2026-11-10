"""A fact's member scope expanded over the atlas tree (spec 2026-11-10 §4). The build loads each cube's member table
into memory once (Tree) and stores the result in fact_scope with `via` = the member the fact names."""
from __future__ import annotations

import sqlite3

from .semantic import Fact


class Tree:
    """One cube's members in memory: placements by name (any case) and children by (dimension, parent)."""

    def __init__(self, db: sqlite3.Connection, cube: str):
        self.by_name: dict[str, list[tuple[str, str, str, int]]] = {}
        self.kids: dict[tuple[str, str], list[str]] = {}
        for dim, name, parent, depth in db.execute(
                "SELECT dimension, name, coalesce(parent, ''), depth FROM member WHERE connection=?", (cube,)):
            self.by_name.setdefault(name.casefold(), []).append((name, dim, parent, depth))
            if parent:
                self.kids.setdefault((dim, parent.casefold()), []).append(name)

    def placements(self, member: str, dims: set[str] | None = None) -> list[tuple[str, str, str, int]]:
        """(name, dimension, parent, depth) of every placement of the member, only those in `dims` when given."""
        rows = self.by_name.get(member.casefold(), [])
        return [r for r in rows if r[1] in dims] if dims else rows

    def canonical(self, member: str) -> str:
        p = self.placements(member)
        return p[0][0] if p else member


class LazyTrees(dict):
    """{cube: Tree}, each loaded on first use (only the cubes facts touch)."""

    def __init__(self, db: sqlite3.Connection):
        super().__init__()
        self.db = db

    def __missing__(self, cube: str) -> Tree:
        self[cube] = tree = Tree(self.db, cube)
        return tree


def descendants(tree: Tree, member: str) -> set[str]:
    out: set[str] = set()
    todo = [(dim, name) for name, dim, _p, _d in tree.placements(member)]
    while todo:
        dim, name = todo.pop()
        for kid in tree.kids.get((dim, name.casefold()), []):
            if kid not in out:
                out.add(kid)
                todo.append((dim, kid))
    return out


def ancestors(tree: Tree, member: str, stops: set[str]) -> set[str]:
    """Up every parent chain (in the dimension of each placement) to the first ancestor in `stops` (included); a chain
    without one gives its first parent. A cycle or a self-parent ends a chain."""
    out: set[str] = set()
    folded = {s.casefold() for s in stops}
    for _name, dim, parent in ((n, d, p) for n, d, p, _ in tree.placements(member)):
        if not parent:
            continue
        stack = [[parent]]
        while stack:
            path = stack.pop()
            node = path[-1]
            if node.casefold() in folded:
                out.update(path)
                continue
            seen = {x.casefold() for x in path}
            nexts = [p for _n, d, p, _ in tree.placements(node) if d == dim and p and p.casefold() not in seen]
            if not nexts:
                out.add(path[0])
            stack.extend([*path, p] for p in dict.fromkeys(nexts))
    return out


def expand(fact: Fact, trees: dict[str, Tree] | sqlite3.Connection,
           stops: dict[str, set[str]]) -> dict[tuple[str, str], str]:
    """{(cube, member): via} for the fact's member scope; `via` is the member the fact itself names. Every named
    member maps to itself; the atlas spelling is stored when the member exists."""
    if isinstance(trees, sqlite3.Connection):
        trees = LazyTrees(trees)
    out: dict[tuple[str, str], str] = {}
    for cube, members in fact.scope_members.items():
        tree = trees[cube]
        named = {m: tree.canonical(m) for m in members}
        for c in named.values():
            out[(cube, c)] = c
        for m, c in named.items():
            reach: set[str] = set()
            if fact.inherit in ("down", "both"):
                reach |= descendants(tree, m)
            if fact.inherit in ("up", "both"):
                reach |= ancestors(tree, m, stops.get(cube, set()))
            for x in reach:
                out.setdefault((cube, x), c)
    return out
