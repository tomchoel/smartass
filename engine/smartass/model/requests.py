# engine/smartass/model/requests.py
"""What to fetch: one grid request per query and series (value + each comparison), with every account the
query's metrics need and the entity subtree as deep as its blocks go."""
from __future__ import annotations

from dataclasses import dataclass

from .periods import Closed, closed_for, label
from .resolver import Resolver
from .series import months, scenario
from .spec import AnalysisSpec, DriversBlock, MatrixBlock, RankingBlock, TableBlock


@dataclass(frozen=True)
class GridRequest:
    query: str
    cube: str
    accounts: tuple[str, ...]
    entities: tuple[str, ...]
    periods: tuple[str, ...]
    scenario: str
    version: str
    balance: str = "Periodic"


def _depth(spec: AnalysisSpec, qid: str) -> int:
    d = 0
    for b in spec.blocks:
        if getattr(b, "query", None) != qid:
            continue
        if isinstance(b, DriversBlock):
            d = max(d, b.depth)
        elif isinstance(b, (RankingBlock, MatrixBlock)) or (isinstance(b, TableBlock) and b.rows == "entities"):
            d = max(d, 1)
    return d


def _subtree(r: Resolver, cube: str, root: str, depth: int) -> tuple[str, ...]:
    out, frontier = [root], [root]
    for _ in range(depth):
        frontier = [c for e in frontier for c in r.children(e, cube)]
        out += frontier
    return tuple(dict.fromkeys(out))


def requests(spec: AnalysisSpec, r: Resolver, closed: Closed) -> list[GridRequest]:
    reqs = []
    for q in spec.queries:
        by_balance: dict[str, list[str]] = {}
        for m in q.metrics:
            by_balance.setdefault(r.balance(m), []).extend(r.members(m, q.cube))
        ents = _subtree(r, q.cube, q.entity, _depth(spec, q.id))
        for s in ("value", *q.compare):
            scen, ver = scenario(q.cube, s, r)
            for bal, accs in by_balance.items():
                reqs.append(GridRequest(q.id, q.cube, tuple(dict.fromkeys(accs)), ents,
                                        tuple(label(m) for m in months(q, s, closed_for(closed, q.cube))), scen, ver, bal))
    return reqs
