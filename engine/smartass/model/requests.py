# engine/smartass/model/requests.py
"""What to fetch: one grid request per query and series (value + each comparison), with every account the
query's metrics need, the entity subtree as deep as its blocks go and the query's products (if any)."""
from __future__ import annotations

from collections import deque
from dataclasses import dataclass

from .periods import Closed, closed_for, label
from .resolver import Resolver
from .series import months, scenario
from .spec import AnalysisSpec, DriversBlock, MatrixBlock, RankingBlock, TableBlock, data_block

CONCEPT_BALANCE = "Periodic"


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
    products: tuple[str, ...] = ()        # empty: no product dimension (the cube's product total)


def _depth(spec: AnalysisSpec, qid: str) -> int:
    d = 0
    for b in map(data_block, spec.blocks):          # an own-data chart fetches like its table
        if getattr(b, "query", None) != qid:
            continue
        if isinstance(b, DriversBlock):
            d = max(d, b.depth)
        elif (isinstance(b, MatrixBlock) or (isinstance(b, RankingBlock) and b.rows == "entities")
              or (isinstance(b, TableBlock) and b.rows == "entities")):
            d = max(d, 1)
    return d


def _subtree(r: Resolver, cube: str, root: str, depth: int) -> tuple[str, ...]:
    """The root and its descendants as deep as the blocks go. A single-child pass-through node costs no depth,
    exactly as the drivers walk treats it, so the fetch reaches every entity the walk shows."""
    out = [root]
    queue = deque([(root, 0)])                              # breadth first: the fetch order stays level by level
    while queue:
        entity, level = queue.popleft()
        kids = r.children(entity, cube)
        if level >= depth or not kids:
            continue
        step = 0 if len(kids) == 1 else 1
        for k in kids:
            out.append(k)
            queue.append((k, level + step))
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
                                        tuple(label(m) for m in months(q, s, closed_for(closed, q.cube))), scen, ver, bal,
                                        tuple(dict.fromkeys(q.products))))
    from .blocks.concept import concept_requests              # the concept blocks' own grids, after the queries'
    for q, g, s in concept_requests(spec, r):
        scen, ver = scenario(q.cube, s, r)
        reqs.append(GridRequest(g.tag, q.cube, g.accounts, g.entities,
                                tuple(label(m) for m in months(q, s, closed_for(closed, q.cube))), scen, ver,
                                CONCEPT_BALANCE, g.products))
    return reqs
