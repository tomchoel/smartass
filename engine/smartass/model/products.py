# engine/smartass/model/products.py
"""Product hierarchy overlap & shared-member analysis: ensures parent+child or shared multi-parent
combinations never double-count in 'Total products', Share %, Other, or metric queries."""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class ProductOverlap:
    ancestor: str
    descendant: str
    weight: float


@dataclass
class ProductPlan:
    all_products: tuple[str, ...]
    dedup_products: tuple[str, ...]
    overlaps: list[ProductOverlap] = field(default_factory=list)
    parent_of: dict[str, str] = field(default_factory=dict)          # covered child -> nearest listed ancestor
    diamond_pairs: list[tuple[str, str, str]] = field(default_factory=list)   # (p1, p2, shared_descendant)
    warnings: list[str] = field(default_factory=list)


_OP_W = {"+": 1.0, "-": -1.0, "~": 0.0}


def _load_db_graph(r, cube: str) -> tuple[dict[str, list[tuple[str, float]]], dict[str, list[tuple[str, float]]]] | None:
    """Load the full All Products graph + operators in 2 SQL queries (avoids per-node round-trips)."""
    db = getattr(r, "db", None)
    if db is None:
        return None
    op_map: dict[tuple[str, str], float] = {}
    card_cube = "FinMax" if cube == "FinLite" else cube
    kdb = getattr(r, "kdb", None)
    if kdb is not None:
        for m, p, op in kdb.execute(
            "SELECT member, coalesce(parent, ''), operator FROM card WHERE cube=? AND dimension='All Products'",
            (card_cube,),
        ):
            if op in _OP_W:
                op_map[(m.casefold(), p.casefold())] = _OP_W[op]
    else:
        for m, p, op in db.execute(
            "SELECT name, coalesce(parent, ''), operator FROM card WHERE connection=? AND dimension='All Products'",
            (cube,),
        ):
            if op in _OP_W:
                op_map[(m.casefold(), p.casefold())] = _OP_W[op]

    children_of: dict[str, list[tuple[str, float]]] = {}
    parents_of: dict[str, list[tuple[str, float]]] = {}
    seen_edges: set[tuple[str, str]] = set()
    for name, par in db.execute(
        "SELECT name, coalesce(parent, '') FROM member WHERE connection=? AND dimension='All Products'",
        (cube,),
    ):
        if not par:
            continue
        ek = (par.casefold(), name.casefold())
        if ek in seen_edges:
            continue
        seen_edges.add(ek)
        w = op_map.get((name.casefold(), par.casefold()), 1.0)
        children_of.setdefault(par, []).append((name, w))
        parents_of.setdefault(name, []).append((par, w))
    return children_of, parents_of


def _weighted_descendants(root: str, get_children) -> tuple[dict[str, float], bool]:
    """Traverse downwards from `root` along nonzero consolidation edges (memoized on state to avoid DAG blowup)."""
    net: dict[str, float] = {}
    stack: list[tuple[str, float]] = [(root, 1.0)]
    seen_states: set[tuple[str, float]] = {(root, 1.0)}
    on_cycle = False
    while stack:
        cur, mult = stack.pop()
        for kid, w in get_children(cur):
            if w == 0:
                continue
            if kid == root:
                on_cycle = True
                continue
            eff = mult * w
            net[kid] = net.get(kid, 0.0) + eff
            st = (kid, eff)
            if st not in seen_states:
                seen_states.add(st)
                stack.append(st)
    return net, on_cycle


def analyze_products(products: list[str] | tuple[str, ...], r, cube: str) -> ProductPlan:
    """Analyze hierarchy overlaps (direct, transitive, multi-parent shared, negative weights, diamond sharing)."""
    unique = tuple(dict.fromkeys(products))
    if len(unique) <= 1:
        return ProductPlan(unique, unique)

    db_graph = _load_db_graph(r, cube)
    if db_graph is not None:
        children_of, parents_of = db_graph
        def get_children(node: str):
            return children_of.get(node, ())
        def get_parents(node: str):
            return parents_of.get(node, ())
    else:
        def get_children(node: str):
            return [(k, r.weight(k, cube, node)) for k in r.children(node, cube)]
        def get_parents(node: str):
            if not hasattr(r, "parents"):
                return []
            return [(p, r.weight(node, cube, p)) for p in r.parents(node, cube)]

    desc_map: dict[str, dict[str, float]] = {}
    unverified = False
    for p in unique:
        dmap, bad = _weighted_descendants(p, get_children)
        desc_map[p] = dmap
        unverified = unverified or bad

    listed_set = set(unique)
    for child in unique:
        up_stack: list[tuple[str, float]] = [(child, 1.0)]
        seen_up: set[tuple[str, float]] = {(child, 1.0)}
        while up_stack:
            cur, mult = up_stack.pop()
            for par, w in get_parents(cur):
                if w == 0 or par == child:
                    continue
                eff = mult * w
                if par in listed_set and child not in desc_map.get(par, {}):
                    desc_map.setdefault(par, {})[child] = eff
                st = (par, eff)
                if st not in seen_up:
                    seen_up.add(st)
                    up_stack.append(st)

    overlaps: list[ProductOverlap] = []
    covered_by: dict[str, str] = {}
    neg_overlaps: list[ProductOverlap] = []

    for anc in unique:
        for desc in unique:
            if anc == desc:
                continue
            w = desc_map.get(anc, {}).get(desc, 0.0)
            if w > 0:
                ov = ProductOverlap(anc, desc, w)
                overlaps.append(ov)
                covered_by.setdefault(desc, anc)
            elif w < 0:
                ov = ProductOverlap(anc, desc, w)
                overlaps.append(ov)
                neg_overlaps.append(ov)

    dedup = tuple(p for p in unique if p not in covered_by)

    diamonds: list[tuple[str, str, str]] = []
    for i, p1 in enumerate(dedup):
        d1 = {k for k, v in desc_map.get(p1, {}).items() if v > 0}
        if not d1:
            continue
        for p2 in dedup[i + 1:]:
            d2 = {k for k, v in desc_map.get(p2, {}).items() if v > 0}
            shared = sorted(d1 & d2)
            for s in shared:
                diamonds.append((p1, p2, s))
                break

    warnings: list[str] = []
    if overlaps:
        pairs = ", ".join(f"{o.ancestor} ⊃ {o.descendant}" for o in overlaps[:6])
        more = f" (+{len(overlaps) - 6} more)" if len(overlaps) > 6 else ""
        warnings.append(
            f"Overlapping products detected ({pairs}{more}): counting each amount once in totals "
            f"({len(dedup)} root product(s) summed instead of {len(unique)})."
        )
    if neg_overlaps:
        pairs = ", ".join(f"{o.ancestor} (-{o.descendant})" for o in neg_overlaps[:4])
        warnings.append(f"Negative consolidation weight between listed products ({pairs}); verify net total.")
    if diamonds:
        pairs = ", ".join(f"{a} & {b} share {s}" for a, b, s in diamonds[:3])
        warnings.append(f"Shared descendant counted under multiple non-nested products ({pairs}); total may overlap.")
    if unverified:
        warnings.append("Product hierarchy contains a cycle or unverified placement; verify totals.")

    return ProductPlan(unique, dedup, overlaps, covered_by, diamonds, warnings)
