# engine/smartass/model/findings.py
"""Material lines (they need a driver in the commentary) and key_values (the numbers the chat may quote)."""
from __future__ import annotations

from .blocks.common import BlockOut, ck
from .expr import Cells
from .resolver import Resolver


def material(outs: list[BlockOut], cells: Cells, r: Resolver) -> list[dict]:
    """A line is material when |variance| >= the metric's absolute threshold, or when |variance %| >= its
    percentage threshold and |variance| >= a tenth of the absolute threshold."""
    hits = []
    for o in outs:
        if "Variance" not in o.headers:
            continue
        vi = o.headers.index("Variance")
        pi = o.headers.index("Variance %") if "Variance %" in o.headers else None
        for ln in o.lines:
            if ln.kind == "check" or not ln.metric:
                continue
            v = cells.get(ck(o.key, ln.key, vi))
            if v is None:
                continue
            p = cells.get(ck(o.key, ln.key, pi)) if pi is not None else None
            a, pct = r.materiality(ln.metric)
            if abs(v) >= a or (p is not None and abs(p) >= pct and abs(v) >= a / 10):
                hits.append({"block": o.title, "line": ln.label, "key": f"{o.key}.{ln.key}", "variance": v,
                             "variance_pct": p})
    return hits


def key_values(outs: list[BlockOut], cells: Cells, limit: int = 40) -> dict[str, dict[str, float | None]]:
    kv: dict[str, dict[str, float | None]] = {}
    for o in outs:
        for ln in o.lines:
            if ln.kind == "check":
                continue
            if len(kv) >= limit:
                return kv
            kv[f"{o.title} · {ln.label}"] = {h: cells.get(ck(o.key, ln.key, j)) for j, h in enumerate(o.headers)}
    return kv
