# engine/smartass/model/findings.py
"""Material lines (they need a driver in the commentary) and key_values (the numbers the chat may quote)."""
from __future__ import annotations

from .blocks.common import GAP, BlockOut, base_of, ck, group_start
from .expr import Cells
from .resolver import Resolver
from .validate import TOLERANCE, Finding


def material(outs: list[BlockOut], cells: Cells, r: Resolver) -> list[dict]:
    """A line is material when |variance| >= the metric's absolute threshold, or when |variance %| >= its
    percentage threshold and |variance| >= a tenth of the absolute threshold. A block is read through its column
    codes: it needs a 'var' column (a table with only 'var_pct' is skipped); each var column is judged against its
    own base column, together with the first var_pct against the same base. With column groups only the last group
    (YTD) is judged, so a line is not listed twice."""
    hits = []
    seen: set[tuple] = set()
    for o in outs:
        last = group_start(o.cols, len(o.cols))
        for vi, c in enumerate(o.cols):
            if c != "var" or vi < last:
                continue
            base = o.cols[base_of(o.cols, vi)]
            pi = next((j for j, x in enumerate(o.cols)
                       if j >= last and x == "var_pct" and o.cols[base_of(o.cols, j)] == base), None)
            for ln in o.lines:
                if ln.kind == "check" or not ln.metric:
                    continue
                v = cells.get(ck(o.key, ln.key, vi))
                if v is None:
                    continue
                p = cells.get(ck(o.key, ln.key, pi)) if pi is not None else None
                a, pct = r.materiality(ln.metric)
                if abs(v) >= a or (p is not None and abs(p) >= pct and abs(v) >= a / 10):
                    if ln.entity is not None:       # the same cell in another block is not listed twice
                        ident = ((o.query, ln.metric, ln.entity, base) if o.type != "concept"
                                 else (o.query, o.title, ln.key, base))     # a concept line is its own cell
                        if ident in seen:
                            continue
                        seen.add(ident)
                    hits.append({"block": o.title, "line": ln.label, "key": f"{o.key}.{ln.key}", "variance": v,
                                 "variance_pct": p, "base": base})
    return hits


def bridge_findings(outs: list[BlockOut], cells: Cells, r: Resolver) -> list[Finding]:
    """A bridge whose 'Other / unexplained' is as large as the total metric's materiality has not explained the
    movement: the steps are missing a driver."""
    out: list[Finding] = []
    for o in outs:
        if o.type != "bridge":
            continue
        end = next((ln for ln in o.lines if ln.key == "end"), None)
        v = cells.get(ck(o.key, "other", 0))
        if end is None or not end.metric or v is None:
            continue
        limit = r.materiality(end.metric)[0]
        if abs(v) >= limit:
            out.append(Finding("warning", f"{o.title}: Other / unexplained is {v:,.0f}, as large as the "
                                          f"materiality of {end.metric} ({limit:,.0f}); add the missing steps",
                               o.key))
    return out


SKIP = (GAP, "material")     # spacer and marker columns are not numbers to quote


def key_values(outs: list[BlockOut], cells: Cells, limit: int = 250) -> dict[str, dict[str, float | None]]:
    """'<block title> · <line label>' -> {stable column key: value}, at most limit lines. The cap is filled with the
    lines given first and the chain lines complete_chain added after them (an added line never pushes a given one
    out); the lines kept are listed in block order."""
    lines = [(o, ln) for o in outs for ln in o.lines if ln.kind != "check"]
    keep: set[str] = set()
    for added in (False, True):
        for o, ln in lines:
            name = f"{o.title} · {ln.label}"
            if ln.added == added and name not in keep and len(keep) < limit:
                keep.add(name)
    kv: dict[str, dict[str, float | None]] = {}
    for o, ln in lines:
        name = f"{o.title} · {ln.label}"
        if name not in keep or (ln.added and name in kv):
            continue
        codes = o.cols if len(o.cols) == len(o.keys) else [None] * len(o.keys)
        kv[name] = {k: cells.get(ck(o.key, ln.key, j)) for j, k in enumerate(o.keys) if codes[j] not in SKIP}
    return kv


def lines_findings(outs: list[BlockOut], cells: Cells) -> list[Finding]:
    """A lines block whose ledger lines do not add up to the cube value: the drill and the cube disagree."""
    out: list[Finding] = []
    for o in outs:
        if o.type != "lines" or not any(ln.key == "gap" for ln in o.lines):
            continue
        gap = cells.get(ck(o.key, "gap", 0))
        if gap is not None and abs(gap) >= TOLERANCE:
            out.append(Finding("warning", f"{o.title}: Ledger lines do not add up to the cube: gap {gap:,.0f} NOK "
                                          "(e.g. lines posted after the drill, or rounding)", o.key))
    return out
