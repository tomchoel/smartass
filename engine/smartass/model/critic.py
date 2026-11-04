# engine/smartass/model/critic.py
"""The critic brief (sub-spec 2 §5): a compact view of one answer for the critic pass - the plan, the material lines
and whether a driver backs each, the comparisons, the warn/block caveats, checks, findings, the commentary drivers
and the gaps the critic must see. Pure: built from the spec and the computed Result, no Excel, no EPM."""
from __future__ import annotations

from .analyze import Result
from .spec import AnalysisSpec, CommentaryBlock, LinesBlock

SERIOUS = ("block", "warn")
ACTUALS = ("FinMax", "FinLite")
MATCH_CHARS = 30


def _lines_by_key(result: Result) -> dict[str, tuple]:
    return {f"{o.key}.{ln.key}": (o, ln) for o in result.outs for ln in o.lines}


def _coverage(spec: AnalysisSpec, result: Result) -> list[tuple[str, str | None, set[str] | None]]:
    """(evidence, metric, entities) per drivers or lines block. A drivers block covers its metric for every
    entity in its tree; a lines block covers the entity of its cell (any metric: the ledger lines explain it)."""
    out: list[tuple[str, str | None, set[str] | None]] = []
    for o in result.outs:
        if o.type == "drivers":
            metric = next((ln.metric for ln in o.lines if ln.metric), None)
            out.append((f"drivers block {o.key}", metric, {ln.entity for ln in o.lines if ln.entity}))
    for i, b in enumerate(spec.blocks):
        if isinstance(b, LinesBlock):
            ent = b.cell.get("All Entities")
            out.append((f"lines block b{i}", None, {ent} if ent else None))
    return out


def _commentary(spec: AnalysisSpec) -> list[CommentaryBlock]:
    return [b for b in spec.blocks if isinstance(b, CommentaryBlock)]


def _names(m: dict) -> set[str]:
    """The ways a commentary driver may name a material line: its key, its label, 'block · label'."""
    return {s.casefold() for s in (m["key"], m["line"], f"{m['block']} · {m['line']}")}


def _backing(m: dict, line, cover, drivers) -> list[str]:
    ev: list[str] = []
    if line is not None and line.metric:
        for what, metric, ents in cover:
            if metric not in (None, line.metric):
                continue
            if ents is None or line.entity is None or line.entity in ents:
                ev.append(what)
    names = _names(m)
    ev += [d.evidence for d in drivers if d.basis == "data" and d.line.strip().casefold() in names]
    return ev


def _mentioned(cav: dict, text: str) -> bool:
    head = (cav.get("text") or "")[:MATCH_CHARS].strip().casefold()
    return (cav.get("id") or "").casefold() in text or bool(head) and head in text


def brief(spec: AnalysisSpec, result: Result, caveats: list[dict], checks: dict,
          stock: set[str] | frozenset = frozenset()) -> dict:
    """stock: the stock metrics (FTE); a query of only stocks is not a P&L query."""
    by_key = _lines_by_key(result)
    cover = _coverage(spec, result)
    comm = _commentary(spec)
    drivers = [d for c in comm for d in c.drivers]
    gaps: list[str] = []
    material = []
    for m in result.material:
        _, line = by_key.get(m["key"], (None, None))
        ev = _backing(m, line, cover, drivers)
        material.append({"key": m["key"], "line": m["line"], "variance": m["variance"], "backed": bool(ev),
                         "evidence": ev})
        if not ev:
            gaps.append(f"material line '{m['line']}' has no driver")
    text = " ".join([c.title for c in comm] + [p for c in comm for p in c.key_points]
                    + [f"{d.line} {d.text} {d.evidence}" for d in drivers]).casefold()
    serious = [c for c in caveats if c.get("severity") in SERIOUS]
    gaps += [f"{c['severity']} caveat '{c.get('id')}' not mentioned" for c in serious if not _mentioned(c, text)]
    for q in spec.queries:
        if q.cube in ACTUALS and "ly" not in q.compare and any(x not in stock for x in q.metrics):
            gaps.append(f"query '{q.id}' (P&L) has no last-year comparison")
    return {"purpose": spec.plan.purpose, "depth": spec.plan.depth, "hypotheses": list(spec.plan.hypotheses),
            "material": material,
            "comparisons": list(dict.fromkeys(c for q in spec.queries for c in q.compare)),
            "caveats": serious, "checks": dict(checks),
            "findings": [{"severity": f.severity, "text": f.text, "where": f.where} for f in result.findings],
            "commentary": {"key_points": sum(len(c.key_points) for c in comm),
                           "drivers": [{"line": d.line, "basis": d.basis} for d in drivers]},
            "gaps": gaps}
