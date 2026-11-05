# engine/smartass/model/chain.py
"""The P&L chain of a metric-row table: Revenue, COGS, Gross margin, Opex, EBITDA before other items, Other items,
EBITDA (semantic metric names). A new question's tables that show EBITDA get the missing chain metrics added to the
spec (complete_chain, before prepare), so they are fetched, stored and seen by refresh, patch and redraw; build_table
orders and types the chain on every path."""
from __future__ import annotations

import copy

from .spec import AnalysisSpec, TableBlock

CHAIN = ["Revenue", "COGS", "Gross margin", "Opex", "EBITDA before other items", "Other items", "EBITDA"]
KINDS = {"Gross margin": "sum", "EBITDA before other items": "sum", "EBITDA": "grand"}
TOP = "EBITDA"
# FinPlan Target carries no other items (live 05.11.2026: 2WHS rejected Total Other Items), so the chain stops at
# EBITDA there unless the query asks for them
NOT_IN_FINPLAN = ("Other items",)


def _resolves(r, name: str, cube: str) -> bool:
    try:
        r.members(name, cube)
    except KeyError:                    # UnknownMetric: not defined (in this cube)
        return False
    return True


def chained(b: TableBlock, labels: list[str]) -> bool:
    """Does the chain apply to this table: metric rows, chain on, and a metric that resolves to EBITDA."""
    return b.rows == "metrics" and b.chain and TOP in labels


def complete_chain(spec: AnalysisSpec, r) -> AnalysisSpec:
    """A new spec in which every chained table holds the whole chain its cube resolves: the missing chain metrics
    are appended (in chain order, after the metrics given, so their m{i} keys do not move) to the table and its
    query, and recorded in the table's chain_added. Idempotent; the spec given is not changed."""
    from .blocks.common import metric_label     # blocks.table imports this module
    out = spec.model_copy(deep=True)
    queries = {q.id: q for q in out.queries}
    for b in out.blocks:
        if not isinstance(b, TableBlock):
            continue
        q = queries[b.query]
        ms = list(b.metrics or q.metrics)
        labels = [metric_label(r, m, q.cube) for m in ms]
        if not chained(b, labels):
            continue
        in_query = {metric_label(r, m, q.cube): m for m in q.metrics}
        add = [in_query.get(c, c) for c in CHAIN if c not in labels and (c in in_query or _resolves(r, c, q.cube))
               and not (q.cube == "FinPlan" and c in NOT_IN_FINPLAN and c not in in_query)]
        if not add:
            continue
        b.metrics = ms + add
        b.chain_added = list(dict.fromkeys([*b.chain_added, *add]))
        q.metrics = [*q.metrics, *(m for m in add if m not in q.metrics)]
    return AnalysisSpec.model_validate(out.model_dump())


def restore_chain(raw: dict) -> dict:
    """A raw spec in which every table's query holds the chain metrics the table recorded in chain_added again: a
    follow-up patch that resends queries without them would otherwise fail validation ('not in query') before
    complete_chain runs. Only the record is used (no resolver); a copy, the raw given is not changed."""
    out = copy.deepcopy(raw)
    queries = {q.get("id"): q for q in out.get("queries") or [] if isinstance(q, dict)}
    for b in out.get("blocks") or []:
        if not isinstance(b, dict) or b.get("type") != "table" or not b.get("chain_added"):
            continue
        q = queries.get(b.get("query"))
        if q is None:
            continue
        have = list(q.get("metrics") or [])
        q["metrics"] = have + [m for m in b["chain_added"] if m not in have and m in (b.get("metrics") or [m])]
    return out
