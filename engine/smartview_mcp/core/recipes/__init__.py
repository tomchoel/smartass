"""Recipes (design §6): pure functions that turn a few parameters into a draft answer spec v1 - pulls, layout
(pull cells use the pull ids), checks. Not tools: sv_answer expands {"recipe": {...}} server-side, applies the
recipe's overrides as a follow-up patch and validates the result like any other spec. Every recipe except
'concept' must emit at least one check (the concept library's own checks are passed through)."""
from __future__ import annotations

import copy
from dataclasses import dataclass, field
from types import SimpleNamespace
from typing import Callable

from .. import contract


@dataclass
class RecipeContext:
    closed: str                                  # FinMax Current Reporting Month, e.g. 'Oct 26'
    cat: object                                  # catalog.Catalog
    label: Callable[[str, str], str]             # (connection, member) -> readable label (templates.member_label)
    pnl_facts: Callable | None = None            # (area, period, compare) -> (res, mode, fills, notes, entity)
    now: str = ""
    lines: dict[str, list[dict]] | None = None   # drilled ledger lines by alias (lines_ranking: d1 this year, d2 LY)


@dataclass
class Draft:
    pulls: list[dict]
    layout: dict
    assumptions: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    result: dict = field(default_factory=dict)


def pull_dict(ps) -> dict:
    """templates.PullSpec -> contract pull dict (no title: source.sheet_title names it Q<n> <connection> <id>)."""
    def ax(a):
        return {"dimension": a.dimension, "members": list(a.members)}
    return {"id": ps.id, "connection": ps.connection, "rows": [ax(a) for a in ps.rows],
            "columns": [ax(a) for a in ps.cols], **({"pov": dict(ps.pov)} if ps.pov else {})}


def stub_pulls(ids) -> dict:
    """Stand-ins for a template's answer(P) before the pull: sheet = the pull id, no labels, no values (so pull
    cells come out in pull-id form and value-based key points as 'n/a', which the recipe drops)."""
    return {i: SimpleNamespace(sheet=i, row_labels={}, col_labels={},
                               layout=SimpleNamespace(col_tuples=[], nr=0), parsed=SimpleNamespace(rows=[]))
            for i in ids}


def relabel(lines: list[dict], connection: str, ctx: RecipeContext) -> None:
    """Template labels fall back to member codes on stubs ('ac.C30', 'ac.C30 — Accumulated Oct 26'): replace the
    leading code with its atlas label, keep the rest."""
    for ln in lines:
        code, _, rest = (ln.get("label") or "").partition(" ")
        if code.startswith(("ac.", "un.", "pr.")):
            ln["label"] = ctx.label(connection, code) + (f" {rest}" if rest else "")


def _registry() -> dict:
    from . import concept, cost, lines, plan, pnl, trend
    return {"pnl": pnl.pnl, "actual_vs_plan": plan.actual_vs_plan, "monthly_trend": trend.monthly_trend,
            "cost_base": cost.cost_base, "concept": concept.concept, "lines_ranking": lines.lines_ranking}


RECIPES: dict[str, Callable] = {}


def _module(name: str | None):
    from . import concept, cost, lines, plan, pnl, trend
    return {"pnl": pnl, "actual_vs_plan": plan, "monthly_trend": trend, "cost_base": cost,
            "concept": concept, "lines_ranking": lines}.get(name or "")


def departures(name: str | None, params: dict) -> list[str]:
    """The finance defaults a recipe's params switch off (D3/D4). Each recipe module's DEFAULTS maps a param to
    (the default value, or a test the value must pass; why it is the default)."""
    rules = getattr(_module(name), "DEFAULTS", {})
    out = []
    for key, (ok, why) in rules.items():
        if key in params and not (ok(params[key]) if callable(ok) else params[key] == ok):
            out.append(f"{name} {key}={params[key]!r} ({why})")
    return out


def layout_departures(layout: dict | None) -> list[str]:
    """A free-form ranking departs from the standard (D22) when it is not the top 10 + an 'Other' line."""
    out = []
    for sec in (layout or {}).get("sections") or []:
        top = ((sec.get("repeat") or {}).get("over") or {}).get("top")
        if top is None:
            continue
        if top != 10:
            out.append(f"a top-{top} ranking (the standard is the top 10 + Other)")
        if not any("other" in str(ln.get("label", "")).lower() for ln in sec.get("lines") or []):
            out.append(f"a top-{top} ranking without an 'Other' line")
    return out


def followups(name: str | None, params: dict, period: str | None, fill: dict | None = None) -> list[str]:
    """At most three natural next questions for a recipe's answer (free-form specs: none). fill: extra
    {placeholders} from the recipe's result; a question whose placeholder is unknown is left out."""
    entity = next((params[k] for k in ("entity", "area", "parent", "concept") if params.get(k)), "the same unit")
    keys = {**(fill or {}), "entity": entity, "period": period or "the same period"}
    out = []
    for f in getattr(_module(name), "FOLLOWUPS", []):
        try:
            out.append(f.format_map(keys))
        except (KeyError, IndexError):
            continue
    return out[:3]


def value_findings(findings: list[dict] | None, values: dict[str, list]) -> list[dict]:
    """A recipe's value findings ({severity, line, col, above, text with {value}}) -> the review's extra findings
    for the values the sheet calculated: only where |value| > above."""
    out = []
    for f in findings or []:
        row = values.get(f["line"]) or []
        v = row[f.get("col", 0)] if f.get("col", 0) < len(row) else None
        if isinstance(v, (int, float)) and abs(v) > f.get("above", 0):
            out.append({"severity": f.get("severity", "Warning"), "text": f["text"].format(value=v)})
    return out


def expand(raw: dict, ctx: RecipeContext) -> tuple[dict, dict]:
    """A spec with a recipe and no layout -> (the full spec, extra tool-result keys). Anything else unchanged."""
    rec = raw.get("recipe")
    if not rec or raw.get("layout"):
        return raw, {}
    if "pnl" not in RECIPES:                     # lazy: the recipe modules import templates/fin_spec (atlas)
        RECIPES.update({k: v for k, v in _registry().items() if k not in RECIPES})
    fn = RECIPES.get(rec.get("name"))
    if fn is None:
        raise ValueError(f"unknown recipe '{rec.get('name')}'; known: {', '.join(sorted(RECIPES))}")
    d = fn(dict(rec.get("params") or {}), ctx)
    have = list(raw.get("assumptions") or [])
    out = {**copy.deepcopy(raw), "pulls": d.pulls, "layout": d.layout,
           "assumptions": have + [a for a in d.assumptions if a not in have]}
    if rec.get("overrides"):
        out = contract.patched({k: v for k, v in out.items() if k != "recipe"}, rec["overrides"]) | {"recipe": rec}
    return out, {**d.result, "warnings": list(d.warnings)}
