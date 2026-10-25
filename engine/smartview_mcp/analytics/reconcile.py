"""Reconciliation of warehouse values (pure): the catalogue's identity checks per entity and month, the sum of the
Periodic months against Accumulated, and the areas against the company total (information only: eliminations and
units outside the areas exist)."""
from __future__ import annotations

import re

from .frame import Frame

TOL = 1.0                                              # NOK


def _eval(expr: str, vals: dict[str, float]) -> float | None:
    out = 0.0
    for sign, key in re.findall(r"([+-]?)\s*([a-z_]\w*)", expr):
        if key not in vals:
            return None
        out += -vals[key] if sign == "-" else vals[key]
    return out


def _n(v: float) -> str:
    return f"{v:,.0f}".replace(",", " ")


def identities(frame: Frame, checks, tol: float = TOL) -> tuple[int, list[str]]:
    """(checks evaluated, problems) for every (entity, period, scenario) of the frame that has the check's lines."""
    groups: dict[tuple[str, str, str], dict[str, float]] = {}
    for (ent, line, period, scen), v in frame.values.items():
        groups.setdefault((ent, period, scen), {})[line] = 0.0 if v is None else v
    n, problems = 0, []
    for (ent, period, scen), vals in sorted(groups.items()):
        for c in checks:
            d = _eval(c.expr, vals)
            if d is None:
                continue
            n += 1
            if abs(d) >= tol:
                problems.append(f"{frame.cube} {ent} {period} {scen}: {c.label} is off by {_n(d)}")
    return n, problems


def months_vs_ytd(periodic: dict, accumulated: dict, months: list[str], tol: float = TOL) -> tuple[int, list[str]]:
    """Accumulated at months[-1] = the sum of the Periodic months, per (entity, account). Both are warehouse rows
    {key: value}; a pair without every month is not checked."""
    sums: dict[tuple[str, str], list[float]] = {}
    for k, v in periodic.items():
        if k[5] in months:
            sums.setdefault((k[3], k[4]), []).append(0.0 if v is None else v)
    n, problems = 0, []
    for k, v in sorted(accumulated.items()):
        parts = sums.get((k[3], k[4]))
        if k[5] != months[-1] or parts is None or len(parts) != len(months):
            continue
        n += 1
        d = (0.0 if v is None else v) - sum(parts)
        if abs(d) >= tol:
            problems.append(f"{k[0]} {k[3]} {k[4]}: Accumulated {months[-1]} differs from the sum of "
                            f"{months[0]}..{months[-1]} by {_n(d)}")
    return n, problems


def areas_vs_total(frame: Frame, areas: list[str], company: str, line: str, period: str,
                   scenario: str) -> str | None:
    """'The areas against the company total' as an information line, or None when a value is missing."""
    vals = [frame.get(a, line, period, scenario) for a in [*areas, company]]
    if any(v is None for v in vals):
        return None
    diff = vals[-1][0] - sum(v[0] for v in vals[:-1])
    return (f"{frame.cube} {line} {period} {scenario}: the company total less the {len(areas)} areas is {_n(diff)} "
            "(units outside the areas and eliminations)")
