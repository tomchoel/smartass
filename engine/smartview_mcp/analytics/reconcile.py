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
    """Accumulated at months[-1] = the sum of the Periodic months, per (cube, scenario, version, entity, account).
    Both are warehouse rows {key: value}; a series without every month is not checked."""
    series: dict[tuple, dict[str, float]] = {}
    for k, v in periodic.items():
        if k[5] in months:
            series.setdefault(k[:5], {})[k[5]] = 0.0 if v is None else v
    n, problems = 0, []
    for k, v in sorted(accumulated.items()):
        got = series.get(k[:5])
        if not months or k[5] != months[-1] or got is None or any(m not in got for m in months):
            continue
        n += 1
        d = (0.0 if v is None else v) - sum(got[m] for m in months)
        if abs(d) >= tol:
            who = " ".join(x for x in (k[0], k[1], k[2], k[3], k[4]) if x)
            problems.append(f"{who}: Accumulated {months[-1]} differs from the sum of "
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
