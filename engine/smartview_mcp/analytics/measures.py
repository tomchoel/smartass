"""Measures over a Frame: a flow sums its months; a stock (FTE) takes the period-end level; the average of a stock is
the opening level (the previous December) plus each month end - n+1 points (owner decision 2026-10-25)."""
from __future__ import annotations

import math

from . import periods
from .facts import Fact, missing
from .frame import Frame


def _unit(frame: Frame, line: str) -> str:
    return "fte" if frame.line(line).measure == "stock" else "nok"


def _span(months: list[str]) -> str:
    if not months:
        return "none"
    return months[0] if len(months) == 1 else f"{months[0]}..{months[-1]}"


def check_contiguous(months: list[str]) -> None:
    for a, b in zip(months, months[1:]):
        if periods.shift(a, 1) != b:
            raise ValueError(f"months must be contiguous and ordered: {a} is followed by {b}")


def _values(frame: Frame, entity: str, line: str, scenario: str, months: list[str], id: str, definition: str):
    """((values, fact keys), None), or (None, a missing Fact naming the first cell never read)."""
    vals, keys = [], []
    for m in months:
        got = frame.get(entity, line, m, scenario)
        if got is None:
            return None, missing(id, _unit(frame, line), definition,
                                 f"no value read for {entity} {line} {m} ({scenario})", keys)
        if not math.isfinite(got[0]):
            return None, missing(id, _unit(frame, line), definition,
                                 f"non-finite value for {entity} {line} {m} ({scenario})", keys)
        vals.append(got[0])
        keys.append(got[1])
    return (vals, tuple(keys)), None


def flow(frame: Frame, entity: str, line: str, scenario: str, months: list[str], id: str | None = None) -> Fact:
    """The sum of the months (0 for no months, e.g. the remaining months after December)."""
    id = id or f"{scenario}.{entity}.{line}.{_span(months)}"
    definition = f"{frame.cube} {line} {scenario} {entity}, sum of {_span(months)}"
    got, gone = _values(frame, entity, line, scenario, months, id, definition)
    if gone:
        return gone
    vals, keys = got
    return Fact(id, float(sum(vals)), _unit(frame, line), definition, keys)


def level(frame: Frame, entity: str, line: str, scenario: str, period: str, id: str | None = None) -> Fact:
    id = id or f"{scenario}.{entity}.{line}.{period}"
    definition = f"{frame.cube} {line} {scenario} {entity}, level at the end of {period}"
    got, gone = _values(frame, entity, line, scenario, [period], id, definition)
    if gone:
        return gone
    vals, keys = got
    return Fact(id, vals[0], _unit(frame, line), definition, keys)


def average_level(frame: Frame, entity: str, line: str, scenario: str, months: list[str],
                  id: str | None = None) -> Fact:
    """The average of the opening level (the month before months[0]: the previous December for a YTD window) and
    each month end: n+1 points. A blank point counts as 0 and stays a point."""
    id = id or f"{scenario}.{entity}.{line}.avg.{_span(months)}"
    if not months:
        return missing(id, _unit(frame, line), f"{frame.cube} {line} {scenario} {entity}, average level",
                       "no months")
    check_contiguous(list(months))
    points = [periods.shift(months[0], -1)] + list(months)
    definition = (f"{frame.cube} {line} {scenario} {entity}, average of {len(points)} levels "
                  f"({points[0]} opening + {len(months)} month ends)")
    got, gone = _values(frame, entity, line, scenario, points, id, definition)
    if gone:
        return gone
    vals, keys = got
    return Fact(id, sum(vals) / len(vals), _unit(frame, line), definition, keys)


def amount(frame: Frame, entity: str, line: str, scenario: str, months: list[str], id: str | None = None) -> Fact:
    """A line over a period by its measure: flow = the sum of the months, stock = the level at the last month."""
    check_contiguous(list(months))
    if frame.line(line).measure == "stock":
        if not months:
            return missing(id or f"{scenario}.{entity}.{line}.none", "fte", f"{frame.cube} {line} level", "no months")
        return level(frame, entity, line, scenario, months[-1], id)
    return flow(frame, entity, line, scenario, months, id)
