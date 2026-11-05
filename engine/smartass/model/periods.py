"""Months as (year, month) tuples, labelled like Smart View ('Sep 26'), and the relative period expressions a
saved spec uses so it can roll forward to the next closed month."""
from __future__ import annotations

import re

from .errors import GuardError

Month = tuple[int, int]
ABBR = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
_LABEL = re.compile(r"([A-Z][a-z]{2}) (\d{2})")


Closed = Month | dict[str, Month]   # one closed month, or one per cube


def closed_for(closed: Closed, cube: str) -> Month:
    if isinstance(closed, dict):
        if cube not in closed:
            raise GuardError(f"no closed month for cube {cube}", "pass the closed month for every cube used")
        return closed[cube]
    return closed


def label(m: Month) -> str:
    return f"{ABBR[m[1] - 1]} {m[0] % 100:02d}"


def parse(text: str) -> Month:
    hit = _LABEL.fullmatch(text.strip())
    if not hit or hit.group(1) not in ABBR:
        raise ValueError(f"not a month label: {text!r} (use 'Sep 26')")
    return (2000 + int(hit.group(2)), ABBR.index(hit.group(1)) + 1)


def shift(m: Month, n: int) -> Month:
    i = m[0] * 12 + (m[1] - 1) + n
    return (i // 12, i % 12 + 1)


def span(a: Month, b: Month) -> list[Month]:
    if b < a:
        raise ValueError(f"period ends before it starts: {label(a)}..{label(b)}")
    out, cur = [], a
    while cur <= b:
        out.append(cur)
        cur = shift(cur, 1)
    return out


def resolve(expr: str, closed: Month) -> list[Month]:
    """'closed', 'closed-ytd', 'closed-12m', 'fy' (the closed month's year), 'Sep 26' or 'Jan 26..Sep 26'."""
    e = expr.strip()
    if e == "closed":
        return [closed]
    if e == "closed-ytd":
        return span((closed[0], 1), closed)
    if e == "closed-12m":
        return span(shift(closed, -11), closed)
    if e == "fy":
        return span((closed[0], 1), (closed[0], 12))
    if ".." in e:
        a, b = e.split("..", 1)
        return span(parse(a), parse(b))
    return [parse(e)]


def period_label(months: list[Month]) -> str:
    """A header label for a set of months: 'Oct 26' for one month, 'YTD Oct 26' for Jan..m of one year, else the
    span 'Feb–Sep 26' or, across years, 'Nov 25–Oct 26'."""
    ms = sorted(months)
    a, b = ms[0], ms[-1]
    if a == b:
        return label(a)
    if a[0] == b[0]:
        return f"YTD {label(b)}" if a[1] == 1 else f"{ABBR[a[1] - 1]}–{label(b)}"
    return f"{label(a)}–{label(b)}"
