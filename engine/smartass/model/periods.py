"""Months as (year, month) tuples, labelled like Smart View ('Sep 26'), and the relative period expressions a
saved spec uses so it can roll forward to the next closed month."""
from __future__ import annotations

import re

Month = tuple[int, int]
ABBR = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
_LABEL = re.compile(r"([A-Z][a-z]{2}) (\d{2})")


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
