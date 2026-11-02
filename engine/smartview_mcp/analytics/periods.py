"""Period arithmetic on FinMax month names ('Sep 26'), reusing templates' helpers. FinPlan months are stored the same
way in the fact store ('Sep 26' = Period Sep of Years FY26)."""
from __future__ import annotations

from ..templates import MONTHS, months_to, same_period_ly, split_period


def label(month: int, yy: int) -> str:
    return f"{MONTHS[month - 1]} {yy:02d}"


def shift(p: str, n: int) -> str:
    """The month n months after p (before it for negative n)."""
    mo, yy = split_period(p)
    i = yy * 12 + mo - 1 + n
    return label(i % 12 + 1, i // 12)


def index(p: str) -> int:
    mo, yy = split_period(p)
    return yy * 12 + mo


def after(a: str, b: str) -> bool:
    return index(a) > index(b)


def ytd(p: str) -> list[str]:
    return months_to(p)


def ly(p: str) -> str:
    return same_period_ly(p)


def year(yy: int) -> list[str]:
    return [label(m, yy) for m in range(1, 13)]


def remaining(p: str) -> list[str]:
    """The months after p in its year (none after December)."""
    mo, yy = split_period(p)
    return [label(m, yy) for m in range(mo + 1, 13)]


def opening(p: str) -> str:
    """The previous December: the opening level of a stock (FTE) for p's year."""
    return label(12, split_period(p)[1] - 1)


def window(closed: str, n: int = 36) -> list[str]:
    """The n months ending with the closed month, oldest first."""
    return [shift(closed, -k) for k in range(n - 1, -1, -1)]


def between(a: str, b: str) -> list[str]:
    """The months after a up to and including b (empty if b is not after a)."""
    return [shift(a, k) for k in range(1, index(b) - index(a) + 1)]


def finplan(p: str) -> tuple[str, str]:
    """'Sep 26' -> ('Sep', 'FY26'): FinPlan's Period and Years members."""
    mo, yy = split_period(p)
    return MONTHS[mo - 1], f"FY{yy:02d}"


def from_finplan(period: str, years: str) -> str:
    """('Sep', 'FY26') -> 'Sep 26'."""
    return f"{period} {years[2:]}"
