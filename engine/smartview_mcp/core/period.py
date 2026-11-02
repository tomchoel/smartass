"""Period phrases in a question (English and Norwegian) -> one Period, relative to the last closed month.
Conventions (design §3.1): no period = YTD to the last closed month (stated); 'this year' alone is asked (YTD or full
year); a month after the last closed month is `open` (no actuals yet)."""
from __future__ import annotations

import re
from dataclasses import dataclass

_MONTHS = {
    1: ("jan", "january", "januar"), 2: ("feb", "february", "februar"), 3: ("mar", "march", "mars"),
    4: ("apr", "april"), 5: ("may", "mai"), 6: ("jun", "june", "juni"), 7: ("jul", "july", "juli"),
    8: ("aug", "august"), 9: ("sep", "sept", "september"), 10: ("oct", "okt", "october", "oktober"),
    11: ("nov", "november"), 12: ("dec", "des", "december", "desember"),
}
_ABBR = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
_NAME = {n: m for m, names in _MONTHS.items() for n in names}
# short forms that are also ordinary words ('may', 'mar', 'des') only count as months with a year after them
_NEEDS_YEAR = {"may", "mai", "mar", "des"}
_YEAR = r"(?:\s*['’]?\s*(\d{4}|\d{2})(?!\d))"
_MONTH_RX = re.compile(r"(?<![\w])(" + "|".join(sorted(_NAME, key=len, reverse=True)) + r")\.?" + _YEAR + "?(?![\w])",
                       re.I)
_YTD_RX = re.compile(r"\b(ytd|year[- ]to[- ]date|hittil i år|so far this year|hiå|akkumulert)\b", re.I)
_FY_RX = re.compile(r"\b(full[- ]year|whole year|fy\s*['’]?(\d{2}|\d{4})?|helår|helårs\w*|årsprognose\w*)(?![\w])",
                    re.I)
_Q_RX = re.compile(r"(?<![\w])(?:q([1-4])|([1-4])\.\s*kvartal)" + _YEAR + "?(?![\w])", re.I)
_THIS_YEAR_RX = re.compile(r"\b(this year|i år|current year|inneværende år)\b", re.I)
_BARE_YEAR_RX = re.compile(r"(?<![\w.])(20\d{2})(?![\w])")


@dataclass(frozen=True)
class Period:
    kind: str                       # month | ytd | quarter | fy | year | this_year | none
    label: str | None               # 'Sep 26', 'YTD Sep 26', 'Q3 26', 'FY 26', '2026'; None while it is a gap
    gap: str | None = None          # the question to ask when the period is not clear
    open: bool = False              # (ends in) a month after the last closed month: no actuals yet
    convention: bool = False        # the label comes from a convention, not the question: state it
    end: tuple[int, int] | None = None   # (year, month) the period ends in


def label(year: int, month: int) -> str:
    return f"{_ABBR[month - 1]} {year % 100:02d}"


def parse_month(text: str) -> tuple[int, int]:
    """'Sep 26' / 'September 2026' -> (2026, 9)."""
    m = _MONTH_RX.search(text)
    if not m or not m.group(2):
        raise ValueError(f"not a month: {text!r}")
    return _year(m.group(2)), _NAME[m.group(1).lower()]


def _year(s: str | None, default: int | None = None) -> int:
    if not s:
        return default
    y = int(s)
    return y if y > 99 else 2000 + y


def _months(text: str, closed_year: int) -> list[tuple[int, int]]:
    out = []
    for m in _MONTH_RX.finditer(text):
        name, yr = m.group(1).lower(), m.group(2)
        if name in _NEEDS_YEAR and not yr:
            continue
        out.append((_year(yr, closed_year), _NAME[name]))
    return out


def parse(text: str, closed: str) -> Period:
    cy, cm = parse_month(closed)
    months = _months(text, cy)
    after = lambda ym: ym > (cy, cm)  # noqa: E731

    if _YTD_RX.search(text):
        y, m = months[0] if months else (cy, cm)
        return Period("ytd", "YTD " + label(y, m), open=after((y, m)), convention=not months, end=(y, m))
    fy = _FY_RX.search(text)
    if fy:
        bare = _BARE_YEAR_RX.search(text)
        y = _year(fy.group(2)) or (int(bare.group(1)) if bare else cy)
        return Period("fy", f"FY {y % 100:02d}", open=after((y, 12)), end=(y, 12))
    q = _Q_RX.search(text)
    if q:
        n = int(q.group(1) or q.group(2))
        y = _year(q.group(3), cy)
        return Period("quarter", f"Q{n} {y % 100:02d}", open=after((y, 3 * n)), end=(y, 3 * n))
    if months:
        y, m = months[0]
        return Period("month", label(y, m), open=after((y, m)), end=(y, m))
    if _THIS_YEAR_RX.search(text):
        return Period("this_year", None, gap=f"YTD to {label(cy, cm)}, or full year (actual + forecast)?")
    bare = _BARE_YEAR_RX.search(text)
    if bare:
        y = int(bare.group(1))
        return Period("year", str(y), end=(y, cm if y == cy else 12))
    return Period("none", "YTD " + label(cy, cm), convention=True, end=(cy, cm))
