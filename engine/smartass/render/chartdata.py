# engine/smartass/render/chartdata.py
"""Chart helper columns (sub-spec 5 §3.3): the stacked waterfall and the variance chart draw from small formula
columns beside their table, under a grey 'Chart helper' header. They are not model cells (no p.addr, no readback):
they only reshape model cells for the chart, and stay live (formulas over the table's cells).

Waterfall. An Excel stacked column stacks a point's positive values up from 0 and its negative values down from 0,
in series order. A bar from the previous running end P to the new one N therefore needs an invisible base above
zero (base), an invisible base below zero (neg_base) and the visible part split at zero (pos, neg), so a bridge that
crosses zero still draws one bar from P to N. Increases and decreases are separate series (green, red); start and
end are totals (navy). waterfall_bars mirrors the formulas in Python (the tests check the maths through it)."""
from __future__ import annotations

WATERFALL_HEADS = ["Run", "Base", "Base <0", "Up", "Up <0", "Down", "Down <0", "Total"]
WATERFALL_ROLES = [None, "base", "base", "up", "up", "down", "down", "total"]   # None: not a series
VARIANCE_HEADS = ["Better", "Worse"]
VARIANCE_ROLES = ["fav", "unfav"]


def waterfall_helpers(rows: list[int], amount: str, cols: list[str]) -> dict[tuple[int, str], object]:
    """{(row, helper column letter): formula or 0} for a bridge whose rows are start, steps..., end and whose
    amounts are in column `amount`; cols are the 8 helper letters (WATERFALL_HEADS order). Each step's running end
    is the previous row's running end plus the step."""
    run, base, nbase, up, upn, down, downn, total = cols
    out: dict[tuple[int, str], object] = {}
    for i, r in enumerate(rows):
        a = f"{amount}{r}"
        if i in (0, len(rows) - 1):                      # start and end: a total bar from zero
            out[(r, run)] = f"={a}"
            for c in (base, nbase, up, upn, down, downn):
                out[(r, c)] = 0
            out[(r, total)] = f"={a}"
            continue
        p, n = f"{run}{rows[i - 1]}", f"{run}{r}"
        lo, hi = f"MIN({p},{n})", f"MAX({p},{n})"
        pos = f"MAX({hi},0)-MAX({lo},0)"
        neg = f"MIN({lo},0)-MIN({hi},0)"
        out[(r, run)] = f"={p}+{a}"
        out[(r, base)] = f"=MAX({lo},0)"
        out[(r, nbase)] = f"=MIN({hi},0)"
        out[(r, up)] = f"=IF({a}>=0,{pos},0)"
        out[(r, upn)] = f"=IF({a}>=0,{neg},0)"
        out[(r, down)] = f"=IF({a}<0,{pos},0)"
        out[(r, downn)] = f"=IF({a}<0,{neg},0)"
        out[(r, total)] = 0
    return out


def waterfall_bars(amounts: list[float]) -> list[list[float]]:
    """The helper values of a bridge (start, steps..., end) by the same rules as waterfall_helpers: one list per
    point in WATERFALL_HEADS order."""
    bars = []
    prev = 0.0
    for i, a in enumerate(amounts):
        if i in (0, len(amounts) - 1):
            bars.append([a, 0, 0, 0, 0, 0, 0, a])
            prev = a
            continue
        run = prev + a
        lo, hi = min(prev, run), max(prev, run)
        pos, neg = max(hi, 0) - max(lo, 0), min(lo, 0) - min(hi, 0)
        bars.append([run, max(lo, 0), min(hi, 0), pos if a >= 0 else 0, neg if a >= 0 else 0,
                     pos if a < 0 else 0, neg if a < 0 else 0, 0])
        prev = run
    return bars


def stacked_spans(bar: list[float], roles: list[str | None] = WATERFALL_ROLES) -> tuple[float, float] | None:
    """How Excel stacks one point of a stacked column: positives up from 0 and negatives down from 0, in series
    order. Returns the (low, high) span of the visible (non-base) series, or None when nothing is visible."""
    up = down = 0.0
    spans = []
    for v, role in zip(bar, roles):
        if role is None:
            continue
        if v >= 0:
            a, b = up, up + v
            up = b
        else:
            a, b = down + v, down
            down = a
        if role != "base" and v != 0:
            spans.append((a, b))
    if not spans:
        return None
    return min(s[0] for s in spans), max(s[1] for s in spans)


def variance_helpers(rows: list[int], var: str, cols: list[str]) -> dict[tuple[int, str], object]:
    """Favourable (Δ >= 0: FinMax signs make a positive Δ better for costs and revenue alike) and unfavourable parts
    of a variance column, one series each so they colour by effect without InvertIfNegative."""
    fav, unfav = cols
    out: dict[tuple[int, str], object] = {}
    for r in rows:
        out[(r, fav)] = f"=MAX({var}{r},0)"
        out[(r, unfav)] = f"=MIN({var}{r},0)"
    return out
