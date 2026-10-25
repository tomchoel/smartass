"""Readable formula builders (skill tn-formelstil) and the formula checker that runs on every built sheet.

Rules, agreed with the user 2026-10-15:
- sums: SUM over ranges (=SUM(C13:C21), =SUM(C10,C20:C25)); a plain '+' only for exactly two terms;
- variance %: =IF(D13=0,"",(C13-D13)/ABS(D13)) - blank on a zero base, never 0;
- lookups into Smart View data sheets: label-keyed XLOOKUP (see userexcel.lookup_formula).
"""
from __future__ import annotations

import logging
import re
from pathlib import Path

from openpyxl.utils import get_column_letter as L

log = logging.getLogger("smartview.formulas")

CHECKER_PATH = Path(__file__).resolve().parent / "vendor" / "formelsjekk.py"   # pinned copy of tn-formelstil's checker
_checker = None


def _runs(nums: list[int]) -> list[tuple[int, int]]:
    """[3,4,5,9,11,12] -> [(3,5),(9,9),(11,12)], keeping the given order of runs."""
    runs: list[list[int]] = []
    for n in nums:
        if runs and n == runs[-1][1] + 1:
            runs[-1][1] = n
        else:
            runs.append([n, n])
    return [(a, b) for a, b in runs]


def sum_rows(col: str, rows: list[int]) -> str:
    """Sum of one column over the given rows: '=C11+C12' (two terms) or '=SUM(C13:C21)' / '=SUM(C10,C20:C25)'."""
    if not rows:
        return "=0"
    if len(rows) == 1:
        return f"={col}{rows[0]}"
    if len(rows) == 2:                                     # Revenue + COGS: two terms read best as +
        return f"={col}{rows[0]}+{col}{rows[1]}"
    parts = [f"{col}{a}" if a == b else f"{col}{a}:{col}{b}" for a, b in _runs(sorted(rows))]
    return f"=SUM({','.join(parts)})"


def avg_rows(col: str, rows: list[int]) -> str:
    """Average of one column over the given rows: '=AVERAGE(C13:C21)' / '=AVERAGE(C10,C20:C25)' (e.g. the FTE
    average: the opening level and each month end)."""
    if not rows:
        raise ValueError("an average needs at least one line")
    parts = [f"{col}{a}" if a == b else f"{col}{a}:{col}{b}" for a, b in _runs(sorted(rows))]
    return f"=AVERAGE({','.join(parts)})"


def sum_cols(row: int, cols: list[int]) -> str:
    """Sum of one row over the given column numbers (1-based)."""
    if not cols:
        return "=0"
    if len(cols) == 1:
        return f"={L(cols[0])}{row}"
    if len(cols) == 2:
        return f"={L(cols[0])}{row}+{L(cols[1])}{row}"
    parts = [f"{L(a)}{row}" if a == b else f"{L(a)}{row}:{L(b)}{row}" for a, b in _runs(sorted(cols))]
    return f"=SUM({','.join(parts)})"


def variance(a: str, b: str) -> str:
    return f"={a}-{b}"


def variance_pct(a: str, b: str) -> str:
    """Positive = better also for costs held as negatives; blank when the comparison is 0 or empty."""
    return f'=IF({b}=0,"",({a}-{b})/ABS({b}))'


_PLUS_ONLY = re.compile(r"\s*[A-Za-z_][A-Za-z0-9_]*(\s*\+\s*[A-Za-z_][A-Za-z0-9_]*){2,}\s*")


def expr_keys_if_plain_sum(expr: str) -> list[str] | None:
    """'a+b+c' (three or more keys, only '+') -> ['a','b','c']: written as SUM instead of a + chain."""
    if _PLUS_ONLY.fullmatch(expr):
        return [k.strip() for k in expr.split("+")]
    return None


# ------------------------------------------------------------------ checker
def _load_checker():
    global _checker
    if _checker is None:
        from .vendor import formelsjekk
        _checker = formelsjekk
    return _checker


def check(cells: list[tuple[str, str, str]]) -> dict:
    """Run tn-formelstil's checker over (sheet, A1 address, formula) triples. Never raises: a broken checker is
    reported, it must not break an analysis."""
    try:
        mod = _load_checker()
        findings = mod.check_cells(cells)
    except Exception as e:   # noqa: BLE001
        log.exception("formula checker failed")
        return {"ok": None, "error": f"formula checker failed: {e}", "block": [], "warn": []}
    short = lambda t: t if len(t) <= 140 else t[:137] + "..."
    fmt = lambda f: {"cell": f"{f.sheet}!{f.cell}", "rule": f.rule, "message": f.message, "formula": short(f.formula),
                     **({"suggestion": short(f.suggestion)} if f.suggestion else {})}
    counts: dict[str, int] = {}
    for f in findings:
        counts[f"{f.rule} ({f.severity})"] = counts.get(f"{f.rule} ({f.severity})", 0) + 1
    block = [f for f in findings if f.severity == "block"]
    warn = [f for f in findings if f.severity != "block"]

    def sample(fs, n):                       # a few examples per rule, not every cell of a copied-across row
        out, per = [], {}
        for f in fs:
            if per.get(f.rule, 0) < n:
                per[f.rule] = per.get(f.rule, 0) + 1
                out.append(fmt(f))
        return out
    return {"ok": not block, "checked": len(cells), "counts": counts,
            "block_cells": sorted({f"{f.sheet}!{f.cell}" for f in block})[:60],
            "block": sample(block, 3), "warn": sample(warn, 2)}


def sheet_formulas(ws, com) -> list[tuple[str, str, str]]:
    """All formulas on a worksheet as (sheet, address, Formula2), read in one block over COM."""
    used = com(lambda: ws.UsedRange)
    r0, c0 = com(lambda: used.Row), com(lambda: used.Column)
    data = com(lambda: used.Formula2)
    if not isinstance(data, tuple):
        data = ((data,),)
    name = com(lambda: ws.Name)
    out = []
    for i, row in enumerate(data):
        for j, v in enumerate(row):
            if isinstance(v, str) and v.startswith("="):
                out.append((name, f"{L(c0 + j)}{r0 + i}", v))
    return out


__all__ = ["sum_rows", "avg_rows", "sum_cols", "variance", "variance_pct", "expr_keys_if_plain_sum", "check", "sheet_formulas"]
