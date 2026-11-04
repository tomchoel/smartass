# engine/smartass/render/readback.py
"""'Sheet = model': every model cell's value on the sheet equals the model's value. The model's None (missing)
equals Excel's 0 or blank (SUMIFS over no rows is 0; an empty ratio is ""). An Excel error comes back from Value2
as a negative int code and is always a mismatch. Amounts that are both below NEAR_ZERO are equal: float noise of
SUMIFS over large NOK amounts (a bridge 'Other' of 0 read back as -8.8e-07). Fractions (variance %, share; the
plan's `fractions`) keep the strict 1e-6, so a wrong % is never hidden."""
from __future__ import annotations

TOL_ABS = 0.5          # NOK
TOL_REL = 1e-9
NEAR_ZERO = 1e-3      # NOK: an amount this small is zero
EXCEL_ERRORS = {-2146826281, -2146826246, -2146826259, -2146826288, -2146826265, -2146826273, -2146826252}


def _pos(a: str) -> tuple[int, int]:
    col = "".join(ch for ch in a if ch.isalpha())
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return int(a[len(col):]), n


def compare(addr: dict[str, str], cells, sheet: tuple, fractions=frozenset()) -> list[str]:
    """sheet: Value2 of the block from A1; fractions: the keys of percentage cells (every other cell is an amount).
    Returns 'C4: sheet x, model y' per mismatch."""
    out = []
    for key, a in addr.items():
        r, c = _pos(a)
        try:
            s = sheet[r - 1][c - 1]
        except IndexError:
            s = None
        m = cells.get(key)
        if isinstance(s, int) and s in EXCEL_ERRORS:
            out.append(f"{a}: sheet shows an Excel error, model {m}")
            continue
        sv = 0.0 if s in (None, "") else s
        mv = 0.0 if m is None else m
        if not isinstance(sv, (int, float)):
            out.append(f"{a}: sheet {sv!r}, model {mv}")
            continue
        if key not in fractions and abs(sv) < NEAR_ZERO and abs(mv) < NEAR_ZERO:
            continue
        # amounts in NOK: half a krone; fractions (variance %, share): 1e-6, so a wrong % is never hidden
        tol = TOL_ABS if abs(mv) >= 1 or abs(sv) >= 1 else 1e-6
        if abs(sv - mv) > max(tol, TOL_REL * abs(mv)):
            out.append(f"{a}: sheet {sv:g}, model {mv:g}")
    return out
