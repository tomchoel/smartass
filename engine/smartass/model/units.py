# engine/smartass/model/units.py
"""The unit class of a metric, so a chart never mixes units and a trend only totals what adds up: nok (an amount),
count (subscriptions, customers, volume), per_unit (ARPU: NOK per subscription), pct (a share or a ratio), fte, or
other (a FinPlan member whose name says nothing; it gets a chart of its own). A semantic metric says its unit; a
FinMax/FinLite account is NOK; a FinPlan KPI member without a semantic metric is read from its name."""
from __future__ import annotations

import re

NOK, COUNT, PER_UNIT, PCT, FTE, OTHER = "nok", "count", "per_unit", "pct", "fte", "other"
UNIT_WORD = {NOK: "NOK", COUNT: "volume", PER_UNIT: "NOK per unit", PCT: "%", FTE: "FTE", OTHER: ""}
SEMANTIC = {"NOK": NOK, "FTE": FTE, "count": COUNT, "ratio": PCT}
_ARPU = re.compile(r"\barpu\b|\barpa\b", re.I)
_PCT = re.compile(r"%|\bshare\b|\bandel\b|\bpercent\b|\bratio\b|\brate\b", re.I)
_COUNT = re.compile(r"\bsubscriptions?\b|\bsubs\b|\bcustomers?\b|\bvolumes?\b|\busers?\b|\blines\b|\babonnement",
                    re.I)
_FTE = re.compile(r"\bfte\b|årsverk", re.I)
SHORT_MAX = 28


def unit_class(r, token: str, cube: str) -> str:
    """The unit class of a metric token (id, name, alias or raw member) in a cube."""
    find = getattr(r, "metric", None)
    mm = getattr(r, "_member_metric", None)
    m = (find(token) if find is not None else None) or (mm(token) if mm is not None else None)
    if m is not None:
        if _ARPU.search(m.name) or _ARPU.search(token):
            return PER_UNIT
        return SEMANTIC.get(getattr(m, "unit", "NOK"), NOK)
    if token.startswith("ac.") or cube in ("FinMax", "FinLite"):
        return NOK
    for rx, unit in ((_ARPU, PER_UNIT), (_PCT, PCT), (_FTE, FTE), (_COUNT, COUNT)):
        if rx.search(token):
            return unit
    return OTHER


def short_name(name: str, cap: int = SHORT_MAX) -> str:
    """A metric name short enough for a legend, a header, a chart title or a key point: whole leading words dropped
    until it is cap characters or fewer ('Fixed Access Independent Streamix Subscriptions' -> 'Streamix
    Subscriptions'; the last words tell FinPlan members apart). Never an ellipsis; one long word stays whole."""
    words = name.split()
    while len(words) > 1 and len(" ".join(words)) > cap:
        words = words[1:]
    return " ".join(words)
