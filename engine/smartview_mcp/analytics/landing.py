"""Year-end landing: YTD actual + the remaining months of the latest forecast; and a run-rate alternative, the YTD
scaled by last year's seasonality (last year's YTD share of its full year, taken on EBITDA before other items - the
caller passes ly_ytd and ly_fy for that line), or YTD / n x 12 when that share is unusable: missing, not positive, or
outside 0.5x-1.5x of n/12 (cost units have a negative EBITDA before other items: a share of two negatives
counts). Both are shown with their gap to target. A forecast STALE_QUARTERS (2) or more quarters behind the closed
month's quarter is still shown, with a warning."""
from __future__ import annotations

import re
from dataclasses import dataclass

from ..templates import InputError
from .facts import Fact, combine, missing
from .frame import Frame
from .measures import flow
from .periods import split_period

STALE_QUARTERS = 2                                    # a forecast this many quarters behind the closed month
SHARE_BAND = (0.5, 1.5)                                # accepted share, as a multiple of n/12


@dataclass(frozen=True)
class Landing:
    forecast: Fact                    # YTD actual + remaining forecast
    run_rate: Fact
    target: Fact
    gap_forecast: Fact                # forecast landing - target
    gap_run_rate: Fact
    seasonal: bool                    # True: the run rate uses last year's YTD share
    method: str
    stale: bool = False               # the forecast is STALE_QUARTERS or more quarters behind
    warning: str | None = None


def remaining_forecast(frame: Frame, entity: str, line: str, scenario: str, months: list[str],
                       id: str | None = None) -> Fact:
    """The forecast summed over the remaining months: 0 when none remain (December closed), missing - never 0 -
    when every remaining month is blank."""
    fact = flow(frame, entity, line, scenario, months, id)
    if fact.value is not None and months and all(frame.values.get((entity, line, m, scenario)) is None
                                                 for m in months):
        return missing(fact.id, fact.unit, fact.definition, f"every remaining {scenario} month is blank",
                       fact.inputs)
    return fact


def _stale(name: str | None, closed: str | None) -> str | None:
    """A warning when the forecast's quarter is STALE_QUARTERS or more behind the closed month's quarter."""
    f = re.search(r"Q(\d)(\d\d)\b", name or "")
    if not f or not closed:
        return None
    try:
        mo, yy = split_period(closed)
    except InputError:
        return None
    if (yy * 4 + (mo - 1) // 3 + 1) - (int(f.group(2)) * 4 + int(f.group(1))) >= STALE_QUARTERS:
        return f"{name} is older than the closed month {closed}"
    return None


def landing(ytd_actual: Fact, remaining_forecast: Fact, target_fy: Fact, ly_ytd: Fact, ly_fy: Fact,
            months_done: int, forecast_name: str | None = None, closed_month: str | None = None,
            id: str = "land") -> Landing:
    src = f"the latest forecast ({forecast_name})" if forecast_name else "the latest forecast"
    fc = combine(f"{id}.forecast", "nok", f"YTD actual + remaining months of {src}",
                 lambda a, b: a + b, ytd_actual, remaining_forecast)
    lo, hi = (x * months_done / 12 for x in SHARE_BAND)
    share, why = None, None
    if ly_ytd.value is None or ly_fy.value is None:
        why = "last year's EBITDA before other items is missing"
    elif ly_fy.value == 0 or ly_ytd.value / ly_fy.value <= 0:
        why = "last year's EBITDA before other items is zero or has opposite signs (YTD against full year)"
    else:
        share = ly_ytd.value / ly_fy.value
        if not lo <= share <= hi:
            why = f"last year's YTD share {share:.1%} is outside {lo:.1%}-{hi:.1%} (0.5x-1.5x of {months_done}/12)"
            share = None
    if share is not None:
        rr = combine(f"{id}.run_rate", "nok", "YTD actual / last year's YTD share of its full year",
                     lambda a, y, f: a / (y / f), ytd_actual, ly_ytd, ly_fy)
        seasonal = True
        method = f"seasonality: last year's YTD was {share:.1%} of its full year (EBITDA before other items)"
    else:
        rr = combine(f"{id}.run_rate", "nok", f"YTD actual / {months_done} x 12",
                     lambda a: a / months_done * 12, ytd_actual)
        seasonal = False
        method = f"plain run rate (YTD / {months_done} x 12): {why}"
    warning = _stale(forecast_name, closed_month)
    if forecast_name:
        method += f"; forecast: {forecast_name}"
    gap_fc = combine(f"{id}.gap_forecast", "nok", "forecast landing - target", lambda a, b: a - b, fc, target_fy)
    gap_rr = combine(f"{id}.gap_run_rate", "nok", "run-rate landing - target", lambda a, b: a - b, rr, target_fy)
    return Landing(fc, rr, target_fy, gap_fc, gap_rr, seasonal, method, warning is not None, warning)
