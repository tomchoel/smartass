"""Year-end landing: YTD actual + the remaining months of the latest forecast; and a run-rate alternative, the YTD
scaled by last year's seasonality (last year's YTD share of its full year), or YTD / n x 12 when last year gives no
usable share. Both are shown with their gap to target."""
from __future__ import annotations

from dataclasses import dataclass

from .facts import Fact, combine


@dataclass(frozen=True)
class Landing:
    forecast: Fact                    # YTD actual + remaining forecast
    run_rate: Fact
    target: Fact
    gap_forecast: Fact                # forecast landing - target
    gap_run_rate: Fact
    seasonal: bool                    # True: the run rate uses last year's YTD share
    method: str


def landing(ytd_actual: Fact, remaining_forecast: Fact, target_fy: Fact, ly_ytd: Fact, ly_fy: Fact,
            months_done: int, id: str = "land") -> Landing:
    fc = combine(f"{id}.forecast", "nok", "YTD actual + remaining months of the latest forecast",
                 lambda a, b: a + b, ytd_actual, remaining_forecast)
    share = ly_ytd.value / ly_fy.value if ly_ytd.value is not None and ly_fy.value else None
    if share is not None and share > 0:
        rr = combine(f"{id}.run_rate", "nok", "YTD actual / last year's YTD share of its full year",
                     lambda a, y, f: a / (y / f), ytd_actual, ly_ytd, ly_fy)
        seasonal, method = True, f"seasonality: last year's YTD was {share:.1%} of its full year"
    else:
        rr = combine(f"{id}.run_rate", "nok", f"YTD actual / {months_done} x 12",
                     lambda a: a / months_done * 12, ytd_actual)
        seasonal = False
        method = f"plain run rate (YTD / {months_done} x 12): last year gives no usable seasonality share"
    gap_fc = combine(f"{id}.gap_forecast", "nok", "forecast landing - target", lambda a, b: a - b, fc, target_fy)
    gap_rr = combine(f"{id}.gap_run_rate", "nok", "run-rate landing - target", lambda a, b: a - b, rr, target_fy)
    return Landing(fc, rr, target_fy, gap_fc, gap_rr, seasonal, method)
