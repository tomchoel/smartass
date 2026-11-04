# engine/smartass/render/charts.py
"""Charts for v2 answers, in v1's style (answer_write._chart / _multi_chart): a line chart per trend block (one
series per value column, months from column B), a waterfall per bridge block (column B labels + column C amounts,
start and end as totals; falls back to a column chart where Excel has no waterfall). Every chart is named
'smartass chart' (v1's owned.CHART_NAME) so a follow-up can drop it. A chart that fails is a warning, never a
failed answer."""
from __future__ import annotations

from ..xl.owned import CHART_NAME
from ..xl.jobs import com

from .excel import col_number
from .layout import ChartSpec, SheetPlan

XL_LINE = 4
XL_WATERFALL = 119
XL_COLUMN = 51
WIDTH = 480.0


def _place(ws, spec: ChartSpec, xl_type: int):
    cell = ws.Range(f"{spec.anchor_col}{spec.anchor_row or spec.header_row}")
    left, top = com(lambda: cell.Left), com(lambda: cell.Top)
    shape = com(lambda: ws.Shapes.AddChart2(-1, xl_type, left, top, WIDTH, spec.height_pt))
    com(lambda: setattr(shape, "Name", CHART_NAME))
    return shape, com(lambda: shape.Chart)


def _title(chart, text: str) -> None:
    com(lambda: setattr(chart, "HasTitle", True))
    com(lambda: setattr(chart.ChartTitle, "Text", text))


def _millions(chart) -> None:
    """Value axis in MNOK via the axis display unit (v1: format codes with ',,' garble in nb-NO)."""
    try:
        ax = com(lambda: chart.Axes(2))
        com(lambda: setattr(ax, "DisplayUnit", -6))          # xlMillions
        com(lambda: setattr(ax, "HasDisplayUnitLabel", True))
        com(lambda: setattr(ax.DisplayUnitLabel, "Text", "MNOK"))
        com(lambda: setattr(ax.TickLabels, "NumberFormat", "0"))
    except Exception:   # noqa: BLE001 - cosmetic
        pass


def _drop_auto_series(chart) -> None:
    """AddChart2 guesses a series from the selection; drop it, as v1 answer_write does."""
    for _ in range(50):
        if not com(lambda: chart.SeriesCollection().Count):
            return
        com(lambda: chart.SeriesCollection(1).Delete())


def _line(ws, spec: ChartSpec) -> None:
    shape, chart = _place(ws, spec, XL_LINE)
    _drop_auto_series(chart)
    r1, r2 = spec.rows[0], spec.rows[-1]
    for col in spec.series_cols:
        s = com(lambda: chart.SeriesCollection().NewSeries())
        com(lambda: setattr(s, "Values", ws.Range(f"{col}{r1}:{col}{r2}")))
        com(lambda: setattr(s, "XValues", ws.Range(f"B{r1}:B{r2}")))
        head = com(lambda: ws.Cells(spec.header_row, col_number(col)).Value)   # series named by its header, as v1
        com(lambda: setattr(s, "Name", str(head or "")))
    _title(chart, spec.title)
    com(lambda: setattr(chart, "HasLegend", len(spec.series_cols) > 1))
    if spec.millions:
        _millions(chart)


def _waterfall(ws, spec: ChartSpec) -> None:
    r1, r2 = spec.rows[0], spec.rows[-1]
    col = spec.series_cols[0]
    shape = None
    try:
        shape, chart = _place(ws, spec, XL_WATERFALL)
        com(lambda: chart.SetSourceData(ws.Range(f"B{r1}:{col}{r2}")))
        series = com(lambda: chart.FullSeriesCollection(1))
        for idx in spec.totals:
            try:                                           # one point failing keeps the waterfall (as v1)
                com(lambda: setattr(series.Points(idx), "IsTotal", True))
            except Exception:   # noqa: BLE001
                pass
    except Exception:   # noqa: BLE001 - no waterfall in this Excel: a column chart says the same
        if shape is not None:
            try:
                com(lambda: shape.Delete())                    # no half-drawn waterfall left behind (as v1)
            except Exception:   # noqa: BLE001
                pass
        shape, chart = _place(ws, spec, XL_COLUMN)
        com(lambda: chart.SetSourceData(ws.Range(f"B{r1}:{col}{r2}")))
    _title(chart, spec.title)
    com(lambda: setattr(chart, "HasLegend", False))
    if spec.millions:
        _millions(chart)


def write_charts(ws, plan: SheetPlan) -> list[str]:
    warnings = []
    for spec in plan.charts:
        try:
            (_waterfall if spec.type == "waterfall" else _line)(ws, spec)
        except Exception as e:   # noqa: BLE001 - a chart never fails the answer
            warnings.append(f"chart '{spec.title}' not drawn: {e}")
    return warnings
