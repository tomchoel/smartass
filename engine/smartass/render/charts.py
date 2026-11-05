# engine/smartass/render/charts.py
"""IBCS charts for v2 answers (sub-spec 5 §3.3), one per ChartSpec at its anchor in the Charts zone. Each chart is
added empty-handed (the series AddChart2 guesses are dropped), gets its series explicitly (NewSeries, Values,
XValues), and only THEN its formats, because a type change resets formatting.

- Scenarios: AC solid navy, PY solid grey, BU outlined (white, navy border), FC hatched; variances green or red by
  effect; highlighted points in the TN accent, the rest grey.
- Labels instead of axes: data labels (MNOK '#,##0.0,,'), no gridlines, no tick marks, no value axis labels; a legend
  only with more than 3 visible series (on top).
- Waterfall: a stacked column over the chart helper columns (render/chartdata.py) with invisible bases; never
  Excel's native waterfall (its COM model does not reliably set totals). Variance: favourable and unfavourable
  series with 100 % overlap (no InvertIfNegative). Bar: ReversePlotOrder keeps the first (largest) row on top.
Every chart is named 'smartass chart' (owned.CHART_NAME) so a follow-up can drop it. A chart that fails is one
lint line with its fix; the others are still drawn."""
from __future__ import annotations

from ..xl.jobs import com
from ..xl.owned import CHART_NAME
from ..xl.styling import GREEN
from ..xl.vendor import tn_rapportstil as tn
from .layout import ChartSpec, SheetPlan

XL_TYPE = {"column": 51, "bar": 57, "line": 4, "stacked": 52, "stacked100": 53, "combo": 51, "variance": 57,
           "waterfall": 52}
XL_LINE = 4
XL_NONE = -4142                  # xlNone / xlTickMarkNone / xlTickLabelPositionNone
XL_LEGEND_TOP = -4160
XL_VALUE, XL_CATEGORY = 2, 1
MSO_TRUE, MSO_FALSE = -1, 0
MSO_PATTERN_WIDE_UPWARD_DIAGONAL = 26      # MsoPatternType (Office typelib)
GAP_WIDTH = 50

NAVY = tn.rgb(tn.NAVY)
GREY = tn.rgb((166, 166, 166))             # PY and the not-highlighted points
WHITE = tn.rgb((255, 255, 255))
GOOD = tn.rgb(GREEN)
BAD = tn.rgb(tn.BREACH_FONT)
ACCENT = tn.rgb((0, 160, 220))             # TN accent (not in the vendored palette): highlighted points
PCT_LINE = tn.rgb(tn.RULE_DARK)
PARTS = [tn.rgb(c) for c in (tn.NAVY, (84, 110, 150), (140, 160, 190), (190, 200, 215), (166, 166, 166))]

FILL = {"ac": NAVY, "py": GREY, "var": GOOD, "fav": GOOD, "unfav": BAD, "up": GOOD, "down": BAD, "total": NAVY}
HIDDEN = ("base",)               # waterfall bases: no fill, no line, no label


def label_format(role: str, millions: bool) -> str:
    """Data label format: MNOK with the scaling commas, zeros hidden (helper series are 0 off their points)."""
    if role == "pct":
        return '0.0 %;-0.0 %;;'
    unit = "#,##0.0,," if millions else "#,##0.0"
    if role in ("down", "unfav"):          # the visible part of a decrease is positive in the helper column
        return f"-{unit};{unit};;" if role == "down" else f"{unit};-{unit};;"
    return f"{unit};-{unit};;"


def _place(ws, spec: ChartSpec):
    cell = ws.Range(spec.anchor)
    left, top = com(lambda: cell.Left), com(lambda: cell.Top)
    shape = com(lambda: ws.Shapes.AddChart2(-1, XL_TYPE[spec.kind], left, top, spec.width, spec.height))
    com(lambda: setattr(shape, "Name", CHART_NAME))
    return shape, com(lambda: shape.Chart)


def _drop_auto_series(chart) -> None:
    """AddChart2 guesses a series from the selection; drop it, as v1 answer_write does."""
    for _ in range(50):
        if not com(lambda: chart.SeriesCollection().Count):
            return
        com(lambda: chart.SeriesCollection(1).Delete())


def _add_series(ws, chart, spec: ChartSpec) -> list:
    out = []
    for name, rng, role in spec.series:
        s = com(lambda: chart.SeriesCollection().NewSeries())
        com(lambda: setattr(s, "Values", ws.Range(rng)))
        com(lambda: setattr(s, "XValues", ws.Range(spec.cat_range)))
        com(lambda: setattr(s, "Name", name))
        if spec.kind == "combo" and role == "pct":      # type and axis before any format
            com(lambda: setattr(s, "ChartType", XL_LINE))
            com(lambda: setattr(s, "AxisGroup", 2))
        out.append((s, role))
    return out


def _fill(s, role: str, k: int, line: bool) -> None:
    if line:                                             # a line series: the line is the scenario colour
        colour = PCT_LINE if role == "pct" else FILL.get(role, NAVY)
        com(lambda: setattr(s.Format.Line, "Visible", MSO_TRUE))
        com(lambda: setattr(s.Format.Line.ForeColor, "RGB", colour))
        com(lambda: setattr(s.Format.Line, "Weight", 2.25))
        return
    f, ln = s.Format.Fill, s.Format.Line
    if role in HIDDEN:
        com(lambda: setattr(f, "Visible", MSO_FALSE))
        com(lambda: setattr(ln, "Visible", MSO_FALSE))
        return
    com(lambda: setattr(f, "Visible", MSO_TRUE))
    if role == "bu":                                     # outlined: white with a navy border
        com(lambda: f.Solid())
        com(lambda: setattr(f.ForeColor, "RGB", WHITE))
        com(lambda: setattr(ln, "Visible", MSO_TRUE))
        com(lambda: setattr(ln.ForeColor, "RGB", NAVY))
        com(lambda: setattr(ln, "Weight", 1.0))
    elif role == "fc":                                   # hatched
        com(lambda: f.Patterned(MSO_PATTERN_WIDE_UPWARD_DIAGONAL))
        com(lambda: setattr(f.ForeColor, "RGB", NAVY))
        com(lambda: setattr(f.BackColor, "RGB", WHITE))
        com(lambda: setattr(ln, "Visible", MSO_TRUE))
        com(lambda: setattr(ln.ForeColor, "RGB", NAVY))
    else:
        colour = PARTS[k % len(PARTS)] if role == "part" else FILL.get(role, NAVY)
        com(lambda: f.Solid())
        com(lambda: setattr(f.ForeColor, "RGB", colour))
        com(lambda: setattr(ln, "Visible", MSO_FALSE))


def _labels(s, role: str, millions: bool) -> None:
    if role in HIDDEN:
        return
    com(lambda: setattr(s, "HasDataLabels", True))
    dl = com(lambda: s.DataLabels())
    com(lambda: setattr(dl, "NumberFormatLinked", False))
    com(lambda: setattr(dl, "NumberFormat", label_format(role, millions)))


def _highlight(s, points: list[int], n: int) -> None:
    """The highlighted points in the accent, every other point grey."""
    for i in range(1, n + 1):
        colour = ACCENT if i in points else GREY
        pt = com(lambda: s.Points(i))
        com(lambda: setattr(pt.Format.Fill.ForeColor, "RGB", colour))


def _axes(chart, spec: ChartSpec) -> None:
    """No gridlines, no tick marks, no value axis labels (the data labels carry the numbers)."""
    val = com(lambda: chart.Axes(XL_VALUE))
    com(lambda: setattr(val, "HasMajorGridlines", False))
    com(lambda: setattr(val, "HasMinorGridlines", False))
    com(lambda: setattr(val, "MajorTickMark", XL_NONE))
    com(lambda: setattr(val, "TickLabelPosition", XL_NONE))
    com(lambda: setattr(val.Format.Line, "Visible", MSO_FALSE))
    cat = com(lambda: chart.Axes(XL_CATEGORY))
    com(lambda: setattr(cat, "MajorTickMark", XL_NONE))
    if spec.kind in ("bar", "variance"):
        com(lambda: setattr(cat, "ReversePlotOrder", spec.sort != "asc"))


def _title(chart, text: str) -> None:
    st = tn.style_defs()["TN Seksjon"]                   # the TN section title: Arial bold navy
    com(lambda: setattr(chart, "HasTitle", True))
    com(lambda: setattr(chart.ChartTitle, "Text", text))
    font = com(lambda: chart.ChartTitle.Font)
    com(lambda: setattr(font, "Name", tn.FONT))
    com(lambda: setattr(font, "Size", st["size"] + 1))
    com(lambda: setattr(font, "Bold", True))
    com(lambda: setattr(font, "Color", tn.rgb(st["color"])))


def _draw(ws, spec: ChartSpec) -> None:
    shape, chart = _place(ws, spec)
    try:
        _drop_auto_series(chart)
        series = _add_series(ws, chart, spec)
        # formats only after every series and type is set
        com(lambda: setattr(chart.ChartArea.Format.TextFrame2.TextRange.Font, "Name", tn.FONT))
        _title(chart, spec.title)
        visible = [r for _, r in series if r not in HIDDEN]
        com(lambda: setattr(chart, "HasLegend", spec.kind != "waterfall" and len(visible) > 3))
        if spec.kind != "waterfall" and len(visible) > 3:
            com(lambda: setattr(chart.Legend, "Position", XL_LEGEND_TOP))
        _axes(chart, spec)
        if spec.kind != "line":
            group = com(lambda: chart.ChartGroups(1))
            com(lambda: setattr(group, "GapWidth", GAP_WIDTH))
            if spec.kind in ("variance", "waterfall"):
                com(lambda: setattr(group, "Overlap", 100))
        n = _points(spec.cat_range)
        for k, (s, role) in enumerate(series):
            line = spec.kind == "line" or (spec.kind == "combo" and role == "pct")
            _fill(s, role, k, line)
            _labels(s, role, spec.millions)
        if spec.highlight and series:
            _highlight(series[0][0], spec.highlight, n)
    except Exception:
        try:
            com(lambda: shape.Delete())                  # no half-drawn chart left behind
        except Exception:   # noqa: BLE001
            pass
        raise


def _points(rng: str) -> int:
    """Number of cells in a one-column (or one-row) range text: 'C5:C7,C9' -> 4."""
    n = 0
    for part in rng.split(","):
        a, _, b = part.partition(":")
        if not b:
            n += 1
            continue
        ra = int("".join(ch for ch in a if ch.isdigit()))
        rb = int("".join(ch for ch in b if ch.isdigit()))
        n += abs(rb - ra) + 1
    return n


def write_charts(ws, plan: SheetPlan) -> list[str]:
    """Draw every chart with series; returns the layout's chart lint plus one line per chart that failed."""
    warnings = list(plan.chart_lint)
    for spec in plan.charts:
        if not spec.series:             # the layout said why (chart_lint); its zone stays reserved
            continue
        try:
            _draw(ws, spec)
        except Exception as e:   # noqa: BLE001 - a chart never fails the answer
            warnings.append(f"chart '{spec.title}' failed: {e}; fix: check the source block's rows and columns, "
                            "or patch the chart block (kind, series)")
    return warnings
