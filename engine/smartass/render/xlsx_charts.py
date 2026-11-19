# engine/smartass/render/xlsx_charts.py
"""The answer's charts in the built file (render/xlsx.py): one native Excel chart per ChartSpec with series, written
with XlsxWriter on the Answer sheet at spec.anchor, 480 x 260 pt, named 'smartass chart'. The rules are
render/charts.py's, reused from there: the kinds (XL_TYPE), the hues and scenario fills (HUES, tint, hue, FILL,
PARTS), the label formats (label_format, en-US codes as they are), the legend rule, the axes; only the drawing moved
from COM to the file (docs/audit/2026-11-19-render-map.md section 5).

Series and category ranges point at the Answer sheet itself, absolute, a range with gaps as a union
('Q3 Answer'!$D$5:$D$7,'Q3 Answer'!$D$9). Its cells are empty when the file is built (the contents are written after
the move-in); the categories are cached as text, the values as not yet known, and Excel reads both from the cells.

Every option is put together before the chart is added, and XlsxWriter's warnings count as errors, so a chart that
cannot be built is one lint line (write_charts' wording) and the others are still drawn.

A per-point label setting is not written to the file: a custom c:dLbl as XlsxWriter writes it carries no number
format, and Excel then shows the raw value (live, 19.11.2026: 11580000 beside 10,3). In the file a line's labels
keep the series position, its unlabelled points deleted (a deleted label shows nothing, so it needs no format), and
a thin waterfall step is labelled like its neighbours. What COM must set on single points after the move-in
(charts._labels) is returned as LabelFix: a line's last label to the right of its point and a label the layout
moves clear of another line; a thin step's label in ink, lifted 10 pt above its bar."""
from __future__ import annotations

import warnings
from dataclasses import dataclass

from xlsxwriter.utility import xl_cell_to_rowcol

from ..xl.vendor import tn_rapportstil as tn
from . import look
from .charts import (ACCENT, FILL, GAP_WIDTH, GREY, HIDDEN, INK, NAVY, PARTS, PCT_LINE, THIN_LIFT, WATERFALL_STEPS,
                     WHITE, XL_TYPE, hue, label_format, legend_shown, tint, union_ref)
from .layout import ChartSpec, SheetPlan

CHART_NAME = "smartass chart"
XLSX_TYPE = {51: ("column", None), 52: ("column", "stacked"), 53: ("column", "percent_stacked"), 57: ("bar", None),
             4: ("line", None)}                       # XL_TYPE's chart types as XlsxWriter's type and subtype
LINE_WIDTH = 2.25                                     # pt
DASH = {"bu": "dash", "fc": "round_dot"}              # msoLineDash, msoLineRoundDot
PX_PER_PT = 4 / 3                                     # 96 dpi
TEXT_GREY = "#595959"                                 # AddChart2's (Office 2013) axis, legend and label text
AXIS_LINE = "#D9D9D9"                                 # AddChart2's category axis line
TEXT = {"name": tn.FONT, "size": 9, "color": TEXT_GREY}
_SECTION = tn.style_defs()["TN Seksjon"]
TITLE = {"name": tn.FONT, "size": _SECTION["size"] + 1, "bold": True, "color": look._hex(_SECTION["color"])}
CLUSTERED_OVERLAP = -27                               # AddChart2's (Office 2013) clustered column and bar default
HIDDEN_AXIS = {"major_gridlines": {"visible": False}, "major_tick_mark": "none", "label_position": "none",
               "line": {"none": True}}
FAILED = ("chart '{}' failed: {}; fix: check the source block's rows and columns, or patch the chart block (kind, "
          "series)")


def _hex(bgr: int) -> str:
    """A COM colour (tn.rgb: r + 256 g + 65536 b) as '#RRGGBB'."""
    return f"#{bgr & 0xFF:02X}{bgr >> 8 & 0xFF:02X}{bgr >> 16 & 0xFF:02X}"


def _ref(sheet: str, rng: str) -> str:
    """An absolute reference to rng on sheet; a range with gaps as a union."""
    u = union_ref(sheet, rng)
    return u if "," in rng else "=" + u[2:-1]


def _count(rng: str) -> int:
    """Cells in a one-column or one-row range text, gaps included: 'C5:C7,C9' -> 4, 'D10:F10' -> 3."""
    n = 0
    for part in rng.split(","):
        a, _, b = part.partition(":")
        (r1, c1), (r2, c2) = xl_cell_to_rowcol(a), xl_cell_to_rowcol(b or a)
        n += (abs(r2 - r1) + 1) * (abs(c2 - c1) + 1)
    return n


@dataclass(frozen=True)
class LabelFix:
    """One data label setting COM makes on a single point after the move-in (the file cannot carry it)."""
    chart: int                      # 1-based, among the charts drawn on the sheet, in drawing order
    series: int                     # 1-based series index inside that chart, as Excel numbers them
    point: int                      # 1-based point
    position: str | None = None     # "right" | "above" | "below" | "left" (line charts)
    font_color: str | None = None   # "#262626" (thin waterfall steps)
    lift_pt: float = 0.0            # 10.0 on thin waterfall steps


def _look(spec: ChartSpec, role: str, k: int, line: bool) -> dict:
    """Series k's fill and outline (charts._fill): the metric's hue styled by scenario."""
    base = hue(spec, k)
    main = look._hex(base)
    if line:
        colour = _hex(PCT_LINE) if role == "pct" else look._hex(tint(base)) if role == "py" else main
        ln = {"color": colour, "width": LINE_WIDTH} | ({"dash_type": DASH[role]} if role in DASH else {})
        return {"line": ln, "marker": {"type": "none"}}
    if role in HIDDEN:
        return {"fill": {"none": True}, "border": {"none": True}}
    if role == "bu":                                      # outlined: white with a 1 pt border in the hue
        return {"fill": {"color": _hex(WHITE)}, "border": {"color": main, "width": 1.0}}
    if role == "fc":                                      # hatched in the hue
        return {"pattern": {"pattern": "wide_upward_diagonal", "fg_color": main, "bg_color": _hex(WHITE)},
                "border": {"color": main}}
    colour = (_hex(PARTS[k % len(PARTS)]) if role == "part" else main if role == "ac"
              else look._hex(tint(base)) if role == "py" else _hex(FILL.get(role, NAVY)))
    return {"fill": {"color": colour}, "border": {"none": True}}


def _labels(spec: ChartSpec, role: str, k: int, n: int) -> dict | None:
    """Series k's data labels (charts._labels): the label format, and the position and font of its kind."""
    if role in HIDDEN:
        return None
    lab = {"value": True, "num_format": label_format(role, spec.millions, spec.whole), "font": dict(TEXT)}
    if spec.kind == "line" and k < len(spec.label_points):     # only the chosen points; the others deleted
        keep = set(spec.label_points[k])
        below = k < len(spec.label_below) and spec.label_below[k]
        lab["position"] = "below" if below else "above"
        lab["custom"] = [None if i in keep else {"delete": True} for i in range(1, n + 1)]
    elif spec.kind in ("stacked", "stacked100") and role == "part":   # inside the part: white on the dark ones
        lab["position"] = "center"
        lab["font"]["color"] = _hex(WHITE if k % len(PARTS) < 2 else INK)
    elif spec.kind == "combo" and role == "pct":
        lab["position"] = "above"
    elif spec.kind == "waterfall" and (role == "total" or role in WATERFALL_STEPS):
        lab["position"] = "inside_end" if role == "total" else "center"
        lab["font"] |= {"color": _hex(WHITE), "bold": True}
    return lab


def _point_fixes(spec: ChartSpec, role: str, k: int, n: int) -> list[tuple[int, dict]]:
    """(point, setting) for series k where a single point's label differs from its series (charts._labels): on a
    line the last labelled point to the right, a point the layout moves clear of another line above / left / below
    when that is not the series position; on a waterfall each thin step's label in ink, lifted off its bar."""
    if spec.kind == "line" and k < len(spec.label_points) and spec.label_points[k]:
        keep = set(spec.label_points[k])
        series_pos = "below" if k < len(spec.label_below) and spec.label_below[k] else "above"
        moved = spec.label_pos[k] if k < len(spec.label_pos) else {}
        out = []
        for i in sorted(keep):
            if i == max(keep):
                out.append((i, {"position": "right"}))
            elif i in moved:
                pos = moved[i] if moved[i] in ("above", "left") else "below"
                if pos != series_pos:
                    out.append((i, {"position": pos}))
        return out
    if spec.kind == "waterfall" and role in WATERFALL_STEPS:
        return [(i, {"font_color": _hex(INK), "lift_pt": THIN_LIFT}) for i in spec.thin_steps if 1 <= i <= n]
    return []


def _highlight(spec: ChartSpec, series: list[dict], roles: list[str], n: int) -> None:
    """Column and bar: the AC (or first part) series grey with the highlighted points in the accent; a comparison
    series white with a 1 pt grey border (charts._draw)."""
    target = next((i for i, r in enumerate(roles) if r in ("ac", "part")), None)
    if not spec.highlight or target is None or spec.kind not in ("column", "bar"):
        return
    series[target]["fill"] = {"color": _hex(GREY)}
    series[target]["points"] = [{"fill": {"color": _hex(ACCENT if i in spec.highlight else GREY)}}
                                for i in range(1, n + 1)]
    for i, r in enumerate(roles):
        if i != target and r in ("py", "bu", "fc"):
            series[i].pop("pattern", None)
            series[i] |= {"fill": {"color": _hex(WHITE)}, "border": {"color": _hex(GREY), "width": 1.0}}


def _axes(spec: ChartSpec) -> tuple[dict, dict]:
    """(value axis, category axis) (charts._axes): the value axis hidden, reversed and truncated as the spec says;
    the category labels at its low end (high when reversed), first row on top of a bar chart."""
    value = dict(HIDDEN_AXIS) | {"num_font": dict(TEXT)}
    if spec.reverse_values:
        value["reverse"] = True
    if spec.axis_min is not None:
        value["max" if spec.axis_min < 0 else "min"] = spec.axis_min
    category = {"major_tick_mark": "none", "label_position": "high" if spec.reverse_values else "low",
                "num_font": dict(TEXT), "line": {"color": AXIS_LINE}}
    if spec.kind in ("bar", "variance"):
        category["reverse"] = spec.sort != "asc"
    return value, category


def _options(spec: ChartSpec, sheet: str, chart_no: int) -> tuple[list[dict], list[dict], list[LabelFix]]:
    """Every series' options, split into the chart's own and the combo's % line, and the label fixes of chart
    chart_no. Excel numbers the series in the order XlsxWriter writes them (c:idx = c:order): the chart's own
    first, then the combo's lines; a waterfall's invisible bases count."""
    if spec.kind not in XL_TYPE:
        raise ValueError(f"unknown chart kind {spec.kind!r}")
    n = _count(spec.cat_range)
    series, roles = [], []
    for k, (name, rng, role) in enumerate(spec.series):
        line = spec.kind == "line" or (spec.kind == "combo" and role == "pct")
        s = {"name": name, "values": _ref(sheet, rng), "values_data": [None] * _count(rng),
             "categories": _ref(sheet, spec.cat_range), "categories_data": [""] * n} | _look(spec, role, k, line)
        lab = _labels(spec, role, k, n)
        if lab:
            s["data_labels"] = lab
        series.append(s)
        roles.append(role)
    _highlight(spec, series, roles, n)
    bars = [s for s, r in zip(series, roles) if not (spec.kind == "combo" and r == "pct")]
    lines = [s | {"y2_axis": True} for s, r in zip(series, roles) if spec.kind == "combo" and r == "pct"]
    if not bars:
        raise ValueError("a combo chart needs a column series beside its % line")
    if spec.kind != "line":                               # stacked: XlsxWriter's overlap 100, as AddChart2's
        overlap = 100 if spec.kind in ("variance", "waterfall") else (
            CLUSTERED_OVERLAP if XLSX_TYPE[XL_TYPE[spec.kind]][1] is None else None)
        bars[0] |= {"gap": GAP_WIDTH} | ({"overlap": overlap} if overlap is not None else {})
    pct = [spec.kind == "combo" and r == "pct" for r in roles]
    order = [k for k, p in enumerate(pct) if not p] + [k for k, p in enumerate(pct) if p]
    fixes = [LabelFix(chart_no, order.index(k) + 1, i, **setting)
             for k, role in enumerate(roles) for i, setting in _point_fixes(spec, role, k, n)]
    return bars, lines, fixes


def _add(workbook, worksheet, sheet: str, spec: ChartSpec, chart_no: int) -> list[LabelFix]:
    """Build one chart and place it; returns its label fixes. Nothing is placed when it fails."""
    bars, lines, fixes = _options(spec, sheet, chart_no)
    kind, subtype = XLSX_TYPE[XL_TYPE[spec.kind]]
    chart = workbook.add_chart({"type": kind, "name": CHART_NAME} | ({"subtype": subtype} if subtype else {}))
    for s in bars:
        chart.add_series(s)
    value, category = _axes(spec)
    if kind == "bar":                                     # XlsxWriter: a bar chart's x axis is its value axis
        chart.set_x_axis(value)
        chart.set_y_axis(category)
    else:
        chart.set_y_axis(value)
        chart.set_x_axis(category)
    if lines:                                             # combo: the % line on a hidden secondary axis
        pct = workbook.add_chart({"type": "line"})
        for s in lines:
            pct.add_series(s)
        pct.set_y2_axis(dict(HIDDEN_AXIS))
        chart.combine(pct)
    chart.set_title({"name": spec.title, "font": dict(TITLE)})
    chart.set_legend({"position": "top", "font": dict(TEXT)} if legend_shown(spec) else {"none": True})
    chart.set_chartarea({"border": {"none": True}})
    chart.set_size({"width": round(spec.width * PX_PER_PT), "height": round(spec.height * PX_PER_PT)})
    row, col = xl_cell_to_rowcol(spec.anchor)
    worksheet.insert_chart(row, col, chart)
    return fixes


def add_charts(workbook, worksheet, sheet_name: str, plan: SheetPlan) -> tuple[list[str], list[str], list[LabelFix]]:
    """Every chart of the plan with series on worksheet (the Answer sheet, named sheet_name). Returns (the titles
    drawn, the lint, the label fixes COM makes after the move-in). The lint is the layout's chart lint
    (plan.chart_lint) first, then a line per chart that failed, as charts.write_charts reports them; the caller adds
    no chart_lint of its own. A skipped or failed chart has no fixes and does not count in LabelFix.chart."""
    drawn: list[str] = []
    fixes: list[LabelFix] = []
    lint = list(plan.chart_lint)
    for spec in plan.charts:
        if not spec.series:             # the layout said why (chart_lint); its zone stays reserved
            continue
        try:
            with warnings.catch_warnings():
                warnings.filterwarnings("error", module="xlsxwriter")   # it warns and drops what it refuses
                chart_fixes = _add(workbook, worksheet, sheet_name, spec, len(drawn) + 1)
            drawn.append(spec.title)
            fixes += chart_fixes
        except Exception as e:   # noqa: BLE001 - a chart never fails the answer
            lint.append(FAILED.format(spec.title, e))
    return drawn, lint, fixes
