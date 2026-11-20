# engine/smartass/render/charts.py
"""The rules of the IBCS charts of v2 answers (sub-spec 5 §3.3), one per ChartSpec at its anchor in the Charts zone;
render/xlsx_charts.py draws them in the built file, and xl/movein.py sets what only COM can set after the move-in.

- Colours: each metric gets its own hue (HUES: navy, teal, amber, purple, grey-blue); within a metric the scenario
  styling is IBCS: AC solid, PY a lighter tint of the same hue, BU outlined (bars) or dashed (lines), FC hatched
  (bars) or dotted (lines); variances green or red by effect; highlighted points in the TN accent, the rest grey.
- Labels instead of axes: data labels (MNOK '#,##0.0,,', en-US codes in the file), no
  gridlines, no tick marks, no value axis labels; a line labels only its last point (plus min and max on 12 points
  or fewer: spec.label_points). A legend on top whenever 2 or more series are visible (not on a waterfall or a
  variance chart, whose bars are labelled by effect).
- Waterfall: a stacked column over the chart helper columns (render/chartdata.py) with invisible bases; never
  Excel's native waterfall (its COM model does not reliably set totals). Variance: favourable and unfavourable
  series with 100 % overlap (no InvertIfNegative). Bar: ReversePlotOrder keeps the first (largest) row on top.
Every chart is named 'smartass chart' (xlsx_charts.CHART_NAME). A chart that fails is one lint line with its fix;
the others are still drawn. Pure: no Excel here."""
from __future__ import annotations

from ..xl.styling import GREEN
from ..xl.vendor import tn_rapportstil as tn
from .layout import ChartSpec

XL_TYPE = {"column": 51, "bar": 57, "line": 4, "stacked": 52, "stacked100": 53, "combo": 51, "variance": 57,
           "waterfall": 52}
GAP_WIDTH = 50

NAVY = tn.rgb(tn.NAVY)
GREY = tn.rgb((166, 166, 166))             # PY and the not-highlighted points
WHITE = tn.rgb((255, 255, 255))
GOOD = tn.rgb(GREEN)
BAD = tn.rgb(tn.BREACH_FONT)
ACCENT = tn.rgb((28, 22, 197))             # TN template blue #1C16C5 (dk2, the highlight colour; Tom 06.11.2026)
PCT_LINE = tn.rgb(tn.RULE_DARK)
INK = tn.rgb(tn.INK)                       # a waterfall label beside a step too thin to hold white text
THIN_LIFT = 10.0                           # pt: that label moved up off its 1 px bar
PARTS = [tn.rgb(c) for c in (tn.NAVY, (84, 110, 150), (140, 160, 190), (190, 200, 215), (166, 166, 166))]

HUES = [tn.NAVY, (0, 128, 128), (214, 137, 16), (112, 48, 160), (84, 110, 150)]   # navy, teal, amber, purple,
# grey-blue: one per metric in a chart
PY_TINT = 0.55                             # PY: the hue blended this far towards white
NO_LEGEND = ("waterfall", "variance")


def tint(c: tuple[int, int, int], t: float = PY_TINT) -> tuple[int, int, int]:
    """The colour blended t of the way towards white."""
    return tuple(round(x + (255 - x) * t) for x in c)


def hue(spec: ChartSpec, k: int) -> tuple[int, int, int]:
    """The RGB tuple of series k's metric."""
    h = spec.hues[k] if k < len(spec.hues) else 0
    return HUES[h % len(HUES)]


FILL = {"ac": NAVY, "py": GREY, "var": GOOD, "fav": GOOD, "unfav": BAD, "up": GOOD, "up_neg": GOOD, "down": BAD,
        "down_neg": BAD, "total": NAVY}
HIDDEN = ("base",)               # waterfall bases: no fill, no line, no label


MINUS = '"−"'      # the typographic minus as a quoted literal: data labels read −16,2, never -16,2


def label_format(role: str, millions: bool, whole: bool = False) -> str:
    """Data label format: MNOK with the scaling commas, whole NOK for amounts under 1 m, one decimal for stocks
    (FTE); zeros hidden (helper series are 0 off their points)."""
    if role == "pct":
        return f'0.0 %;{MINUS}0.0 %;;'
    unit = "#,##0.0,," if millions else "#,##0" if whole else "#,##0.0"
    if role in ("up", "up_neg"):           # an increase: its size, above or below zero
        return f"{unit};{unit};;"
    if role in ("down", "down_neg"):       # a decrease: minus its size, above or below zero
        return f"{MINUS}{unit};{MINUS}{unit};;"
    return f"{unit};{MINUS}{unit};;"


def legend_shown(spec: ChartSpec) -> bool:
    """The legend rule: stacked parts are named by the legend; a waterfall or a variance chart never has one;
    otherwise a single series is named by the title, unless it is not AC (a budget-only line must say BU)."""
    visible = [r for _, _, r in spec.series if r not in HIDDEN]
    return spec.kind in ("stacked", "stacked100") or (spec.kind not in NO_LEGEND and (
        len(visible) >= 2 or any(r not in ("ac", "part", "fav", "unfav", "var") for r in visible)))


def union_ref(sheet: str, rng: str) -> str:
    """A range with gaps ('D5:D7,D9') as one absolute union reference on sheet."""
    name = "'" + sheet.replace("'", "''") + "'"

    def absolute(a: str) -> str:
        col = "".join(ch for ch in a if ch.isalpha())
        return f"${col}${a[len(col):]}"
    parts = [":".join(absolute(x) for x in part.split(":")) for part in rng.split(",")]
    return "=(" + ",".join(f"{name}!{p}" for p in parts) + ")"


WATERFALL_STEPS = ("up", "up_neg", "down", "down_neg")
