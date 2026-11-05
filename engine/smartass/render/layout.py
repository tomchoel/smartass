# engine/smartass/render/layout.py
"""Q<n> Answer as a SheetPlan: every cell (text, or a formula from the model's expression trees) plus the style
roles styling.apply_report understands. Fixed zones (sub-spec 5 §1.1): Title, Subtitle, Key points, Charts, the
tables (the main table first), Checks, Notes (conventions, departures, thresholds, findings, caveats), Chart data and
Source. Column B holds codes (and every title, section and note), C the labels, values start in D. Two passes: rows
and addresses first, then formulas, so a Ref may point forwards (bridge 'Other' -> 'end', a check -> the tolerance
in Notes). Pure: no Excel."""
from __future__ import annotations

import math
import textwrap
from collections.abc import Callable
from dataclasses import dataclass, field

from ..model.analyze import Result
from ..model.blocks.common import GAP, MAT_ABS, MAT_PCT, MAT_SMALL, MARK_CODES, ck, group_start
from ..model.expr import Const
from ..model.frame import FILTER_COLUMNS
from ..model.periods import label as month_label
from ..model.spec import AnalysisSpec, CommentaryBlock, TableBlock
from ..xl.vendor import tn_rapportstil as tn

TEXT_COL = 2             # titles, subtitles, sections, notes and text: left-aligned with the code
CODE_COL = 2
LABEL_COL = 3
FIRST_COL = 4
NOTE_WIDTH = 78          # characters per note row (the width styling lint allows)
PCT_CODES = {"var_pct", "share", "key", "key_ly"}
CHART_ROWS = 18          # sheet rows per row of (at most two) charts
CHART_WIDTH = 480.0      # pt
CHART_GAP = 8.0          # pt between two charts in one zone row
ROW_PT = 15.0
LABEL_MIN, LABEL_MAX = 46, 80
LINE_KINDS = ("line", "sum", "grand", "check")
W_NUM, W_GAP, W_MARK, W_A, W_CODE, W_CODE_MAX = 16, 2, 9, 2, 9, 24
DEFAULT_WIDTH = 8.43
CONVENTIONS = ("NOK. FinMax signs: costs are negative. Δ = current − comparison; ▲ better, ▼ worse. Variance % "
               "blank when |base| < floor. ● = material (thresholds below).")
UNIT_MNOK, TOLERANCE = "unit_mnok", "tolerance"
TOLERANCE_NOK = 1.0
NON_TABLE = ("chart", "chart_data")    # outs that are not laid out in the tables zone

KeyPoints = Callable[..., list[tuple[str, str | None]]]


def col_letter(c: int) -> str:
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def col_number(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


@dataclass
class ChartSpec:
    type: str                 # chart kind: "line" | "waterfall" | "column" | ... (Task 7 draws the rest)
    title: str
    header_row: int
    rows: list[int]           # data rows (trend: months only, no Total; bridge: start..end); [] = not drawn yet
    series_cols: list[str]    # one series per column
    totals: list[int]         # waterfall points (1-based) that are totals
    anchor_col: str = "B"     # top-left cell of the chart in the Charts zone: anchor_col + anchor_row
    anchor_row: int = 0
    millions: bool = True     # value axis in MNOK; False for stocks (FTE) and percentages
    source: str = ""          # the key of the out whose rows the chart draws (b<i>)

    @property
    def height_pt(self) -> float:
        return (CHART_ROWS - 1) * ROW_PT          # fits its zone row, one empty row below

    @property
    def rows_tall(self) -> int:
        return math.ceil(self.height_pt / ROW_PT)


@dataclass
class SheetPlan:
    cells: dict[tuple[int, int], object] = field(default_factory=dict)
    rows: list[dict] = field(default_factory=list)
    addr: dict[str, str] = field(default_factory=dict)
    last_col: str = "C"
    last_row: int = 1
    charts: list[ChartSpec] = field(default_factory=list)
    fractions: set[str] = field(default_factory=set)        # keys of percentage cells (readback: no near-zero rule)
    params: dict[str, str] = field(default_factory=dict)    # threshold cells (Param keys) -> absolute address
    zones: list[tuple[str, int]] = field(default_factory=list)   # (zone name, first row), in sheet order
    status_cell: str = ""                                   # the empty 'Sheet = model' value cell (written later)
    widths: dict[str, float] = field(default_factory=dict)  # column letter -> width (A, B, C, numbers, gap, marker)

    def at(self, key: str) -> tuple[int, int]:
        a = self.addr[key]
        col = "".join(ch for ch in a if ch.isalpha())
        return int(a[len(col):]), col_number(col)


def _num(code: str, stock: bool) -> str:
    if code in MARK_CODES:
        return "markor"
    if code == "var":
        return "tall1" if stock else "avvik"
    if code == "var_pct":
        return "avvikpst"
    if code in PCT_CODES:
        return "prosent1"
    return "tall1" if stock else "heltall"


def _millions(o, stock: set[str]) -> bool:
    """MNOK on the value axis only for NOK amounts: not for a stock metric (FTE) and not for a %/ratio block."""
    metrics = set(o.col_metrics) | {ln.metric for ln in o.lines if ln.metric}
    if metrics & stock:
        return False
    return not (o.cols and all(c in PCT_CODES for c in o.cols))


def constant_columns(frame) -> dict[str, str]:
    """Filter columns that hold one single value over the whole Data table: a criterion on it matches every row."""
    out = {}
    for col in FILTER_COLUMNS:
        vals = {getattr(x, col.lower()) for x in frame.facts}
        if len(vals) == 1:
            out[col] = next(iter(vals))
    return out


def label_width(p: SheetPlan) -> int:
    """Column C fits the longest table line label (indent included): max(46, min(80, longest + 2)). Titles, text
    and notes are not counted (they may run over the empty cells to their right)."""
    rows = {r["row"] for r in p.rows if r.get("kind") in LINE_KINDS}
    longest = max((len(v) for (r, c), v in p.cells.items() if c == LABEL_COL and r in rows and isinstance(v, str)),
                  default=0)
    return max(LABEL_MIN, min(LABEL_MAX, longest + 2))


def code_width(p: SheetPlan) -> int:
    """Column B is 9 wide, wider only when a code would be cut off (EBITDA_BOI in bold on a sum row), at most 24."""
    rows = {r["row"] for r in p.rows if r.get("kind") in LINE_KINDS}
    longest = max((tn.text_units(v, 10, True) for (r, c), v in p.cells.items()
                   if c == CODE_COL and r in rows and isinstance(v, str)), default=0.0)
    return max(W_CODE, min(W_CODE_MAX, math.ceil(longest + 1)))


def width_pt(chars: float) -> float:
    """Column width in characters -> points (Arial 10: 7 px per character + 5 px padding, 0.75 pt per px)."""
    return (chars * 7 + 5) * 0.75


def main_table(res: Result):
    """The first table-like out in spec order (not a chart, not chart data, with lines), else None."""
    return next((o for o in res.outs if o.type not in NON_TABLE and o.lines), None)


def _last_period(o) -> str:
    """The period of the out's last column group, from its value header ('AC YTD Oct 26' -> 'YTD Oct 26')."""
    if o is None or "value" not in o.cols:
        return ""
    start = group_start(o.cols, len(o.cols) - 1)
    j = o.cols.index("value", start) if "value" in o.cols[start:] else o.cols.index("value")
    h = o.headers[j]
    return h[3:] if h.startswith("AC ") else h


def answer_subtitle(spec: AnalysisSpec, res: Result, closed: dict) -> str:
    """'NOK · <cubes> · <entity of the first query> · <period of the main table's last group> · closed <month>'."""
    cubes = sorted({q.cube for q in spec.queries}) or ["FinLite"]
    if not isinstance(closed, dict):           # one closed month for every cube
        closed = {c: closed for c in cubes}
    used = {c: m for c, m in closed.items() if c in cubes}
    labels = {month_label(m) for m in used.values()}
    when = (f"closed {labels.pop()}" if len(labels) == 1 else
            "closed " + ", ".join(f"{c} {month_label(m)}" for c, m in sorted(used.items())) if used else "")
    parts = ["NOK", " + ".join(cubes), res.entity_label, _last_period(main_table(res)), when]
    return " · ".join(x for x in parts if x)


def _no_key_points(main, res, drivers, addr) -> list[tuple[str, str | None]]:
    """The Key points hook (filled in plan 5 Task 6): (formula, cause or None) per point. The number of points must
    not depend on addr: the layout calls the hook once with a placeholder addr to reserve the rows, then again with
    the real one."""
    return []


def _param_label(key: str, label: str, stock: set[str]) -> str:
    """'(NOK)' only on NOK metrics: a stock metric (FTE) threshold has no unit."""
    metric = key.split(".", 1)[-1]
    if key.startswith((MAT_ABS, MAT_SMALL)) and metric in stock:
        return label.replace(" (NOK)", "")
    return label


def layout(q: int, spec: AnalysisSpec, res: Result, stock: set[str], subtitle: str,
           const: dict[str, str] | None = None, caveats: list[dict] | None = None,
           key_points: KeyPoints | None = None) -> SheetPlan:
    p = SheetPlan()
    table = f"Data_Q{q}"
    hook = key_points or _no_key_points
    outs = {o.key: o for o in res.outs}
    tables = [o for o in res.outs if o.type not in NON_TABLE]       # spec order (compute builds charts last)
    order = {f"b{i}": i for i in range(len(spec.blocks))}
    tables.sort(key=lambda o: order.get(o.key, len(order)))
    chart_data = [o for o in res.outs if o.type == "chart_data"]
    chart_data.sort(key=lambda o: order.get(o.key, len(order)))
    main = main_table(res)
    drivers = [d for b in spec.blocks if isinstance(b, CommentaryBlock) for d in b.drivers]
    widest = max([len(o.headers) for o in tables + chart_data] or [1])
    col_kind: dict[int, set[str]] = {}            # sheet column -> the column codes the tables put there
    placed: list[tuple[int, object, object]] = []          # (row, BlockOut, Line) per model line, for pass 2
    block_rows: dict[str, tuple[int, list[tuple[int, str]], list[str]]] = {}   # out key -> header row, rows, letters
    r = 1

    def zone(name: str, row: int) -> None:
        p.zones.append((name, row))

    def text(row: int, value: str, kind: str) -> None:
        p.cells[(row, TEXT_COL)] = value
        p.rows.append({"row": row, "kind": kind})

    def notes(row: int, value: str) -> int:
        """A note longer than the lint width continues on following note rows, split at word boundaries."""
        for part in textwrap.wrap(value, NOTE_WIDTH) or [""]:
            text(row, part, "note")
            row += 1
        return row

    def param(row: int, key: str, label: str, value: float, num: str) -> None:
        p.cells[(row, LABEL_COL)] = label
        p.cells[(row, FIRST_COL)] = value
        p.params[key] = f"${col_letter(FIRST_COL)}${row}"
        p.rows.append({"row": row, "kind": "line", "num": {col_letter(FIRST_COL): num}})

    def place(o, r: int) -> int:
        text(r, o.title, "section")
        r += 1
        letters = [col_letter(FIRST_COL + j) for j in range(len(o.headers))]
        for j, h in enumerate(o.headers):
            p.cells[(r, FIRST_COL + j)] = h
            col_kind.setdefault(FIRST_COL + j, set()).add(o.cols[j] if j < len(o.cols) else "")
        nums = [letters[j] for j in range(len(o.headers)) if j >= len(o.cols) or o.cols[j] != GAP]
        p.rows.append({"row": r, "kind": "header", "num_cols": nums})
        header_row = r
        r += 1
        line_rows: list[tuple[int, str]] = []
        for ln in o.lines:
            if ln.code:
                p.cells[(r, CODE_COL)] = ln.code
            p.cells[(r, LABEL_COL)] = ("   " * ln.level) + ln.label
            is_stock = ln.metric in stock if ln.metric else False
            num = {letters[j]: _num(o.cols[j], is_stock) for j in range(len(o.headers)) if o.cols[j] != GAP}
            p.rows.append({"row": r, "kind": ln.kind, "num": num})
            for j, e in enumerate(ln.cells):
                if e is not None:
                    p.addr[ck(o.key, ln.key, j)] = f"{letters[j]}{r}"
                    if o.cols[j] in PCT_CODES:
                        p.fractions.add(ck(o.key, ln.key, j))
            placed.append((r, o, ln))
            line_rows.append((r, ln.kind))
            r += 1
        block_rows[o.key] = (header_row, line_rows, letters)
        return r + 1

    # 1-2. Title and subtitle
    zone("title", r)
    text(r, spec.plan.title or spec.plan.purpose or spec.question, "title")
    zone("subtitle", r + 1)
    text(r + 1, subtitle, "subtitle")
    r += 3

    # 3. Key points (rows reserved now, formulas in pass 2: they point at the tables below)
    zone("key_points", r)
    text(r, "Key points", "section")
    r += 1
    points = hook(main, res, drivers, lambda key: "A1") if main is not None else []
    kp_row = r
    r += sum(1 + (cause is not None) for _, cause in points) + 1

    # 4. Charts: two per zone row, in spec order (a trend/bridge chart at its block, chart blocks, the implicit one)
    zone("charts", r)
    chart_outs = []
    for o in sorted(res.outs, key=lambda o: order.get(o.key, len(order))):
        if o.type == "chart" or o.type == "chart_data" or (o.chart and o.type in ("trend", "bridge")):
            chart_outs.append(o)
    charts_row = r
    r += math.ceil(len(chart_outs) / 2) * CHART_ROWS

    # 5. Tables, in spec order: the main table first
    zone("tables", r)
    for o in tables:
        r = place(o, r)

    # 6. Checks: Sheet = model (written after the readback), then every check line as OK / DIFF
    zone("checks", r)
    text(r, "Checks", "section")
    r += 1
    p.cells[(r, LABEL_COL)] = "Sheet = model"
    p.rows.append({"row": r, "kind": "check", "num": {}})
    p.status_cell = f"{col_letter(FIRST_COL)}{r}"
    r += 1
    check_rows: list[tuple[int, object, object]] = []
    for o in tables + chart_data:
        for ln in o.lines:
            if ln.kind == "check" and any(e is not None for e in ln.cells):
                p.cells[(r, LABEL_COL)] = f"{o.title}: {ln.label}"
                p.rows.append({"row": r, "kind": "check", "num": {}, "status_col": col_letter(FIRST_COL)})
                check_rows.append((r, o, ln))
                r += 1
    r += 1

    # 7. Notes: conventions, departures, thresholds, findings, caveats
    zone("notes", r)
    text(r, "Notes", "section")
    r += 1
    r = notes(r, CONVENTIONS)
    for b in spec.blocks:
        if isinstance(b, TableBlock) and b.rows == "metrics" and not b.chain:
            r = notes(r, f"{b.title or 'Table'}: the P&L chain is switched off (chain: false); the table shows a "
                         "deliberate subset of the lines.")
    r += 1
    text(r, "Thresholds", "section")
    r += 1
    if res.floors:
        r = notes(r, "Variance % is left blank when |base| is below the floor of its metric.")
    if any(k.startswith(MAT_PCT) for k in res.params):
        r = notes(r, "Material (●) when |variance| >= the materiality, or when |variance %| >= the materiality % "
                     "and |variance| >= the materiality with %.")
    for key, (label, value) in res.params.items():
        metric = key.split(".", 1)[-1]
        num = "prosent1" if key.startswith(MAT_PCT) else _num("value", metric in stock)
        param(r, key, _param_label(key, label, stock), value, num)
        r += 1
    param(r, UNIT_MNOK, "Unit MNOK (NOK)", 1_000_000, "heltall")
    r += 1
    param(r, TOLERANCE, "Tolerance (NOK)", TOLERANCE_NOK, "heltall")
    r += 2
    if res.findings:
        text(r, "Findings", "section")
        r += 1
        for f in res.findings:
            r = notes(r, f"{f.severity.capitalize()}: {f.text}")
        r += 1
    if caveats:                         # the knowledge facts the answer touched (sub-spec 2 §4), wrapped like notes
        text(r, "Caveats", "section")
        r += 1
        for c in caveats:
            r = notes(r, f"{c['severity'].capitalize()}: {c['text']}")
        r += 1

    # 8. Chart data: own chart data as tables (titled 'Chart data: <title>' by the model)
    if chart_data:
        zone("chart_data", r)
        for o in chart_data:
            r = place(o, r)

    # 9. Source
    zone("source", r)
    cubes = " + ".join(sorted({x.cube for x in spec.queries}) or ["FinLite"])
    r = notes(r, f"Source: Oracle EPM Cloud ({cubes}) · table {table} (sheet Q{q} Data) · Sources sheet") - 1

    # pass 2: formulas
    def addr(key: str) -> str:
        return p.addr[key] if key in p.addr else p.params[key]

    for row, o, ln in placed:
        for j, e in enumerate(ln.cells):
            if isinstance(e, Const):              # an input (a 100 % key share): a value, never a formula
                p.cells[(row, FIRST_COL + j)] = e.value
            elif e is not None:
                p.cells[(row, FIRST_COL + j)] = "=" + e.to_excel(table, addr, const)
    tol = p.params[TOLERANCE]
    for row, o, ln in check_rows:
        refs = [p.addr[ck(o.key, ln.key, j)] for j, e in enumerate(ln.cells) if e is not None]
        size = f"ABS({refs[0]})" if len(refs) == 1 else "MAX(" + ",".join(f"ABS({a})" for a in refs) + ")"
        p.cells[(row, FIRST_COL)] = f'=IF({size}<={tol},"OK","DIFF")'
    if main is not None:
        rr = kp_row
        for formula, cause in hook(main, res, drivers, addr):
            text(rr, formula, "text")
            rr += 1
            if cause is not None:
                text(rr, cause, "text")
                rr += 1
    p.last_col = col_letter(FIRST_COL + widest - 1)
    p.last_row = r
    p.widths = _widths(p, col_kind)
    p.charts = _charts(chart_outs, outs, block_rows, stock, charts_row, p.widths)
    return p


def _widths(p: SheetPlan, col_kind: dict[int, set[str]]) -> dict[str, float]:
    """A 2, B 9 (code), C the label width, numbers 16; a gap column 2 and the Material column 9 only where every
    table that uses the column has a gap (or the marker) there."""
    w: dict[str, float] = {"A": W_A, "B": code_width(p), "C": label_width(p)}
    for c in range(FIRST_COL, col_number(p.last_col) + 1):
        kinds = col_kind.get(c, set())
        w[col_letter(c)] = (W_GAP if kinds == {GAP} else W_MARK if kinds and kinds <= MARK_CODES else W_NUM)
    return w


def _charts(chart_outs, outs, block_rows, stock, top: int, widths: dict[str, float]) -> list[ChartSpec]:
    """One ChartSpec per chart: anchors two per zone row (B, and the first column 480 pt + a gap to the right of
    B); rows and series for the kinds charts.py draws today (a trend line, a bridge waterfall), [] for the rest
    (plan 5 Task 7 draws them)."""
    left = {}
    x = 0.0
    for c in range(1, 80):
        left[c] = x
        x += width_pt(widths.get(col_letter(c), DEFAULT_WIDTH))
    want = left[CODE_COL] + CHART_WIDTH + CHART_GAP
    second = col_letter(next(c for c in range(CODE_COL + 1, 80) if left[c] >= want))
    specs = []
    for i, o in enumerate(chart_outs):
        src = o.chart_of if o.type == "chart" else o.key
        data = outs.get(src)
        kind = o.chart or "column"
        rows: list[int] = []
        series: list[str] = []
        totals: list[int] = []
        header_row = 0
        if data is not None and data.key in block_rows:
            header_row, line_rows, letters = block_rows[data.key]
            if kind == "waterfall" and data.type == "bridge":
                rows = [row for row, _ in line_rows]
                series, totals = letters[:1], [1, len(rows)] if rows else []
            elif kind == "line" and data.type == "trend":
                rows = [row for row, k in line_rows if k not in ("grand", "check")]
                series = letters
        millions = _millions(data, stock) if data is not None else True
        specs.append(ChartSpec(kind, o.title, header_row, rows, series, totals,
                               anchor_col="B" if i % 2 == 0 else second, anchor_row=top + (i // 2) * CHART_ROWS,
                               millions=millions, source=src or ""))
    return specs
