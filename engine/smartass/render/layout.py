# engine/smartass/render/layout.py
"""Q<n> Answer as a SheetPlan: every cell (text, or a formula from the model's expression trees) plus the style
roles styling.apply_report understands. Fixed zones (sub-spec 5 §1.1): Title, Subtitle, Key points, Charts, the
tables (the main table first), Checks, Notes (conventions, departures, thresholds, findings, caveats), Chart data and
Source. Column B holds codes (and every title, section and note), C the labels, values start in D. Two passes: rows
and addresses first, then formulas, so a Ref may point forwards (bridge 'Other' -> 'end', a check -> the tolerance
in Notes). Pure: no Excel."""
from __future__ import annotations

import math
import re
import textwrap
from collections.abc import Callable
from dataclasses import dataclass, field

from ..model.analyze import Result
from ..model.blocks.common import GAP, MAT_ABS, MAT_PCT, MAT_SMALL, MARK_CODES, ck, group_start
from ..model.expr import Const, Var
from ..model.frame import FILTER_COLUMNS
from ..model.periods import label as month_label
from ..model.spec import AnalysisSpec, CommentaryBlock, LinesBlock, TableBlock
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
CHART_HEIGHT = 260.0     # pt (fits a zone row of 18 rows x 15 pt)
ROW_PT = 15.0
LABEL_MIN, LABEL_MAX = 46, 80
LINE_KINDS = ("line", "sum", "grand", "check")
W_NUM, W_GAP, W_MARK, W_A, W_CODE, W_CODE_MAX, W_HELPER = 16, 2, 9, 2, 9, 24, 16   # helpers hold NOK bn
HELPER = "helper"           # col_kind of a chart helper column
DEFAULT_WIDTH = 8.43
CONVENTIONS = ("NOK. FinMax signs: costs are negative. Δ = current − comparison; ▲ better, ▼ worse. Variance % "
               "blank when |base| < floor. ● = material (thresholds below).")
UNIT_MNOK, TOLERANCE, HUNDRED = "unit_mnok", "tolerance", "hundred"
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
    """One chart (ChartSpec v2, sub-spec 5 §3.3): what charts.write_charts draws. series are (name, range, role)
    with role ac | py | bu | fc | var | pct (scenario fills), fav | unfav (variance), base | up | down | total
    (stacked waterfall) or part (stacked); [] = not drawn (the zone stays reserved, the lint says why)."""
    kind: str                 # waterfall | line | column | bar | stacked | stacked100 | combo | variance
    title: str
    cat_range: str = ""       # category labels (column C), e.g. 'C12:C16'
    series: list[tuple[str, str, str]] = field(default_factory=list)
    anchor: str = "B1"        # top-left cell of the chart in the Charts zone
    width: float = CHART_WIDTH
    height: float = CHART_HEIGHT
    millions: bool = True     # data labels in MNOK; False for stocks (FTE), percentages and amounts under 1 m
    whole: bool = False       # NOK amounts under 1 m (millions False): labels in whole NOK ('#,##0')
    reverse_values: bool = False   # every plotted value <= 0 (costs): the value axis reversed, bars grow upwards
    axis_min: float | None = None  # a truncated bridge axis (waterfall): MinimumScale (MaximumScale when negative)
    thin_steps: list[int] = field(default_factory=list)  # waterfall steps (1-based) too thin for a white label
    sort: str = "none"        # desc | asc | none (bar: the order on the sheet; desc keeps the largest on top)
    highlight: list[int] = field(default_factory=list)   # 1-based points coloured with the accent, the rest grey
    source: str = ""          # the key of the out whose rows the chart draws (b<i>)
    helpers: dict[tuple[int, int], object] = field(default_factory=dict)   # chart helper cells (not model cells)
    hues: list[int] = field(default_factory=list)   # per series: the metric's hue (charts.HUES); [] = the first
    label_points: list[list[int]] = field(default_factory=list)   # per series (line): 1-based labelled points
    label_below: list[bool] = field(default_factory=list)   # per series (line): every value <= 0, labels below it
    unit: str = "nok"         # the unit class of every series (model/units.py): a chart never mixes units
    auto_title: bool = False  # an implicit chart (or trend/bridge sugar): the engine names it by metric and base
    _meta: object = field(default=None, repr=False, compare=False)   # (rows, picked columns) for the unit split

    @property
    def anchor_col(self) -> str:
        return "".join(ch for ch in self.anchor if ch.isalpha())

    @property
    def anchor_row(self) -> int:
        return int(self.anchor[len(self.anchor_col):])

    @property
    def height_pt(self) -> float:
        return self.height


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
    chart_lint: list[str] = field(default_factory=list)     # chart problems found while laying out (reply lint)
    helper_cols: list[str] = field(default_factory=list)    # chart helper columns (grey, right of every table)
    print_last_col: str = ""                                # the tables' last column (text rows; helpers excluded)
    print_area_col: str = ""                                # the print area's last column: the tables or the charts
    footer: str = ""                                        # the print footer (pipeline: source and fetch date)
    literal: set[tuple[int, int]] = field(default_factory=set)   # text cells: always written as text, never parsed

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


def sheet_label(ln) -> str:
    """The label shown on the sheet: without the code when it repeats it ('4P&C_T People & Competence' with code
    4P&C_T -> 'People & Competence'; '2WHS Wholesale' with code 2WHS -> 'Wholesale'). The model label
    (key_values keys) is never changed."""
    if getattr(ln, "display", ""):          # a sheet label of its own (a concept item without '(C30) (booked)')
        return ln.display
    code = getattr(ln, "code", "") or ""
    label = ln.label
    if code and label.startswith(code + " ") and label[len(code) + 1:].strip():
        return label[len(code) + 1:].lstrip()
    return label


MILLION = 1_000_000.0


def unit_of(value) -> str:
    """'MNOK' for an amount of at least 1 m (or unknown), 'NOK' below (no '-0,0 MNOK' for 7 271 NOK)."""
    return "NOK" if isinstance(value, (int, float)) and abs(value) < MILLION else "MNOK"


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


def key_point_table(tables: list):
    """The main table of the Key points: the first out in spec order with a value column and lines that is not a
    bridge or a trend (a table, kpi, ranking, concept ...), else None (then only the agent's free points)."""
    return next((o for o in tables if o.lines and "value" in o.cols and o.type not in ("bridge", "trend")), None)


def _point_rows(points) -> list[str]:
    from .keypoints import point_rows
    return point_rows(points)


def _last_period(o) -> str:
    """The period of the out: its own (a trend 'Jan 25–Oct 26', a concept), else its last column group's from the
    value header ('AC YTD Oct 26' -> 'YTD Oct 26')."""
    if o is not None and getattr(o, "period", ""):
        return o.period
    if o is None or "value" not in o.cols:
        return ""
    start = group_start(o.cols, len(o.cols) - 1)
    j = o.cols.index("value", start) if "value" in o.cols[start:] else o.cols.index("value")
    h = o.headers[j]
    return h[3:] if h.startswith("AC ") else h


def answer_subtitle(spec: AnalysisSpec, res: Result, closed: dict) -> str:
    """'NOK · <cubes> · <entity of the first query> · <period of the main table's last group> · closed <month>'.
    NOK only when every metric shown is an amount; otherwise no unit word (the headers carry the units)."""
    cubes = sorted({q.cube for q in spec.queries}) or ["FinLite"]
    if not isinstance(closed, dict):           # one closed month for every cube
        closed = {c: closed for c in cubes}
    used = {c: m for c, m in closed.items() if c in cubes}
    labels = {month_label(m) for m in used.values()}
    when = (f"closed {labels.pop()}" if len(labels) == 1 else
            "closed " + ", ".join(f"{c} {month_label(m)}" for c, m in sorted(used.items())) if used else "")
    units = set((getattr(res, "units", None) or {}).values())
    entity, period = res.entity_label, _last_period(main_table(res))
    cell = next((b.cell for b in spec.blocks if isinstance(b, LinesBlock)), None)
    if cell and not entity:                    # a ledger-lines answer: its drill cell's account, entity and month
        short = [str(cell.get(d, "")).removeprefix("ac.").removeprefix("un.SY.")
                 for d in ("All Accounts", "All Entities")]
        entity = " · ".join(x for x in short if x)
        period = period or str(cell.get("All Time Periods", ""))
    parts = ["NOK" if units <= {"nok"} else "", " + ".join(cubes), entity, period, when]
    return " · ".join(x for x in parts if x)


def _no_key_points(main, res, drivers, addr) -> list[tuple[str, str | None]]:
    """The Key points hook (filled in plan 5 Task 6): (formula, cause or None) per point. The number of points must
    not depend on addr: the layout calls the hook once with a placeholder addr to reserve the rows, then again with
    the real one."""
    return []


FLOOR = "floor."                                        # model.analyze.FLOOR_KEY
THRESHOLD_KINDS = (FLOOR, MAT_ABS, MAT_PCT, MAT_SMALL)   # the columns of the thresholds table, in this order
THRESHOLD_HEADS = ("Var % floor", "Materiality", "Materiality %", "Materiality with %")
PARAM_KIND = "param"                                    # col_kind of a thresholds table column


def threshold_cols(kinds: dict[int, set[str]]) -> list[int]:
    """The 4 sheet columns of the thresholds table: the first value columns from D that no table uses only as a
    gap (width 2) or a Material marker (width 9), so a 5 000 000 threshold never shows as ###."""
    out, c = [], FIRST_COL
    while len(out) < len(THRESHOLD_KINDS):
        k = kinds.get(c, set())
        if not (k == {GAP} or (k and k <= MARK_CODES)):
            out.append(c)
        c += 1
    return out


def _threshold_table(p: SheetPlan, r: int, params: dict, stock: set[str], cols: list[int]) -> int:
    """The thresholds as one compact table: a header row (Metric | Var % floor | Materiality | Materiality % |
    Materiality with %), then one row per metric with a cell per threshold it has (no cell otherwise). The unit sits
    on the metric: 'Opex (NOK)', a stock metric (FTE) without one. Returns the next free row."""
    rows: dict[str, dict[int, tuple[str, float]]] = {}
    for key, (_, value) in params.items():
        k = next((i for i, x in enumerate(THRESHOLD_KINDS) if key.startswith(x)), None)
        if k is not None:
            rows.setdefault(key[len(THRESHOLD_KINDS[k]):], {})[k] = (key, value)
    if not rows:
        return r
    kinds = sorted({k for have in rows.values() for k in have})   # only the kinds with a value: no empty column
    pos = {k: i for i, k in enumerate(kinds)}
    letters = [col_letter(c) for c in cols]
    p.cells[(r, LABEL_COL)] = "Metric"
    for k in kinds:
        p.cells[(r, cols[pos[k]])] = THRESHOLD_HEADS[k]
    p.rows.append({"row": r, "kind": "header", "num_cols": letters[:len(kinds)]})
    r += 1
    for metric, have in rows.items():
        is_stock = metric in stock
        p.cells[(r, LABEL_COL)] = metric if is_stock else f"{metric} (NOK)"
        num = {}
        for k, (key, value) in sorted(have.items()):
            p.cells[(r, cols[pos[k]])] = value
            p.params[key] = f"${letters[pos[k]]}${r}"
            num[letters[pos[k]]] = "prosent1" if THRESHOLD_KINDS[k] == MAT_PCT else _num("value", is_stock)
        p.rows.append({"row": r, "kind": "line", "num": num})
        r += 1
    return r


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
    main = (key_point_table(tables) or next((o for o in tables if o.type == "lines" and o.lines), None)
            or next((o for o in tables if o.type == "trend" and o.lines), None))
    drivers = [d for b in spec.blocks if isinstance(b, CommentaryBlock) for d in b.drivers]
    widest = max([len(o.headers) for o in tables + chart_data] or [1])
    col_kind: dict[int, set[str]] = {}            # sheet column -> the column codes the tables put there
    used: dict[int, set[str]] = {}
    for o in tables + chart_data:
        for j in range(len(o.headers)):
            used.setdefault(FIRST_COL + j, set()).add(o.cols[j] if j < len(o.cols) else "")
    t_cols = threshold_cols(used) if any(k.startswith(THRESHOLD_KINDS) for k in res.params) else []
    for c in t_cols:                              # the thresholds table: value columns, inside the print area
        col_kind.setdefault(c, set()).add(PARAM_KIND)
    widest = max([widest, *(c - FIRST_COL + 1 for c in t_cols)])
    placed: list[tuple[int, object, object]] = []          # (row, BlockOut, Line) per model line, for pass 2
    block_rows: dict[str, tuple[int, list[tuple[int, object]], list[str]]] = {}   # out key -> header row, (row, Line), letters
    r = 1

    def zone(name: str, row: int) -> None:
        p.zones.append((name, row))

    def text(row: int, value: str, kind: str, literal: bool = True) -> None:
        """literal: a wrapped note can start with '=' ('… YTD Sep' / '= 78 million …', live 05.11.2026), which Excel
        would read as a formula; text cells are always written as text."""
        p.cells[(row, TEXT_COL)] = value
        p.rows.append({"row": row, "kind": kind})
        if literal:
            p.literal.add((row, TEXT_COL))
        else:
            p.literal.discard((row, TEXT_COL))

    def note_width() -> int:
        """Characters per note row: the width of the tables under the note (B code + C label + the value columns,
        a gap or a marker column at its own width), at least NOTE_WIDTH - a note no longer wraps at 40 % of a wide
        table. 0.95 char per column unit stays inside the styling lint's Arial measure (italic 9 pt)."""
        units = W_CODE + LABEL_MIN
        for c in range(FIRST_COL, FIRST_COL + widest):
            k = col_kind.get(c, set()) | used.get(c, set())
            units += W_GAP if k == {GAP} else W_MARK if (k and k <= MARK_CODES) else W_NUM
        return max(NOTE_WIDTH, int(units * 0.95))

    def notes(row: int, value: str) -> int:
        """A note longer than the row width continues on following note rows, split at word boundaries."""
        for part in textwrap.wrap(value, note_width()) or [""]:
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
        line_rows: list[tuple[int, object]] = []
        for k, ln in enumerate(o.lines):
            if ln.kind == "heading":           # a concept sub-heading: navy bold, a spacer row before it
                if k:
                    r += 1
                p.cells[(r, LABEL_COL)] = sheet_label(ln)
                p.rows.append({"row": r, "kind": "subhead"})
                placed.append((r, o, ln))
                line_rows.append((r, ln))
                r += 1
                continue
            if ln.code:
                p.cells[(r, CODE_COL)] = ln.code
            p.cells[(r, LABEL_COL)] = ("   " * ln.level) + sheet_label(ln)
            is_stock = ln.metric in stock if ln.metric else False
            num = {letters[j]: _num(o.cols[j], is_stock) for j in range(len(o.headers)) if o.cols[j] != GAP}
            p.rows.append({"row": r, "kind": ln.kind, "num": num})
            for j, e in enumerate(ln.cells):
                if e is not None:
                    p.addr[ck(o.key, ln.key, j)] = f"{letters[j]}{r}"
                    if o.cols[j] in PCT_CODES:
                        p.fractions.add(ck(o.key, ln.key, j))
            placed.append((r, o, ln))
            line_rows.append((r, ln))
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
    kp_row = r
    asked: set[str] = set()              # the param cells the key points use (Unit MNOK, Percent factor)
    r += len(_point_rows(hook(main, res, drivers, lambda key: asked.add(key) or "A1"))) + 1

    # 4. Charts: two per zone row, in spec order (a trend/bridge chart at its block, chart blocks, the implicit one)
    zone("charts", r)
    chart_outs = []
    charted = {o.chart_of for o in res.outs if o.type == "chart" and o.chart_of}   # a chart block on it: its own
    for o in sorted(res.outs, key=lambda o: order.get(o.key, len(order))):       # chart sugar is not drawn too
        if o.type == "chart" or o.type == "chart_data" or (o.chart and o.type in ("trend", "bridge")
                                                           and o.key not in charted):
            chart_outs.append(o)
    charts_row = r
    dry = _dry_rows(tables + chart_data)     # a chart may split by unit: count the charts on provisional rows
    n_charts = sum(1 for o in chart_outs for cs in _charts(o, spec, res, outs, dry, stock, []) if _drawn(cs))
    r += math.ceil(n_charts / 2) * CHART_ROWS

    # 5. Tables, in spec order: the main table first
    zone("tables", r)
    for o in tables:
        r = place(o, r)

    # 6. Checks: Sheet = model (written after the readback), then every check line as OK / DIFF
    zone("checks", r)
    text(r, "Checks", "section")
    r += 1
    p.cells[(r, LABEL_COL)] = "Sheet = model"
    p.rows.append({"row": r, "kind": "check", "num": {}, "status_col": col_letter(FIRST_COL)})
    p.status_cell = f"{col_letter(FIRST_COL)}{r}"
    r += 1
    check_rows: list[tuple[int, object, object]] = []
    checked = [o for o in tables + chart_data if any(ln.kind == "check" for ln in o.lines)]
    for o in tables + chart_data:
        for ln in o.lines:
            if ln.kind == "check" and any(e is not None for e in ln.cells):
                what = ln.label.removeprefix("Check: ")          # the zone is called Checks
                p.cells[(r, LABEL_COL)] = what if len(checked) == 1 else f"{o.title}: {what}"
                p.rows.append({"row": r, "kind": "check", "num": {}, "status_col": col_letter(FIRST_COL)})
                check_rows.append((r, o, ln))
                r += 1
    r += 1

    # 7. Notes: conventions, departures, thresholds, findings, caveats
    zone("notes", r)
    text(r, "Notes", "section")
    r += 1
    r = notes(r, conventions(spec, res, tables + chart_data))
    for b in spec.blocks:
        if isinstance(b, TableBlock) and b.rows == "metrics" and not b.chain:
            r = notes(r, f"{b.title or 'Table'}: the P&L chain is switched off (chain: false); the table shows a "
                         "deliberate subset of the lines.")
    for line in _axis_notes(chart_outs, outs, res, stock):
        r = notes(r, line)
    r += 1
    # Thresholds: only the params a formula on the sheet uses (the model's floors and materiality, Unit MNOK and
    # Percent factor when a key point uses them, Tolerance when there is a check line); none: no section
    own = [(k, lab, v) for k, lab, v in ((UNIT_MNOK, "NOK per MNOK", 1_000_000),
                                         (TOLERANCE, "Tolerance (NOK)", TOLERANCE_NOK),
                                         (HUNDRED, "Percent factor", 100))
           if k in asked or (k == TOLERANCE and check_rows)]
    plain = [(key, label, value) for key, (label, value) in res.params.items()   # a param of no known kind
             if not key.startswith(THRESHOLD_KINDS)] + own
    has_thresholds = any(k.startswith(THRESHOLD_KINDS) for k in res.params)
    if has_thresholds or plain:
        text(r, "Thresholds" if has_thresholds else "Parameters", "section")
        r += 1
        if res.floors:
            r = notes(r, "Variance % is left blank when |base| is below the floor of its metric.")
        if any(k.startswith(MAT_PCT) for k in res.params):
            r = notes(r, "Material (●) when |variance| >= the materiality, or when |variance %| >= the "
                         "materiality % and |variance| >= the materiality with %.")
        r = _threshold_table(p, r, res.params, stock, t_cols)
        if plain:                                   # the formula constants: their own Parameter | Value table
            if has_thresholds:
                r += 1
            p.cells[(r, LABEL_COL)] = "Parameter"
            p.cells[(r, FIRST_COL)] = "Value"
            p.rows.append({"row": r, "kind": "header", "num_cols": [col_letter(FIRST_COL)]})
            r += 1
        for key, label, value in plain:
            param(r, key, label, value, "heltall")
            r += 1
        r += 1
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
            elif isinstance(e, Var):              # its base is missing on every line: blank while the base is 0
                base = e.b.to_excel(table, addr, const)
                p.cells[(row, FIRST_COL + j)] = f'=IF({base}=0,"",{e.to_excel(table, addr, const)})'
            elif e is not None:
                p.cells[(row, FIRST_COL + j)] = "=" + e.to_excel(table, addr, const)
    tol = p.params.get(TOLERANCE, "")
    for row, o, ln in check_rows:
        refs = [p.addr[ck(o.key, ln.key, j)] for j, e in enumerate(ln.cells) if e is not None]
        size = f"ABS({refs[0]})" if len(refs) == 1 else "MAX(" + ",".join(f"ABS({a})" for a in refs) + ")"
        p.cells[(row, FIRST_COL)] = f'=IF({size}<={tol},"OK","DIFF")'
    for i, row_text in enumerate(_point_rows(hook(main, res, drivers, addr))):
        text(kp_row + i, row_text, "text", literal=not row_text.startswith('="'))
    p.print_last_col = col_letter(FIRST_COL + widest - 1)
    p.last_col = p.print_last_col
    p.last_row = r
    specs = [cs for o in chart_outs for cs in _charts(o, spec, res, outs, block_rows, stock, p.chart_lint)]
    # the chart helpers go right of the tables AND of the charts' right edge, so the print area (which reaches the
    # charts) never takes them in: anchors from provisional widths first (every column right of the tables as wide
    # as a number or helper column, as it ends up), which the helpers do not move
    p.last_col = col_letter(FIRST_COL + widest + 60)
    p.widths = _widths(p, col_kind)
    p.charts = _anchor(specs, charts_row, p.widths)
    extent = col_number(_print_area_col(p))
    p.last_col = p.print_last_col
    _place_helpers(p, specs, max(FIRST_COL + widest + 1, extent + 2), col_kind)
    p.widths = _widths(p, col_kind)
    p.charts = _anchor(specs, charts_row, p.widths)
    p.print_area_col = _print_area_col(p)
    return p


def conventions(spec: AnalysisSpec, res: Result, outs) -> str:
    """The conventions note, only what holds for this sheet: NOK and FinMax signs only when a FinMax/FinLite amount
    is shown; FinPlan amounts without a sign claim (FinPlan's convention is not verified); FinPlan KPIs read Periodic
    with their scenarios; Δ/▲▼ only with variance columns, the floor with variance %, ● with a Material column."""
    cube = {q.id: q.cube for q in spec.queries}
    units = getattr(res, "units", None) or {}
    fm_nok = fp_nok = fp_kpi = False
    actual_cubes: set[str] = set()
    codes: set[str] = set()
    for o in outs:
        c = cube.get(o.query or "", "FinLite")
        ms = [*o.col_metrics, *(ln.metric for ln in o.lines if ln.metric)]
        nok = not ms or any(units.get(m, "nok") == "nok" for m in ms)
        kpi = any(units.get(m, "nok") != "nok" for m in ms)
        codes |= set(o.cols)
        if c == "FinPlan":
            fp_nok |= nok
            fp_kpi |= kpi
        else:
            fm_nok |= nok
            if nok:
                actual_cubes.add(c)
    parts = []
    if fm_nok:
        parts.append(f"NOK. {'/'.join(sorted(actual_cubes)) or 'FinMax'} signs: costs are negative.")
    if fp_nok:
        parts.append(("" if fm_nok else "NOK. ") + "FinPlan amounts as FinPlan holds them (its sign convention is "
                     "not verified).")
    if fp_kpi:
        scen = [x for code, x in (("value", "AC = actual (Working)"), ("budget", "BU = Target (Final)"),
                                  ("forecast", "FC = forecast (Final)")) if code in codes]
        parts.append("FinPlan KPIs read Periodic per month" + ("; " + ", ".join(scen) if scen else "") + ".")
    elif fp_nok and "budget" in codes:
        parts.append("BU = Target (Final).")
    if "var" in codes or "var_pct" in codes:
        parts.append("Δ = current − comparison; ▲ better, ▼ worse.")
    if "var_pct" in codes:
        parts.append("Variance % blank when |base| < floor.")
    if MARK_CODES & codes:
        parts.append("● = material (thresholds below).")
    return " ".join(parts) or "NOK."


def _print_area_col(p: SheetPlan) -> str:
    """The print area's last column: the tables', or further right to the column under a chart's right edge (two
    charts side by side may be wider than the tables; one page wide still holds, Excel scales down)."""
    last = col_number(p.print_last_col)
    left, x = {}, 0.0
    for c in range(1, 200):
        left[c] = x
        x += width_pt(p.widths.get(col_letter(c), DEFAULT_WIDTH))
    for cs in p.charts:
        if not (cs.series or cs.helpers):       # not drawn (a waterfall gets its series with its helpers)
            continue
        right = left[col_number(cs.anchor_col)] + cs.width
        c = next(c for c in range(1, 199) if left[c + 1] >= right)
        last = max(last, c)
    return col_letter(last)


def _widths(p: SheetPlan, col_kind: dict[int, set[str]]) -> dict[str, float]:
    """A 2, B 9 (code), C the label width, numbers 16; a gap column 2 and the Material column 9 only where every
    table that uses the column has a gap (or the marker) there."""
    w: dict[str, float] = {"A": W_A, "B": code_width(p), "C": label_width(p)}
    for c in range(FIRST_COL, col_number(p.last_col) + 1):
        kinds = col_kind.get(c, set())
        w[col_letter(c)] = (W_GAP if kinds == {GAP} else W_MARK if kinds and kinds <= MARK_CODES
                            else W_HELPER if kinds == {HELPER} else W_NUM)
    return w



# --- charts (ChartSpec v2) ----------------------------------------------------------------------------------------

ROLE = {"value": "ac", "amount": "ac", "ly": "py", "budget": "bu", "forecast": "fc", "var": "var", "var_pct": "pct"}
ALIAS = {"value": "amount", "amount": "value"}       # a ledger lines / bridge block has amount where others have value
IMPLICIT_SERIES = {"drivers": ["var"]}                 # an implicit bar on a drivers block shows the variance
SORTED_SOURCES = ("ranking", "lines")                  # rows already ranked by size (descending)
DROP_KEYS = {"lines": {"total", "cube", "gap"}}        # ledger lines rows that are controls, not parts
MAX_PARTS = 5


TRUNCATE_RATIO = 0.25       # a bridge axis is truncated when its largest |step| < 25 % of its smallest |total|
TRUNCATE_AT = 0.9            # ... at a round number below 90 % of its level nearest zero
THIN_STEP = 0.05             # a step under 5 % of the value span: a bar too thin for a white label (dark instead)


def nice_floor(x: float) -> float:
    """x > 0 rounded down to two significant digits: 3 330 000 000 -> 3 300 000 000."""
    unit = 10.0 ** (math.floor(math.log10(x)) - 1)
    return math.floor(x / unit + 1e-9) * unit


def bridge_axis(amounts: list[float]) -> float | None:
    """The truncated value axis of a bridge (start, steps..., end), the consulting convention when the steps are
    small next to the totals: a round number below every running level (MinimumScale; for a bridge below zero the
    MaximumScale, negative). None when the steps are large enough to read from zero, or the levels cross zero."""
    from .chartdata import waterfall_bars
    if len(amounts) < 3:
        return None
    totals, steps = (amounts[0], amounts[-1]), amounts[1:-1]
    smallest = min(abs(t) for t in totals)
    if not smallest or max(abs(s) for s in steps) >= TRUNCATE_RATIO * smallest:
        return None
    levels = [b[0] for b in waterfall_bars(amounts)]
    if all(v > 0 for v in levels):
        return nice_floor(TRUNCATE_AT * min(levels))
    if all(v < 0 for v in levels):
        return -nice_floor(TRUNCATE_AT * min(abs(v) for v in levels))
    return None


def _bridge_amounts(data, res, rows=None) -> list[float] | None:
    if data is None or data.type != "bridge" or "amount" not in data.cols:
        return None
    j = data.cols.index("amount")
    lines = [ln for _, ln in rows] if rows is not None else data.lines     # the chart's categories, else all
    vals = [res.cells.get(ck(data.key, ln.key, j)) for ln in lines]
    return [float(v) if isinstance(v, (int, float)) else 0.0 for v in vals]


def amount_text(v: float, millions: bool) -> str:
    """An amount for a note: '3 300 MNOK' (one decimal only when needed) or '45 000 NOK'."""
    if not millions:
        return f"{v:,.0f}".replace(",", " ") + " NOK"
    m = v / MILLION
    return (f"{m:,.0f}" if abs(m - round(m)) < 1e-9 else f"{m:,.1f}").replace(",", " ") + " MNOK"


def _axis_notes(chart_outs, outs, res, stock) -> list[str]:
    """'Bridge axis starts at <x> MNOK' for every waterfall drawn on a truncated axis (named when there are more)."""
    found = []
    for o in chart_outs:
        if (o.chart or "column") != "waterfall":
            continue
        data = outs.get(o.chart_of if o.type == "chart" else o.key)
        amounts = _bridge_amounts(data, res)
        axis = bridge_axis(amounts) if amounts else None
        if axis is not None:
            millions = _millions(data, stock) and max(abs(a) for a in amounts) >= MILLION
            found.append((o.title, amount_text(axis, millions)))
    if len(found) == 1:
        return [f"Bridge axis starts at {found[0][1]}."]
    return [f"Bridge axis starts at {x} ({title})." for title, x in found]


def _runs(letter: str, rows: list[int]) -> str:
    """A column's cells on the given rows as one range text: 'D5:D7' or, with gaps, 'D5:D6,D8'."""
    out, start = [], None
    for k, row in enumerate(rows):
        if start is None:
            start = row
        if k + 1 == len(rows) or rows[k + 1] != row + 1:
            out.append(f"{letter}{start}" if start == row else f"{letter}{start}:{letter}{row}")
            start = None
    return ",".join(out)


def _across(letters: list[str], row: int) -> str:
    """One row's cells in the given columns: 'D5:E5' when adjacent, else 'D5,F5'."""
    nums = [col_number(c) for c in letters]
    if nums == list(range(nums[0], nums[0] + len(nums))):
        return f"{letters[0]}{row}" if len(letters) == 1 else f"{letters[0]}{row}:{letters[-1]}{row}"
    return ",".join(f"{c}{row}" for c in letters)


def _dry_rows(outs) -> dict:
    """Provisional block rows (every out from row 1, columns from D) for counting the charts before the tables are
    placed: _charts only needs the rows' order, not where they end up."""
    return {o.key: (0, [(i + 1, ln) for i, ln in enumerate(o.lines)],
                    [col_letter(FIRST_COL + j) for j in range(len(o.headers))]) for o in outs}


def _val(res, key: str) -> float | None:
    v = res.cells.get(key)
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _empty(vals) -> bool:
    return all(v is None or v == 0 for v in vals)


def _concept_amounts(data, res, rows) -> list[float]:
    j = value_col(data)
    return [_val(res, ck(data.key, ln.key, j)) or 0.0 for _, ln in rows]


def value_col(data) -> int:
    """The value column of the block's last column group (a bridge: its amount)."""
    if "amount" in data.cols:
        return data.cols.index("amount")
    last = group_start(data.cols, len(data.cols) - 1)
    return data.cols.index("value", last) if "value" in data.cols[last:] else data.cols.index("value")


def series_metric(data, j: int, rows) -> str | None:
    """The metric behind series column j: the trend's column metric, else the rows' metric when they share one."""
    if j < len(data.col_metrics) and data.col_metrics[j]:
        return data.col_metrics[j]
    ms = {ln.metric for _, ln in rows if ln.metric}
    return ms.pop() if len(ms) == 1 else None


def unit_group_name(unit: str) -> str:
    from ..model.units import UNIT_WORD
    return UNIT_WORD.get(unit, "") or unit


def _charts(o, spec, res, outs, block_rows, stock, lint: list[str]) -> list[ChartSpec]:
    """The ChartSpecs of one chart out: one, or one per unit group when its series (a trend's metrics) or its rows
    (a table's metric lines) mix units (NOK, counts, ARPU, %, FTE): a chart never mixes units. Titles of a split
    chart name the group ('<title> — <metric>'). A problem is one lint line and a chart without series."""
    from ..model.units import NOK
    cs = _chart(o, spec, res, outs, block_rows, stock, lint)
    data = outs.get(cs.source)
    meta = getattr(cs, "_meta", None)
    if not cs.series or data is None or meta is None:
        return [cs]
    rows, picked = meta
    if cs.kind not in ("column", "bar", "line", "combo"):
        return [cs]
    units = getattr(res, "units", {}) or {}
    groups: dict[str, tuple[list[int], list]] = {}
    if cs.kind in ("column", "bar", "line"):
        by_col = [(series_metric(data, j, rows), j) for j in picked]
        if data.col_metrics:                                   # a trend: the series carry the metric
            for m, j in by_col:
                u = units.get(m, NOK) if m else NOK
                groups.setdefault(u, ([], rows))[0].append(j)
        else:                                                  # a table: the rows carry the metric
            for r, ln in rows:
                u = units.get(ln.metric, NOK) if ln.metric else NOK
                groups.setdefault(u, (list(picked), []))[1].append((r, ln))
    if len(groups) <= 1:
        unit = next(iter(groups), NOK)
        return [_finish(cs, data, res, rows, picked, unit, stock, lint, None)]
    out = []
    for unit, (js, rs) in groups.items():
        part = ChartSpec(cs.kind, cs.title, source=cs.source, sort=cs.sort)
        part.highlight = []
        if data.col_metrics:
            ms = list(dict.fromkeys(series_metric(data, j, rs) for j in js))
            name = data.headers[js[0]].rsplit(" ", 1)[0] if len(ms) == 1 else unit_group_name(unit)
        else:
            name = sheet_label(rs[0][1]).strip() if len(rs) == 1 else unit_group_name(unit)
        part.auto_title = cs.auto_title
        part.series = [(data.headers[j], _runs(block_rows[data.key][2][j], [r for r, _ in rs]), ROLE.get(
            data.cols[j], "ac")) for j in js]
        part.cat_range = _runs(col_letter(LABEL_COL), [r for r, _ in rs])
        part.millions = cs.millions
        out.append(_finish(part, data, res, rs, js, unit, stock, lint, name))
    return out


BASE_WORD = {"py": "last year", "bu": "budget", "fc": "forecast"}


def chart_title(cs: ChartSpec, data, kept: list[int], split_name: str | None) -> str:
    """An agent's chart title stays (a split adds ' — <metric>'). An implicit chart is named by what it shows:
    '<metric> vs budget' (last year, forecast), '<metric> per month' for a trend without a base; one or two metrics
    by name, more (or a table's rows) by the block title. Never '<block title> — ...'."""
    if not cs.auto_title:
        return f"{cs.title} — {split_name}" if split_name else cs.title
    names = list(dict.fromkeys(data.headers[j].rsplit(" ", 1)[0] for j in kept)) if data.col_metrics else []
    subject = " and ".join(names) if 1 <= len(names) <= 2 else (split_name or data.title)
    bars = getattr(cs, "bars", None)          # 'Top 15 by Opex' over 11 bars (zero ones left out): 'Top 11 ...'
    if bars and re.match(r"Top \d+ ", subject):
        subject = re.sub(r"^Top \d+ ", f"Top {bars} ", subject)
    roles = [r for _, _, r in cs.series]
    base = next((BASE_WORD[r] for r in roles if r in BASE_WORD), None)
    monthly = " per month" if data.type == "trend" else ""
    if base and "ac" in roles:
        return f"{subject.split(' vs ')[0]} vs {base}"      # 'Actual vs plan' + budget: 'Actual vs budget'
    if base:
        return f"{subject} {base}{monthly}"
    return f"{subject}{monthly}"


def _finish(cs: ChartSpec, data, res, rows, picked: list[int], unit: str, stock, lint,
            split_name: str | None = None) -> ChartSpec:
    """Per chart after the unit split: drop series with no data (lint), drop a trend's leading empty months, colour
    each metric with its own hue, label a line's last point (and min/max on 12 points or fewer), units."""
    from ..model.units import COUNT, NOK, PER_UNIT
    kept, series = [], []
    for j, s in zip(picked, cs.series):
        vals = [_val(res, ck(data.key, ln.key, j)) for _, ln in rows]
        if _empty(vals):
            lint.append(f"chart '{cs.title}': series '{s[0]}' has no data; left out")
            continue
        kept.append(j)
        series.append(s)
    if not kept:
        lint.append(f"chart '{cs.title}': no series has data; no chart drawn")
        cs.series = []
        return cs
    if data.type == "trend":            # a product launched in Nov 25: the chart starts at its first month with data
        first = next((i for i, (_, ln) in enumerate(rows)
                      if not _empty([_val(res, ck(data.key, ln.key, j)) for j in kept])), 0)
        rows = rows[first:]
    row_nums = [r for r, _ in rows]
    cols = [_letter(rng) for _, rng, _ in series]
    cs.series = [(name, _runs(cols[k], row_nums), role) for k, (name, _, role) in enumerate(series)]
    cs.cat_range = _runs(col_letter(LABEL_COL), row_nums)
    metrics = [series_metric(data, j, rows) for j in kept]
    # a metric keeps its hue across the charts a split makes (subscriptions navy, ARPU teal)
    hue_of = {m: i for i, m in enumerate(dict.fromkeys([*data.col_metrics, *metrics]))}
    cs.hues = [hue_of[m] for m in metrics]
    if cs.kind == "line":
        n = len(rows)
        cs.label_points = []
        for j in kept:
            vals = [_val(res, ck(data.key, ln.key, j)) for _, ln in rows]
            idx = [i for i, v in enumerate(vals) if v is not None]
            pts = {idx[-1] + 1} if idx else set()
            if idx and n <= LABEL_ALL_MAX:
                pts |= {max(idx, key=lambda i: vals[i]) + 1, min(idx, key=lambda i: vals[i]) + 1}
            cs.label_points.append(sorted(pts))
            cs.label_below.append(bool(idx) and all(vals[i] <= 0 for i in idx))
    if unit != NOK:
        cs.millions, cs.whole = False, unit in (COUNT, PER_UNIT)
    _scale(cs, data, res, rows, kept)
    cs.unit = unit
    cs.title = chart_title(cs, data, kept, split_name)
    return cs


LABEL_ALL_MAX = 12          # a line chart labels its last point, plus min and max on 12 points or fewer


def _letter(rng: str) -> str:
    """The column of a one-column range text: 'F12:F20,F22' -> 'F'."""
    return "".join(ch for ch in rng.split(":")[0].split(",")[0] if ch.isalpha())


def _chart(o, spec, res, outs, block_rows, stock, lint: list[str]) -> ChartSpec:
    """The ChartSpec of one chart out (rows and columns of the block it draws, its last column group); a problem is
    one lint line and a chart without series (never a failed answer). A concept draws the waterfall of the items in
    its total (headings, memo, missing items and checks are never categories)."""
    from ..model.blocks.chart import default_series
    from ..model.spec import ChartBlock
    kind = o.chart or "column"
    src = o.chart_of if o.type == "chart" else o.key
    data = outs.get(src)
    if data is not None and data.type == "concept" and kind != "variance":
        kind = "waterfall"
    cs = ChartSpec(kind, o.title, source=src or "", sort="desc" if kind == "bar" else "none")
    try:
        cb = None
        if o.type in ("chart", "chart_data") and o.key.startswith("b") and o.key[1:].isdigit():
            b = spec.blocks[int(o.key[1:])]
            cb = b if isinstance(b, ChartBlock) else None
        if cb is not None:
            cs.sort = cb.sort
        cs.auto_title = cb is None
        if cs.auto_title and data is not None and data.type == "concept":   # 'Direct contribution X YTD Oct 26'
            total = next((ln for ln in data.lines if ln.key == "total"), None)
            if total is not None:
                cs.title = " ".join(x for x in (sheet_label(total), data.period) if x)
        if data is None or data.key not in block_rows:
            lint.append(f"chart '{o.title}': its source has no rows on the sheet; fix: point source.block at a "
                        "table, trend, bridge, ranking or kpi block")
            return cs
        cs.millions = _millions(data, stock)
        if cb is None and data.type not in SORTED_SOURCES:
            cs.sort = "none"           # an implicit bar keeps the table's order (drivers: the tree)
        header_row, line_rows, letters = block_rows[data.key]
        cols = data.cols
        last = group_start(cols, len(cols) - 1) if cols else 0
        cand = [j for j in range(last, len(cols)) if cols[j] != GAP and cols[j] not in MARK_CODES]
        have = ", ".join(dict.fromkeys(cols[j] for j in cand)) or "none"
        if data.type == "concept":
            steps = {k: i for i, k in enumerate(data.steps)}
            rows = sorted(((r, ln) for r, ln in line_rows if ln.key in steps), key=lambda x: steps[x[1].key])
        elif kind == "waterfall":
            rows = list(line_rows)
        elif data.type == "drivers":    # the bottom of the tree: the top and middle levels are sums of it
            from ..model.blocks.drivers import leaf_keys
            leaves = leaf_keys(data)
            rows = [(r, ln) for r, ln in line_rows if ln.key in leaves]
        else:
            drop = DROP_KEYS.get(data.type, set())
            rows = [(r, ln) for r, ln in line_rows if ln.kind not in ("grand", "check") and ln.key not in drop]
            if any(ln.kind == "line" for _, ln in rows):
                rows = [(r, ln) for r, ln in rows if ln.kind != "sum"]
        if not rows:
            lint.append(f"chart '{o.title}': the source block has no rows to chart; fix: point the chart at a "
                        "block with lines")
            return cs
        if cb is not None and cb.series and data.type != "concept":
            want = list(cb.series)
        elif cb is None and data.type in IMPLICIT_SERIES:
            want = list(IMPLICIT_SERIES[data.type])
        elif data.type == "concept":
            want = ["amount"] if kind == "waterfall" else ["var"]
        else:
            want = default_series(kind, cols)
        picked: list[int] = []
        for code in want:
            hits = [j for j in cand if cols[j] == code] or [j for j in cand if cols[j] == ALIAS.get(code)]
            if not hits and (code in data.dropped or ALIAS.get(code) in data.dropped):
                continue                  # the model dropped it as empty on every line (no AC, no PY): not an error
            if not hits:
                lint.append(f"chart '{o.title}': series '{code}' is not in the source (columns: {have}); fix: add "
                            f"'{code}' to the source block's columns or drop it from the chart's series")
                continue
            picked += [j for j in (hits if data.type == "trend" else hits[-1:]) if j not in picked]
        if not picked:
            lint.append(f"chart '{o.title}': no series to draw, chart skipped; fix: give series the source has "
                        f"({have})")
            return cs
        full = rows
        rows = _no_zero_categories(rows, data, res, picked, kind, cs.millions)
        if len(rows) < len(full):
            cs.bars = sum(1 for _, ln in rows if ln.key != "other")
        row_nums = [r for r, _ in rows]
        cs.cat_range = _runs(col_letter(LABEL_COL), row_nums)
        labels = [ln.label.strip().lower() for _, ln in rows]
        shown = [sheet_label(ln).strip().lower() for _, ln in rows]     # as column C shows it
        every = {x for _, ln in full for x in (ln.label.strip().lower(), sheet_label(ln).strip().lower())}
        for h in (cb.highlight if cb is not None else []):
            want_h = h.strip().lower()
            if want_h in labels or want_h in shown:
                cs.highlight.append((labels.index(want_h) if want_h in labels else shown.index(want_h)) + 1)
            elif want_h not in every:         # a zero line left out of the chart has no bar to highlight
                lint.append(f"chart '{o.title}': highlight '{h}' is not a line of the source; fix: use a line "
                            "label as shown in column C")
        if cb is not None and cb.highlight and kind not in ("column", "bar"):
            lint.append(f"chart '{o.title}': highlight is drawn on column and bar charts only; fix: use kind "
                        "'column' or 'bar', or drop highlight")
            cs.highlight = []
        if kind == "waterfall":
            if data.type not in ("bridge", "concept") or (data.type == "bridge" and "amount" not in cols):
                lint.append(f"chart '{o.title}': a waterfall needs a bridge block; fix: source {{block: <bridge "
                            "index>}")
                return cs
            cs.helpers = {"kind": "waterfall", "rows": row_nums, "col": letters[value_col(data)],
                          "header_row": header_row}
        elif kind == "variance":
            v = next((j for j in picked if cols[j] == "var"), None)
            if v is None:
                lint.append(f"chart '{o.title}': a variance chart needs a var column; fix: add 'var' to the "
                            "source's columns")
                return cs
            cs.helpers = {"kind": "variance", "rows": row_nums, "col": letters[v], "header_row": header_row,
                          "name": data.headers[v]}
        elif kind in ("stacked", "stacked100"):
            if len(rows) > MAX_PARTS:
                lint.append(f"chart '{o.title}': {len(rows)} parts (at most {MAX_PARTS} read well); fix: chart a "
                            "block with fewer lines, or use a bar chart")
            sel = [letters[j] for j in picked]
            cs.cat_range = _across(sel, header_row)
            cs.series = [(sheet_label(ln).strip(), _across(sel, r), "part") for r, ln in rows]
        else:
            if kind == "bar" and cs.sort != "none" and data.type not in SORTED_SOURCES:
                vals = [res.cells.get(ck(data.key, ln.key, picked[0])) for _, ln in rows]
                vals = [v if isinstance(v, (int, float)) else 0.0 for v in vals]
                if vals != sorted(vals, reverse=cs.sort == "desc"):
                    lint.append(f"chart '{o.title}': sort '{cs.sort}' not applied, the bars keep the table's "
                                "order; fix: chart a ranking block (rows sorted by size) or set sort 'none'")
            cs.series = [(data.headers[j], _runs(letters[j], row_nums), ROLE.get(cols[j], "ac")) for j in picked]
        cs._meta = (rows, picked)
        if kind not in ("column", "bar", "line", "combo"):
            _scale(cs, data, res, rows, picked)
    except Exception as e:   # noqa: BLE001 - a chart never fails the answer
        cs.series, cs.helpers = [], {}
        cs._meta = None
        lint.append(f"chart '{o.title}' failed: {e}; fix: patch the chart block (kind, source, series)")
    return cs


ZERO_MNOK, ZERO_NOK = 50_000, 0.5     # under this a data label reads 0,0 (MNOK, one decimal) or 0 (whole NOK)


def _no_zero_categories(rows, data, res, picked: list[int], kind: str, millions: bool):
    """A category whose every plotted value would print as 0 (0,0 MNOK, or 0 NOK) is no bar, only an empty slot or
    a '-0,0' label: left out of the chart (the table keeps the line). A waterfall keeps its start and end bars and
    drops zero steps (the running total is unchanged); a stacked chart keeps its parts. Nothing is dropped when it
    would leave no category."""
    if kind in ("stacked", "stacked100", "line"):
        return rows
    cols = [value_col(data)] if kind == "waterfall" else [j for j in picked if data.cols[j] not in PCT_CODES]
    if not cols:
        return rows
    every = [v for _, ln in rows for j in cols
             if isinstance(v := res.cells.get(ck(data.key, ln.key, j)), (int, float))]
    big = millions and any(abs(v) >= MILLION for v in every)      # _scale: MNOK labels only from 1 m
    tiny = ZERO_MNOK if big else ZERO_NOK

    def zero(ln) -> bool:
        if kind == "waterfall" and ln.kind in ("sum", "grand"):
            return False
        vals = [res.cells.get(ck(data.key, ln.key, j)) for j in cols]
        return all(not isinstance(v, (int, float)) or abs(v) < tiny for v in vals)

    kept = [(r, ln) for r, ln in rows if not zero(ln)]
    return kept or rows


def _scale(cs: ChartSpec, data, res, rows, picked: list[int]) -> None:
    """Units and value axis from the model values the chart plots: MNOK labels only from 1 m (|max|), whole NOK
    below; a value axis reversed when every plotted value is <= 0 (costs: the bars grow away from the category
    labels, which keep their minus sign); a bridge's truncated axis (bridge_axis; a concept's waterfall always
    starts at zero)."""
    from .chartdata import waterfall_bars
    if cs.kind == "waterfall":
        vals = (_concept_amounts(data, res, rows) if data.type == "concept" else _bridge_amounts(data, res, rows)) or []
        cs.axis_min = bridge_axis(vals) if vals and data.type == "bridge" else None
        levels = [b[0] for b in waterfall_bars(vals)] if vals else []
        if levels:
            reach = max(abs(v) for v in levels)
            span = reach - abs(cs.axis_min) if cs.axis_min is not None else max(levels + [0]) - min(levels + [0])
            cs.thin_steps = [i + 1 for i in range(1, len(vals) - 1) if span and abs(vals[i]) < THIN_STEP * span]
    else:
        if cs.kind == "variance":
            picked = [j for j in picked if data.cols[j] == "var"][:1]
        js = [j for j in picked if data.cols[j] not in PCT_CODES]
        vals = [res.cells.get(ck(data.key, ln.key, j)) for _, ln in rows for j in js]
        vals = [float(v) for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
        levels = vals
    if cs.millions and vals and max(abs(v) for v in vals) < MILLION:
        cs.millions, cs.whole = False, True
    cs.reverse_values = (cs.kind != "variance" and bool(levels) and all(v <= 0 for v in levels)
                         and any(v < 0 for v in levels))     # a variance keeps its sign's side (red left)


def _place_helpers(p: SheetPlan, specs: list[ChartSpec], first: int, col_kind) -> None:
    """Chart helper columns (render/chartdata.py) to the right of every table (first = widest table + one empty
    column), on the rows of the block they reshape, under a grey 'Chart helper' header; the chart's series point
    at them. They widen p.last_col (owned range, styling, readback block) but not the print area. Not model cells:
    never in p.addr."""
    from .chartdata import (VARIANCE_HEADS, VARIANCE_ROLES, WATERFALL_HEADS, WATERFALL_ROLES, variance_helpers,
                            waterfall_helpers)
    by_row = {d["row"]: d for d in p.rows}
    cursor = first
    for cs in specs:
        h = cs.helpers
        if not h:
            continue
        wf = h["kind"] == "waterfall"
        heads, roles = (WATERFALL_HEADS, WATERFALL_ROLES) if wf else (VARIANCE_HEADS, VARIANCE_ROLES)
        letters = [col_letter(cursor + k) for k in range(len(heads))]
        cells = (waterfall_helpers if wf else variance_helpers)(h["rows"], h["col"], letters)
        hr = h["header_row"]
        p.cells[(hr - 1, cursor)] = "Chart helper"
        for k, name in enumerate(heads):
            p.cells[(hr, cursor + k)] = name
            col_kind.setdefault(cursor + k, set()).add(HELPER)
        by_row[hr].setdefault("num_cols", []).extend(letters)
        num = "heltall" if cs.millions else "tall1"
        for (row, letter), v in cells.items():
            p.cells[(row, col_number(letter))] = v
            by_row[row].setdefault("num", {})[letter] = num
        cs.series = [(heads[k] if wf else f"{h['name']} {heads[k].lower()}", _runs(letters[k], h["rows"]), role)
                     for k, role in enumerate(roles) if role]      # the same rows (runs) as the categories
        cs.helpers = {(row, col_number(letter)): v for (row, letter), v in cells.items()}
        p.helper_cols += letters
        cursor += len(heads)
    if p.helper_cols:
        p.last_col = p.helper_cols[-1]


def _anchor(specs: list[ChartSpec], top: int, widths: dict[str, float]) -> list[ChartSpec]:
    """Anchors two per zone row: B, and the first column 480 pt + a gap to the right of B."""
    left = {}
    x = 0.0
    for c in range(1, 80):
        left[c] = x
        x += width_pt(widths.get(col_letter(c), DEFAULT_WIDTH))
    want = left[CODE_COL] + CHART_WIDTH + CHART_GAP
    second = col_letter(next(c for c in range(CODE_COL + 1, 80) if left[c] >= want))
    i = 0
    for cs in specs:            # a chart that is not drawn (no data: the lint says why) takes no slot
        cs.anchor = f"{'B' if i % 2 == 0 else second}{top + (i // 2) * CHART_ROWS}"
        i += _drawn(cs)
    return specs


def _drawn(cs: ChartSpec) -> bool:
    """A chart write_charts will draw: it has series, or helpers that give it its series (waterfall, variance)."""
    return bool(cs.series or cs.helpers)
