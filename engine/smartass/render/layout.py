# engine/smartass/render/layout.py
"""Q<n> Answer as a SheetPlan: every cell (text, or a formula from the model's expression trees) plus the style
roles v1's styling.apply_report understands. Column B holds labels, values start in C. Two passes: rows and
addresses first, then formulas, so a Ref may point forwards (bridge 'Other' -> 'end'). Pure: no Excel."""
from __future__ import annotations

import math
import textwrap
from dataclasses import dataclass, field

from ..model.analyze import FLOOR_KEY, Result
from ..model.blocks.common import ck
from ..model.expr import Const
from ..model.frame import FILTER_COLUMNS
from ..model.spec import AnalysisSpec, CommentaryBlock

FIRST_COL = 3
LABEL_COL = 2
NOTE_WIDTH = 78          # characters per note row (the width styling lint allows)
PCT_CODES = {"var_pct", "share", "key", "key_ly"}


def col_letter(c: int) -> str:
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


@dataclass
class ChartSpec:
    type: str                 # "line" | "waterfall"
    title: str
    header_row: int
    rows: list[int]           # data rows (trend: months only, no Total; bridge: start..end)
    series_cols: list[str]    # one series per column
    totals: list[int]         # waterfall points (1-based) that are totals
    anchor_col: str           # top-left cell of the chart: anchor_col + anchor_row
    anchor_row: int = 0       # 0 = the header row; the layout moves it down below the previous chart
    millions: bool = True     # value axis in MNOK; False for stocks (FTE) and percentages

    @property
    def height_pt(self) -> float:
        return max(180.0, 15.0 * (len(self.rows) + 3))

    @property
    def rows_tall(self) -> int:
        return math.ceil(self.height_pt / 15.0)


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

    def at(self, key: str) -> tuple[int, int]:
        a = self.addr[key]
        col = "".join(ch for ch in a if ch.isalpha())
        n = 0
        for ch in col:
            n = n * 26 + ord(ch) - 64
        return int(a[len(col):]), n


def _num(code: str, stock: bool) -> str:
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


def layout(q: int, spec: AnalysisSpec, res: Result, stock: set[str], subtitle: str,
           const: dict[str, str] | None = None, caveats: list[dict] | None = None) -> SheetPlan:
    p = SheetPlan()
    table = f"Data_Q{q}"
    widest = max([len(o.headers) for o in res.outs] or [1])
    r = 1

    def text(row: int, value: str, kind: str) -> None:
        p.cells[(row, LABEL_COL)] = value
        p.rows.append({"row": row, "kind": kind})

    def notes(row: int, value: str) -> int:
        """A note longer than the lint width continues on following note rows, split at word boundaries."""
        for part in textwrap.wrap(value, NOTE_WIDTH) or [""]:
            text(row, part, "note")
            row += 1
        return row

    text(r, spec.plan.purpose or spec.question, "title")
    text(r + 1, subtitle, "subtitle")
    r += 3
    text(r, "Plan", "section")
    r += 1
    text(r, f"Question: {spec.question}", "text")
    r += 1
    for h in spec.plan.hypotheses:
        text(r, f"•  Hypothesis: {h}", "text")
        r += 1
    r += 1

    placed: list[tuple[int, object, object]] = []          # (row, BlockOut, Line) per model line, for pass 2
    outs = {o.key: o for o in res.outs}
    for i, b in enumerate(spec.blocks):
        o = outs.get(f"b{i}")
        if isinstance(b, CommentaryBlock):
            text(r, b.title, "section")
            r += 1
            for kp in b.key_points:
                text(r, f"•  {kp}", "text")
                r += 1
            for d in b.drivers:
                ev = f"; evidence {d.evidence}" if d.evidence else ""
                basis = "Hypothesis - not verified in data" if d.basis == "hypothesis" else "data"
                text(r, f"•  {d.line}: {d.text} ({basis}{ev})", "text")
                r += 1
            r += 1
            continue
        if o is None:
            continue
        text(r, o.title, "section")
        r += 1
        letters = [col_letter(FIRST_COL + j) for j in range(len(o.headers))]
        for j, h in enumerate(o.headers):
            p.cells[(r, FIRST_COL + j)] = h
        p.rows.append({"row": r, "kind": "header", "num_cols": letters})
        header_row = r
        r += 1
        line_rows: list[tuple[int, str]] = []
        for ln in o.lines:
            p.cells[(r, LABEL_COL)] = ("   " * ln.level) + ln.label
            is_stock = ln.metric in stock if ln.metric else False
            num = {letters[j]: _num(o.cols[j], is_stock) for j in range(len(o.headers))}
            p.rows.append({"row": r, "kind": ln.kind, "num": num})
            for j, e in enumerate(ln.cells):
                if e is not None:
                    p.addr[ck(o.key, ln.key, j)] = f"{letters[j]}{r}"
                    if o.cols[j] in PCT_CODES:
                        p.fractions.add(ck(o.key, ln.key, j))
            placed.append((r, o, ln))
            line_rows.append((r, ln.kind))
            r += 1
        if o.chart:
            rows = [row for row, kind in line_rows if o.chart == "waterfall" or kind not in ("grand", "check")]
            if rows:
                line = o.chart == "line"
                p.charts.append(ChartSpec(o.chart, o.title, header_row, rows, letters if line else letters[:1],
                                          [] if line else [1, len(rows)], "", millions=_millions(o, stock)))
        r += 1

    if res.floors:                      # the variance % floors the formulas reference: labelled input cells
        text(r, "Thresholds", "section")
        r += 1
        r = notes(r, "Variance % is left blank when |base| is below the floor of its metric.")
        for metric, floor in res.floors.items():
            p.cells[(r, LABEL_COL)] = f"Variance % floor: {metric}"
            p.cells[(r, FIRST_COL)] = floor
            p.params[FLOOR_KEY + metric] = f"${col_letter(FIRST_COL)}${r}"
            p.rows.append({"row": r, "kind": "line", "num": {col_letter(FIRST_COL): _num("value", metric in stock)}})
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
    r = notes(r, f"Source: Oracle EPM Cloud via Smart View. Every number is a formula on table {table} "
                 f"(sheet Q{q} Data).") - 1

    def addr(key: str) -> str:
        return p.addr[key] if key in p.addr else p.params[key]

    for row, o, ln in placed:
        for j, e in enumerate(ln.cells):
            if isinstance(e, Const):              # an input (a 100 % key share): a value, never a formula
                p.cells[(row, FIRST_COL + j)] = e.value
            elif e is not None:
                p.cells[(row, FIRST_COL + j)] = "=" + e.to_excel(table, addr, const)
    p.last_col = col_letter(FIRST_COL + widest - 1)
    bottom = 0
    for c in p.charts:                  # right of the widest table, at its block's header row, never over the last chart
        c.anchor_col = col_letter(FIRST_COL + widest + 1)
        c.anchor_row = max(c.header_row, bottom + 1)
        bottom = c.anchor_row + c.rows_tall
    p.last_row = r
    return p
