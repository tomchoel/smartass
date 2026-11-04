# engine/smartass/render/layout.py
"""Q<n> Answer as a SheetPlan: every cell (text, or a formula from the model's expression trees) plus the style
roles v1's styling.apply_report understands. Column B holds labels, values start in C. Two passes: rows and
addresses first, then formulas, so a Ref may point forwards (bridge 'Other' -> 'end'). Pure: no Excel."""
from __future__ import annotations

from dataclasses import dataclass, field

from ..model.analyze import Result
from ..model.blocks.common import ck
from ..model.spec import AnalysisSpec, CommentaryBlock

FIRST_COL = 3
LABEL_COL = 2
PCT_CODES = {"var_pct", "share"}


def col_letter(c: int) -> str:
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


@dataclass
class SheetPlan:
    cells: dict[tuple[int, int], object] = field(default_factory=dict)
    rows: list[dict] = field(default_factory=list)
    addr: dict[str, str] = field(default_factory=dict)
    last_col: str = "C"
    last_row: int = 1

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


def layout(q: int, spec: AnalysisSpec, res: Result, stock: set[str], subtitle: str) -> SheetPlan:
    p = SheetPlan()
    table = f"Data_Q{q}"
    widest = max([len(o.headers) for o in res.outs] or [1])
    r = 1

    def text(row: int, value: str, kind: str) -> None:
        p.cells[(row, LABEL_COL)] = value
        p.rows.append({"row": row, "kind": kind})

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
        r += 1
        for ln in o.lines:
            p.cells[(r, LABEL_COL)] = ("   " * ln.level) + ln.label
            is_stock = ln.metric in stock if ln.metric else False
            num = {letters[j]: _num(o.cols[j], is_stock) for j in range(len(o.headers))}
            p.rows.append({"row": r, "kind": ln.kind, "num": num})
            for j, e in enumerate(ln.cells):
                if e is not None:
                    p.addr[ck(o.key, ln.key, j)] = f"{letters[j]}{r}"
            placed.append((r, o, ln))
            r += 1
        r += 1

    if res.findings:
        text(r, "Findings", "section")
        r += 1
        for f in res.findings:
            text(r, f"{f.severity.capitalize()}: {f.text}", "note")
            r += 1
        r += 1
    text(r, f"Source: Oracle EPM Cloud via Smart View. Every number is a formula on table {table} "
            f"(sheet Q{q} Data).", "note")

    def addr(key: str) -> str:
        return p.addr[key]

    for row, o, ln in placed:
        for j, e in enumerate(ln.cells):
            if e is not None:
                p.cells[(row, FIRST_COL + j)] = "=" + e.to_excel(table, addr)
    p.last_col = col_letter(FIRST_COL + widest - 1)
    p.last_row = r
    return p
