# engine/smartass/render/look.py
"""The Answer's look as a pure function of its plan: what styling.apply_report does to Q<n> Answer through COM
(named TN styles per row kind, number formats laid over them, the grey code and helper columns, row heights, column
widths, merges, the check colours and the page breaks), computed without Excel so that a file writer can put it into
the .xlsx it builds. apply_report stays as it is (xl_tidy uses it) and is this module's oracle (tests/v2/test_look.py).

Two of its decisions read Excel's calculated values today; here they get a pure rule:
- a Key point's row height and merge use the calculated sentence when the caller passes it in texts, else the
  formula's own text (its top-level string literals: a lower bound); text_heights gives the final heights from the
  block read back from Excel;
- a value in a table column without a number kind is TN Heltall when it is a number or a formula whose text is not
  given, TN Etikett otherwise.
No COM, no Excel, no file: plain data in, plain data out.

style_findings gives back what apply_report's lint (styling._lint_fast) told the reply before 0.3.0, as far as the
look can know it: a text that does not fit its merged row and a text cut off or spilling over its column, in the
same words, from the read-back block. The lint's other findings read Excel itself - the sheet's font, vertical
borders and content outside the table width anywhere on the sheet: the look writes Arial with top and bottom edges
only, inside its block, and what Excel or the user does to the sheet afterwards it cannot see."""
from __future__ import annotations

from dataclasses import dataclass, replace
from functools import cache

from ..xl import styling
from ..xl.vendor import tn_rapportstil as tn
from .excel import answer_roles, keep_spans
from .layout import DEFAULT_WIDTH, SheetPlan, col_letter, col_number, width_pt

ROW_STYLE = {"title": "TN Tittel", "subtitle": "TN Undertittel", "section": "TN Seksjon", "note": "TN Note",
             "blank": "TN Tom"}                                    # one style across the whole row
ROW_HEIGHT = {"title": 28, "subtitle": 16, "section": 22}           # every other row 15 pt, unless its text wraps
EMPH = {"line": "", "check": "", "sum": "Sum", "grand": "Hovedsum"}


@dataclass(frozen=True)
class CellLook:
    style: str                      # a TN named style, e.g. "TN Heltall · Sum"; "TN Tom" for a blank cell
    num_format: str | None = None   # en-US code laid over the style (map 4.2 overrides)
    font_color: str | None = None   # "#767676" on code cells and helper columns
    align: str | None = None        # "left" on code cells
    wrap: bool | None = None        # True on a merged text or title cell
    valign: str | None = None       # "top" on the first cell of a merge


@dataclass
class SheetLook:
    cells: dict[tuple[int, int], CellLook]      # every cell of A1:<last_col><last_row>, 1-based (row, col)
    heights: dict[int, float]                   # row -> points
    widths: dict[str, float]                    # column letter -> characters
    merges: list[tuple[int, int, int]]          # (row, first col, last col)
    status_runs: list[tuple[int, int, int]]     # (col, first row, last row): the OK / not-OK conditional formats
    text_rows: list[int]                        # Key point rows whose height depends on the calculated text
    print_area: str                             # "$A$1:$N$60"
    breaks: list[int]                           # rows that start a new printed page
    footer: str                                 # plan.footer


# ------------------------------------------------------------------ named styles as concrete formats
def _hex(c: tuple[int, int, int]) -> str:
    return "#{:02X}{:02X}{:02X}".format(*c)


GREY = _hex(tn.GREY)
_ALIGN = {tn.H_LEFT: "left", tn.H_RIGHT: "right"}                       # H_GEN: no key (general)
_VALIGN = {tn.V_TOP: "top", tn.V_CENTER: "vcenter", tn.V_BOTTOM: "bottom"}
_BORDER = {tn.THIN: 1, tn.MEDIUM: 2, tn.DOUBLE: 6}                      # XlsxWriter border codes


@cache
def _defs() -> dict:
    return tn.style_defs()


def style_props(style: str) -> dict:
    """The XlsxWriter format properties of a TN named style (tn.style_defs(), map 4.2): font, alignment, wrap,
    indent, number format, top/bottom border and colour, fill."""
    p = _defs()[style]
    out = {"font_name": tn.FONT, "font_size": p.get("size", 10), "bold": bool(p.get("bold")),
           "italic": bool(p.get("italic")), "font_color": _hex(p.get("color", tn.INK)),
           "valign": _VALIGN[p.get("v", tn.V_CENTER)], "text_wrap": bool(p.get("wrap"))}
    if p.get("h", tn.H_GEN) in _ALIGN:
        out["align"] = _ALIGN[p["h"]]
    if p.get("underline"):
        out["underline"] = 1
    if p.get("indent"):
        out["indent"] = p["indent"]
    if p.get("num"):
        out["num_format"] = p["num"]
    if p.get("fill"):
        out["bg_color"] = _hex(p["fill"])
    for edge in ("top", "bottom"):
        if p.get(edge):
            weight, colour = p[edge]
            out[edge] = _BORDER[weight]
            out[f"{edge}_color"] = _hex(colour)
    return out


# ------------------------------------------------------------------ what a cell shows, without Excel
def _formula(plan: SheetPlan, rc) -> bool:
    v = plan.cells.get(rc)
    return isinstance(v, str) and v.startswith("=") and rc not in plan.literal


def _static_text(formula: str) -> str:
    """The text a formula shows whatever its inputs: its top-level string literals joined
    ('="•  Opex "&FIXED(D5,0)&" NOK"' -> '•  Opex  NOK'). Literals inside a function's arguments are left out."""
    out, depth, i, n = [], 0, 0, len(formula)
    while i < n:
        ch = formula[i]
        if ch == '"':
            j, buf = i + 1, []
            while j < n and not (formula[j] == '"' and formula[j + 1:j + 2] != '"'):
                buf.append(formula[j])
                j += 2 if formula[j] == '"' else 1          # a doubled quote is one quote
            if depth == 0:
                out.append("".join(buf))
            i = j + 1
            continue
        depth += {"(": 1, ")": -1}.get(ch, 0)
        i += 1
    return "".join(out)


class _Shown:
    """A cell's displayed value as far as the plan and the given texts tell it."""

    def __init__(self, plan: SheetPlan, texts: dict[tuple[int, int], str]):
        self.plan, self.texts = plan, texts

    def text(self, rc):
        """The text a height is measured on: the given text, else a formula's own text, else the cell's value."""
        if rc in self.texts:
            return self.texts[rc]
        v = self.plan.cells.get(rc)
        return _static_text(v) if _formula(self.plan, rc) else v

    def empty(self, rc) -> bool:
        v = self.texts[rc] if rc in self.texts else self.plan.cells.get(rc)
        return v in (None, "")

    def numeric(self, rc) -> bool:
        if rc in self.texts:
            return False
        v = self.plan.cells.get(rc)
        return (isinstance(v, (int, float)) and not isinstance(v, bool)) or _formula(self.plan, rc)


def _text_height(txt, width: float) -> float:
    """A Key point row: 13 pt per wrapped line + 2, at least 15 (4 lines when a word does not fit)."""
    n = (tn.wrap_lines(str(txt), width) or 4) if txt else 1
    return max(15, n * 13 + 2)


# ------------------------------------------------------------------ the look
def answer_look(plan: SheetPlan, texts: dict[tuple[int, int], str] | None = None) -> SheetLook:
    """The look apply_report(ws, answer_roles(plan)) gives the sheet. texts: the displayed text of cells whose
    height depends on it (title, Key points); without it a Key point row gets the height of its formula's own text."""
    shown = _Shown(plan, texts or {})
    roles = answer_roles(plan)
    last = col_number(roles["last_col"])
    text_last = col_number(roles.get("print_last_col") or roles["last_col"])   # text rows end at the print area
    labels = sorted(col_number(c) for c in roles["label_cols"])
    first_label = labels[0]
    given = {c: w for c, w in styling._parse_widths(roles["col_widths"]).items() if c <= last}
    widths = {c: DEFAULT_WIDTH for c in range(1, last + 1)} | given
    text_width = sum(widths[c] for c in range(first_label, text_last + 1))
    rows = {int(r["row"]): r for r in roles["rows"]}
    last_row = max(max(rows, default=0), int(roles["last_row"]))
    codes = {col_number(c) for c in roles["code_cols"]}
    helpers = sorted(col_number(c) for c in roles.get("helper_cols", []))
    helper_set = set(helpers)
    bounded = bool(roles.get("print_last_col") or roles.get("print_footer")
                   or any("end_col" in x for x in roles["rows"]))

    cells: dict[tuple[int, int], CellLook] = {}
    heights: dict[int, float] = {}
    merges: list[tuple[int, int, int]] = []
    checks: list[tuple[int, int]] = []
    text_rows: list[int] = []
    for r in range(1, last_row + 1):
        spec = rows.get(r, {"kind": "blank"})
        kind = spec["kind"]
        span_end = col_number(spec["end_col"]) if spec.get("end_col") else (text_last if bounded else last)
        if kind in ROW_STYLE:
            st = ROW_STYLE[kind]
            if kind == "blank":
                row = {c: CellLook("TN Tom") for c in range(1, last + 1)}
            elif bounded:
                end = span_end if kind == "section" else text_last
                row = {c: CellLook(st if first_label <= c <= end and c not in helper_set else "TN Tom")
                       for c in range(1, last + 1)}
            else:
                row = {c: CellLook(st) for c in range(1, last + 1)}
            heights[r] = ROW_HEIGHT.get(kind, 15)
            txt = shown.text((r, first_label)) if kind == "title" else None
            if txt:                       # a long message title wraps inside the table width
                n = tn.wrap_lines(str(txt), text_width, 16, True) or 1
                if n > 1:
                    heights[r] = n * 21 + 4
                    merges.append((r, first_label, text_last))
        elif kind == "subhead":
            end = span_end if bounded else last
            row = {c: CellLook("TN Underseksjon" if first_label <= c <= end and c not in helper_set else "TN Tom")
                   for c in range(1, last + 1)}
            heights[r] = 15
        elif kind == "text":
            end = text_last if bounded else last
            row = {c: CellLook("TN Tekst" if first_label <= c <= end and c not in helper_set else "TN Tom")
                   for c in range(1, last + 1)}
            heights[r] = _text_height(shown.text((r, first_label)), text_width)
            if _formula(plan, (r, first_label)):
                text_rows.append(r)
            if all(shown.empty((r, c)) for c in range(first_label + 1, text_last + 1)):
                merges.append((r, first_label, text_last))
        elif kind == "header":
            nc = {col_number(c) for c in spec["num_cols"]} if "num_cols" in spec else \
                set(range(labels[-1] + 1, (span_end if bounded else last) + 1))
            if bounded:
                row = {c: CellLook(("TN Kolonne tall" if c in nc else "TN Kolonne")
                                   if first_label <= c <= span_end and c not in helper_set else "TN Tom")
                       for c in range(1, last + 1)}
            else:
                row = {c: CellLook("TN Kolonne tall" if c in nc else "TN Kolonne") for c in range(1, last + 1)}
            lines = 1
            for c in range(1, last + 1):
                t = shown.text((r, c))
                if t not in (None, ""):
                    lines = max(lines, tn.wrap_lines(str(t), widths[c], 9, True) or 4)
            heights[r] = max(26, lines * 12 + 3)
        elif kind in EMPH:
            row = _line_row(r, spec, last, first_label, labels, codes, shown, span_end, helper_set, bounded)
            heights[r] = 15
            if spec.get("status_col"):
                checks.append((r, col_number(spec["status_col"])))
        else:
            raise ValueError(f"row {r}: unknown kind {kind!r}")
        cells |= {(r, c): cl for c, cl in row.items()}

    if helpers:                           # chart helper columns: grey, they are not the answer
        for r in range(1, last_row + 1):
            for c in range(helpers[0], helpers[-1] + 1):
                cells[(r, c)] = replace(cells[(r, c)], font_color=GREY)
    for r, c1, c2 in merges:              # merged: the style again (it resets the cells), wrapped, text at the top
        for c in range(c1, c2 + 1):
            cells[(r, c)] = CellLook(cells[(r, c)].style, wrap=True, valign="top" if c == c1 else None)

    area_col = plan.print_area_col or plan.print_last_col or plan.last_col
    out = SheetLook(cells=cells, heights=heights, widths={col_letter(c): w for c, w in sorted(given.items())},
                    merges=sorted(merges), status_runs=_runs(checks), text_rows=text_rows,
                    print_area=f"$A$1:${area_col}${last_row}", breaks=[], footer=plan.footer)
    out.breaks = page_breaks(plan, out)
    return out


def page_breaks(plan: SheetPlan, look: SheetLook, heights: dict[int, float] | None = None) -> list[int]:
    """The rows that start a new printed page (styling.keep_breaks over the print area's width) for the look's row
    heights with heights laid over them: text_heights after the read-back gives the breaks of the calculated
    sheet. Every row of the block has a height in look.heights."""
    area_col = plan.print_area_col or plan.print_last_col or plan.last_col
    area_width = sum(width_pt(look.widths.get(col_letter(c), DEFAULT_WIDTH))
                     for c in range(1, col_number(area_col) + 1))
    return styling.keep_breaks(look.heights | (heights or {}), max(look.heights, default=0), keep_spans(plan),
                               area_width)


def _line_row(r: int, spec: dict, last: int, first_label: int, labels: list[int], codes: set[int],
              shown: _Shown, span_end: int, helper_set: set[int], bounded: bool) -> dict[int, CellLook]:
    """A table row (line, sum, grand, check): the number kind's style and format, the status column, the labels,
    the code column grey and left."""
    emph = EMPH[spec["kind"]]
    num = {col_number(k): v for k, v in (spec.get("num") or {}).items()}
    status = col_number(spec["status_col"]) if spec.get("status_col") else None
    label = tn.style_name("etikett", emph)
    row: dict[int, CellLook] = {}
    for c in range(1, last + 1):
        if c < first_label:
            row[c] = CellLook("TN Tom")
        elif c in num:
            k = num[c]
            fmt = styling.TALL1_THOUSANDS if k == "tall1" else styling.NUM_FORMATS.get(k)
            row[c] = CellLook(styling._style_for(k, emph), num_format=fmt)
        elif c == status:
            row[c] = CellLook("TN Tekst tallkolonne" if not emph else label)
        elif bounded and (c > span_end or c in helper_set):
            row[c] = CellLook("TN Tom")
        elif c in labels:
            row[c] = CellLook(label)
        elif shown.numeric((r, c)):
            row[c] = CellLook(styling._style_for("heltall", emph))
        else:
            row[c] = CellLook(label)
    for c in codes - set(num):            # the code column (member codes): grey, left-aligned
        if c in row:
            row[c] = replace(row[c], font_color=GREY, align="left")
    return row


def _runs(cells: list[tuple[int, int]]) -> list[tuple[int, int, int]]:
    """(row, col) check status cells -> (col, first row, last row) runs of consecutive rows, as apply_report sets
    the check colours."""
    out: list[list[int]] = []
    for r, c in sorted(set(cells), key=lambda x: (x[1], x[0])):
        if out and out[-1][0] == c and out[-1][2] == r - 1:
            out[-1][2] = r
        else:
            out.append([c, r, r])
    return [(c, a, b) for c, a, b in out]


def style_findings(plan: SheetPlan, block: tuple) -> list[str]:
    """The text findings of apply_report's lint (styling._lint_fast) for the drawn Answer, in its words: a merged
    text (a Key point, a wrapped title) that does not fit its row's height, and a text that is not right-aligned or
    wrapped and is wider than its column and the empty columns to its right. block: the read-back of A1:<last col>
    <last row> (Value2, as the lint read it); the look's styles, widths and merges, and its heights with the
    calculated Key points' (text_heights). Pure: no Excel call."""
    lk = answer_look(plan)
    heights = lk.heights | text_heights(plan, lk, block)
    roles = answer_roles(plan)
    last = col_number(roles["last_col"])
    widths = {c: lk.widths.get(col_letter(c), DEFAULT_WIDTH) for c in range(1, last + 1)}
    merged = {(r, c1): c2 for r, c1, c2 in lk.merges}
    defs = _defs()
    out: list[str] = []
    for r, row in enumerate(block[:plan.last_row], 1):
        for c, v in enumerate(row[:last], 1):
            if not isinstance(v, str) or not v or v.startswith("="):
                continue
            p = defs.get(lk.cells[(r, c)].style, {})
            size, bold, a = p.get("size", 10), bool(p.get("bold")), f"{col_letter(c)}{r}"
            if (r, c) in merged:
                lines = tn.wrap_lines(v, sum(widths[k] for k in range(c, merged[(r, c)] + 1)), size, bold)
                h = heights.get(r, 15)
                if lines is None or lines * (size + 3) > h + 1:
                    out.append(f"{a}: text does not fit ({lines or 'word too long'} lines, {h:.0f} pt)")
            elif p.get("h", tn.H_LEFT) != tn.H_RIGHT and not p.get("wrap"):
                avail = widths[c]
                for k in range(c + 1, last + 1):
                    if row[k - 1] not in (None, ""):
                        break
                    avail += widths[k]
                need = tn.text_units(v, size, bold)
                if need > avail:
                    out.append(f"{a}: text is cut off or spills over (~{need:.0f} > {avail:.0f} chars)")
    return out


def text_heights(plan: SheetPlan, look: SheetLook, block: tuple) -> dict[int, float]:
    """The final heights of the Key point rows from the block read back from Excel (Value2 of the rectangle from
    A1): only the rows whose height differs from look.heights."""
    roles = answer_roles(plan)
    first = sorted(col_number(c) for c in roles["label_cols"])[0]
    text_last = col_number(roles.get("print_last_col") or roles["last_col"])
    width = sum(look.widths.get(col_letter(c), DEFAULT_WIDTH) for c in range(first, text_last + 1))
    out = {}
    for r in look.text_rows:
        row = block[r - 1] if r <= len(block) else ()
        h = _text_height(row[first - 1] if first <= len(row) else None, width)
        if h != look.heights[r]:
            out[r] = h
    return out
