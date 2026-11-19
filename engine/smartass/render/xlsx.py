# engine/smartass/render/xlsx.py
"""An answer's sheets built as one .xlsx file, without Excel: Q<n> Answer (its look only: every cell's format,
widths, heights, merges, the check colours, the page setup, and its charts (render/xlsx_charts.py); no contents),
Q<n> Data (the provenance header's look, the facts as values) and one Q<n> Lines <alias> per line set (title, cube
value, the lines as values). The file holds no Excel Table: a sheet that contains one cannot be copied between
workbooks, so the tables are created after the sheets are copied into the user's workbook, and the formulas (they
name the tables) are written then too, each block with one Range.Formula2. build returns both as TableSpec and
FormulaBlock, and the charts drawn with their lint and label fixes (BuiltFile.charts, chart_lint, label_fixes).

Every text is written with write_string (XlsxWriter's string conversions are off), so 'Sep 26', '0047', 'TRUE' or
'=SUM(A1)' stay text. XlsxWriter has no named cell styles: every cell gets a direct format (look.style_props plus the
CellLook overrides), one format object per distinct combination. Values are today's COM writers' (render/excel.py,
xl/drill_xl.py, pipeline._write_lines/_lines_header); docs/audit/2026-11-19-render-map.md sections 2, 4.3, 4.4, 6.1."""
from __future__ import annotations

import datetime as dt
import math
import re
from dataclasses import dataclass, field
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_cell_to_rowcol

from ..core import config
from ..epm.drill_lines import KEY_LABELS
from ..model.lines import CUBE_CELL
from ..xl import styling
from ..xl.drill_xl import _EXCEL_EPOCH, AMOUNT_LABELS, AMT_FMT, DATE_LABELS, TEXT_LABELS, WIDTHS, header_width
from ..xl.vendor import tn_rapportstil as tn
from . import look, xlsx_charts
from .datasheet import LINES_HEADER, Header
from .excel import HEADER_COLS, HEADER_MIN, LABEL_MAX, SHEET_FOOTER, VALUE_FORMAT, _text, answer_grid
from .layout import SheetPlan, col_letter, col_number

MAX_ROWS = 1_048_576                    # rows of an Excel sheet
ANSWER_TAB, DATA_TAB = "#1F3864", "#A6A6A6"
PRINT_HEADER = '&"Arial"&8Intern'       # styling.print_setup's CenterHeader
MARGIN = 1.5 / 2.54                     # 1.5 cm left and right, in inches
DATE_FMT = "dd.mm.yyyy"                 # drill_xl._local_formats' date format as an en-US code
VALUE_BLUE = "#0000C0"                  # render.excel.VALUE_BLUE (BGR 0xC00000)
LINES_ROW = 4                           # the lines table header row (drill_xl.write_lines, pipeline.LINES_TABLE_ROW)
OPTIONS = {"strings_to_formulas": False, "strings_to_numbers": False, "strings_to_urls": False}
INK, GREY, NAVY = look._hex(tn.INK), look._hex(tn.GREY), look._hex(tn.NAVY)
OK_LOOK = {"font_color": look._hex(styling.GREEN), "bold": True}
NOT_OK_LOOK = {"font_color": look._hex(tn.BREACH_FONT), "bold": True, "bg_color": look._hex(tn.BREACH_FILL)}


class TooManyRows(ValueError):
    """A sheet would need more rows than Excel has."""


@dataclass
class TableSpec:
    sheet: str
    ref: str                            # "A14:J120": the header row included
    name: str
    stripes: bool                       # row stripes: False for Data, True for Lines (as today)


@dataclass
class FormulaBlock:
    sheet: str
    ref: str
    grid: list[list]                    # written with ONE Range.Formula2 after the tables exist (text with its ')


@dataclass
class DataSheet:
    name: str
    header: Header
    grid: list[tuple]                   # data_grid(frame): the column row, then one row per fact


@dataclass
class LinesSheet:
    name: str
    table: str
    title: str
    source: str
    columns: list[str]
    rows: list[dict]
    cube_value: float | None
    header_block: list[list]            # the F1:K3 rectangle _lines_header writes today (text with its ')


@dataclass
class BuiltFile:
    path: Path
    sheets: list[str]
    tables: list[TableSpec]
    formulas: list[FormulaBlock]
    charts: list[str] = field(default_factory=list)         # the titles of the charts drawn (the reply's charts)
    chart_lint: list[str] = field(default_factory=list)     # plan.chart_lint and the charts that failed
    label_fixes: list[xlsx_charts.LabelFix] = field(default_factory=list)   # set by COM after the move-in


def check_rows(sheet: str, last_row: int) -> None:
    if last_row > MAX_ROWS:
        need = f"{last_row:,}".replace(",", " ")
        raise TooManyRows(f"sheet {sheet!r} would need {need} rows; an Excel sheet has 1 048 576")


def value_text(v) -> str:
    """A Value cell as VALUE_FORMAT shows it ('#,##0;(#,##0);"–"'): what the column width is measured on."""
    if v is None or v == "":
        return ""
    if not isinstance(v, (int, float)) or isinstance(v, bool):
        return str(v)
    if v == 0:
        return "–"
    shown = f"{abs(v):,.0f}"
    return f"({shown})" if v < 0 else shown


def _lines_cols(ls: LinesSheet) -> list[str]:
    """Above config.DRILL_XL_FULL lines only the key columns go to Excel (drill_xl.write_lines)."""
    if len(ls.rows) <= config.DRILL_XL_FULL:
        return list(ls.columns)
    return [c for c in ls.columns if c in KEY_LABELS]


def _table_end(first: int, body: int) -> int:
    """The last row of a table whose header row is 'first': a table keeps at least one body row."""
    return first + max(body, 1)


def build(path: Path, answer_name: str, plan: SheetPlan, data: DataSheet, lines: list[LinesSheet]) -> BuiltFile:
    """Write the answer's sheets to path, in the order Answer, Data, the Lines sheets. Refuses (TooManyRows,
    ValueError) before any file is written. Returns the tables and formula blocks to create after the move-in."""
    m = re.fullmatch(r"Q(\d+) Answer", answer_name)
    if m is None:
        raise ValueError(f"answer sheet {answer_name!r}: the name must be 'Q<n> Answer'")
    q = int(m.group(1))
    check_rows(data.name, _table_end(data.header.table_row, len(data.grid) - 1))
    for ls in lines:
        check_rows(ls.name, _table_end(LINES_ROW, len(ls.rows)))
    tables: list[TableSpec] = []
    formulas: list[FormulaBlock] = []
    with xlsxwriter.Workbook(str(path), OPTIONS) as wb:
        fmt = _Formats(wb)
        formulas.append(_answer(wb, fmt, answer_name, plan))
        charts, chart_lint, label_fixes = xlsx_charts.add_charts(wb, wb.get_worksheet_by_name(answer_name),
                                                                 answer_name, plan)
        table, block = _data(wb, fmt, q, data)
        tables.append(table)
        formulas.append(block)
        for ls in lines:
            table, block = _lines(wb, fmt, ls)
            tables.append(table)
            formulas.append(block)
    return BuiltFile(path=Path(path), sheets=[answer_name, data.name, *(ls.name for ls in lines)], tables=tables,
                     formulas=formulas, charts=charts, chart_lint=chart_lint, label_fixes=label_fixes)


class _Formats:
    """One XlsxWriter format per distinct set of properties."""

    def __init__(self, wb):
        self.wb, self.cache = wb, {}

    def __call__(self, props: dict):
        key = tuple(sorted(props.items()))
        if key not in self.cache:
            self.cache[key] = self.wb.add_format(dict(props))
        return self.cache[key]

    def look(self, cl: look.CellLook):
        p = look.style_props(cl.style)
        for name, value in (("num_format", cl.num_format), ("font_color", cl.font_color), ("align", cl.align),
                            ("text_wrap", cl.wrap), ("valign", cl.valign)):
            if value is not None:
                p[name] = value
        return self(p)


def _page(ws, area: str, footer: str, breaks=(), title_row: int | None = None) -> None:
    """styling.print_setup in the file: landscape, centred, 1.5 cm left and right, the Intern header, the footer
    centred. Without title_row one page wide; with it (a Lines sheet) 100 % and that row on every page."""
    ws.print_area(area)
    ws.set_landscape()
    if title_row:
        ws.repeat_rows(title_row - 1)
    else:
        ws.fit_to_pages(1, 0)
    ws.center_horizontally()
    ws.set_margins(left=MARGIN, right=MARGIN)
    ws.set_header("&C" + PRINT_HEADER)
    if footer:
        ws.set_footer("&C" + footer)
    if breaks:
        ws.set_h_pagebreaks([r - 1 for r in breaks])          # XlsxWriter: the break below that (0-based) row


def _row_height(ws, r: int, points: float) -> None:
    """A fixed row height in points, as apply_report sets it (RowHeight: a custom height, so the row never fits
    itself to its font). set_row would round the points to whole pixels (28 -> 27.75), and XlsxWriter writes no
    height for a row at the sheet's default (15 pt); a pixel value one float step above the default is written as
    ht="15" customHeight="1"."""
    px = points / 0.75
    if px == ws.default_row_height:
        px = math.nextafter(px, math.inf)
    ws.set_row_pixels(r - 1, px)


def _answer(wb, fmt: _Formats, name: str, plan: SheetPlan) -> FormulaBlock:
    """The Answer's look (look.answer_look) on every cell of its block, no contents; its formulas as one block."""
    ws = wb.add_worksheet(name)
    lk = look.answer_look(plan)
    for (r, c), cl in lk.cells.items():
        ws.write_blank(r - 1, c - 1, None, fmt.look(cl))
    for col, w in lk.widths.items():
        n = col_number(col) - 1
        ws.set_column(n, n, w)
    for r, h in lk.heights.items():
        _row_height(ws, r, h)
    for r, c1, c2 in lk.merges:                                  # the first cell's look; the others keep their own
        ws.merge_range(r - 1, c1 - 1, r - 1, c2 - 1, "", fmt.look(lk.cells[(r, c1)]))
        for c in range(c1 + 1, c2 + 1):
            ws.write_blank(r - 1, c - 1, None, fmt.look(lk.cells[(r, c)]))
    for c, a, b in lk.status_runs:
        for criteria, props in (("==", OK_LOOK), ("!=", NOT_OK_LOOK)):
            ws.conditional_format(a - 1, c - 1, b - 1, c - 1, {"type": "cell", "criteria": criteria,
                                                               "value": '"OK"', "format": fmt(props)})
    ws.set_tab_color(ANSWER_TAB)
    ws.hide_gridlines(2)
    _page(ws, lk.print_area, lk.footer, lk.breaks)
    r0, c0, g = answer_grid(plan)
    ref = f"{col_letter(c0)}{r0}:{col_letter(c0 + len(g[0]) - 1)}{r0 + len(g) - 1}"
    return FormulaBlock(name, ref, g)


def _area(last_col: int, last_row: int) -> str:
    return "$A$1:$" + col_letter(last_col) + "$" + str(last_row)


def _data(wb, fmt: _Formats, q: int, data: DataSheet) -> tuple[TableSpec, FormulaBlock]:
    """Q<n> Data: the header block's look (its contents formulas on the table: a block), the facts as values at
    A<table_row> (render.excel.write_data's look), the widths from the texts, frozen below the table header."""
    ws = wb.add_worksheet(data.name)
    header, g = data.header, data.grid
    h, r0 = len(header), header.table_row
    numeric = {i + 1 for i, (lab, _) in enumerate(header) if lab in ("Control total", "Tolerance (NOK)")}
    base = {"font_name": tn.FONT, "font_size": 9, "font_color": INK}
    for r in range(1, h + 1):
        for c in range(1, HEADER_COLS + 1):
            p = dict(base)
            if (r, c) == (1, 1):
                p |= {"bold": True, "font_size": 14, "font_color": NAVY}
            elif c == 1:
                p["font_color"] = GREY
            if r in numeric and c in (2, 3):
                p["num_format"] = VALUE_FORMAT
            ws.write_blank(r - 1, c - 1, None, fmt(p))
    w = len(g[0])
    table = {"font_name": tn.FONT, "font_size": 9}
    head = [fmt(table | {"bold": True} | ({"num_format": "@"} if j < w - 1 else {})) for j in range(w)]
    label = fmt(table | {"num_format": "@"})
    value = fmt(table | {"num_format": VALUE_FORMAT, "font_color": VALUE_BLUE, "align": "right"})
    widths = [0.0] * w
    for j, name in enumerate(g[0]):
        ws.write_string(r0 - 1, j, str(name), head[j])
        widths[j] = tn.text_units(str(name), 9)
    for i, row in enumerate(g[1:], 1):
        for j, v in enumerate(row):
            last = j == w - 1
            f = value if last else label
            if v is None or v == "":
                ws.write_blank(r0 - 1 + i, j, None, f)
                continue
            if last and isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.write_number(r0 - 1 + i, j, v, f)
            else:
                ws.write_string(r0 - 1 + i, j, str(v), f)
            widths[j] = max(widths[j], tn.text_units(value_text(v) if last else str(v), 9))
    end = _table_end(r0, len(g) - 1)
    for i in range(len(g), end - r0 + 1):                         # no facts: the empty body row, formatted
        for j in range(w):
            ws.write_blank(r0 - 1 + i, j, None, value if j == w - 1 else label)
    for j, width in enumerate(widths):                            # AutoFit over the table, at most 40, A:D at least 16
        fit = min(width, LABEL_MAX)
        ws.set_column(j, j, max(fit, HEADER_MIN) if j < HEADER_COLS else fit)
    ws.freeze_panes(r0, 0)
    ws.set_tab_color(DATA_TAB)
    ws.hide_gridlines(2)
    _page(ws, _area(max(HEADER_COLS, w), end), SHEET_FOOTER)
    return (TableSpec(data.name, f"A{r0}:{col_letter(w)}{end}", f"Data_Q{q}", stripes=False),
            FormulaBlock(data.name, f"A1:{col_letter(HEADER_COLS)}{h}", _header_grid(header)))


def _header_grid(header: Header) -> list[list]:
    """The header rows as render.excel._write_header writes them: the title text in A1, label in A, values from B;
    text with its leading apostrophe."""
    g = [["", "", "", ""] for _ in header]
    for i, (lab, value) in enumerate(header):
        if i == 0:
            g[0][0] = _text(value)
            continue
        g[i][0] = _text(lab)
        for j, x in enumerate(value if isinstance(value, tuple) else (value,)):
            g[i][1 + j] = _text(x)
    return g


def _line_value(ws, r: int, c: int, label: str, v, f) -> None:
    """One line value as drill_xl._cell gives it: a datetime as a serial number (time zone dropped), text and the
    TEXT_LABELS columns as text, numbers as numbers."""
    if v is None:
        ws.write_blank(r, c, None, f)
    elif isinstance(v, dt.datetime):
        ws.write_number(r, c, (v.replace(tzinfo=None) - _EXCEL_EPOCH) / dt.timedelta(days=1), f)
    elif isinstance(v, dt.date):
        ws.write_number(r, c, (dt.datetime(v.year, v.month, v.day) - _EXCEL_EPOCH) / dt.timedelta(days=1), f)
    elif isinstance(v, str) or label in TEXT_LABELS:
        ws.write_string(r, c, str(v), f)
    elif isinstance(v, bool):
        ws.write_boolean(r, c, v, f)
    elif isinstance(v, (int, float)):
        ws.write_number(r, c, v, f)
    else:
        ws.write_string(r, c, str(v), f)


def _lines(wb, fmt: _Formats, ls: LinesSheet) -> tuple[TableSpec, FormulaBlock]:
    """Q<n> Lines <alias>: A1 title (bold), A2 source, the cube value beside its label, the column names in row 4,
    the lines from row 5 (drill_xl.write_lines + pipeline._write_lines); F1:K3's look, its contents a block."""
    ws = wb.add_worksheet(ls.name)
    cols = _lines_cols(ls)
    ws.write_string(0, 0, ls.title, fmt({"bold": True}))
    ws.write_string(1, 0, ls.source, fmt({"font_color": "#000000"}))
    cr, cc = xl_cell_to_rowcol(CUBE_CELL)
    ws.write_string(cr, cc - 1, "Cube value")
    amount = fmt({"num_format": AMT_FMT})
    if ls.cube_value is None:
        ws.write_blank(cr, cc, None, amount)
    else:
        ws.write_number(cr, cc, ls.cube_value, amount)
    first, last = LINES_HEADER.split(":")
    (r1, c1), (r2, c2) = xl_cell_to_rowcol(first), xl_cell_to_rowcol(last)
    base = {"font_name": tn.FONT, "font_size": 9, "font_color": INK}
    grey = {(r, c1) for r in range(r1, r2 + 1)} | {(r2, c1 + 2), (r2, c1 + 4)}       # F1:F3, H3, J3
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            p = base | ({"font_color": GREY} if (r, c) in grey else {})
            if (r, c) == (r2, c1 + 1):                                               # G3: the lines total
                p["num_format"] = VALUE_FORMAT
            ws.write_blank(r, c, None, fmt(p))
    for j, c in enumerate(cols):
        ws.write_string(LINES_ROW - 1, j, c)
    date = fmt({"num_format": DATE_FMT})
    col_fmt = [amount if c in AMOUNT_LABELS else date if c in DATE_LABELS else None for c in cols]
    for i, row in enumerate(ls.rows):
        for j, c in enumerate(cols):
            _line_value(ws, LINES_ROW + i, j, c, row.get(c), col_fmt[j])
    for j, c in enumerate(cols):
        ws.set_column(j, j, max(WIDTHS.get(c, 14), header_width(c)))
    ws.freeze_panes(LINES_ROW, 0)
    ws.set_tab_color(DATA_TAB)
    end = _table_end(LINES_ROW, len(ls.rows))
    _page(ws, _area(max(len(cols), c2 + 1), end), SHEET_FOOTER, title_row=LINES_ROW)
    return (TableSpec(ls.name, f"A{LINES_ROW}:{col_letter(len(cols))}{end}", ls.table, stripes=True),
            FormulaBlock(ls.name, LINES_HEADER, ls.header_block))
