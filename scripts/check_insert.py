r"""Live proof for light and fast phase 3b, Task 2a: the built file's sheets inserted into the workbook with ONE
Sheets.Add(Type:=<path of the built .xlsx>) give the same workbook as the Workbooks.Open + Worksheets.Copy + Close
+ BreakLink they replaced. Everything runs in a hidden Excel of this script's own (DispatchEx, refused unless its pid is new;
check_fastcom.start_excel) and nothing is saved: the targets are copies of the new-workbook seed in a temporary
folder, and SMARTASS_HOME points there too.

The copy route is kept here (copy_route: the move-in as it was before the switch, the links broken by the file's
name); the insert route is the engine's movein.move_in. No EPM: the answers are built from the plans and fixed data
of engine/tests/v2 (test_look.CASES, an answer with every chart kind and thin waterfall steps, a matrix with kpi
tiles and Key points, the Lines sheet of test_xlsx), and the whole jobs from pipeline.prepare over the tests' fake
provider.

1. Per fixture: the file is built once and copied byte for byte under two names; a fresh copy of the same target
   (user sheets around, a hidden and a very hidden one) gets it by the copy route, another by the insert route, and
   a third by the copy route again (the control: what differs between two copies is noise, never a finding). Both
   are compared right after the move-in and again after movein.finish + calculate (tables, formula blocks, the
   label fixes), on: cell values, Formula2, number formats, fonts, fills, borders (all four edges), alignment, wrap
   and indent, merged areas, row heights, column widths, hidden rows / columns, page setup, manual page breaks, the
   window (freeze panes, zoom, gridlines, scroll), the sheet (visibility, tab colour, order), the defined names
   (name, RefersTo, scope), conditional formats, data validation, the styles, the charts (type, position, size,
   title, legend, groups, axes, every series' formula, fills, data labels and every labelled point) and the links
   (LinkSources of every type empty, no formula, name or series formula mentioning the file).
2. Where the sheets land (after a named sheet, after the last, after a hidden / very hidden one, first, before the
   first visible), what is active afterwards, a taken name (refused before the insert; what Excel itself would
   do), a file holding a table (refused), a missing or broken file (what Excel leaves), the temporary file deleted
   right after the insert, and the Excel around it (calculation, the active workbook, windows, the recent list,
   the undo state, visible windows of the process).
3. Whole jobs (exceljob.excel_job, a second question with a Lines sheet, then exceljob.redraw_job: a follow-up's
   replacement), since Task 2b the job's batches (xl/inproc) once through the COM executor (side 'copy', and
   'control' for the noise) and once through the in-process add-in (side 'insert'), each in a workbook of its
   own, compared sheet by sheet; sheet = model, the warnings and layout findings per step, the draw's seconds and
   batches. (A quiet() around the routes is gone with Task 2b: the parts above put the four settings off and back
   as a batch does, with no landing.)
   Failures (Task 2b): a follow-up and a new answer whose finish batch fails, through both executors - the
   workbook, the four settings and the active sheet as they were before, no 'old ...' sheet left.
   Labels (Task 2b): every fixed chart label's Top in the job's order against the standalone one, and where Excel
   itself puts the lifted ones (printed as a measurement).
   Re-pointing (Task 3): a summary sheet's formulas pointing at Q1 (a cell, a table column, a cross-sheet SUM, the
   Answer, a legacy array, a dynamic array, defined names - one with a function and a decimal in the user's
   notation -, a chart series, a chart sheet) and a cell on Q2's Answer survive two follow-ups and a refresh-style
   redraw through both executors - as the user wrote them, showing their targets' values, no link, no 'old ...'
   sheet - and a follow-up whose drop fails points them back; with a merged cell and a table's calculated column.
   --parts repoint-bench (not in the default parts): a follow-up with 500 cells pointing at it against none, both
   executors (the close-out's writes run with Calculation manual).
   New workbook (phase 3c, --parts newbook): the first answer of a new workbook made today's way (Workbooks.Add of
   the TN template, label, SaveAs, the insert, the Sources log furnished over COM) against one whose built file is
   the workbook (Excel only opens it), each followed by a second answer with a Lines sheet and a follow-up, through
   the add-in, today's way twice (the noise); compared item by item, the expected differences named (EXPECTED_NEW),
   the files the jobs saved labelled (a first answer saved in the job, each later one's save after the reply waited
   for: xl/inproc/later.py); and a built
   new workbook whose finish fails is closed unsaved and its file deleted.
4. Timing: the move-in alone, both routes, --rounds rounds each.

Prints OK / DIFF per item (never a cell value: addresses and property names only); exit 0 only when all are OK.

Usage (from the repo root):
    engine\.venv\Scripts\python.exe scripts\check_insert.py [--rounds 5]
"""
from __future__ import annotations

import argparse
import contextlib
import dataclasses
import datetime as dt
import logging
import math
import os
import re
import shutil
import statistics
import sys
import tempfile
import threading
import time
import uuid
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TMP = Path(tempfile.mkdtemp(prefix="check_insert_"))
os.environ["SMARTASS_HOME"] = str(TMP / "home")          # before smartass.core.config is imported
SEED = Path(os.environ.get("LOCALAPPDATA", "")) / "smartass" / "new_workbook_seed.xlsx"   # read, never written
sys.path[:0] = [str(ROOT / "engine" / "tests"), str(ROOT / "engine" / "tests" / "v2")]   # the fixtures

import pythoncom  # noqa: E402
import pywintypes  # noqa: E402
import xlsxwriter  # noqa: E402

from check_fastcom import QUIT_WAIT_S, _wait_gone, kill_if_ours, start_excel  # noqa: E402
from smartass import exceljob, pipeline  # noqa: E402
from smartass.core import config  # noqa: E402
from smartass.render import xlsx  # noqa: E402
from smartass.render.layout import ChartSpec  # noqa: E402
from smartass.xl import fastcom, jobs, movein  # noqa: E402
from smartass.xl.hostproc import windows_of  # noqa: E402
from smartass.xl.inproc import later, runner  # noqa: E402
from smartass.xl.inproc.batch import MISSING, Batch  # noqa: E402
from smartass.xl.sheetorder import ANSWER_TAB, DATA_TAB  # noqa: E402
from smartass.xl.target import Target  # noqa: E402

ERR = "<error>"
XL_NONE = -4142
XL_VALUE, XL_CATEGORY = 2, 1
EN_US = 0x409
XML_SPREADSHEET = 11
SS = "{urn:schemas-microsoft-com:office:spreadsheet}"
LINK_TYPES = (1, 2, 5, 6)                 # xlExcelLinks, xlOLELinks, xlPublishers, xlSubscribers
MANUAL_BREAK = -4135
ALL_VALIDATION = -4174                    # xlCellTypeAllValidation
STAMP = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}(?::\d{2})?(?: ?[+-]\d{2}:?\d{2})?")
USER_SHEETS = ("Intro", "Notes", "Sources", "_svmeta", "_vh")    # the target: _svmeta hidden, _vh very hidden
results: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> bool:
    results.append((name, ok, detail))
    print(f"{'OK  ' if ok else 'DIFF'} {name}{'  ' + detail if detail else ''}", flush=True)
    return ok


def info(text: str) -> None:
    print(f"     {text}", flush=True)


# ---------------------------------------------------------------------------------------------------- the routes
def _break_links(wb, file_name: str) -> None:
    for link in wb.LinkSources(1) or ():
        if re.split(r"[\\/]", str(link))[-1].casefold() == file_name.casefold():
            wb.BreakLink(link, 1)


def copy_route(t, built, after, first=False, book=None, **_):
    """The move-in before Task 2a: the names checked, the file opened read-only, its sheets copied in one
    Worksheets.Copy, the file closed, the links the copy left broken."""
    book = book or movein.read_book(t)
    found = movein.taken(book, built.sheets, [s.name for s in built.tables], [])
    if found:
        raise movein.NameTaken(", ".join(found))
    tables = movein.file_parts(built.path)[1]
    if tables:
        raise ValueError(f"the built file holds an Excel Table ({', '.join(tables)})")
    before, after_sheet = movein._anchor(book, after, first)
    path, names = str(built.path), list(built.sheets)
    ws = lambda s: t.wb.Worksheets(s.name)    # noqa: E731
    src = jobs.com(lambda: t.app.Workbooks.Open(path, 0, True), limit=120)
    try:
        t.sheets_added.extend(names)
        if before is not None:
            jobs.com(lambda: src.Worksheets(names).Copy(ws(before)))
        else:
            jobs.com(lambda: src.Worksheets(names).Copy(None, ws(after_sheet)))
    finally:
        jobs.com(lambda: src.Close(False), limit=30, check=False)
    k = book.sheets.index(before) if before is not None else book.sheets.index(after_sheet) + 1
    book.sheets[k:k] = [movein.Sheet(n, -1, ANSWER_TAB if i == 0 else DATA_TAB) for i, n in enumerate(names)]
    jobs.com(lambda: _break_links(t.wb, built.path.name))



MOVE_IN = movein.move_in                 # the engine's (the jobs below put a route in its place)


def insert_route(t, built, after, first=False, book=None, **_):
    """The move-in as the engine does it now: one Sheets.Add(Type:=path)."""
    return MOVE_IN(t, built, after, first, book=book)


ROUTES = {"copy": copy_route, "insert": insert_route, "control": copy_route}


# ---------------------------------------------------------------------------------------------------- reading
def g(fn):
    """One read; ERR when Excel refuses it (recorded, never fatal)."""
    try:
        return _plain(jobs.retry_busy(fn, 20))
    except Exception:   # noqa: BLE001
        return ERR


def obj(fn):
    """A COM object to read further from (never through _plain, which would make it text); None when Excel refuses
    it, so every read from it is ERR."""
    try:
        return jobs.retry_busy(fn, 20)
    except Exception:   # noqa: BLE001
        return None


def _plain(v):
    if isinstance(v, float) and math.isnan(v):          # a label of an empty point: NaN Top/Left (NaN != NaN)
        return "NaN"
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, tuple):
        return tuple(_plain(x) for x in v)
    if isinstance(v, (bool, int, float, str)) or v is None:
        return v
    return str(v)


def col_letter(c: int) -> str:
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def col_number(s: str) -> int:
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def a1(r: int, c: int) -> str:
    return f"{col_letter(c)}{r}"


def parse_rect(addr: str) -> tuple[int, int, int, int]:
    parts = addr.replace("$", "").split(":")
    out = []
    for p in parts:
        m = re.fullmatch(r"([A-Z]+)(\d+)", p)
        out.append((int(m.group(2)), col_number(m.group(1))))
    (r1, c1), (r2, c2) = out[0], out[-1]
    return r1, c1, r2, c2


def rect_text(r) -> str:
    return f"{a1(r[0], r[1])}:{a1(r[2], r[3])}"


def union(a, b):
    if a is None:
        return b
    if b is None:
        return a
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


class _Raw:
    """Property reads straight on the IDispatch by a cached DISPID, in en-US (scripts/compare_renders.py, 84f1d6f)."""

    def __init__(self):
        self.ids = {}

    def __call__(self, disp, kind: str, name: str, *args):
        dispid = self.ids.get((kind, name))
        if dispid is None:
            dispid = self.ids[(kind, name)] = disp.GetIDsOfNames(name)
        return disp.Invoke(dispid, EN_US, pythoncom.DISPATCH_PROPERTYGET, True, *args)

    def item(self, disp, index):
        return disp.Invoke(0, EN_US, pythoncom.DISPATCH_PROPERTYGET, True, index)


class _Rect:
    def __init__(self, raw: _Raw, ws, addr: str):
        self.raw, self.rng, self.subs = raw, jobs.retry_busy(lambda: raw(ws, "Worksheet", "Range", addr), 20), {}

    def get(self, name: str):
        return jobs.retry_busy(lambda: self.raw(self.rng, "Range", name), 20)

    def sub(self, obj: str, name: str):
        if obj not in self.subs:
            self.subs[obj] = self.get(obj)
        return jobs.retry_busy(lambda: self.raw(self.subs[obj], obj, name), 20)


def _fill(x: _Rect):
    pattern = x.sub("Interior", "Pattern")
    if pattern is None:
        return None
    if pattern == XL_NONE:
        return "none"
    colour = x.sub("Interior", "Color")
    return None if colour is None else f"{int(pattern)}:{int(colour)}"


def _style(x: _Rect):
    s = x.get("Style")
    return None if s is None else str(jobs.retry_busy(lambda: x.raw(s, "Style", "Name"), 20))


CELL_PROPS = {        # item -> {property: reader}
    "number formats": {"number format": lambda x: x.get("NumberFormat")},
    "fonts": {"font name": lambda x: x.sub("Font", "Name"), "font size": lambda x: x.sub("Font", "Size"),
              "bold": lambda x: x.sub("Font", "Bold"), "italic": lambda x: x.sub("Font", "Italic"),
              "underline": lambda x: x.sub("Font", "Underline"), "font colour": lambda x: x.sub("Font", "Color")},
    "fills": {"fill": _fill},
    "alignment and wrap": {"horizontal": lambda x: x.get("HorizontalAlignment"),
                           "vertical": lambda x: x.get("VerticalAlignment"), "wrap": lambda x: x.get("WrapText"),
                           "indent": lambda x: x.get("IndentLevel"), "shrink": lambda x: x.get("ShrinkToFit"),
                           "orientation": lambda x: x.get("Orientation")},
    "merged areas": {"merged": lambda x: x.get("MergeCells")},
    "cell styles": {"style": _style},
}
PROP_ITEM = {p: item for item, props in CELL_PROPS.items() for p in props}
ALL_PROPS = {p: fn for props in CELL_PROPS.values() for p, fn in props.items()}


def _read_prop(fn, x):
    try:
        return fn(x)
    except Exception:   # noqa: BLE001
        return ERR


def _walk(raw: _Raw, ws, rect, props: dict, out: dict) -> None:
    """Each property over rect: one read when the rect is uniform (a mixed range reads None), else split in two
    until it is; out[prop] gets one entry per cell."""
    r1, c1, r2, c2 = rect
    single = r1 == r2 and c1 == c2
    try:
        x = _Rect(raw, ws, rect_text(rect))
    except Exception:   # noqa: BLE001
        x = None
    rest = {}
    for name, fn in props.items():
        v = ERR if x is None else _read_prop(fn, x)
        if v is None and not single or (v == ERR and not single):
            rest[name] = fn
        else:
            v = _plain(v)
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    out[name][a1(r, c)] = v
    if not rest:
        return
    if c2 > c1:
        m = (c1 + c2) // 2
        halves = ((r1, c1, r2, m), (r1, m + 1, r2, c2))
    else:
        m = (r1 + r2) // 2
        halves = ((r1, c1, m, c2), (m + 1, c1, r2, c2))
    for h in halves:
        _walk(raw, ws, h, rest, out)


def _style_borders(styles) -> dict:
    own, parent = {}, {}
    for st in styles:
        sid = st.get(SS + "ID")
        parent[sid] = st.get(SS + "Parent")
        b = st.find(SS + "Borders")
        if b is not None:
            own[sid] = {x.get(SS + "Position"): ":".join(x.get(SS + k, "") for k in ("LineStyle", "Weight", "Color"))
                        for x in b.findall(SS + "Border")}
    out = {}
    for sid in parent:
        s, seen = sid, set()
        while s is not None and s not in own and s not in seen:
            seen.add(s)
            s = parent.get(s)
        out[sid] = own.get(s, {})
    return out


def _borders(raw: _Raw, ws, rect) -> dict:
    """Every cell's own four borders, from the range as Excel's XML Spreadsheet (one read; compare_renders' way)."""
    import xml.etree.ElementTree as ET
    r1, c1, r2, c2 = rect
    try:
        root = ET.fromstring(jobs.retry_busy(lambda: raw(raw(ws, "Worksheet", "Range", rect_text(rect)), "Range",
                                                          "Value", XML_SPREADSHEET), 20))
    except Exception:   # noqa: BLE001
        return {"read": ERR}
    borders = _style_borders(root.iter(SS + "Style"))
    table = next(root.iter(SS + "Table"))
    width = c2 - c1 + 1
    col_style = []
    for col in table.findall(SS + "Column"):
        i = int(col.get(SS + "Index", len(col_style) + 1))
        col_style += ["Default"] * (i - 1 - len(col_style))
        col_style += [col.get(SS + "StyleID", "Default")] * (1 + int(col.get(SS + "Span", 0)))
    col_style += ["Default"] * (width - len(col_style))
    rows, r = {}, 0
    for row in table.findall(SS + "Row"):
        r = int(row.get(SS + "Index", r + 1))
        span = 1 + int(row.get(SS + "Span", 0))
        styles = [row.get(SS + "StyleID") or col_style[j] for j in range(width)]
        c = 0
        for cell in row.findall(SS + "Cell"):
            c = int(cell.get(SS + "Index", c + 1))
            sid = cell.get(SS + "StyleID", row.get(SS + "StyleID") or col_style[c - 1])
            across = int(cell.get(SS + "MergeAcross", 0))
            for j in range(c, min(c + across, width) + 1):
                styles[j - 1] = sid
            c += across
        for k in range(span):
            rows[r + k] = styles
        r += span - 1
    out = {}
    for i in range(1, r2 - r1 + 2):
        styles = rows.get(i) or col_style
        for j, sid in enumerate(styles, 1):
            b = borders.get(sid, {})
            edges = tuple(b.get(p, "none") for p in ("Top", "Bottom", "Left", "Right"))
            if edges != ("none",) * 4:
                out[a1(r1 + i - 1, c1 + j - 1)] = edges
    return out


def _bisect(get_value, a: int, b: int, out: dict, key) -> None:
    v = get_value(a, b)
    if v is not None and (v != ERR or a == b):
        out[key(a, b)] = round(v, 2) if isinstance(v, float) else v
        return
    m = (a + b) // 2
    _bisect(get_value, a, m, out, key)
    _bisect(get_value, m + 1, b, out, key)


PAGE = ("PrintArea", "Orientation", "PaperSize", "Zoom", "FitToPagesWide", "FitToPagesTall", "PrintTitleRows",
        "PrintTitleColumns", "LeftMargin", "RightMargin", "TopMargin", "BottomMargin", "HeaderMargin", "FooterMargin",
        "CenterHorizontally", "CenterVertically", "LeftHeader", "CenterHeader", "RightHeader", "LeftFooter",
        "CenterFooter", "RightFooter", "PrintGridlines", "PrintHeadings", "Order", "FirstPageNumber", "BlackAndWhite",
        "Draft", "ScaleWithDocHeaderFooter", "AlignMarginsHeaderFooter", "DifferentFirstPageHeaderFooter",
        "OddAndEvenPagesHeaderFooter")
WINDOW = ("FreezePanes", "SplitRow", "SplitColumn", "ScrollRow", "ScrollColumn", "Zoom", "DisplayGridlines",
          "DisplayHeadings", "DisplayZeros", "DisplayFormulas", "View", "GridlineColorIndex")


def _texts(v, mask: bool):
    return STAMP.sub("<stamp>", v) if mask and isinstance(v, str) else v


def _contents(ws, rect, mask: bool) -> tuple[dict, dict]:
    """({address: Value2} of every non-empty cell, {address: Formula2} of every formula)."""
    rng = ws.Range(rect_text(rect))
    f2, v2 = g(lambda: rng.Formula2), g(lambda: rng.Value2)
    grid = lambda x: x if isinstance(x, tuple) else ((x,),)     # noqa: E731 - one cell reads a scalar
    values, formulas = {}, {}
    if f2 == ERR or v2 == ERR:
        return {"read": ERR}, {"read": ERR}
    r1, c1 = rect[0], rect[1]
    for i, (frow, vrow) in enumerate(zip(grid(f2), grid(v2))):
        for j, (f, v) in enumerate(zip(frow, vrow)):
            addr = a1(r1 + i, c1 + j)
            if isinstance(f, str) and f.startswith("="):
                formulas[addr] = _texts(f, mask)
            if v is not None and v != "":
                values[addr] = _texts(v, mask)
    return values, formulas


def _labels_of(s, k: int) -> dict:
    out = {}
    count = g(lambda: s.Points().Count)
    for i in range(1, (count if isinstance(count, int) else 0) + 1):
        p = obj(lambda i=i: s.Points(i))
        key = f"series {k} point {i}"
        out[key + " fill"] = (g(lambda: p.Format.Fill.Visible), g(lambda: p.Format.Fill.ForeColor.RGB))
        if g(lambda: p.HasDataLabel) is True:
            lab = obj(lambda: p.DataLabel)
            out[key + " label"] = tuple(g(lambda a=a: getattr(lab, a)) for a in (
                "Position", "Top", "Left", "NumberFormat", "ShowValue", "ShowCategoryName", "ShowSeriesName"))
            out[key + " label font"] = tuple(g(lambda a=a: getattr(lab.Font, a)) for a in ("Name", "Size", "Color",
                                                                                         "Bold"))
    return out


def _chart(co, mask: bool) -> dict:
    ch = obj(lambda: co.Chart)
    d = {"object": tuple(g(lambda a=a: getattr(co, a)) for a in ("Name", "Left", "Top", "Width", "Height",
                                                                "Placement", "Visible")),
         "top left": g(lambda: co.TopLeftCell.Address), "bottom right": g(lambda: co.BottomRightCell.Address),
         "type": g(lambda: ch.ChartType), "has title": g(lambda: ch.HasTitle),
         "has legend": g(lambda: ch.HasLegend), "plot area": tuple(g(lambda a=a: getattr(ch.PlotArea, a)) for a in (
             "InsideLeft", "InsideTop", "InsideWidth", "InsideHeight")),
         "chart area border": g(lambda: ch.ChartArea.Format.Line.Visible),
         "plot visible only": g(lambda: ch.PlotVisibleOnly), "display blanks as": g(lambda: ch.DisplayBlanksAs)}
    if d["has title"] is True:
        d["title"] = (_texts(g(lambda: ch.ChartTitle.Text), mask), g(lambda: ch.ChartTitle.Font.Name),
                      g(lambda: ch.ChartTitle.Font.Size), g(lambda: ch.ChartTitle.Font.Bold),
                      g(lambda: ch.ChartTitle.Font.Color))
    if d["has legend"] is True:
        d["legend"] = (g(lambda: ch.Legend.Position), g(lambda: ch.Legend.Font.Name), g(lambda: ch.Legend.Font.Size),
                       g(lambda: ch.Legend.Font.Color))
    groups = g(lambda: ch.ChartGroups().Count)
    for i in range(1, (groups if isinstance(groups, int) else 0) + 1):
        grp = obj(lambda i=i: ch.ChartGroups(i))
        d[f"group {i}"] = (g(lambda: grp.Overlap), g(lambda: grp.GapWidth), g(lambda: grp.AxisGroup))
    for axis_group in (1, 2):
        for kind, label in ((XL_CATEGORY, "category"), (XL_VALUE, "value")):
            if g(lambda kind=kind, ag=axis_group: ch.HasAxis(kind, ag)) is not True:
                d[f"{label} axis {axis_group}"] = "none"
                continue
            ax = obj(lambda kind=kind, ag=axis_group: ch.Axes(kind, ag))
            d[f"{label} axis {axis_group}"] = tuple(g(lambda a=a: getattr(ax, a)) for a in (
                "MinimumScaleIsAuto", "MaximumScaleIsAuto", "ReversePlotOrder", "TickLabelPosition", "MajorTickMark",
                "HasMajorGridlines", "HasMinorGridlines", "HasTitle", "Crosses"))
            d[f"{label} axis {axis_group} scale"] = (g(lambda: ax.MinimumScale), g(lambda: ax.MaximumScale))
            d[f"{label} axis {axis_group} labels"] = (g(lambda: ax.TickLabels.NumberFormat),
                                                      g(lambda: ax.TickLabels.Font.Name),
                                                      g(lambda: ax.TickLabels.Font.Size),
                                                      g(lambda: ax.TickLabels.Font.Color))
            d[f"{label} axis {axis_group} line"] = (g(lambda: ax.Format.Line.Visible),
                                                    g(lambda: ax.Format.Line.ForeColor.RGB))
    n = g(lambda: ch.SeriesCollection().Count)
    d["series count"] = n
    for k in range(1, (n if isinstance(n, int) else 0) + 1):
        s = obj(lambda k=k: ch.SeriesCollection(k))
        d[f"series {k} formula"] = _texts(g(lambda: s.Formula), mask)
        d[f"series {k}"] = tuple(g(lambda a=a: getattr(s, a)) for a in ("ChartType", "AxisGroup", "PlotOrder",
                                                                          "HasDataLabels", "MarkerStyle"))
        d[f"series {k} fill"] = (g(lambda: s.Format.Fill.Visible), g(lambda: s.Format.Fill.ForeColor.RGB),
                                 g(lambda: s.Format.Fill.Pattern))
        d[f"series {k} line"] = (g(lambda: s.Format.Line.Visible), g(lambda: s.Format.Line.ForeColor.RGB),
                                 g(lambda: s.Format.Line.Weight), g(lambda: s.Format.Line.DashStyle))
        if d[f"series {k}"][3] is True:
            labs = obj(lambda: s.DataLabels())
            d[f"series {k} labels"] = tuple(g(lambda a=a: getattr(labs, a)) for a in (
                "NumberFormat", "Position", "ShowValue"))
            d[f"series {k} labels font"] = tuple(g(lambda a=a: getattr(labs.Font, a)) for a in (
                "Name", "Size", "Color", "Bold"))
        d.update(_labels_of(s, k))
    return d


def read_sheet(app, wb, name: str, rect, mask: bool = False) -> dict:
    """item -> {key: value} for one sheet over rect (from A1)."""
    ws = wb.Worksheets(name)
    raw = _Raw()
    out: dict = {}
    out["cell values"], out["Formula2"] = _contents(ws, rect, mask)
    cells: dict = defaultdict(dict)
    _walk(raw, ws._oleobj_, rect, ALL_PROPS, cells)
    for prop, by_cell in cells.items():
        item = out.setdefault(PROP_ITEM[prop], {})
        for addr, v in by_cell.items():
            item[f"{addr} {prop}"] = v
    out["borders"] = _borders(raw, ws._oleobj_, rect)
    r1, c1, r2, c2 = rect
    sizes, hidden = {}, {}
    _bisect(lambda a, b: g(lambda: ws.Range(f"{a}:{b}").RowHeight), 1, r2, sizes, lambda a, b: f"rows {a}:{b} height")
    _bisect(lambda a, b: g(lambda: ws.Range(f"{col_letter(a)}:{col_letter(b)}").ColumnWidth), 1, c2, sizes,
            lambda a, b: f"columns {col_letter(a)}:{col_letter(b)} width")
    _bisect(lambda a, b: g(lambda: ws.Range(f"{a}:{b}").EntireRow.Hidden), 1, r2 + 5, hidden,
            lambda a, b: f"rows {a}:{b}")
    _bisect(lambda a, b: g(lambda: ws.Range(f"{col_letter(a)}:{col_letter(b)}").EntireColumn.Hidden), 1, c2 + 5,
            hidden,
            lambda a, b: f"columns {col_letter(a)}:{col_letter(b)}")
    out["row heights"] = {k: v for k, v in sizes.items() if k.startswith("rows")}
    out["column widths"] = {k: v for k, v in sizes.items() if k.startswith("columns")}
    out["hidden rows and columns"] = hidden
    page = obj(lambda: ws.PageSetup)
    out["page setup"] = {a: (ERR if page is None else g(lambda a=a: getattr(page, a))) for a in PAGE}
    out["manual page breaks"] = {
        "rows": g(lambda: tuple(int(b.Location.Row) for b in ws.HPageBreaks if int(b.Type) == MANUAL_BREAK)),
        "columns": g(lambda: tuple(int(b.Location.Column) for b in ws.VPageBreaks if int(b.Type) == MANUAL_BREAK))}
    out["sheet"] = {"visible": g(lambda: ws.Visible), "tab colour": g(lambda: ws.Tab.Color),
                    "tab tint": g(lambda: ws.Tab.TintAndShade), "standard width": g(lambda: ws.StandardWidth),
                    "standard height": g(lambda: ws.StandardHeight), "used range": g(lambda: ws.UsedRange.Address),
                    "display page breaks": g(lambda: ws.DisplayPageBreaks),
                    "enable calculation": g(lambda: ws.EnableCalculation), "code name set": bool(g(lambda: ws.CodeName))}
    out["sheet-level names"] = {g(lambda n=n: n.Name): (_texts(g(lambda n=n: n.RefersTo), mask),
                                                        g(lambda n=n: n.Visible)) for n in ws.Names}
    fcs = g(lambda: ws.Cells.FormatConditions.Count)
    cf = {"count": fcs}
    for i in range(1, (fcs if isinstance(fcs, int) else 0) + 1):
        fc = obj(lambda i=i: ws.Cells.FormatConditions(i))
        cf[f"rule {i}"] = tuple(g(lambda a=a: getattr(fc, a)) for a in ("Type", "Operator", "Formula1", "StopIfTrue",
                                                                         "Priority"))
        cf[f"rule {i} applies to"] = g(lambda: fc.AppliesTo.Address)
        cf[f"rule {i} look"] = (g(lambda: fc.Font.Color), g(lambda: fc.Font.Bold), g(lambda: fc.Interior.Color),
                                g(lambda: fc.Interior.Pattern))
    out["conditional formats"] = cf
    try:
        dv = str(ws.Cells.SpecialCells(ALL_VALIDATION).Address)
    except Exception:   # noqa: BLE001 - Excel: "no cells were found"
        dv = "none"
    out["data validation"] = {"cells": dv}
    if out["sheet"]["visible"] == -1:             # the window's settings are the active sheet's
        g(lambda: ws.Activate())
        win = obj(lambda: app.ActiveWindow)
        out["window"] = {a: g(lambda a=a: getattr(win, a)) for a in WINDOW}
    else:
        out["window"] = {"hidden sheet": True}
    count = g(lambda: ws.ChartObjects().Count)
    charts = {"count": count}
    for i in range(1, (count if isinstance(count, int) else 0) + 1):
        for k, v in _chart(obj(lambda i=i: ws.ChartObjects(i)), mask).items():
            charts[f"chart {i} {k}"] = v
    out["charts"] = charts
    return out


def used(wb, name: str):
    addr = g(lambda: wb.Worksheets(name).UsedRange.Address)
    return None if addr == ERR else parse_rect(addr)


def read_book(app, wb, rects: dict, mask: bool = False) -> dict:
    """The whole workbook: 'sheet:<name>' -> read_sheet, and the workbook's own items."""
    out = {}
    for name, rect in rects.items():
        out[f"sheet:{name}"] = read_sheet(app, wb, name, rect, mask)
    names = {}
    for n in wb.Names:
        parent = g(lambda n=n: n.Parent.Name)
        scope = "workbook" if parent == g(lambda: wb.Name) else f"sheet {parent}"
        names[f"{g(lambda n=n: n.Name)} ({scope})"] = (_texts(g(lambda n=n: n.RefersTo), mask),
                                                       g(lambda n=n: n.Visible))
    out["workbook"] = {
        "sheet order and visibility": {"sheets": tuple((str(ws.Name), int(ws.Visible)) for ws in wb.Sheets)},
        "workbook names": names,
        "styles": {"count": g(lambda: wb.Styles.Count),
                   "names": g(lambda: tuple(sorted(str(s.Name) for s in wb.Styles)))},
        "links": {f"type {k}": g(lambda k=k: wb.LinkSources(k)) for k in LINK_TYPES},
    }
    return out


def book_rects(wbs: list, names: list[str]) -> dict:
    """Per sheet the union of every side's used range, from A1 (one cell more each way)."""
    out = {}
    for name in names:
        r = (1, 1, 1, 1)
        for wb in wbs:
            r = union(r, used(wb, name))
        out[name] = (1, 1, r[2] + 1, r[3] + 1)
    return out


def _at(snap: dict, where: str, item: str, key):
    return snap.get(where, {}).get(item, {}).get(key, "<absent>")


def diff(a: dict, b: dict) -> dict:
    """item -> [(where, key) that differ] (keys only: never a value)."""
    out = defaultdict(list)
    for where in sorted(set(a) | set(b)):
        sa, sb = a.get(where, {}), b.get(where, {})
        for item in sorted(set(sa) | set(sb)):
            ia, ib = sa.get(item, {}), sb.get(item, {})
            for key in sorted(set(ia) | set(ib), key=str):
                if ia.get(key, "<absent>") != ib.get(key, "<absent>"):
                    out[item].append((where, key))
    return out


def shown(where: str, key) -> str:
    return f"{where.removeprefix('sheet:')}!{key}"


def route_diff(snaps: dict) -> tuple[dict, dict]:
    """(item -> [keys where the insert differs from the copy], item -> [keys masked as noise]). A key the copy and the
    control (the same copy route twice) disagree on is noise only when the insert has the copy's or the control's
    value there (it falls inside the copy route's own spread); otherwise it is a difference too."""
    found, masked = defaultdict(list), defaultdict(list)
    for item, keys in diff(snaps["copy"], snaps["insert"]).items():
        for where, key in keys:
            ins = _at(snaps["insert"], where, item, key)
            if ins == _at(snaps["control"], where, item, key):
                masked[item].append(shown(where, key))
            else:
                found[item].append(shown(where, key))
    return found, masked


def _has_err(v) -> bool:
    return v == ERR or (isinstance(v, tuple) and ERR in v)


def _prop(key, v) -> str:
    """A key without its cell, sheet, chart or series numbers, with the tuple positions that failed: the property."""
    name = re.sub(r"\b[A-Z]{1,3}\d+\b", "<cell>", str(key))
    name = re.sub(r"\d+", "#", name)
    if isinstance(v, tuple):
        name += f" [{','.join(str(i) for i, x in enumerate(v) if x == ERR)}]"
    return name


def unread(snaps: dict) -> dict:
    """item -> {property: how many keys} that could not be read on EVERY route (never a value)."""
    out: dict = defaultdict(lambda: defaultdict(int))
    first, *rest = snaps.values()
    for where, items in first.items():
        for item, keys in items.items():
            for key, v in keys.items():
                if _has_err(v) and all(_has_err(s.get(where, {}).get(item, {}).get(key)) for s in rest):
                    out[item][_prop(key, v)] += 1
    return out


# Properties a hidden Excel of our own cannot read on any route, and why that is no gap (printed with the count).
NOT_READABLE: dict[str, str] = {}


def unread_findings(bad: dict) -> tuple[list, list]:
    """(the failing [item: property x n], the accepted ones) of unread()."""
    fail, ok = [], []
    for item, props in bad.items():
        for prop, n in props.items():
            (ok if any(re.fullmatch(p, f"{item}: {prop}") for p in NOT_READABLE) else fail).append(
                f"{item}: {prop} x{n}")
    return fail, ok


# ---------------------------------------------------------------------------------------------------- fixtures
def _chart_kinds_plan():
    """test_look.small_sheet's plan (rows 5:8: labels in C, numbers in D:G) with one chart of every kind the
    renderer has over those rows; the waterfall has a thin step (its labels are fixed after the move-in) and the
    line labels points the layout moves (also fixed after the move-in)."""
    import test_look
    import test_xlsx
    p, _ = test_look.small_sheet()
    p.footer = test_xlsx.FOOTER
    cat = "C5:C8"
    kinds = [
        ("column", [("AC", "D5:D8", "ac"), ("PY", "E5:E8", "py")], {}),
        ("bar", [("AC", "D5:D8", "ac"), ("BU", "E5:E8", "bu")], {"sort": "desc", "highlight": [2]}),
        ("line", [("AC", "D5:D8", "ac"), ("PY", "E5:E8", "py")],
         {"label_points": [[1, 2, 4], [3, 4]], "label_pos": [{2: "below"}, {3: "left"}]}),
        ("stacked", [("A", "D5:D8", "part"), ("B", "E5:E8", "part")], {}),
        ("stacked100", [("A", "D5:D8", "part"), ("B", "E5:E8", "part")], {}),
        ("combo", [("AC", "D5:D8", "ac"), ("Pct", "G5:G8", "pct")], {}),
        ("variance", [("Better", "F5:F8", "fav"), ("Worse", "G5:G8", "unfav")], {}),
        ("waterfall", [("Base", "D5:D8", "base"), ("Up", "E5:E8", "up"), ("Down", "F5:F8", "down"),
                       ("Total", "G5:G8", "total")], {"thin_steps": [2]}),
    ]
    p.charts = [ChartSpec(kind, f"{kind} chart", cat_range=cat, series=series, anchor=f"B{20 + 20 * i}", **kw)
                for i, (kind, series, kw) in enumerate(kinds)]
    return p


def fixtures() -> dict:
    """name -> build(path) -> BuiltFile, every one an 'Q3 Answer' with its Data sheet (and Lines where it has
    them), from the tests' plans and fixed data."""
    import test_look
    import test_xlsx
    from test_v5_layout import GROUPED
    lines = [test_xlsx._line(i) for i in range(3)]

    def build(plan, with_lines=False):
        plan.footer = plan.footer or test_xlsx.FOOTER
        return lambda p: xlsx.build(p, "Q3 Answer", plan, test_xlsx._data(),
                                    [test_xlsx._lines(lines)] if with_lines else [])
    out = {"answer with a Lines sheet": build(test_xlsx._plan(), True)}
    for case in test_look.CASES:
        out[case.__name__] = build(case()[0], case.__name__ == "lines_ranking")
    out["every chart kind"] = build(_chart_kinds_plan())
    out["matrix, kpi tiles, key points"] = build(test_look._laid_out([
        GROUPED, {"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA", "Opex"]},
        {"type": "matrix", "query": "fm", "metric": "Opex"},
        {"type": "commentary", "key_points": ["A point of the agent's own"]}])[0])
    return out


# ---------------------------------------------------------------------------------------------------- targets
class Excel:
    """The hidden Excel of this script: app (the engine's wrapper), raw, pid; every workbook opened is closed."""

    def __init__(self):
        self.raw, self.pid, self.ctime = start_excel()
        self.app = fastcom.wrap(self.raw)
        self.books: list = []

    def fresh(self, tag: str, sheets=USER_SHEETS):
        """A copy of the new-workbook seed opened, with the user's sheets around 'Sources' (_svmeta hidden, _vh
        very hidden); its Sources sheet is the seed's own."""
        folder = TMP / "targets"
        folder.mkdir(exist_ok=True)
        path = folder / f"{tag}_{uuid.uuid4().hex[:8]}.xlsx"
        if SEED.exists():
            shutil.copyfile(SEED, path)
            wb = self.app.Workbooks.Open(str(path))
        else:
            wb = self.app.Workbooks.Add()
            wb.Worksheets(1).Name = "Sources"
        self.books.append(wb)
        have = [str(ws.Name) for ws in wb.Worksheets]
        for name in sheets:
            if name in have:
                continue
            ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
            ws.Name = name
            ws.Range("A1").Value2 = f"the user's {name}"
            ws.Range("B2").Formula2 = "=LEN(A1)*2"
        order = [n for n in sheets] + [n for n in have if n not in sheets]
        for i, name in enumerate(order[1:], 1):
            wb.Worksheets(name).Move(None, wb.Worksheets(order[i - 1]))
        for name, vis in (("_svmeta", 0), ("_vh", 2)):
            if name in sheets:
                wb.Worksheets(name).Visible = vis
        first = next(n for n in order if n not in ("_svmeta", "_vh"))
        wb.Worksheets(first).Activate()
        return wb

    def close(self, wb) -> None:
        with contextlib.suppress(Exception):
            wb.Close(False)
        if wb in self.books:
            self.books.remove(wb)

    def quit(self) -> None:
        dog = threading.Timer(QUIT_WAIT_S, lambda: kill_if_ours(self.pid, self.ctime))
        dog.start()
        try:
            for wb in list(self.raw.Workbooks):
                wb.Close(False)
            self.raw.Quit()
        except pywintypes.com_error:
            pass
        finally:
            dog.cancel()
        self.raw = self.app = None
        self.books.clear()
        fastcom._IFACE.clear()
        _wait_gone(self.pid, self.ctime)
        kill_if_ours(self.pid, self.ctime)


def _copy_of(built, tag: str):
    """The built file under a new name (byte for byte): each route gets the same file."""
    path = built.path.parent / f"{uuid.uuid4().hex}.xlsx"
    shutil.copyfile(built.path, path)
    return dataclasses.replace(built, path=path)


def _names(wb) -> list[str]:
    return [str(ws.Name) for ws in wb.Sheets]


def _mentions_file(snap: dict, built) -> list[str]:
    """Where a formula, a name or a series formula mentions the built file (its name, '.xlsx]' or a path)."""
    stem = built.path.stem.casefold()
    found = []
    for where, items in snap.items():
        for item in ("Formula2", "sheet-level names", "workbook names", "charts"):
            for key, v in items.get(item, {}).items():
                text = " ".join(str(x) for x in v) if isinstance(v, tuple) else str(v)
                low = text.casefold()
                if stem in low or ".xlsx" in low or (item == "charts" and "formula" in key and "[" in text):
                    found.append(f"{where}!{key}")
    return found


def report_items(part: str, items, found: dict, noise: dict, errs: dict) -> None:
    """One OK / DIFF per item: DIFF when the insert differs from the copy outside the copy route's own spread, or
    when a property could not be read on any route (an unread property compares equal and would pass unseen),
    unless NOT_READABLE says why that is no gap. Masked keys are printed as unproven."""
    for item in sorted(items):
        keys = found.get(item, [])
        fail, accepted = unread_findings({item: errs.get(item, {})})
        detail = (f"{len(keys)} differ: {keys[:4]}" if keys else "") + (f" unread on every route: {fail[:6]}"
                                                                          if fail else "")
        check(f"{part}: {item}", not keys and not fail, detail.strip())
        if accepted:
            info(f"unread on every route, accepted ({item}): {accepted}")
        if noise.get(item):
            info(f"unproven (the copy route differs from itself here; the insert has the control's value) in {item}: "
                 f"{len(noise[item])}, e.g. {noise[item][:3]}")


# ---------------------------------------------------------------------------------------------------- 1. the files
def compare_fixtures(xl: Excel) -> None:
    """Per fixture: the copy route, the insert route and the control (a second copy route), compared after the
    move-in and after finish + calculate."""
    found: dict = defaultdict(list)            # item -> ["fixture phase: key", ...] (route differences)
    noise: dict = defaultdict(list)            # item -> [...] (masked: the insert has the control's value)
    errs: dict = defaultdict(lambda: defaultdict(int))     # item -> property -> reads failed on every route
    items: set = set()
    extra = {"links": [], "temporary file deleted right after": [], "the Book learns the order": []}
    folder = TMP / "build"
    folder.mkdir(exist_ok=True)
    for fix, build in fixtures().items():
        built0 = build(folder / f"{uuid.uuid4().hex}.xlsx")
        if len(built0.chart_lint):
            info(f"{fix}: chart lint {len(built0.chart_lint)} line(s)")
        sides = {}
        for side, route in ROUTES.items():
            built = _copy_of(built0, side)
            wb = xl.fresh(side)
            t = Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0)
            book = movein.read_book(t)
            with _settings_off(xl):
                route(t, built, "Sources", book=book)
            if book.names() != [n for n in _names(wb)]:
                extra["the Book learns the order"].append(f"{fix} {side}")
            try:
                built.path.unlink()
            except OSError as e:
                extra["temporary file deleted right after"].append(f"{fix} {side}: {type(e).__name__}")
            sides[side] = (t, built)
        names = _names(sides["copy"][0].wb)
        for phase in ("after the move-in", "after finish and calculate"):
            if phase != "after the move-in":
                for t, built in sides.values():
                    with _settings_off(xl):
                        movein.finish(t, built)            # the job's order: labels, calculate, lifts

            rects = book_rects([t.wb for t, _ in sides.values()], names)
            snaps = {side: read_book(xl.app, t.wb, rects) for side, (t, _) in sides.items()}
            for snap in snaps.values():
                for where in snap.values():
                    items.update(where)
            real, masked = route_diff(snaps)
            for item, keys in real.items():
                found[item] += [f"{fix} {phase}: {k}" for k in keys]
            for item, keys in masked.items():
                noise[item] += [f"{fix} {phase}: {k}" for k in keys]
            for item, props in unread(snaps).items():
                for prop, n in props.items():
                    errs[item][prop] += n
            t, built = sides["insert"]
            links = snaps["insert"]["workbook"]["links"]
            if any(v is not None for v in links.values()):                 # a link, or LinkSources unreadable
                extra["links"].append(f"{fix} {phase}: LinkSources {sorted(k for k, v in links.items() if v)}")
            extra["links"] += [f"{fix} {phase}: {k}" for k in _mentions_file(snaps["insert"], built)]
        for t, _ in sides.values():
            xl.close(t.wb)
        print(f"     compared: {fix} ({len(names)} sheets)", flush=True)
    report_items("files", items, found, noise, errs)
    check("files: links (none of any type; no formula, name or series formula mentions the file)",
          not extra["links"], str(extra["links"][:4]))
    check("files: the temporary file can be deleted right after the move-in",
          not extra["temporary file deleted right after"], str(extra["temporary file deleted right after"][:4]))
    check("files: the Book knows where the sheets landed", not extra["the Book learns the order"],
          str(extra["the Book learns the order"][:4]))


# ---------------------------------------------------------------------------------------------------- 2. placement and refusals
PLACES = [  # (label, sheets of the target, after, first)
    ("after a named sheet", USER_SHEETS, "Notes", False),
    ("after the last sheet (very hidden: back to the last visible)", USER_SHEETS, None, False),
    ("after a hidden sheet", USER_SHEETS, "_svmeta", False),
    ("after a very hidden sheet", USER_SHEETS, "_vh", False),
    ("first (before the first visible sheet)", USER_SHEETS, None, True),
    ("first with a hidden first sheet", ("_svmeta", "Intro", "Notes", "Sources"), None, True),
    ("nothing visible before a hidden anchor (before the first visible)", ("_svmeta", "Notes", "Sources"),
     "_svmeta", False),
]


def placement(xl: Excel, built0) -> None:
    for label, sheets, after, first in PLACES:
        got = {}
        for side in ("copy", "insert"):
            built = _copy_of(built0, side)
            wb = xl.fresh(side, sheets)
            if sheets[0] == "_svmeta":
                wb.Worksheets("_svmeta").Move(wb.Worksheets(1))
                wb.Worksheets("_svmeta").Visible = 0
            t = Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0)
            book = movein.read_book(t)
            with _settings_off(xl):
                ROUTES[side](t, built, after, first, book=book)
            got[side] = (_names(wb), [int(ws.Visible) for ws in wb.Sheets], book.names() == _names(wb))
            xl.close(wb)
            built.path.unlink(missing_ok=True)
        ok = got["copy"] == got["insert"] and got["insert"][2]
        check(f"lands: {label}", ok, "" if ok else f"copy {got['copy'][0]} | insert {got['insert'][0]}")


def _where(xl: Excel, wb) -> tuple[str, str]:
    """(which workbook is active: 'target' or another's first sheet, the active sheet)."""
    active = xl.app.ActiveWorkbook
    return ("target" if str(active.Name) == str(wb.Name) else "another workbook", str(xl.app.ActiveSheet.Name))


@contextlib.contextmanager
def _settings_off(xl: Excel):
    """A quiet batch's four settings (xl/inproc/runner.QUIET) around a route: what the insert itself leaves active.
    (Task 2b: the job's own landing is its close-out's, no longer a quiet() around the route.)"""
    before = [getattr(xl.app, name) for name, _ in runner.QUIET]
    for name, value in runner.QUIET:
        setattr(xl.app, name, value)
    try:
        yield
    finally:
        for (name, _), value in reversed(list(zip(runner.QUIET, before))):
            setattr(xl.app, name, value)


def active_after(xl: Excel, built0) -> None:
    """What is active right after the route itself (a quiet batch's settings, no landing), with the target the
    active workbook and with another workbook active."""
    for label, other_active in (("the target active", False), ("another workbook active", True)):
        for how, around in (("right after the route", _settings_off),):
            got = {}
            for side in ("copy", "insert"):
                built = _copy_of(built0, side)
                wb = xl.fresh(side)
                other = xl.fresh("other", ("Mine",)) if other_active else None
                if other is not None:
                    other.Activate()
                t = Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0)
                try:
                    with around(xl):
                        ROUTES[side](t, built, "Sources", book=movein.read_book(t))
                    got[side] = _where(xl, wb)
                except Exception as e:   # noqa: BLE001 - recorded: the copy route's BreakLink, see the report
                    got[side] = ("failed", f"{type(e).__name__} {getattr(e, 'excepinfo', None) or e}"[:160])
                for b in (wb, other):
                    if b is not None:
                        xl.close(b)
                built.path.unlink(missing_ok=True)
            info(f"active {how}, {label}: copy {got['copy']} | insert {got['insert']}")
            if got["copy"][0] == "failed":
                check(f"active {how}, {label} (insert route; the copy route failed)", got["insert"][0] != "failed",
                      f"insert {got['insert']}")
            else:
                check(f"active {how}, {label}", got["copy"] == got["insert"],
                      f"copy {got['copy']} | insert {got['insert']}")


def refusals(xl: Excel, built0) -> None:
    # a taken name: refused before anything is inserted
    wb = xl.fresh("taken", USER_SHEETS + ("Q3 Data",))
    t = Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0)
    before = (_names(wb), int(wb.Sheets.Count))
    built = _copy_of(built0, "taken")
    try:
        ROUTES["insert"](t, built, "Sources", book=movein.read_book(t))
        refused = False
    except movein.NameTaken:
        refused = True
    check("a taken name is refused before the insert", refused and (_names(wb), int(wb.Sheets.Count)) == before
          and t.sheets_added == [], f"refused {refused}, sheets {_names(wb)}")
    with _settings_off(xl):
        wb.Sheets.Add(None, wb.Worksheets("Sources"), None, str(built.path))
    info(f"what Excel itself does with a taken name: {[n for n in _names(wb) if n not in before[0]]}")
    xl.close(wb)
    built.path.unlink(missing_ok=True)
    # a file holding a table; a file holding one sheet more than the answer's; the answer's sheets in another order
    def odd(names: list[str], table: bool = False):
        path = TMP / "build" / f"{uuid.uuid4().hex}.xlsx"
        with xlsxwriter.Workbook(str(path)) as book:
            for name in names:
                ws = book.add_worksheet(name)
                ws.write_number(0, 0, 1.0)
                if table and name == "Q3 Data":
                    ws.add_table("C1:D3", {"name": "Data_Q3"})
        return dataclasses.replace(built0, path=path)
    sheets = list(built0.sheets)
    for label, bad, words in (("a file holding a Table", odd(sheets, table=True), "Table"),
                              ("a file holding a sheet more", odd([*sheets, "Scratch"]), "sheets"),
                              ("a file holding the sheets in another order", odd(sheets[::-1]), "sheets")):
        wb = xl.fresh("odd")
        t = Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0)
        before = _names(wb)
        try:
            ROUTES["insert"](t, bad, "Sources", book=movein.read_book(t))
            refused = ""
        except ValueError as e:
            refused = str(e)
        check(f"{label} is refused before the insert", words in refused and _names(wb) == before
              and t.sheets_added == [], f"sheets {_names(wb)}" if _names(wb) != before else "")
        xl.close(wb)
    # a file that is not there, a file that is not a workbook: what Excel leaves when the insert itself fails (the
    # zip guard would refuse both before Excel is asked, so Sheets.Add is called directly here)
    broken = TMP / "build" / "broken.xlsx"
    broken.write_bytes(b"not a workbook")
    for label, p in (("a missing file", TMP / "build" / "missing.xlsx"), ("a broken file", broken)):
        wb = xl.fresh("fail")
        before = _names(wb)
        err = None
        try:
            with _settings_off(xl):
                jobs.com(lambda wb=wb, p=p: wb.Sheets.Add(None, wb.Worksheets("Sources"), None, str(p)), limit=30)
        except Exception as e:   # noqa: BLE001
            err = type(e).__name__
        left = [n for n in _names(wb) if n not in before]
        check(f"the insert of {label}: an error, nothing inserted, Excel answers", err is not None and not left
              and int(xl.app.Workbooks.Count) == len(xl.books), f"error {err}, left {left}")
        xl.close(wb)


def excel_around(xl: Excel, built0) -> None:
    """The Excel around the insert: calculation, the active workbook, the windows, the recent list, the undo
    state, the workbooks, the visible windows of the process; EnableEvents off during it (quiet)."""
    def state(wb):
        def recent():
            r = xl.app.RecentFiles
            return (int(r.Count), str(r(1).Name) if int(r.Count) else None)
        undo = g(lambda: xl.app.CommandBars.GetEnabledMso("Undo"))
        return {"calculation": int(xl.app.Calculation), "active workbook": str(xl.app.ActiveWorkbook.Name),
                "windows": int(xl.app.Windows.Count), "workbooks": int(xl.app.Workbooks.Count),
                "recent": g(recent), "undo": undo, "screen updating": bool(xl.app.ScreenUpdating),
                "events": bool(xl.app.EnableEvents), "alerts": bool(xl.app.DisplayAlerts),
                "visible windows": len(windows_of(xl.pid))}
    got = {}
    for side in ("copy", "insert"):
        built = _copy_of(built0, side)
        wb = xl.fresh(side)
        t = Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0)
        book = movein.read_book(t)
        before = state(wb)
        seen = {}
        with _settings_off(xl):
            seen["events during"] = bool(xl.app.EnableEvents)
            ROUTES[side](t, built, "Sources", book=book)
        after = state(wb)
        got[side] = (before, after, seen)
        changed = {k for k in before if before[k] != after[k]}
        check(f"Excel around the {side} route: nothing else changes", not changed - {"undo"},
              f"changed {sorted(changed)}" if changed else "")
        info(f"{side}: undo enabled before {before['undo']} after {after['undo']}; events during {seen}")
        xl.close(wb)
        built.path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------------------------------- 3. whole jobs
class _Now(dt.datetime):
    """One fetch time for every job of both routes (the texts that carry it compare equal)."""

    @classmethod
    def now(cls, tz=None):
        return cls(2026, 11, 20, 9, 30, tzinfo=tz)


@contextlib.contextmanager
def _patched(seams: dict):
    """(module, name) -> value, put back afterwards."""
    old = {k: getattr(*k) for k in seams}
    try:
        for (mod, name), value in seams.items():
            setattr(mod, name, value)
        yield
    finally:
        for (mod, name), value in old.items():
            setattr(mod, name, value)


def _pres():
    from test_file_job import SUPPLIERS
    from test_v2_pipeline import LINE_CELL, FakeProvider, lines_only_spec
    from v2util import facts, resolver, spec_dict

    from smartass.model.lines import LineSet
    ls = LineSet("b0", list(SUPPLIERS), ["AP_CUSTOMER_NAME", "AMOUNT"], cube_value=-100.0, cell=LINE_CELL)
    prep = lambda spec, **kw: pipeline.prepare(spec, resolver(), FakeProvider(facts()), **kw)   # noqa: E731
    return [("first answer", None, prep(spec_dict())),
            ("second answer with a Lines sheet", None, prep(lines_only_spec(), lines={"b0": ls})),
            ("follow-up of Q1 (its sheets replaced)", 1, prep(spec_dict(question="How is Business opex doing?")))]


JOB_SIDES = {"copy": False, "insert": True, "control": False}    # (Task 2b) side -> config.INPROC: the batched job
                                                                  # over COM, in-process, over COM again (the noise)


def compare_jobs(xl: Excel) -> None:
    """Whole jobs (a new answer, a second with a Lines sheet, a follow-up) as batches (Task 2b) through the COM
    executor ('copy' side, and 'control') and the in-process add-in ('insert' side): the same workbook either way."""
    from smartass.core.timing import BATCH_PREFIX
    books, steps = {}, defaultdict(dict)
    for side, inproc in JOB_SIDES.items():
        wb = xl.fresh(side, ("Sources",))
        books[side] = wb
        target = lambda h, target, title, build_new=False, wb=wb: Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0, created=True)  # noqa: E731
        seams = {(exceljob, "_target"): target, (exceljob, "_save"): lambda job: None,
                 (exceljob, "remember"): lambda pre, out: None, (exceljob, "datetime"): _Now,
                 (config, "INPROC"): inproc}
        with _patched(seams):
            for label, q, pre in _pres():
                job = (lambda h, pre=pre: exceljob.excel_job(h, pre)) if q is None else (
                    lambda h, pre=pre, q=q: exceljob.redraw_job(h, pre, q, "this workbook"))
                out, timings = exceljob.with_timings(job)(None)
                steps[label][side] = ((out.get("checks") or {}).get("sheet_equals_model"),
                                      timings.get("draw", 0.0),
                                      int(sum(v for k, v in timings.items() if k.startswith(BATCH_PREFIX)
                                              and "." not in k[len(BATCH_PREFIX):])),
                                      sorted(set(out.get("warnings") or []) | set(out.get("layout") or [])))
    for label, by in steps.items():
        ok = all(v[0] == "OK" for v in by.values())
        check(f"jobs: {label}: sheet = model through both executors", ok, f"{ {k: v[0] for k, v in by.items()} }")
        check(f"jobs: {label}: the same warnings and layout findings", by["copy"][3] == by["insert"][3],
              f"com {by['copy'][3]} | inproc {by['insert'][3]}")
        info(f"{label}: draw over COM {by['copy'][1]:.2f} s, in-process {by['insert'][1]:.2f} s; batches "
             f"{by['copy'][2]} / {by['insert'][2]}")
    names = {side: _names(wb) for side, wb in books.items()}
    check("jobs: the same sheets in the same order", names["copy"] == names["insert"],
          "" if names["copy"] == names["insert"] else f"com {names['copy']} | inproc {names['insert']}")
    rects = book_rects(list(books.values()), names["copy"])
    snaps = {side: read_book(xl.app, wb, rects, mask=True) for side, wb in books.items()}
    found, noise = route_diff(snaps)
    items = {item for snap in snaps.values() for where in snap.values() for item in where}
    report_items("jobs", items, found, noise, unread(snaps))
    for wb in books.values():
        xl.close(wb)


# ---------------------------------------------------------------------------------------------------- 3e. a new workbook
NEW_SIDES = {"copy": False, "insert": True, "control": False}     # (phase 3c) side -> its first answer's built file is
                                                                   # the workbook (else today's: template, SaveAs, insert)
# What a new workbook built as a file has otherwise than one made today's way, by name (item, the keys' pattern, why):
# printed as expected, never masked; any other difference is a DIFF.
EXPECTED_NEW: list[tuple[str, str, str]] = [
    ("styles", r"workbook!(count|names)", "the TN template's cell styles are not in a built file (48 against 102); "
     "xl_tidy and apply_report install them when they need them (styling.ensure_styles merges them from the template)"),
    ("workbook names", r"workbook!_tnStylesVer \(workbook\)", "the template's style marker, added with the styles"),
    ("window", r"Sources!(FreezePanes|ScrollRow|SplitRow)", "this hidden Excel refuses the log's frozen header over "
     "COM (a window without a client area: undone); the built file has it frozen, as a visible Excel freezes it"),
]


def _new_target_seam(xl: Excel, holder: dict):
    """exceljob._target for the new-workbook part: the first question resolves to a new workbook in this script's
    Excel (target._new_target, today's way or the built file's), every later one to that workbook (the session's)."""
    from smartass.xl import target as xltarget

    def target(h, target, title, build_new=False):
        if "t" not in holder:
            holder["t"] = xltarget._new_target(None, title, [(xl.pid, 0, xl.app)], {xl.pid}, build_new)
            return holder["t"]
        return Target(app=xl.app, wb=holder["t"].wb, pid=xl.pid, hwnd=0, created=True)
    return target


def _expected(item: str, key: str) -> str | None:
    return next((why for it, pattern, why in EXPECTED_NEW if it == item and re.fullmatch(pattern, key)), None)


def new_workbook(xl: Excel) -> None:
    """(Phase 3c Task 1) The first answer of a new workbook made today's way (Workbooks.Add of the TN template, the
    label, SaveAs, the insert, the Sources log furnished over COM) against one whose built file is the workbook
    (Excel only opens it), through the add-in, today's way twice (the noise): a first answer followed by a second with
    a Lines sheet and a follow-up, and a first answer with a Lines sheet. Compared item by item; each workbook is saved
    by the job in this script's temporary folder (a file Excel did not write - no label - would show here). Then the
    TN styles installed in a built workbook (styling.ensure_styles, as xl_tidy does) leave its sheets as they were,
    and a built new workbook whose finish fails is closed unsaved, its file deleted."""
    from smartass.core import session
    folder = TMP / "analyser"
    folder.mkdir(exist_ok=True)
    seams = {(exceljob, "remember"): lambda pre, out: None, (exceljob, "datetime"): _Now, (config, "INPROC"): True,
             (session, "project_dir"): lambda: folder, (session, "set_workbook"): lambda p: None}
    pres = _pres()
    lines_first = [("first answer with a Lines sheet", None, pres[1][2])]
    for scenario, questions in (("new workbook", pres), ("new workbook, lines first", lines_first)):
        books = _new_books(xl, seams, questions, scenario)
        if questions is pres:
            _styles_installed(xl, books["insert"])
        for wb in books.values():
            xl.close(wb)
    _failed_new_workbook(xl, seams, folder)


def _new_books(xl: Excel, seams: dict, questions: list, scenario: str) -> dict:
    """The questions answered into a new workbook per side (NEW_SIDES), compared: sheet = model, warnings, the saved
    file labelled, the sheets, every item (the expected differences printed as such). Returns {side: workbook}."""
    from smartass.core.timing import BATCH_PREFIX
    from smartass.xl import target as xltarget
    books, steps, files = {}, defaultdict(dict), {}
    for side, build in NEW_SIDES.items():
        holder: dict = {}
        side_seams = seams | {(exceljob, "_target"): _new_target_seam(xl, holder)}
        if not build:
            side_seams[(exceljob, "_new_book")] = lambda: None
        with _patched(side_seams):
            for label, q, pre in questions:
                job = (lambda h, pre=pre: exceljob.excel_job(h, pre)) if q is None else (
                    lambda h, pre=pre, q=q: exceljob.redraw_job(h, pre, q, "this workbook"))
                out, timings = exceljob.with_timings(job)(None)
                steps[label][side] = ((out.get("checks") or {}).get("sheet_equals_model"),
                                      sum(timings.get(s, 0.0) for s in ("open target", "draw", "save")),
                                      int(sum(v for k, v in timings.items() if k.startswith(BATCH_PREFIX)
                                              and "." not in k[len(BATCH_PREFIX):])),
                                      sorted(set(out.get("warnings") or []) | set(out.get("layout") or [])))
        saves = later.wait(60)                            # (Task 2) Excel's saves after the replies (SV_SAVE=sync: none)
        check(f"{scenario} ({side}): every save after a reply went through",   # (final review I1) not the first's:
              (len(saves) == len(questions) - 1 and all(s["ok"] for s in saves))   # a new workbook's is in the job
              if config.SAVE == "later" else not saves,
              f"{len(saves)} saves, {[s.get('how') for s in saves]}; failed: "
              f"{[s.get('error') for s in saves if not s['ok']]}")
        t = holder["t"]
        books[side], files[side] = t.wb, Path(t.path)
        check(f"{scenario} ({side}): {'its built file is the workbook' if build else 'made from the template'}",
              (t.made is not None) == build, f"made {t.made}")
    for label, by in steps.items():
        check(f"new workbook: {label}: sheet = model either way", all(v[0] == "OK" for v in by.values()),
              f"{ {k: v[0] for k, v in by.items()} }")
        check(f"new workbook: {label}: the same warnings and layout findings", by["copy"][3] == by["insert"][3],
              f"today {by['copy'][3]} | built {by['insert'][3]}")
        info(f"{label}: the Excel job today's way {by['copy'][1]:.2f} s ({by['copy'][2]} batches), its built file "
             f"{by['insert'][1]:.2f} s ({by['insert'][2]} batches), today's way again {by['control'][1]:.2f} s")
    for side, wb in books.items():
        path = files[side]
        check(f"{scenario} ({side}): the file the job saved is labelled", path.stat().st_size > 0
              and xltarget._seed_has_label(path), f"{path.stat().st_size} bytes")
        info(f"{scenario} ({side}): Saved {wb.Saved} after its jobs")
    names = {side: _names(wb) for side, wb in books.items()}
    check(f"{scenario}: the same sheets in the same order", names["copy"] == names["insert"],
          "" if names["copy"] == names["insert"] else f"today {names['copy']} | built {names['insert']}")
    rects = book_rects(list(books.values()), names["copy"])
    snaps = {side: read_book(xl.app, wb, rects, mask=True) for side, wb in books.items()}
    found, noise = route_diff(snaps)
    for item, keys in list(found.items()):
        for key in list(keys):
            why = _expected(item, key)
            if why is not None:
                keys.remove(key)
                info(f"expected difference ({item}) {key}: {why}")
    items = {item for snap in snaps.values() for where in snap.values() for item in where}
    report_items(scenario, items, {k: v for k, v in found.items() if v}, noise, unread(snaps))
    return books


def _styles_installed(xl: Excel, wb) -> None:
    """styling.ensure_styles on a built workbook (what xl_tidy and apply_report do first): the TN styles and their
    marker arrive from the template (one Styles.Merge), with no prompt, and the answer's sheets stay as they were."""
    from smartass.xl import styling
    names = _names(wb)
    rects = book_rects([wb], names)
    before = read_book(xl.app, wb, rects, mask=True)
    t0 = time.perf_counter()
    styling.ensure_styles(wb)
    ms = (time.perf_counter() - t0) * 1000
    after = read_book(xl.app, wb, rects, mask=True)
    changed = {item: keys for item, keys in diff(before, after).items() if item not in ("styles", "workbook names")}
    marker = [k for k in diff(before, after).get("workbook names", []) if "_tnStylesVer" not in str(k)]
    check("new workbook: the TN styles installed in a built workbook leave its sheets as they were",
          styling.has_styles(wb) and int(wb.Styles.Count) > 100 and not changed and not marker,
          f"styles {int(wb.Styles.Count)}, changed {dict(changed)} {marker}")
    info(f"the TN styles installed in a built workbook in {ms:.0f} ms ({int(wb.Styles.Count)} styles)")


def _failed_new_workbook(xl: Excel, seams: dict, folder: Path) -> None:
    """A first answer whose built workbook's finish fails (an operation on a sheet that is not there): the error
    says the workbook was closed without saving and its file deleted; it is not open, the file is gone, and the
    four settings and the active workbook are as before."""
    real = movein.calculate

    def calculate(t, sheets, b=None):
        real(t, sheets, b)
        if b is not None:
            b.workbook(jobs.book_name(t)).Worksheets("No such sheet").Calculate()
    other = xl.fresh("around", ("Sources",))
    before = (int(xl.app.Workbooks.Count), [getattr(xl.app, n) for n, _ in runner.QUIET],
              str(xl.app.ActiveWorkbook.Name), sorted(p.name for p in folder.iterdir()))
    holder: dict = {}
    with _patched(seams | {(exceljob, "_target"): _new_target_seam(xl, holder), (movein, "calculate"): calculate}):
        try:
            exceljob.excel_job(None, _pres()[0][2])
            failed = ""
        except Exception as e:   # noqa: BLE001 - the failure is the point
            failed = str(e)
    after = (int(xl.app.Workbooks.Count), [getattr(xl.app, n) for n, _ in runner.QUIET],
             str(xl.app.ActiveWorkbook.Name), sorted(p.name for p in folder.iterdir()))
    t = holder.get("t")
    check("new workbook: a built one that fails in its finish is closed unsaved and its file deleted",
          "closed without saving and its file deleted" in failed and before == after and t is not None
          and t.made is not None and not t.made.exists(), f"raised {failed[-120:]!r}; before {before}, after {after}")
    xl.close(other)


# ---------------------------------------------------------------------------------------------------- 3b. failed jobs
def _state(xl: Excel, wb) -> dict:
    """What a failed job must leave as it was: the sheets in order with their visibility, the tables, the four quiet
    settings, the active workbook and sheet."""
    return {"sheets": [(str(ws.Name), int(ws.Visible)) for ws in wb.Worksheets],
            "tables": sorted(str(lo.Name) for ws in wb.Worksheets for lo in ws.ListObjects),
            "settings": [getattr(xl.app, name) for name, _ in runner.QUIET],
            "active": (str(xl.app.ActiveWorkbook.Name), str(xl.app.ActiveSheet.Name))}


def failed_jobs(xl: Excel) -> None:
    """(Task 2b) A follow-up and a new answer whose finish batch fails (an operation on a sheet that is not there,
    after the park and the insert went through), through both executors: the clean-up batches (rollback, restore,
    the land on the user's sheet) leave the workbook and Excel as they were, no 'old ...' sheet behind."""
    pres = _pres()
    real = movein.calculate

    def calculate(t, sheets, b=None):
        real(t, sheets, b)
        if b is not None:
            b.workbook(jobs.book_name(t)).Worksheets("No such sheet").Calculate()
    for side, inproc in (("com", False), ("inproc", True)):
        wb = xl.fresh(f"fail_{side}", ("Sources",))
        target = lambda h, target, title, build_new=False, wb=wb: Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0, created=True)  # noqa: E731
        seams = {(exceljob, "_target"): target, (exceljob, "_save"): lambda job: None,
                 (exceljob, "remember"): lambda pre, out: None, (exceljob, "datetime"): _Now,
                 (config, "INPROC"): inproc}
        with _patched(seams):
            for _, _, pre in pres[:2]:
                exceljob.excel_job(None, pre)
            for label, job in (("a follow-up", lambda: exceljob.redraw_job(None, pres[2][2], 1, "this workbook")),
                               ("a new answer", lambda: exceljob.excel_job(None, pres[0][2]))):
                before = _state(xl, wb)
                with _patched({(movein, "calculate"): calculate}):
                    try:
                        job()
                        failed = None
                    except Exception as e:   # noqa: BLE001 - the failure is the point
                        failed = f"{type(e).__name__}: {str(e)[:90]}"
                after = _state(xl, wb)
                left = [n for n, _ in after["sheets"] if n.startswith("old ")]
                changed = [k for k in before if before[k] != after[k]]
                check(f"failures ({side}): {label} that fails in its finish leaves the workbook and Excel as they were",
                      failed is not None and not changed and not left,
                      f"raised {failed}; changed {changed}; left {left}" if changed or left or not failed
                      else f"raised {failed}")
        xl.close(wb)


# ---------------------------------------------------------------------------------------------------- 3d. re-pointing
def _structured(name: str) -> str:
    """A column name as a structured reference writes it (', [, ], # escaped with ')."""
    return "".join("'" + ch if ch in "'[]#" else ch for ch in name)


def _point_at_q1(xl: Excel, wb) -> dict:
    """(Task 3) The user's formulas that point at Q1, as a summary sheet of their own holds them: a cell reference, a
    table reference, a cross-sheet SUM, a reference to the Answer, a legacy array, a dynamic array's anchor, a merged
    cell, a table's calculated column, a defined name that is a plain reference and one with a function and a decimal
    in the user's notation (RefersToLocal), cells using both, a chart series, a chart sheet, and a cell on Q2's
    Answer (another answer). Returns {where: (how it is read, its text as Excel keeps it, how to read its target or
    None)}."""
    data = wb.Worksheets("Q1 Data")
    lo = data.ListObjects("Data_Q1")
    heads = [str(h) for h in lo.HeaderRowRange.Value[0]]
    k = next(i for i in range(1, len(heads) + 1) if _is_num(lo.DataBodyRange.Cells(1, i).Value2))
    cell, column = str(lo.DataBodyRange.Cells(1, k).Address), str(lo.ListColumns(k).DataBodyRange.Address)
    answer = next(a for a in ("A1", "B1", "A2", "B2", "B3") if wb.Worksheets("Q1 Answer").Range(a).Value2 is not None)
    intl = xl.raw.International                                     # read whole: 0-based
    sep, dec = str(intl[4]), str(intl[2])                           # the user's list and decimal separators

    def value(sheet, addr):
        return lambda: wb.Worksheets(sheet).Range(addr).Value2

    def total():
        return sum(v for (v,) in wb.Worksheets("Q1 Data").Range(column).Value2 if _is_num(v))

    def least():
        return min(v for (v,) in wb.Worksheets("Q1 Data").Range(column).Value2 if _is_num(v))

    def plus():
        return value("Q1 Data", cell)() + 1.5
    s = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
    s.Name = "Summary"
    wb.Names.Add("Q1First", f"='Q1 Data'!{cell}")
    wb.Names.Add("Q1Plus", f"='Q1 Data'!{cell}")
    wb.Names("Q1Plus").RefersToLocal = f"=SUM('Q1 Data'!{cell}{sep}1{dec}5)"
    cells = {"Summary!A1": (f"='Q1 Data'!{cell}", value("Q1 Data", cell)),
             "Summary!A2": (f"=SUM(Data_Q1[{_structured(heads[k - 1])}])", total),
             "Summary!A3": (f"=SUM('Q1 Data'!{column})", total),
             "Summary!A4": (f"='Q1 Answer'!{answer}", value("Q1 Answer", answer)),
             "Summary!A5": ("=Q1First", value("Q1 Data", cell)),
             "Summary!A6": ("=Q1Plus", plus),
             "Summary!E1": (f"=SORT('Q1 Data'!{column})", least),
             "Q2 Answer!Z1": (f"='Q1 Data'!{cell}", value("Q1 Data", cell))}
    for where, (formula, _) in cells.items():
        sheet, addr = where.split("!")
        wb.Worksheets(sheet).Range(addr).Formula2 = formula
    s.Range("C1:C2").FormulaArray = f"='Q1 Data'!{cell}*1"
    s.Range("G1:H2").Merge()                                        # a merged cell: its formula is the anchor's
    s.Range("G1").Formula2 = f"='Q1 Data'!{cell}"
    s.Range("J1:K4").Value2 = (("k", "v"), ("a", None), ("b", None), ("c", None))
    s.ListObjects.Add(1, s.Range("J1:K4"), None, 1).Name = "UserTable"
    s.Range("K2:K4").Formula2 = f"='Q1 Data'!{cell}*2"               # a calculated column (an automated write of
                                                                    # one cell does not fill it: live 20.11.2026)
    s.ChartObjects().Add(300, 10, 300, 200).Chart.SetSourceData(data.Range(column))
    sheet_chart = wb.Charts.Add()
    sheet_chart.SetSourceData(data.Range(column))
    sheet_chart.Name = "Q1 Chart"
    s.Activate()
    refs = {where: ("cell", fn) for where, (_, fn) in cells.items()}
    refs["Summary!C1"] = ("array", value("Q1 Data", cell))
    refs["Summary!G1"] = ("cell", value("Q1 Data", cell))

    def double():
        return value("Q1 Data", cell)() * 2
    for addr in ("K2", "K3", "K4"):
        refs[f"Summary!{addr}"] = ("cell", double)
    refs["the name Q1First"] = ("name", None)
    refs["the name Q1Plus"] = ("name", None)
    refs["a chart series on Summary"] = ("series", None)
    refs["a chart series on Q1 Chart"] = ("chart", None)
    return {where: (how, _read(wb, where, how)[0], target) for where, (how, target) in refs.items()}


def _read(wb, where: str, how: str) -> tuple:
    """(the text of the formula at where, the value it shows or None) as Excel keeps them: a cell's Formula2, a legacy
    array's FormulaArray, a name's RefersTo and RefersToLocal, a series' formula."""
    if how == "name":
        n = wb.Names(where.rpartition(" ")[2])
        return f"{n.RefersTo} | {n.RefersToLocal}", None
    if how == "series":
        return str(wb.Worksheets("Summary").ChartObjects(1).Chart.SeriesCollection(1).Formula), None
    if how == "chart":
        return str(wb.Charts("Q1 Chart").SeriesCollection(1).Formula), None
    sheet, addr = where.split("!")
    rng = wb.Worksheets(sheet).Range(addr)
    return str(rng.FormulaArray if how == "array" else rng.Formula2), rng.Value2


def _pointing_problems(wb, refs: dict, q2: bool = True) -> list[str]:
    """What is wrong with the formulas that point at Q1 now: a formula that is not as the user wrote it, a value that
    is not its target's (compared, never printed), a link to another workbook, an 'old ...' sheet left."""
    bad = []
    for where, (how, formula, target) in refs.items():
        if where.startswith("Q2 ") and not q2:
            continue                                    # Q2's own redraw replaced its Answer, and the cell with it
        now, value = _read(wb, where, how)
        if now != formula:
            bad.append(f"{where} reads {now!r}")
        elif target is not None:
            want = target()
            same = (abs(value - want) < 1e-9 * max(1.0, abs(want)) if _is_num(value) and _is_num(want)
                    else value == want)
            if not same:
                bad.append(f"{where} does not show its target's value")
    links = wb.LinkSources(1)
    if links:
        bad.append(f"links to {len(links)} other workbook(s)")
    bad += [f"the sheet {n!r} is left" for n in _names(wb) if n.startswith("old ")]
    return bad


def repoint_bench(xl: Excel) -> None:
    """(Review ruling 2) A follow-up of Q1 with 500 cells pointing at it, against one with none, through both
    executors: the draw's seconds and the close-out's (its writes run with Calculation manual). Then (phase 3c Task 3)
    the 500 cells on a sheet whose used range reaches the last row and the last column (a stray format in
    XFD1048576): the look for them is halved by cells, so the draw costs about the same."""
    pres = _pres()
    for side, inproc in (("com", False), ("inproc", True)):
        for count, wide in ((0, False), (500, False), (500, True)):
            wb = xl.fresh(f"bench_{side}_{count}", ("Sources",))
            seams = {(exceljob, "_target"): _target_of(xl, wb), (exceljob, "_save"): lambda job: None,
                     (exceljob, "remember"): lambda pre, out: None, (exceljob, "datetime"): _Now,
                     (config, "INPROC"): inproc}
            with _patched(seams):
                exceljob.excel_job(None, pres[0][2])
                if count:
                    cell = str(wb.Worksheets("Q1 Data").ListObjects("Data_Q1").DataBodyRange.Cells(1, 1).Address)
                    s = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
                    s.Name = "Summary"
                    s.Range(f"A1:A{count}").Formula2 = f"='Q1 Data'!{cell}"
                    if wide:
                        s.Range("XFD1048576").Interior.Color = 255          # the used range to the last cell
                out, timings = exceljob.with_timings(
                    lambda h: exceljob.redraw_job(h, pres[2][2], 1, "this workbook"))(None)
                left = [] if not count else [a for a, (f,) in enumerate(s.Range(f"A1:A{count}").Formula2, 1)
                                             if f != f"='Q1 Data'!{cell}"]
            where = " on a sheet whose used range reaches XFD1048576" if wide else ""
            info(f"re-pointing bench ({side}): a follow-up with {count} pointing cells{where}: draw "
                 f"{timings.get('draw', 0):.2f} s, close out {timings.get('draw.close out', 0):.2f} s"
                 + (f"; {len(left)} not re-pointed" if left else ""))
            if count:
                check(f"re-pointing bench ({side}): all {count} cells re-pointed{where}", not left)
            xl.close(wb)


def _target_of(xl: Excel, wb):
    """exceljob._target's stand-in: wb as our own analysis file in this script's Excel."""
    return lambda h, target, title, build_new=False: Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0, created=True)


def repointing(xl: Excel) -> None:
    """(Task 3) Formulas elsewhere that point at an answer survive its follow-ups and a refresh-style redraw (Q1, then
    Q2, as fin_refresh does), through both executors: after every step each formula reads as the user wrote it and
    shows its target's value in the new answer. Then a follow-up whose drop fails, after the re-pointing in the same
    batch: the formulas pointed back, the old answer restored, as before."""
    pres = _pres()

    def drop(t, parked, book=None, b=None):             # a drop Excel refuses, after the re-pointing
        b.workbook(jobs.book_name(t)).Worksheets(["No such sheet"]).Delete()
        return b.position - 1
    for side, inproc in (("com", False), ("inproc", True)):
        wb = xl.fresh(f"repoint_{side}", ("Sources",))
        seams = {(exceljob, "_target"): _target_of(xl, wb), (exceljob, "_save"): lambda job: None,
                 (exceljob, "remember"): lambda pre, out: None, (exceljob, "datetime"): _Now,
                 (config, "INPROC"): inproc}
        with _patched(seams):
            exceljob.excel_job(None, pres[0][2])
            exceljob.excel_job(None, pres[1][2])
            refs = _point_at_q1(xl, wb)
            check(f"re-pointing ({side}): the formulas point at Q1 before", not _pointing_problems(wb, refs))
            steps = [("a follow-up", lambda: [exceljob.redraw_job(None, pres[2][2], 1, "this workbook")]),
                     ("a second follow-up", lambda: [exceljob.redraw_job(None, pres[0][2], 1, "this workbook")]),
                     ("a refresh (Q1, then Q2)", lambda: [exceljob.redraw_job(None, pres[0][2], 1, "this workbook"),
                                                          exceljob.redraw_job(None, pres[1][2], 2, "this workbook")])]
            for label, step in steps:
                t0 = time.perf_counter()
                outs = step()
                secs = time.perf_counter() - t0
                bad = _pointing_problems(wb, refs, q2=not label.startswith("a refresh"))
                notes = [w for out in outs for w in out.get("warnings") or [] if "formula" in w]
                check(f"re-pointing ({side}): after {label} every formula shows the new answer", not bad and not notes,
                      f"{secs:.2f} s" if not bad and not notes else "; ".join(bad + notes))
            before = _state(xl, wb)
            with _patched({(movein, "drop"): drop}):
                try:
                    exceljob.redraw_job(None, pres[2][2], 1, "this workbook")
                    failed = None
                except Exception as e:   # noqa: BLE001 - the failure is the point
                    failed = f"{type(e).__name__}: {str(e)[:90]}"
            bad = _pointing_problems(wb, refs, q2=False)
            changed = [k for k in before if before[k] != _state(xl, wb)[k]]
            check(f"re-pointing ({side}): a follow-up whose drop fails points them back and restores the old answer",
                  failed is not None and not bad and not changed,
                  f"raised {failed}" + (f"; {'; '.join(bad)}; changed {changed}" if bad or changed else ""))
        xl.close(wb)


def protected_sheets(xl: Excel) -> None:
    """(Phase 3c Task 3) A protected sheet with a hidden formula cannot be scanned; it is warned about only when a
    cell there shows #REF! after the old sheets are dropped that did not before - a follow-up where the close-out
    runs with Calculation automatic (nothing to re-point) and one where it runs manual (a Summary formula to
    re-point), through both executors: 'Pointing' (its hidden formula points at Q1) is said with its one new #REF!,
    'Other' (a hidden formula that does not, and a #REF! from before) is not."""
    pres = _pres()
    for side, inproc in (("com", False), ("inproc", True)):
        wb = xl.fresh(f"protected_{side}", ("Sources",))
        seams = {(exceljob, "_target"): _target_of(xl, wb), (exceljob, "_save"): lambda job: None,
                 (exceljob, "remember"): lambda pre, out: None, (exceljob, "datetime"): _Now,
                 (config, "INPROC"): inproc}
        with _patched(seams):
            exceljob.excel_job(None, pres[0][2])
            sheets_ = {}
            for name, formulas in (("Pointing", ("='Q1 Answer'!A1",)), ("Other", ("=1+1", "=#REF!"))):
                s = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
                s.Name = name
                for i, f in enumerate(formulas, 2):
                    s.Range(f"B{i}").Formula2 = f
                s.Range("B2").FormulaHidden = True
                s.Protect()
                sheets_[name] = s
            for label, summary in (("automatic", False), ("manual: a formula re-pointed", True)):
                if summary:
                    p = sheets_["Pointing"]
                    p.Unprotect()
                    p.Range("B2").Formula2 = "='Q1 Answer'!A1"      # the first follow-up turned it #REF!
                    p.Protect()
                    s = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
                    s.Name = "Summary"
                    s.Range("A1").Formula2 = "='Q1 Answer'!A1"
                out = exceljob.redraw_job(None, pres[2][2], 1, "this workbook")
                said = [w for w in out.get("warnings") or [] if "could not read the formulas" in w]
                ok = (len(said) == 1 and "'Pointing'" in said[0] and "1 cell(s)" in said[0]
                      and not [w for w in said if "'Other'" in w])
                check(f"protected sheets ({side}, close-out {label}): only the one whose cell turned #REF! is said",
                      ok, "" if ok else "; ".join(said) or "nothing said")
        xl.close(wb)


# ---------------------------------------------------------------------------------------------------- 3c. label order
def label_order(xl: Excel) -> None:
    """(Task 2b, review M3) The chart labels in the job's order (movein.finish: positions and colours, calculate, the
    lifts by nudge right after it, one batch) against 0.3.0's order (positions, the tops read and the lifts set, all
    before the calculation) and against no lift at all (where Excel itself puts the labels): every fixed label's Top
    on the 'every chart kind' answer, as a measurement (in this script's hidden Excel: an Excel started this way
    would not become visible here, 20.11.2026)."""
    build = fixtures()["every chart kind"]
    folder = TMP / "build"
    folder.mkdir(exist_ok=True)
    built0 = build(folder / f"{uuid.uuid4().hex}.xlsx")
    tops = {}
    for side in ("job", "0.3.0", "unlifted"):
        built = _copy_of(built0, side)
        if side == "unlifted":
            built.label_fixes = [dataclasses.replace(f, lift_pt=0.0) for f in built.label_fixes]
        wb = xl.fresh(f"labels_{side.replace('.', '')}")
        t = Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0)
        movein.move_in(t, built, "Sources", book=movein.read_book(t))
        if side != "0.3.0":
            movein.finish(t, built)
        else:                                           # 0.3.0: labels (a read, the lift) before the calculation
            lifted = [f for f in built.label_fixes if f.lift_pt]
            b = Batch()
            wbn = b.workbook(jobs.book_name(t))
            built_nolift = dataclasses.replace(built, label_fixes=[f for f in built.label_fixes])
            for spec in built.tables:
                ws_ = wbn.Worksheets(spec.sheet)
                lo = ws_.ListObjects.Add(movein.SRC_RANGE, ws_.Range(spec.ref), MISSING, movein.HAS_HEADERS)
                lo.Name = spec.name
            for block in built.formulas:
                wbn.Worksheets(block.sheet).Range(block.ref).Formula2 = tuple(tuple(r) for r in block.grid)
            movein.fix_labels(b, wbn, built_nolift)
            ans = wbn.Worksheets(built.sheets[0])
            with b.attempt():
                for i, f in enumerate(lifted):
                    b.ret(movein._label(ans, f).Top, f"top {i}")
            res = runner.run(xl.app, b)
            b = Batch()
            ans = b.workbook(jobs.book_name(t)).Worksheets(built.sheets[0])
            with b.attempt():
                for i, f in enumerate(lifted):
                    top = res.values.get(f"top {i}")
                    if isinstance(top, (int, float)) and top > f.lift_pt:
                        movein._label(ans, f).Top = top - f.lift_pt
            runner.run(xl.app, b)
            movein.calculate(t, built.sheets)
        ws = wb.Worksheets(built.sheets[0])
        tops[side] = [g(lambda f=f: ws.ChartObjects(f.chart).Chart.SeriesCollection(f.series).Points(f.point)
                        .DataLabel.Top) for f in built.label_fixes]
        xl.close(wb)
        built.path.unlink(missing_ok=True)
    lifts = [i for i, f in enumerate(built0.label_fixes) if f.lift_pt]
    where = "a visible" if xl.raw.Visible else "a hidden"
    for side in ("job", "0.3.0"):
        rows = [f"fix {i}: {tops[side][i] - tops['unlifted'][i]:+.2f} pt" for i in lifts]
        info(f"labels ({where} Excel), {side} order against no lift: " + "; ".join(rows)
             + f" (lift {', '.join(f'{built0.label_fixes[i].lift_pt:g}' for i in lifts)} pt)")
    same = [i for i in range(len(built0.label_fixes)) if i not in lifts and tops["job"][i] == tops["unlifted"][i]
            == tops["0.3.0"][i]]
    info(f"labels: the {len(built0.label_fixes) - len(lifts)} labels without a lift equal on every side: "
         f"{len(same) == len(built0.label_fixes) - len(lifts)}")
    check("labels: the job's order lifts every lifted label by its lift",
          all(_is_num(tops['job'][i]) and _is_num(tops['unlifted'][i])
              and abs(tops["unlifted"][i] - tops["job"][i] - built0.label_fixes[i].lift_pt) < 0.01 for i in lifts),
          "; ".join(f"fix {i}: job {tops['job'][i]} unlifted {tops['unlifted'][i]}" for i in lifts))


def _is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# ---------------------------------------------------------------------------------------------------- 4. timing
def timing(xl: Excel, built0, rounds: int) -> None:
    ms, calls = defaultdict(list), defaultdict(list)
    real = jobs.com
    for _ in range(rounds):
        for side in ("copy", "insert"):
            built = _copy_of(built0, side)
            wb = xl.fresh(side)
            t = Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0)
            book = movein.read_book(t)
            n = []

            def counting(call, *a, **k):
                n.append(1)
                return real(call, *a, **k)
            with _settings_off(xl), _patched({(jobs, "com"): counting}):
                t0 = time.perf_counter()
                ROUTES[side](t, built, "Sources", book=book)
                ms[side].append((time.perf_counter() - t0) * 1000)
            calls[side].append(len(n))
            xl.close(wb)
            built.path.unlink(missing_ok=True)
    for side in ("copy", "insert"):
        info(f"timing {side}: median {statistics.median(ms[side]):.0f} ms over {rounds} rounds "
             f"({', '.join(f'{x:.0f}' for x in ms[side])}); COM calls {sorted(set(calls[side]))}")
    check("timing: the insert is not slower than the copy",
          statistics.median(ms["insert"]) <= statistics.median(ms["copy"]))


# ---------------------------------------------------------------------------------------------------- main
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--rounds", type=int, default=5, help="rounds per route for the timing")
    ap.add_argument("--parts", default="files,places,jobs,newbook,failures,repoint,labels,timing",
                    help="what to run (comma-separated; the proof is all of them)")
    args = ap.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")
    logging.basicConfig(level=logging.WARNING, format="  log %(levelname)s %(message)s")
    logging.getLogger("smartview.stage").setLevel(logging.WARNING)        # the jobs' stage lines
    pythoncom.CoInitialize()
    jobs.bind_job()
    xl = Excel()
    print(f"check_insert {dt.datetime.now():%Y-%m-%d %H:%M}: hidden Excel pid {xl.pid}, home {config.HOME}, "
          f"seed {'yes' if SEED.exists() else 'no (Workbooks.Add)'}", flush=True)
    try:
        import test_xlsx
        (TMP / "build").mkdir(exist_ok=True)
        answer = test_xlsx.build_with(TMP / "build")
        parts = set(args.parts.split(","))
        if "files" in parts:
            compare_fixtures(xl)
        if "places" in parts:
            placement(xl, answer)
            active_after(xl, answer)
            refusals(xl, answer)
            excel_around(xl, answer)
        if "jobs" in parts:
            compare_jobs(xl)
        if "newbook" in parts:
            new_workbook(xl)
        if "failures" in parts:
            failed_jobs(xl)
        if "repoint" in parts:
            repointing(xl)
            protected_sheets(xl)
        if "repoint-bench" in parts:
            repoint_bench(xl)
        if "labels" in parts:
            label_order(xl)
        if "timing" in parts:
            timing(xl, answer, args.rounds)
    finally:
        xl.quit()
        runner._EXCELS.clear()                         # the add-in the jobs part loaded went with that Excel
        runner._last.__dict__.clear()
        shutil.rmtree(TMP, ignore_errors=True)
    failed = [n for n, ok, _ in results if not ok]
    print(f"\n{len(results) - len(failed)}/{len(results)} checks OK" + (f"; DIFF: {', '.join(failed)}" if failed else ""))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
