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
3. Whole jobs (pipeline.excel_job, a second question with a Lines sheet, then pipeline.redraw_job: a follow-up's
   replacement) once per route in a workbook of its own, compared sheet by sheet; sheet = model per step, the move
   in's seconds and COM calls.
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
from smartass import pipeline  # noqa: E402
from smartass.core import config  # noqa: E402
from smartass.render import xlsx  # noqa: E402
from smartass.render.layout import ChartSpec  # noqa: E402
from smartass.xl import fastcom, jobs, movein  # noqa: E402
from smartass.xl.hostproc import windows_of  # noqa: E402
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
    found = movein.taken(t.wb, built.sheets, [s.name for s in built.tables], book)
    if found:
        raise movein.NameTaken(", ".join(found))
    tables = movein.file_parts(built.path)[1]
    if tables:
        raise ValueError(f"the built file holds an Excel Table ({', '.join(tables)})")
    before, after_sheet = movein._anchor(book, after, first)
    path, names = str(built.path), list(built.sheets)
    ws = lambda s: s.ws if s.ws is not None else t.wb.Worksheets(s.name)    # noqa: E731
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
    book.sheets[k:k] = [movein.Sheet(n, None, -1, ANSWER_TAB if i == 0 else DATA_TAB) for i, n in enumerate(names)]
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
            with movein.quiet(xl.app, wb):
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
                    with movein.quiet(xl.app, t.wb):
                        movein.finish(t, built)
                        movein.calculate(t, built.sheets)
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
            with movein.quiet(xl.app, wb):
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
    """quiet's four settings without its landing: what the insert itself leaves active."""
    before = [getattr(xl.app, name) for name, _ in movein.QUIET]
    for name, value in movein.QUIET:
        setattr(xl.app, name, value)
    try:
        yield
    finally:
        for (name, _), value in reversed(list(zip(movein.QUIET, before))):
            setattr(xl.app, name, value)


def active_after(xl: Excel, built0) -> None:
    """What is active right after the route itself (quiet's settings, not its landing), with the target the active
    workbook and with another workbook active; and after the job's quiet (which lands on the user's sheet)."""
    for label, other_active in (("the target active", False), ("another workbook active", True)):
        for how, around in (("right after the route", _settings_off), ("after quiet", None)):
            got = {}
            for side in ("copy", "insert"):
                built = _copy_of(built0, side)
                wb = xl.fresh(side)
                other = xl.fresh("other", ("Mine",)) if other_active else None
                if other is not None:
                    other.Activate()
                t = Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0)
                try:
                    with (around(xl) if around else movein.quiet(xl.app, wb)):
                        ROUTES[side](t, built, "Sources", book=movein.read_book(t))
                    got[side] = _where(xl, wb)
                except Exception as e:   # noqa: BLE001 - recorded: the copy route's BreakLink, see the report
                    got[side] = ("failed", f"{type(e).__name__} {getattr(e, 'excepinfo', None) or e}"[:160])
                for b in (wb, other):
                    if b is not None:
                        xl.close(b)
                built.path.unlink(missing_ok=True)
            if around or got["copy"][0] == "failed":
                info(f"active {how}, {label}: copy {got['copy']} | insert {got['insert']}")
            if around:
                continue
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
    with movein.quiet(xl.app, wb):
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
            with movein.quiet(xl.app, wb):
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
        with movein.quiet(xl.app, wb):
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


def compare_jobs(xl: Excel) -> None:
    from smartass.core.timing import COM_PREFIX
    books, steps = {}, defaultdict(dict)
    for side in ROUTES:
        wb = xl.fresh(side, ("Sources",))
        books[side] = wb
        target = lambda h, target, title, wb=wb: Target(app=xl.app, wb=wb, pid=xl.pid, hwnd=0, created=True)  # noqa: E731
        seams = {(pipeline, "_target"): target, (pipeline, "_save"): lambda t: None,
                 (pipeline, "remember"): lambda pre, out: None, (pipeline, "datetime"): _Now,
                 (movein, "move_in"): ROUTES[side]}
        with _patched(seams):
            for label, q, pre in _pres():
                job = (lambda h, pre=pre: pipeline.excel_job(h, pre)) if q is None else (
                    lambda h, pre=pre, q=q: pipeline.redraw_job(h, pre, q, "this workbook"))
                out, timings = pipeline.with_timings(job)(None)
                steps[label][side] = ((out.get("checks") or {}).get("sheet_equals_model"),
                                      timings.get("draw.move in", 0.0),
                                      int(timings.get(COM_PREFIX + "draw.move in", 0)),
                                      int(sum(v for k, v in timings.items() if k.startswith(COM_PREFIX)
                                              and "." not in k[len(COM_PREFIX):])))
    for label, by in steps.items():
        ok = all(v[0] == "OK" for v in by.values())
        check(f"jobs: {label}: sheet = model on both routes", ok, f"{ {k: v[0] for k, v in by.items()} }")
        info(f"{label}: move in copy {by['copy'][1]:.2f} s / {by['copy'][2]} calls, insert {by['insert'][1]:.2f} s"
             f" / {by['insert'][2]} calls; the job's COM calls copy {by['copy'][3]}, insert {by['insert'][3]}")
    names = {side: _names(wb) for side, wb in books.items()}
    check("jobs: the same sheets in the same order", names["copy"] == names["insert"],
          "" if names["copy"] == names["insert"] else f"copy {names['copy']} | insert {names['insert']}")
    rects = book_rects(list(books.values()), names["copy"])
    snaps = {side: read_book(xl.app, wb, rects, mask=True) for side, wb in books.items()}
    found, noise = route_diff(snaps)
    items = {item for snap in snaps.values() for where in snap.values() for item in where}
    report_items("jobs", items, found, noise, unread(snaps))
    for wb in books.values():
        xl.close(wb)


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
            with movein.quiet(xl.app, wb), _patched({(jobs, "com"): counting}):
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
    ap.add_argument("--parts", default="files,places,jobs,timing",
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
        if "timing" in parts:
            timing(xl, answer, args.rounds)
    finally:
        xl.quit()
        shutil.rmtree(TMP, ignore_errors=True)
    failed = [n for n, ok, _ in results if not ok]
    print(f"\n{len(results) - len(failed)}/{len(results)} checks OK" + (f"; DIFF: {', '.join(failed)}" if failed else ""))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
