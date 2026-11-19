r"""Live parity gate of the two renderers (light and fast, phase 3, Task 6): every test-set case drawn once with
SV_RENDER=com (pipeline.excel_job, the reference) into one new workbook and once with SV_RENDER=file
(pipeline.file_job) into another, then each pair of Q sheets read back through the user's Excel and compared.

Usage (from the repo root; Excel open, signed in to EPM):
    engine\.venv\Scripts\python.exe scripts\compare_renders.py [--case b01 ...] [--close]

config.RENDER is read at import, so each renderer draws in its own child process (this script with --child draw,
the cases as bench_answers.py runs them: the three without a chart block skip guard_chart). The first case creates
the side's workbook in the analyser folder (target 'new'), every later case is drawn into that workbook by its path.
A third child reads both workbooks (one user-Excel job per side and case) over the union of each sheet pair's used
ranges and writes one dump per side; the parent diffs the dumps and deletes them (they hold rounded numbers). Nothing
is written to EPM, and nothing is changed in Excel but the two workbooks this run created (reading the freeze panes
activates their sheets; the active sheet and ScreenUpdating are put back). --close closes those two workbooks
afterwards (the engine saved them); without it they stay open to be looked at side by side. --dir keeps the work
folder (the dumps and report.txt); --skip-draw / --skip-read reuse what it holds, --exact reads every cell on its own
(slow; checks the reader's splitting).

Compared, per cell of each Q sheet pair: Formula2 of formulas, Value2 (numbers within 0.005; texts by a hash, the
run's fetch timestamp masked), number format (.NumberFormat, en-US), font name / size / bold / italic / colour, fill,
each cell's own top and bottom border (from Excel's XML Spreadsheet of the range), horizontal and vertical alignment,
wrap, merged, the cell style's name; per sheet: row heights, column widths, merges, freeze panes, gridlines, tab
colour, print area, orientation, zoom / fit, print title rows, manual page breaks, tables, sheet-level names and the
charts (type, Overlap and GapWidth of group 1, legend and its font, the category labels' font, series formulas,
types, fill and line colours, data labels with their number format, font and position, the value axis, title, size
and top-left cell, and the label position and font colour of every labelled point of line and stacked charts); per
answer the reply's keys, sheets, chart titles, layout findings, warnings (numbers masked) and the sheet = model status.

Printed: counts, addresses, property names and format values; never a cell value (a formula only when it is one),
a chart title or a value axis bound (a waterfall's comes from the figures): those only as 'differs'.
Allowed (controller rulings): the cell style's name, the hidden sheet name smartass_pagesetup, a Data column width
1.0 character under to 3.0 over AutoFit's, a chart's width and height within 0.5 pt, a line chart group's Overlap.
Exit 0 when every difference is allowed, 1 otherwise, 2 when a run could not proceed."""
from __future__ import annotations

import argparse
import asyncio
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve()
SIDES = ("com", "file")
ERR = "<error>"                         # a property that could not be read (never None: None means mixed)
STAMP = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}(?: [+-]\d{4})?")      # datasheet.fetched_text: a run artefact
XL_NONE = -4142
EDGE_TOP, EDGE_BOTTOM = 8, 9                                            # xlEdgeTop, xlEdgeBottom
XL_VALUE, XL_CATEGORY = 2, 1                                            # xlValue, xlCategory
EN_US = 0x409                           # the LCID of every raw read: NumberFormat in en-US ('General', 'dd/mm/yyyy')
ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
          -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826238: "#SPILL!",
          -2146826215: "#CALC!"}
POINT_LABEL_TYPES = {4, 65, 52, 53}     # line (and with markers), stacked and 100 % stacked columns
TOLERANCE = 0.005
DATA_NARROWER, DATA_WIDER = 1.0, 3.0    # a Data width may be this many characters under / over AutoFit's (ruled)
CHART_SIZE = 0.5                        # pt: a chart's width and height (XlsxWriter sizes in whole pixels; ruled)
LINE_TYPES = {4, 65}                    # xlLine, xlLineMarkers: a line group's Overlap means nothing (ruled)


# ---------------------------------------------------------------------------------------------------- addresses
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
    """'$A$1:$N$60' (or '$B$4') -> (r1, c1, r2, c2)."""
    parts = addr.replace("$", "").split(":")
    out = []
    for p in parts:
        m = re.fullmatch(r"([A-Z]+)(\d+)", p)
        out.append((int(m.group(2)), col_number(m.group(1))))
    (r1, c1), (r2, c2) = out[0], out[-1]
    return r1, c1, r2, c2


def rect_text(r) -> str:
    return f"{a1(r[0], r[1])}:{a1(r[2], r[3])}"


def normalise(s, q: int):
    """Q<n> (sheet, table and name references) as Q#, so a pair whose q differ compares."""
    if not isinstance(s, str):
        return s
    return re.sub(rf"(?<![A-Za-z0-9])Q{q}(?!\d)", "Q#", s)


def hashed(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:12]


# ---------------------------------------------------------------------------------------------------- reading
def _get(fn):
    """One COM read through the job runner (busy retry); ERR when it fails."""
    from smartass.xl import jobs
    try:
        return jobs.com(fn, limit=20)
    except jobs.JobAborted:
        raise
    except Exception:   # noqa: BLE001 - a property that cannot be read is recorded, never fatal
        return ERR


class _Raw:
    """Property reads straight on the IDispatch by a cached DISPID: one cross-process call per property (a dynamic
    pywin32 object asks Excel for its type information on every new object, about 15 ms a read)."""

    def __init__(self):
        import pythoncom
        self.get_flag, self.ids = pythoncom.DISPATCH_PROPERTYGET, {}

    def __call__(self, disp, kind: str, name: str, *args):
        dispid = self.ids.get((kind, name))
        if dispid is None:
            dispid = self.ids[(kind, name)] = disp.GetIDsOfNames(name)
        return disp.Invoke(dispid, EN_US, self.get_flag, True, *args)

    def item(self, disp, index):
        return disp.Invoke(0, EN_US, self.get_flag, True, index)      # a collection's default member


class _Rect:
    """One rectangle of a sheet, its Font, Interior and Borders fetched once for every property read on it."""

    def __init__(self, raw: _Raw, ws, addr: str):
        self.raw, self.rng, self.subs = raw, _busy(lambda: raw(ws, "Worksheet", "Range", addr)), {}

    def get(self, name: str):
        return _busy(lambda: self.raw(self.rng, "Range", name))

    def sub(self, obj: str, name: str):
        if obj not in self.subs:
            self.subs[obj] = self.get(obj)
        return _busy(lambda: self.raw(self.subs[obj], obj, name))

    def border(self, edge: int, name: str):
        key = f"Border{edge}"
        if key not in self.subs:
            self.subs[key] = _busy(lambda: self.raw.item(self.get("Borders"), edge))
        return _busy(lambda: self.raw(self.subs[key], "Border", name))


def _busy(fn):
    """A read retried while Excel is busy (jobs.retry_busy), not counted as one of the job's COM calls."""
    from smartass.xl import jobs
    return jobs.retry_busy(fn, 20)


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
    return None if s is None else str(_busy(lambda: x.raw(s, "Style", "Name")))


CELL_PROPS = {
    "number format": lambda x: x.get("NumberFormat"),
    "font name": lambda x: x.sub("Font", "Name"),
    "font size": lambda x: x.sub("Font", "Size"),
    "bold": lambda x: x.sub("Font", "Bold"),
    "italic": lambda x: x.sub("Font", "Italic"),
    "font colour": lambda x: x.sub("Font", "Color"),
    "fill": _fill,
    "horizontal alignment": lambda x: x.get("HorizontalAlignment"),
    "vertical alignment": lambda x: x.get("VerticalAlignment"),
    "wrap": lambda x: x.get("WrapText"),
    "merged": lambda x: x.get("MergeCells"),
    "style": _style,
}


def _plain(v):
    """A COM value as JSON (floats that are whole as ints, so 9.0 and 9 compare)."""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, (bool, int, float, str)) or v is None:
        return v
    return str(v)


def _read(fn, x):
    try:
        return fn(x)
    except Exception:   # noqa: BLE001 - a property that cannot be read is recorded, never fatal
        return ERR


def _walk(raw: _Raw, ws, rect, props: dict, out: dict, exact: bool) -> None:
    """Each property over rect: one read when the rect is uniform (a mixed range reads None), else split in two
    until it is, columns first (a table's columns are uniform below its header); out[prop] gets (r1, c1, r2, c2,
    value) rectangles."""
    r1, c1, r2, c2 = rect
    single = r1 == r2 and c1 == c2
    if exact and not single:
        rest = dict(props)
    else:
        try:
            x = _Rect(raw, ws, rect_text(rect))
        except Exception:   # noqa: BLE001
            x = None
        rest = {}
        for name, fn in props.items():
            v = ERR if x is None else _read(fn, x)
            if v is None or (v == ERR and not single):
                if single:
                    out[name].append((r1, c1, r2, c2, None))
                else:
                    rest[name] = fn
            else:
                out[name].append((r1, c1, r2, c2, _plain(v)))
    if not rest:
        return
    if c2 > c1:
        m = (c1 + c2) // 2
        halves = ((r1, c1, r2, m), (r1, m + 1, r2, c2))
    else:
        m = (r1 + r2) // 2
        halves = ((r1, c1, m, c2), (m + 1, c1, r2, c2))
    for h in halves:
        _walk(raw, ws, h, rest, out, exact)


SS = "{urn:schemas-microsoft-com:office:spreadsheet}"
XML_SPREADSHEET = 11                    # xlRangeValueXMLSpreadsheet: the range as Excel's XML Spreadsheet 2003


def _style_borders(styles) -> dict:
    """{style id: {'Top': 'LineStyle:Weight:Color', ...}}: a style without a Borders element takes its parent's."""
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


def _walk_borders(raw: _Raw, ws, rect, out: dict) -> None:
    """Every cell's own top and bottom border, from the range as Excel's XML Spreadsheet (one read). A multi-cell
    range's Borders read a style while some of its cells have none, and xlInsideHorizontal none over rows that have
    one (live, 19.11.2026), so borders cannot be split like the other properties, and one read per cell is too slow
    for a Lines sheet. A cell takes its style's borders, a cell the XML leaves out those of its row's or column's
    style, a merged cell those of the merge's first cell."""
    import xml.etree.ElementTree as ET
    r1, c1, r2, c2 = rect
    try:
        root = ET.fromstring(_busy(lambda: raw(raw(ws, "Worksheet", "Range", rect_text(rect)), "Range", "Value",
                                                XML_SPREADSHEET)))
    except Exception:   # noqa: BLE001 - recorded as unread, never fatal
        for name in ("top border", "bottom border"):
            out[name].append((r1, c1, r2, c2, ERR))
        return
    borders = _style_borders(root.iter(SS + "Style"))
    table = next(root.iter(SS + "Table"))
    width = c2 - c1 + 1
    col_style = []
    for col in table.findall(SS + "Column"):
        i = int(col.get(SS + "Index", len(col_style) + 1))
        col_style += ["Default"] * (i - 1 - len(col_style))
        col_style += [col.get(SS + "StyleID", "Default")] * (1 + int(col.get(SS + "Span", 0)))
    col_style += ["Default"] * (width - len(col_style))
    rows = {}
    r = 0
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
    for i in range(1, r2 - r1 + 2):
        styles = rows.get(i) or col_style
        for name, pos in (("top border", "Top"), ("bottom border", "Bottom")):
            run = None
            for j, sid in enumerate(styles, 1):
                v = borders.get(sid, {}).get(pos, "none")
                if run and run[1] == v:
                    run[0] = j
                    continue
                if run:
                    out[name].append((r1 + i - 1, c1 + run[2] - 1, r1 + i - 1, c1 + run[0] - 1, run[1]))
                run = [j, v, j]
            if run:
                out[name].append((r1 + i - 1, c1 + run[2] - 1, r1 + i - 1, c1 + run[0] - 1, run[1]))


def _contents(ws, rect, q: int) -> dict:
    """{'A1': [kind, x]} for every non-empty cell: ['f', formula] (Q normalised), ['n', rounded number], ['t', text
    hash], ['b', bool], ['e', error name]; and how many texts held the run's timestamp (masked)."""
    rng = _get(lambda: ws.Range(rect_text(rect)))
    f2 = _get(lambda: rng.Formula2)
    v2 = _get(lambda: rng.Value2)
    r1, c1, r2, c2 = rect
    grid = lambda g: g if isinstance(g, tuple) else ((g,),)      # noqa: E731 - one cell reads a scalar
    out, masked = {}, 0
    if f2 == ERR or v2 == ERR:
        return {"error": True, "cells": {}, "masked": 0}
    for i, (frow, vrow) in enumerate(zip(grid(f2), grid(v2))):
        for j, (f, v) in enumerate(zip(frow, vrow)):
            addr = a1(r1 + i, c1 + j)
            if isinstance(f, str) and f.startswith("="):
                out[addr] = ["f", normalise(f, q)]
            elif v is None or v == "":
                continue
            elif isinstance(v, bool):
                out[addr] = ["b", v]
            elif isinstance(v, (int, float)):
                out[addr] = ["e", ERRORS[int(v)]] if int(v) in ERRORS and float(v).is_integer() else ["n", round(v, 3)]
            else:
                text = str(v)
                if STAMP.search(text):
                    masked += 1
                    text = STAMP.sub("<fetched>", text)
                out[addr] = ["t", hashed(normalise(text, q))]
    return {"error": False, "cells": out, "masked": masked}


def _bisect_lines(get_value, a: int, b: int, out: dict) -> None:
    """Row heights / column widths a..b: one read while they agree (a mixed range reads None)."""
    v = get_value(a, b)
    if v is not None and (v != ERR or a == b):
        out[f"{a}:{b}"] = _plain(round(v, 2) if isinstance(v, float) else v)
        return
    m = (a + b) // 2
    _bisect_lines(get_value, a, m, out)
    _bisect_lines(get_value, m + 1, b, out)


def _chart(co, q: int) -> dict:
    ch = _get(lambda: co.Chart)
    d = {"name": _get(lambda: co.Name), "top left": _get(lambda: co.TopLeftCell.Address),
         "width": _get(lambda: round(co.Width, 1)), "height": _get(lambda: round(co.Height, 1)),
         "type": _get(lambda: ch.ChartType), "has title": _get(lambda: ch.HasTitle)}
    if d["has title"] is True:
        d["title"] = _get(lambda: hashed(normalise(str(ch.ChartTitle.Text), q)))
    d["has legend"] = _get(lambda: ch.HasLegend)
    if d["has legend"] is True:
        d["legend position"] = _get(lambda: ch.Legend.Position)
        d["legend font"] = _get(lambda: [str(ch.Legend.Font.Name), ch.Legend.Font.Size, ch.Legend.Font.Color])
    d["category labels font"] = _get(lambda: [str(ch.Axes(XL_CATEGORY).TickLabels.Font.Name),
                                              ch.Axes(XL_CATEGORY).TickLabels.Font.Size,
                                              ch.Axes(XL_CATEGORY).TickLabels.Font.Color])
    group = _get(lambda: ch.ChartGroups(1))
    d["overlap"] = _get(lambda: group.Overlap)
    d["gap width"] = _get(lambda: group.GapWidth)
    axis = _get(lambda: ch.Axes(XL_VALUE))
    d["value axis reversed"] = _get(lambda: axis.ReversePlotOrder)
    for end in ("Minimum", "Maximum"):
        auto = _get(lambda end=end: getattr(axis, f"{end}ScaleIsAuto"))
        d[f"value axis {end.lower()}"] = "auto" if auto is True else _get(lambda end=end: getattr(axis, f"{end}Scale"))
    n = _get(lambda: ch.SeriesCollection().Count)
    series = []
    for k in range(1, (n if isinstance(n, int) else 0) + 1):
        s = _get(lambda k=k: ch.SeriesCollection(k))
        e = {"formula": _get(lambda: normalise(str(s.Formula), q)), "type": _get(lambda: s.ChartType),
             "axis group": _get(lambda: s.AxisGroup)}
        fill_visible = _get(lambda: s.Format.Fill.Visible)
        e["fill"] = _get(lambda: s.Format.Fill.ForeColor.RGB) if fill_visible not in (0, False, ERR) else "none"
        line_visible = _get(lambda: s.Format.Line.Visible)
        e["line"] = _get(lambda: s.Format.Line.ForeColor.RGB) if line_visible not in (0, False, ERR) else "none"
        e["has data labels"] = _get(lambda: s.HasDataLabels)
        if e["has data labels"] is True:
            e["label number format"] = _get(lambda: s.DataLabels().NumberFormat)
            e["label font"] = _get(lambda: [str(s.DataLabels().Font.Name), s.DataLabels().Font.Size,
                                            s.DataLabels().Font.Color, s.DataLabels().Font.Bold])
            e["label position"] = _get(lambda: s.DataLabels().Position)
        if d["type"] in POINT_LABEL_TYPES:
            points = {}
            count = _get(lambda: s.Points().Count)
            for i in range(1, (count if isinstance(count, int) else 0) + 1):
                p = _get(lambda i=i: s.Points(i))
                if _get(lambda: p.HasDataLabel) is True:
                    points[str(i)] = [_get(lambda: p.DataLabel.Position), _get(lambda: p.DataLabel.Font.Color)]
            e["labelled points"] = points
        series.append(e)
    d["series"] = series
    return d


def read_sheet(app, wb, name: str, q: int, union, exact: bool) -> dict:
    ws = _get(lambda: wb.Worksheets(name))
    if ws == ERR:
        return {"missing": True}
    rect = union
    out: dict = {"missing": False, "used": _get(lambda: ws.UsedRange.Address), "rect": list(rect)}
    cells: dict = defaultdict(list)
    raw = _Raw()
    _walk(raw, ws._oleobj_, rect, CELL_PROPS, cells, exact)
    _walk_borders(raw, ws._oleobj_, rect, cells)
    out["cells"] = dict(cells)
    out["contents"] = _contents(ws, rect, q)
    r1, c1, r2, c2 = rect
    heights, widths = {}, {}
    _bisect_lines(lambda a, b: _get(lambda: ws.Range(f"{a}:{b}").RowHeight), 1, r2, heights)
    _bisect_lines(lambda a, b: _get(lambda: ws.Range(f"{col_letter(a)}:{col_letter(b)}").ColumnWidth), 1, c2,
                  widths)
    out["row heights"], out["column widths"] = heights, widths
    merges = set()
    for (mr1, mc1, mr2, mc2, v) in cells["merged"]:
        if v is True:
            for r in range(mr1, mr2 + 1):
                for c in range(mc1, mc2 + 1):
                    if not any(m[0] <= r <= m[2] and m[1] <= c <= m[3] for m in merges):
                        area = _get(lambda r=r, c=c: ws.Cells(r, c).MergeArea.Address)
                        if area != ERR:
                            merges.add(parse_rect(area))
    out["merges"] = sorted(rect_text(m) for m in merges)
    page = _get(lambda: ws.PageSetup)
    for key, attr in (("print area", "PrintArea"), ("orientation", "Orientation"), ("zoom", "Zoom"),
                      ("fit pages wide", "FitToPagesWide"), ("fit pages tall", "FitToPagesTall"),
                      ("print title rows", "PrintTitleRows"), ("centre horizontally", "CenterHorizontally")):
        out[key] = _plain(_get(lambda attr=attr: getattr(page, attr)))
    out["tab colour"] = _plain(_get(lambda: ws.Tab.Color))
    breaks = _get(lambda: [(int(b.Location.Row), int(b.Type)) for b in ws.HPageBreaks if int(b.Type) == -4135])
    out["manual page breaks"] = breaks if breaks == ERR else [r for r, _ in breaks]
    out["tables"] = _get(lambda: sorted([normalise(str(lo.Name), q), str(lo.Range.Address), str(lo.TableStyle.Name),
                                         bool(lo.ShowTableStyleRowStripes)] for lo in ws.ListObjects))
    out["names"] = _get(lambda: sorted([normalise(str(n.Name), q), normalise(str(n.RefersTo), q), bool(n.Visible)]
                                       for n in ws.Names))
    _get(lambda: ws.Activate())                       # freeze panes and gridlines are the window's, of the active sheet
    win = _get(lambda: app.ActiveWindow)
    out["freeze panes"] = _get(lambda: [bool(win.FreezePanes), int(win.SplitRow), int(win.SplitColumn)])
    out["gridlines"] = _get(lambda: bool(win.DisplayGridlines))
    count = _get(lambda: ws.ChartObjects().Count)
    out["chart count"] = count
    out["charts"] = [_chart(_get(lambda i=i: ws.ChartObjects(i)), q) for i in range(1, (count if isinstance(count, int)
                                                                                         else 0) + 1)]
    return out


# ---------------------------------------------------------------------------------------------------- children
def _bench():
    sys.path.insert(0, str(HERE.parent))
    import bench_answers
    return bench_answers


def _reply_summary(out: dict, scrub) -> dict:
    q = out.get("q")
    return {"keys": sorted(out), "sheets": [normalise(s, q) for s in out.get("sheets") or []],
            "charts": [hashed(normalise(c, q)) for c in out.get("charts") or []],
            "layout": [scrub(normalise(x, q)) for x in out.get("layout") or []],
            "warnings": [scrub(normalise(STAMP.sub("<fetched>", str(x)), q)) for x in out.get("warnings") or []],
            "status": ((out.get("checks") or {}).get("sheet_equals_model") or "")}


async def draw(cases: list[dict], out_path: Path) -> int:
    """Child: every case drawn with this process's renderer into ONE new workbook; the record (q, workbook, the
    reply's shape) to out_path. The reply's figures are dropped."""
    bench = _bench()
    from smartass import server
    from smartass.core import config
    from smartass.core import serverkit as kit
    from smartass.model.spec import AnalysisSpec

    async def fin_analyze(spec, target, chart=True):       # bench_answers' run: the server's own path
        if chart:
            return await server._new_question(spec, target)
        server.guard_key_points(spec)
        pre, notes = await server._prepare(spec.model_dump(), server._provider(), spec, target, complete=True)
        job = server._answer_job()
        return server._with_notes(await server._draw_job(lambda h: job(h, pre, target), "fin_analyze",
                                                         spec.question), notes)

    tool = kit.guard(fin_analyze)
    record = {"render": config.RENDER, "workbook": None, "cases": {}}
    target = "new"
    try:
        for case in cases:
            rec: dict = {"ok": False}
            try:
                spec = AnalysisSpec.model_validate(case["spec"])
                chart = bench._has_chart(spec)
                out = await tool(spec, target, chart=chart)
                rec = {"ok": True, "q": out["q"], "workbook": out["workbook"], "chart block": chart,
                       "sheets": list(out.get("sheets") or []), "reply": _reply_summary(out, bench.scrub)}
                if target == "new":
                    record["workbook"] = target = out["workbook"]
            except Exception as e:   # noqa: BLE001 - a case that fails to draw is a finding, the others go on
                rec = {"ok": False, "error": bench.scrub(e)}
            record["cases"][case["id"]] = rec
            state = f"Q{rec['q']} sheet = model {rec['reply']['status']}" if rec["ok"] else f"FAILED {rec['error']}"
            print(f"  {config.RENDER} {case['id']}: {state}", flush=True)
            out_path.write_text(json.dumps(record, indent=1), encoding="utf-8")
    finally:
        bench._shutdown(kit)
    return 0


def _union(a, b):
    if a is None:
        return b
    if b is None:
        return a
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


async def read(folder: Path, close: bool, exact: bool, only: list[str] | None = None) -> int:
    """Child: both sides' Q sheets read (the union of the pair's used ranges, from A1) into <side>.dump.json, one
    user-Excel job per side and case (a job has 600 s); with close, the two workbooks are closed afterwards (only
    when they are this run's own analysis files in the analyser folder)."""
    bench = _bench()
    from smartass import pipeline
    from smartass.core import config
    from smartass.core import serverkit as kit
    from smartass.xl import jobs
    records = {s: json.loads((folder / f"{s}.json").read_text(encoding="utf-8")) for s in SIDES}
    sides = [s for s in SIDES if records[s]["workbook"]]

    def sheets_of(side: str) -> dict:
        """{case: {kind: sheet name}} of the cases drawn; kind is the name without 'Q<n> '."""
        return {case: {s.split(" ", 1)[1]: s for s in c["sheets"]} for case, c in records[side]["cases"].items()
                if c["ok"] and (not only or case in only)}

    by_side = {s: sheets_of(s) for s in sides}

    def used_job(h):
        out = {}
        for s in sides:
            t = pipeline._target(h, records[s]["workbook"], None)
            for case, kinds in by_side[s].items():
                for kind, name in kinds.items():
                    addr = _get(lambda name=name, t=t: t.wb.Worksheets(name).UsedRange.Address)
                    out[(s, case, kind)] = None if addr == ERR else parse_rect(addr)
        return out

    def case_job(h, side: str, case: str, used: dict):
        t = pipeline._target(h, records[side]["workbook"], None)
        app, q = t.app, records[side]["cases"][case]["q"]
        other = next(o for o in SIDES if o != side)
        before = jobs.com(lambda: (app.ScreenUpdating, app.ActiveSheet))
        jobs.com(lambda: setattr(app, "ScreenUpdating", False))
        out = {}
        try:
            for kind, name in by_side[side][case].items():
                rect = _union((1, 1, 1, 1), _union(used.get((side, case, kind)), used.get((other, case, kind))))
                t0 = time.monotonic()
                out[kind] = read_sheet(app, t.wb, name, q, rect, exact)
                print(f"  read {side} {case} {kind} {rect_text(rect)}: {time.monotonic() - t0:.0f} s", flush=True)
        finally:
            ws = before[1]
            if ws is not None:
                _get(lambda: (ws.Parent.Activate(), ws.Activate()))
            _get(lambda: setattr(app, "ScreenUpdating", before[0]))
        return out

    def close_job(h):
        for s in sides:
            path = Path(records[s]["workbook"])
            if path.resolve().parent.is_relative_to(config.ANALYSER_DIR.resolve()):
                t = pipeline._target(h, str(path), None)
                jobs.com(lambda t=t: t.wb.Close(False))
                print(f"  closed {path.name}", flush=True)

    try:
        used = await kit._call_user(used_job, what="compare renders: used ranges")
        for s in sides:
            dump = {"workbook": records[s]["workbook"], "cases": {}}
            for case in by_side[s]:
                dump["cases"][case] = await kit._call_user(lambda h, s=s, case=case: case_job(h, s, case, used),
                                                           what=f"compare renders: read {s} {case}")
            (folder / f"{s}.dump.json").write_text(json.dumps(dump), encoding="utf-8")
        if close:
            await kit._call_user(close_job, what="compare renders: close")
        return 0
    finally:
        bench._shutdown(kit)


# ---------------------------------------------------------------------------------------------------- diff
class Finding:
    def __init__(self, case, sheet, prop, where, a, b, allowed: str | None = None, show: bool = True):
        self.case, self.sheet, self.prop, self.where = case, sheet, prop, where
        self.a, self.b, self.allowed, self.show = a, b, allowed, show

    def key(self):
        return (self.sheet, self.prop, repr(self.a) if self.show else "", repr(self.b) if self.show else "",
                self.allowed or "")


def expand(rects: list) -> dict:
    out = {}
    for r1, c1, r2, c2, v in rects:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                out[(r, c)] = v
    return out


def _lines_map(m: dict) -> dict:
    out = {}
    for k, v in m.items():
        a, b = (int(x) for x in k.split(":"))
        for i in range(a, b + 1):
            out[i] = v
    return out


def _same_value(x, y) -> bool:
    if x is None or y is None:
        return x == y
    if x[0] == "n" and y[0] == "n":
        return abs(x[1] - y[1]) <= TOLERANCE + 1e-9          # the rounding of the dump itself
    return x == y


SHEET_KEYS = ("print area", "orientation", "zoom", "fit pages wide", "fit pages tall", "print title rows",
              "centre horizontally", "tab colour", "manual page breaks", "tables", "freeze panes", "gridlines",
              "chart count", "merges")
CHART_KEYS = ("type", "top left", "width", "height", "has title", "title", "has legend", "legend position",
              "legend font", "category labels font", "overlap", "gap width", "value axis reversed",
              "value axis minimum", "value axis maximum")
SERIES_KEYS = ("formula", "type", "axis group", "fill", "line", "has data labels", "label number format",
               "label font", "label position", "labelled points")
HIDDEN = {"title", "value axis minimum", "value axis maximum"}   # never shown, only that they differ: a title is a
                                        # hash, a waterfall's axis bound comes from the figures


def _chart_allowed(p: str, x: dict, y: dict) -> str | None:
    """The ruled chart allowances: the size within CHART_SIZE pt, the Overlap of a line chart's group."""
    a, b = x.get(p), y.get(p)
    if p in ("width", "height") and all(isinstance(v, (int, float)) for v in (a, b)) and abs(a - b) <= CHART_SIZE:
        return f"chart size within {CHART_SIZE} pt"
    if p == "overlap" and x.get("type") in LINE_TYPES and y.get("type") in LINE_TYPES:
        return "the Overlap of a line chart group"
    return None


def diff_sheet(case: str, kind: str, a: dict, b: dict) -> list[Finding]:
    out: list[Finding] = []
    if a.get("missing") or b.get("missing"):
        if a.get("missing") != b.get("missing"):
            out.append(Finding(case, kind, "sheet", "", "missing" if a.get("missing") else "present",
                               "missing" if b.get("missing") else "present"))
        return out
    props = sorted(set(a["cells"]) | set(b["cells"]))
    for prop in props:
        ma, mb = expand(a["cells"].get(prop, [])), expand(b["cells"].get(prop, []))
        for rc in sorted(set(ma) | set(mb)):
            va, vb = ma.get(rc, "<unread>"), mb.get(rc, "<unread>")
            if va != vb:
                allowed = "the cell style name" if prop == "style" else None
                out.append(Finding(case, kind, prop, a1(*rc), va, vb, allowed))
    ca, cb = a["contents"], b["contents"]
    if ca["error"] or cb["error"]:
        out.append(Finding(case, kind, "contents", "", "unread" if ca["error"] else "read",
                           "unread" if cb["error"] else "read"))
    for addr in sorted(set(ca["cells"]) | set(cb["cells"]), key=lambda x: (int(re.sub(r"\D", "", x)), x)):
        x, y = ca["cells"].get(addr), cb["cells"].get(addr)
        if not _same_value(x, y):
            formula = x is not None and y is not None and x[0] == y[0] == "f"
            out.append(Finding(case, kind, "formula" if formula else "value", addr,
                               x[1] if formula else None, y[1] if formula else None, show=formula))
    for prop in ("row heights", "column widths"):
        la, lb = _lines_map(a[prop]), _lines_map(b[prop])
        for i in sorted(set(la) | set(lb)):
            x, y = la.get(i), lb.get(i)
            if x == y:
                continue
            allowed = None
            if (prop == "column widths" and kind == "Data" and isinstance(x, (int, float))
                    and isinstance(y, (int, float)) and x - DATA_NARROWER <= y <= x + DATA_WIDER):
                allowed = f"Data column width {DATA_NARROWER} under to {DATA_WIDER} over AutoFit"
            out.append(Finding(case, kind, prop, str(i) if prop == "row heights" else col_letter(i), x, y, allowed))
    for key in SHEET_KEYS:
        if a.get(key) != b.get(key):
            out.append(Finding(case, kind, key, "", a.get(key), b.get(key)))
    na = {n[0]: n for n in a.get("names") or [] if isinstance(n, list)}
    nb = {n[0]: n for n in b.get("names") or [] if isinstance(n, list)}
    for name in sorted(set(na) | set(nb)):
        if na.get(name) != nb.get(name):
            allowed = "the hidden sheet name smartass_pagesetup" if name.endswith("!smartass_pagesetup") else None
            out.append(Finding(case, kind, "sheet name", name, na.get(name), nb.get(name), allowed))
    out += diff_charts(case, kind, a.get("charts") or [], b.get("charts") or [])
    return out


def diff_charts(case: str, kind: str, ca: list, cb: list) -> list[Finding]:
    """Charts paired by their top-left cell (then title)."""
    out = []
    key = lambda d: (str(d.get("top left")), str(d.get("title")))       # noqa: E731
    pa, pb = {key(d): d for d in ca}, {key(d): d for d in cb}
    for k in sorted(set(pa) | set(pb)):
        x, y = pa.get(k), pb.get(k)
        where = f"chart at {k[0]}"
        if x is None or y is None:
            out.append(Finding(case, kind, "chart", where, "missing" if x is None else "present",
                               "missing" if y is None else "present"))
            continue
        for p in CHART_KEYS:
            if x.get(p) != y.get(p):
                out.append(Finding(case, kind, f"chart {p}", where, x.get(p), y.get(p), _chart_allowed(p, x, y),
                                   show=p not in HIDDEN))
        sx, sy = x.get("series") or [], y.get("series") or []
        if len(sx) != len(sy):
            out.append(Finding(case, kind, "chart series count", where, len(sx), len(sy)))
        for i, (s, t) in enumerate(zip(sx, sy), 1):
            for p in SERIES_KEYS:
                if s.get(p) != t.get(p):
                    out.append(Finding(case, kind, f"series {p}", f"{where} series {i}", s.get(p), t.get(p)))
    return out


def diff_replies(case: str, a: dict, b: dict) -> list[Finding]:
    out = []
    for key in ("keys", "sheets", "charts", "layout", "warnings", "status"):
        x, y = a.get(key), b.get(key)
        if x != y:
            if isinstance(x, list) and isinstance(y, list):
                only_a = [v for v in x if v not in y]
                only_b = [v for v in y if v not in x]
                out.append(Finding(case, "reply", key, "", only_a, only_b, show=key != "charts"))
            else:
                out.append(Finding(case, "reply", key, "", x, y))
    return out


def diff(records: dict, dumps: dict) -> tuple[list[Finding], dict]:
    """Every difference between the com side (reference) and the file side; and counts (cases, cells masked,
    reads that failed)."""
    out: list[Finding] = []
    stats = {"cases": 0, "masked": 0, "errors": defaultdict(int), "cells": 0}
    ra, rb = records["com"]["cases"], records["file"]["cases"]
    for case in sorted(set(ra) | set(rb)):
        x, y = ra.get(case, {"ok": False, "error": "not run"}), rb.get(case, {"ok": False, "error": "not run"})
        if not (x["ok"] and y["ok"]):
            out.append(Finding(case, "draw", "drawn", "", "ok" if x["ok"] else x.get("error"),
                               "ok" if y["ok"] else y.get("error")))
            continue
        stats["cases"] += 1
        out += diff_replies(case, x["reply"], y["reply"])
        da, db = dumps["com"]["cases"].get(case, {}), dumps["file"]["cases"].get(case, {})
        for kind in sorted(set(da) | set(db)):
            sa, sb = da.get(kind, {"missing": True}), db.get(kind, {"missing": True})
            for side, s in (("com", sa), ("file", sb)):
                if not s.get("missing"):
                    stats["masked"] += s["contents"]["masked"]
                    r = s["rect"]
                    stats["cells"] += (r[2] - r[0] + 1) * (r[3] - r[1] + 1) if side == "com" else 0
                    for prop, rects in s["cells"].items():
                        stats["errors"][(side, prop)] += sum(1 for *_, v in rects if v == ERR)
            out += diff_sheet(case, kind, sa, sb)
    return out, stats


def _shown(prop: str, v):
    if isinstance(v, int) and ("colour" in prop or prop in ("series fill", "series line")) and v >= 0:
        return f"#{v & 0xFF:02X}{v >> 8 & 0xFF:02X}{v >> 16 & 0xFF:02X}"
    return v


def report(findings: list[Finding], stats: dict) -> tuple[str, int]:
    """The findings grouped by sheet kind, property and com -> file value, with counts and the first addresses;
    the number of groups that are not allowed."""
    groups: dict = defaultdict(list)
    for f in findings:
        groups[f.key()].append(f)
    lines, bad = [], 0
    for (sheet, prop, _, _, allowed), fs in sorted(groups.items(), key=lambda kv: (bool(kv[0][4]), kv[0][:2])):
        f0 = fs[0]
        cases = sorted({f.case for f in fs})
        where = ", ".join(f"{f.case} {f.where}".strip() for f in fs[:6]) + (" ..." if len(fs) > 6 else "")
        values = f" | com {_shown(prop, f0.a)!r} -> file {_shown(prop, f0.b)!r}" if f0.show else " | differs"
        tag = f"allowed ({allowed})" if allowed else "DIFF"
        bad += 0 if allowed else 1
        lines.append(f"{tag} | {sheet} | {prop}{values} | {len(fs)} in {' '.join(cases)}: {where}")
    errs = {k: v for k, v in stats["errors"].items() if v}
    head = [f"cases compared: {stats['cases']}; cells per side: {stats['cells']}; texts with the run's timestamp "
            f"masked: {stats['masked']}; failed reads: {sum(errs.values())}"
            + (f" ({', '.join(f'{s} {p}: {n}' for (s, p), n in sorted(errs.items()))})" if errs else ""),
            f"difference groups: {len(groups)}, not allowed: {bad}"]
    return "\n".join(head + lines), bad


# ---------------------------------------------------------------------------------------------------- parent
def _child(args: list[str], env: dict | None = None) -> int:
    cmd = [sys.executable, str(HERE), *args]
    return subprocess.run(cmd, env={**os.environ, **(env or {})}).returncode


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--case", nargs="+", default=None, help="case ids (default: all)")
    ap.add_argument("--close", action="store_true", help="close the two workbooks after they are read")
    ap.add_argument("--dir", default=None, help="work folder kept afterwards (default: a new temp folder, deleted)")
    ap.add_argument("--skip-draw", action="store_true", help="reuse the draw records in --dir")
    ap.add_argument("--skip-read", action="store_true", help="reuse the dumps in --dir")
    ap.add_argument("--exact", action="store_true", help="read every cell on its own (slow; checks the reader)")
    ap.add_argument("--child", choices=("draw", "read"), default=None, help=argparse.SUPPRESS)
    ap.add_argument("--out", default=None, help=argparse.SUPPRESS)
    args = ap.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")
    if args.child == "draw":
        from smartass.testset import runner
        cases = [c for c in runner.load_cases() if not args.case or c["id"] in args.case]
        return asyncio.run(draw(cases, Path(args.out)))
    if args.child == "read":
        return asyncio.run(read(Path(args.out), args.close, args.exact, args.case))
    if (args.skip_draw or args.skip_read) and not args.dir:
        ap.error("--skip-draw and --skip-read need --dir")
    folder = Path(args.dir) if args.dir else Path(tempfile.mkdtemp(prefix="compare_renders_"))
    folder.mkdir(parents=True, exist_ok=True)
    try:
        cases = ["--case", *args.case] if args.case else []
        if not args.skip_draw:
            for side in SIDES:
                print(f"drawing with SV_RENDER={side}", flush=True)
                code = _child(["--child", "draw", "--out", str(folder / f"{side}.json"), *cases], {"SV_RENDER": side})
                if code:
                    print(f"STOPPED: the {side} draw could not proceed (exit {code})")
                    return 2
        records = {s: json.loads((folder / f"{s}.json").read_text(encoding="utf-8")) for s in SIDES}
        for r in records.values():
            r["cases"] = {k: v for k, v in r["cases"].items() if not args.case or k in args.case}
        if not args.skip_read:
            print("reading both workbooks", flush=True)
            code = _child(["--child", "read", "--out", str(folder), *cases] + (["--close"] if args.close else [])
                          + (["--exact"] if args.exact else []))
            if code:
                print(f"STOPPED: the read could not proceed (exit {code})")
                return 2
        dumps = {s: json.loads((folder / f"{s}.dump.json").read_text(encoding="utf-8")) for s in SIDES}
        findings, stats = diff(records, dumps)
        text, bad = report(findings, stats)
        print(f"\ncom workbook: {Path(records['com']['workbook'] or '-').name}")
        print(f"file workbook: {Path(records['file']['workbook'] or '-').name}\n")
        print(text)
        if args.dir:
            (folder / "report.txt").write_text(text, encoding="utf-8")
        return 1 if bad else 0
    finally:
        if not args.dir:
            shutil.rmtree(folder, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
