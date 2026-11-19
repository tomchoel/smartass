r"""Live check of xl/fastcom.py (light and fast phase 3, Task 8A): every Excel idiom the engine uses, run through
pywin32's standard late-bound wrapper and through fastcom on the same hidden Excel, must give equal results (or the
same com_error); then both are timed.

The Excel is one this script starts itself (DispatchEx, refused unless its pid is new) and quits at the end; if Quit
hangs, only that process is killed (pid, start time and "-Embedding" checked). The user's Excel is never touched. The
scratch workbook is never saved; the one file it opens (Workbooks.Open(path, 0, True)) is written with openpyxl under
config.HOME / "tmp" and deleted at the end.

Each idiom is a function of (app, wb, ws, ws2): it starts from the root Application, so every run uses new objects,
and works on its own side's sheets ("StdSide"/"StdSide2" for the standard wrapper, "FastSide"/"FastSide2" for
fastcom; strings are compared with the side's name replaced). The list runs four times - in the order of the list and
reversed, each first cold (fastcom's learnt interfaces cleared) and then warm - because what is shared per interface
shows only on the second object and depends on what was learnt first. Each fast result is compared with the
standard wrapper's of the same round. Then both are timed (median of --runs, warm, the two wrappers' runs interleaved). One line per idiom: equal yes/no
and the two medians; exit 1 on any inequality. The idioms come from the engine's com(lambda: ...) calls.

Usage (from the repo root):
    engine\.venv\Scripts\python.exe scripts\check_fastcom.py [--runs 7] [--only chart]
"""
import argparse
import datetime as dt
import statistics
import sys
import threading
import time

import psutil
import pythoncom
import pywintypes
import win32com.client as wc
import win32process
from win32com.client import dynamic

from smartass.core import config
from smartass.xl import fastcom

SIDES = {"std": ("StdSide", "StdSide2"), "fast": ("FastSide", "FastSide2")}
NEUTRAL = "Neutral"                          # the active sheet while the idioms run
XL_UP, XL_MANUAL = -4162, -4135
QUIT_WAIT_S = 30


# ------------------------------------------------------------------ the idioms
def _err(e: Exception) -> tuple:
    if isinstance(e, pythoncom.com_error):
        info = e.excepinfo or ()
        return ("com_error", e.hresult, info[2] if len(info) > 2 else None, info[5] if len(info) > 5 else None)
    return (type(e).__name__,)


def _try(fn):
    """fn()'s result, or the error it raised as a comparable tuple."""
    try:
        return fn()
    except Exception as e:   # noqa: BLE001 - the error is the result to compare
        return _err(e)


def ranges(app, wb, ws, ws2):
    return (ws.Range("A1:C3").Address, ws.Range(ws.Cells(1, 1), ws.Cells(3, 3)).Address, ws.Cells(2, 3).Address,
            ws.Rows(4).Address, ws.Columns(2).Address, ws.Rows("2:3").Address, (ws.Rows)("5:6").Address,
            (ws.Columns)(3).Address, ws.Range("B2").Offset(2, 3).Address, ws.Range("B2").Resize(2, 3).Address,
            ws.Range("B2").EntireRow.Address, ws.Range("B2").EntireColumn.Address, ws.Range("A1:C3")(2, 2).Address,
            ws.Range("A1:C3").GetAddress(False, False), ws.Range("A1:C3").Rows.Count, ws.Range("A1:C3").Columns.Count,
            ws.Range("B3").Row, ws.Range("B3").Column, ws.Range("A1:C3").Areas.Count, ws.Range("A1").Worksheet.Name,
            ws.Range("B2").Left, ws.Range("B2").Top, ws.Range("A1:B2")[0].Address,
            [c.Address for c in ws.Range("A1:B2")], len(ws.Range("A1:C3")))


def used_and_end(app, wb, ws, ws2):
    ws.Range("A1:C3").Value = ((1, 2, 3), (4, "five", 6.5), (None, True, "x"))
    return (ws.UsedRange.Address, ws.UsedRange.Row, ws.UsedRange.Column, ws.UsedRange.Rows.Count,
            ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row, ws.Range("C1").End(XL_UP).Row, ws.UsedRange.Value,
            ws.UsedRange.Formula2)


def values_scalar(app, wb, ws, ws2):
    r = ws.Range("A10")
    out = []
    for prop, v in (("Value", "text"), ("Value", 12.5), ("Value2", 3), ("Value", True), ("Value", None),
                    ("Formula", "=1+2"), ("Formula2", "=SUM(1,2)"), ("Formula2", "'=not a formula"),
                    ("Value", "'0042"), ("Value", dt.datetime(2026, 11, 19, 12, 30))):
        setattr(r, prop, v)
        out.append((r.Value, r.Value2, r.Formula, r.Formula2, r.Text, r.HasFormula))
    r.Clear()
    return out


def values_grid(app, wb, ws, ws2):
    r = ws.Range("A12:C13")
    r.Value = ((1, "a", None), (2.5, True, "=1+1"))
    a = (r.Value, r.Value2, r.Formula, r.Formula2)
    r.Formula2 = (("'=x", "=A12*2", 7), ("text", None, "=SUM(A12:B12)"))
    b = (r.Value, r.Formula2)
    r.Value2 = ((1.5,) * 3, (dt.datetime(2026, 1, 31),) * 3)
    c = (r.Value, r.Value2)
    r.Clear()
    return a, b, c


def number_formats(app, wb, ws, ws2):
    r = ws.Range("A16")
    out = []
    for prop, v in (("NumberFormat", "$#,##0.00"), ("Value", 12.34),       # a currency format: Value is a Decimal
                    ("NumberFormat", "@"), ("Value", "Sep 26"), ("NumberFormatLocal", "0,0"),
                    ("NumberFormatLocal", "#,##0.0"), ("NumberFormat", "#,##0;-#,##0;\"-\"")):
        out.append((prop, v, _try(lambda: setattr(r, prop, v)), r.Value, r.Value2, r.NumberFormat,
                    r.NumberFormatLocal, r.Text))
    r.ClearFormats()
    out.append(r.NumberFormat)
    r.Clear()
    return out


def fonts_fills_borders(app, wb, ws, ws2):
    f = ws.Range("B20").Font
    f.Name, f.Size, f.Bold, f.Color = "Arial", 11, True, 0x112233
    ws.Range("B21").Interior.Color = 0x00FFFF
    b = ws.Range("B22").Borders(9)
    b.LineStyle, b.Weight, b.Color = 1, 2, 0x445566
    out = (f.Name, f.Size, f.Bold, f.Color, ws.Range("B20").Font.Name, ws.Range("B21").Interior.Color,
           ws.Range("B21").Interior.Pattern, b.LineStyle, b.Weight, b.Color, ws.Range("B22").Borders(9).Color)
    ws.Range("B20:B22").Clear()
    return out


def widths_heights(app, wb, ws, ws2):
    ws.Columns(5).ColumnWidth = 14
    ws.Rows(5).RowHeight = 20
    ws.Range("F1").EntireColumn.ColumnWidth = 9.5
    ws.Range("F1").Value = "a fairly long text in F1"
    ws.Range("F1").Columns.AutoFit()
    out = (ws.Columns(5).ColumnWidth, ws.Rows(5).RowHeight, ws.Range("F1").ColumnWidth, ws.Columns("E:F").ColumnWidth)
    ws.Range("F1").Clear()
    ws.Columns("E:F").ColumnWidth = 8.43
    ws.Rows(5).RowHeight = 15
    return out


def merge(app, wb, ws, ws2):
    r = ws.Range("G20:H21")
    r.Merge()
    a = (r.MergeCells, ws.Range("G20").MergeArea.Address, ws.Range("H21").MergeCells)
    r.UnMerge()
    return a, r.MergeCells


def format_conditions(app, wb, ws, ws2):
    r = ws.Range("A1:C3")
    fc = r.FormatConditions.Add(1, 5, "=0")
    out = (fc.Type, fc.Formula1, fc.Operator, r.FormatConditions.Count, r.FormatConditions(1).Type)
    r.FormatConditions.Delete()
    return out, r.FormatConditions.Count


def sheets_add_move_delete(app, wb, ws, ws2):
    n = wb.Worksheets.Add(None, ws)
    n.Name = ws.Name + "_add"
    n.Tab.Color = 0x00FF00
    n.Visible = 0
    hidden = n.Visible
    n.Visible = -1
    n.Move(None, wb.Worksheets(wb.Worksheets.Count))
    out = (n.Name, hidden, n.Visible, n.Tab.Color, n.Index == wb.Worksheets.Count, wb.Worksheets(n.Name).Name)
    app.DisplayAlerts = False
    n.Delete()
    return out, _try(lambda: wb.Worksheets(ws.Name + "_add"))


def sheets_copy_list(app, wb, ws, ws2):
    before = wb.Worksheets.Count
    wb.Worksheets([ws.Name, ws2.Name]).Copy(None, wb.Worksheets(wb.Worksheets.Count))
    names = [str(wb.Worksheets(i).Name) for i in range(before + 1, wb.Worksheets.Count + 1)]
    app.DisplayAlerts = False
    wb.Worksheets(names).Delete()
    one = ws.Copy(None, ws)
    copied = wb.Worksheets(ws.Index + 1)
    cname = copied.Name
    copied.Delete()
    return names, one, cname, wb.Worksheets.Count == before


def collections_and_iteration(app, wb, ws, ws2):
    own = [str(s.Name) for s in wb.Worksheets if str(s.Name).startswith(ws.Name)]
    return (own, wb.Worksheets.Count, len(wb.Worksheets), wb.Worksheets(1).Name, wb.Worksheets[0].Name,
            wb.Worksheets.Item(1).Name, wb.Worksheets(ws.Name).Index == ws.Index, ws.Parent.Name, ws.Visible,
            app.Workbooks.Count, [str(b.Name) for b in app.Workbooks], ws.Shapes.Count, ws.ChartObjects().Count)


def tables(app, wb, ws, ws2):
    ws.Range("K1:M4").Value = (("k", "v", "w"), ("a", 1, 2), ("b", 3, 4), ("c", 5, 6))
    lo = ws.ListObjects.Add(1, ws.Range("K1:M4"), None, 1)
    lo.Name = ws.Name + "_tbl"
    lo.TableStyle = "TableStyleLight1"
    lo.ShowTableStyleRowStripes = False
    a = (lo.Name, lo.TableStyle, lo.Range.Address, lo.DataBodyRange.Address, lo.HeaderRowRange.Value,
         lo.HeaderRowRange.Row, lo.Range.Rows.Count, lo.Range.Value2, ws.ListObjects(lo.Name).Name,
         [str(t.Name) for t in ws.ListObjects], ws.ListObjects.Count, lo.ShowTableStyleRowStripes,
         lo.HeaderRowRange.Font.Bold)
    lo.Resize(ws.Range("K1:M5"))
    b = (lo.Range.Address, lo.DataBodyRange.Rows.Count)
    lo.Delete()
    return a, b, _try(lambda: ws.ListObjects(ws.Name + "_tbl")), _try(lambda: ws.ListObjects("no such table"))


def names(app, wb, ws, ws2):
    nm = wb.Names.Add(ws.Name + "_nm", "=" + ws.Name + "!$A$1:$B$2")
    out = (nm.Name, nm.RefersTo, nm.RefersToRange.Address, wb.Names(ws.Name + "_nm").Name,
           [str(n.Name) for n in wb.Names if str(n.Name).startswith(ws.Name)])
    nm.Delete()
    return out, [str(n.Name) for n in wb.Names if str(n.Name).startswith(ws.Name)]


def hyperlinks(app, wb, ws, ws2):
    h = ws.Hyperlinks.Add(ws.Range("N1"), "", f"'{ws2.Name}'!A1", "a tip", "go there")
    out = (h.SubAddress, h.TextToDisplay, h.ScreenTip, ws.Hyperlinks.Count, ws.Range("N1").Hyperlinks(1).SubAddress,
           [x.SubAddress for x in ws.Hyperlinks])
    ws.Hyperlinks.Delete()
    ws.Range("N1").Clear()
    return out, ws.Hyperlinks.Count


def autofilter(app, wb, ws, ws2):
    ws.Range("K1:M4").Value = (("k", "v", "w"), ("a", 1, 2), ("b", 3, 4), ("c", 5, 6))
    if ws.AutoFilterMode:
        ws.AutoFilterMode = False
    ws.Range("K1:M4").AutoFilter()
    a = (ws.AutoFilterMode, ws.AutoFilter.Range.Address)
    ws.AutoFilterMode = False
    ws.Range("K1:M4").Clear()
    return a, ws.AutoFilterMode, ws.AutoFilter


def page_setup(app, wb, ws, ws2):
    ps = ws.PageSetup
    ps.Orientation = 2
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = False
    ps.PrintArea = "$A$1:$D$40"
    ps.CenterFooter = "&P / &N"
    out = (ps.Orientation, ps.Zoom, ps.FitToPagesWide, ps.FitToPagesTall, ps.PrintArea, ps.CenterFooter,
           ws.PageSetup.LeftMargin)
    ps.Orientation = 1
    ps.PrintArea = ""
    ws.DisplayPageBreaks = False            # else every later row insert repaginates and the timings are Excel's
    return out


def page_breaks(app, wb, ws, ws2):
    ws.Range("A60").Value = "far down"
    ws.HPageBreaks.Add(ws.Range("A20"))
    a = (ws.HPageBreaks.Count, ws.HPageBreaks(1).Location.Address)
    ws.ResetAllPageBreaks()
    ws.Range("A60").Clear()
    ws.DisplayPageBreaks = False
    return a


def chart(app, wb, ws, ws2):
    ws.Range("H1:I4").Value = (("k", "v"), ("a", 1), ("b", 2), ("c", 3))
    shp = ws.Shapes.AddChart2(-1, 51, 300, 10, 240, 160)
    try:
        ch = shp.Chart
        while ch.SeriesCollection().Count:
            ch.SeriesCollection(1).Delete()
        s = ch.SeriesCollection().NewSeries()
        s.Values = ws.Range("I2:I4")
        s.XValues = ws.Range("H2:H4")
        s.Name = "Series x"
        s.HasDataLabels = True
        dl = s.DataLabels()
        dl.NumberFormatLinked = False
        dl.NumberFormat = "0.0"
        dl.Font.Color = 0xFFFFFF
        s.Points(2).HasDataLabel = False
        s.Points(3).DataLabel.Position = -4108                     # xlLabelPositionCenter
        ch.HasTitle = True
        ch.ChartTitle.Text = "Title"
        ch.ChartTitle.Font.Bold = True
        ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Arial"
        ch.Axes(2).HasMajorGridlines = False
        ch.Axes(2).TickLabelPosition = -4142
        ch.Axes(1).ReversePlotOrder = True
        ch.HasLegend = False
        g = ch.ChartGroups(1)
        g.GapWidth = 50
        s.Format.Fill.ForeColor.RGB = 0x123456
        s.Points(1).Format.Fill.ForeColor.RGB = 0x654321
        return (s.Formula, s.Name, s.Points(3).DataLabel.Text, ch.ChartTitle.Text, ch.Axes(2).HasMajorGridlines,
                g.GapWidth, s.Format.Fill.ForeColor.RGB, s.Points(1).Format.Fill.ForeColor.RGB, dl.NumberFormat,
                ws.ChartObjects().Count, shp.Name == ws.ChartObjects(1).Name, ch.SeriesCollection().Count,
                ws.ChartObjects(1).Chart.SeriesCollection(1).Name)
    finally:
        shp.Delete()
        ws.Range("H1:I4").Clear()


def chart_lines(app, wb, ws, ws2):
    ws.Range("H1:J4").Value = (("k", "a", "b"), ("x", 1, 4), ("y", 2, 5), ("z", 3, 6))
    shp = ws.Shapes.AddChart2(-1, 51, 300, 10, 240, 160)
    try:
        shp.Name = "check chart"
        ch = shp.Chart
        while ch.SeriesCollection().Count:
            ch.SeriesCollection(1).Delete()
        for col in ("I", "J"):
            s = ch.SeriesCollection().NewSeries()
            s.Values = ws.Range(f"{col}2:{col}4")
            s.XValues = ws.Range("H2:H4")
        s2 = ch.SeriesCollection(2)
        s2.ChartType = 4                                             # xlLine
        s2.AxisGroup = 2
        ln = s2.Format.Line
        ln.Visible = -1
        ln.ForeColor.RGB = 0x0000FF
        ln.Weight = 2.25
        ln.DashStyle = 4
        f = ch.SeriesCollection(1).Format.Fill
        f.Visible = -1
        f.Solid()
        f.ForeColor.RGB = 0x00FF00
        f.Patterned(28)
        f.BackColor.RGB = 0xFFFFFF
        lab = ch.SeriesCollection(1).Points(1)
        lab.HasDataLabel = True
        top = lab.DataLabel.Top
        lab.DataLabel.Top = top - 2
        val = ch.Axes(2)
        val.MaximumScale = 10
        val.MajorTickMark = -4142
        sec = ch.Axes(2, 2)
        sec.HasMajorGridlines = False
        sec.Format.Line.Visible = 0
        ch.HasLegend = True
        ch.Legend.Position = -4160                                   # xlLegendPositionTop
        g = ch.ChartGroups(1)
        g.Overlap = 100
        return (shp.Name, s2.ChartType, s2.AxisGroup, ln.Weight, ln.DashStyle, ln.ForeColor.RGB, f.BackColor.RGB,
                val.MaximumScale, sec.HasMajorGridlines, ch.Legend.Position, g.Overlap, ch.ChartGroups().Count,
                abs(lab.DataLabel.Top - (top - 2)) < 1, ws.ChartObjects("check chart").Chart.SeriesCollection().Count)
    finally:
        shp.Delete()
        ws.Range("H1:J4").Clear()


def chart_title_missing(app, wb, ws, ws2):
    """Office answers a property of a title that is not there with E_INVALIDARG: what pywin32 makes of that answer
    for one chart must not change another chart's title."""
    shp = ws.Shapes.AddChart2(-1, 51, 300, 200, 240, 160)
    try:
        ch = shp.Chart
        ch.HasTitle = False
        missing = _try(lambda: ch.ChartTitle.Text)
        again = _try(lambda: shp.Chart.ChartTitle.Text)
        other = ws.ChartObjects(shp.Name).Chart
        other.HasTitle = True
        other.ChartTitle.Text = "Now"
        return missing, again, shp.Chart.ChartTitle.Text, _try(lambda: shp.Chart.Axes(2, 2))
    finally:
        shp.Delete()


def application(app, wb, ws, ws2):
    old = (app.Calculation, app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts)
    app.Calculation = XL_MANUAL
    app.ScreenUpdating = False
    app.EnableEvents = False
    app.DisplayAlerts = False
    mid = (app.Calculation, app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts)
    app.Calculation, app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts = old
    app.StatusBar = "Claude: checking"
    bar = app.StatusBar
    app.StatusBar = False
    intl = app.International
    return (old, mid, intl[4], _try(lambda: app.International(5)), app.GetInternational(5), app.CentimetersToPoints(1),
            app.WorksheetFunction.CountA(ws.Range("A1:C3")), ws.Application.WorksheetFunction.Sum(ws.Range("A1:B2")),
            app.WorksheetFunction.CountA(ws.Cells), app.Name, app.Ready, bar, app.StatusBar)


def root_properties(app, wb, ws, ws2):
    """What host._start does on the root right after DispatchEx."""
    app.Visible = False
    app.IgnoreRemoteRequests = True
    app.DisplayAlerts = False
    app.AutomationSecurity = 1
    app.AskToUpdateLinks = False
    out = (isinstance(app.Hwnd, int), app.Visible, app.IgnoreRemoteRequests, app.DisplayAlerts,
           app.AutomationSecurity, app.AskToUpdateLinks)
    n = wb.Worksheets.Add()                                          # no arguments: before the active sheet
    n.Name = ws.Name + "_bare"
    added = (n.Name, n.Index == wb.ActiveSheet.Index)
    n.Delete()
    wb.Worksheets(NEUTRAL).Activate()
    return out, added


VBA = ('Public Function C_Echo(a, b)\n    C_Echo = a & "|" & b\nEnd Function\n'
       'Public Function C_Pair(a)\n    C_Pair = Array("OK", a * 2, "x")\nEnd Function\n')


def vba_bridge(app, wb, ws, ws2):
    """The bridge as host._load_bridge builds it and Target.run / HiddenExcel._run call it."""
    comp = wb.VBProject.VBComponents.Add(1)
    try:
        comp.Name = "chk" + ws.Name
        comp.CodeModule.AddFromString(VBA)
        mod = f"'{wb.Name}'!{comp.Name}"
        pair = app.Run(f"{mod}.C_Pair", 21)
        return (comp.Name, comp.CodeModule.CountOfLines, app.Run(f"{mod}.C_Echo", "a", 3), pair, pair[0],
                _try(lambda: app.Run(f"{mod}.NoSuchFunction")))
    finally:
        wb.VBProject.VBComponents.Remove(comp)


def union_address(app, wb, ws, ws2):
    sep = ws.Application.International[4]
    return ws.Range(sep.join(["A1", "C3", "E5:F6"])).Address, ws.Range(sep.join(["A1", "C3"])).Areas.Count


def workbooks_add_open_close(app, wb, ws, ws2):
    b = app.Workbooks.Add()
    a = (b.Worksheets.Count, b.Saved, str(b.Name).startswith("Book"), b.Worksheets(1).Range("A1").Value)
    b.Close(False)
    o = app.Workbooks.Open(str(OPEN_FILE), 0, True)
    try:
        c = (o.Name, o.ReadOnly, o.Worksheets(1).Name, o.Worksheets(1).Range("A1:B2").Value, o.LinkSources(1),
             app.Workbooks(o.Name).Name, o.Worksheets(1).ListObjects.Count)
    finally:
        o.Close(False)
    return a, c, _try(lambda: app.Workbooks("no such book.xlsx")), _try(lambda: wb.Worksheets("no such sheet"))


def activation_and_selection(app, wb, ws, ws2):
    wb.Activate()
    ws.Activate()
    app.Goto(ws.Range("B2"))
    a = (app.Selection.Address, app.ActiveSheet.Name, wb.ActiveSheet.Name, ws.Application.Selection.Worksheet.Name,
         app.ActiveWorkbook.Name == wb.Name)
    ws.Range("A1").Select()
    out = a, app.Selection.Address
    wb.Worksheets(NEUTRAL).Activate()       # work on the active sheet is slower in Excel: neither side keeps it
    return out


def special_and_errors(app, wb, ws, ws2):
    ws.Range("A1:C3").Value = ((1, 2, 3), (4, "five", 6.5), (None, True, "x"))
    ws.Range("D1").Formula = "=A1*2"
    return (ws.UsedRange.SpecialCells(2).Count, ws.UsedRange.SpecialCells(-4123).Count,
            _try(lambda: ws.Range("Z100").SpecialCells(-4123)), ws.Range("A1").DirectDependents.Address,
            _try(lambda: ws.Range("B2").DirectDependents), _try(lambda: ws.Range("not an address")),
            _try(lambda: ws.NoSuchProperty), _try(lambda: setattr(ws.Range("A1"), "NoSuchProperty", 1)))


def styles_and_xml(app, wb, ws, ws2):
    r = ws.Range("A30:B31")
    r.Value = (("a", 1), ("b", 2))
    r.Style = "Normal"
    xml = r.GetValue(11)
    r.Clear()
    r.SetValue(11, xml)
    out = (ws.Range("A30").Style.Name, xml, r.Value, ws.Range("A30").Style.Name)
    r.Clear()
    return out


def insert_delete_clear(app, wb, ws, ws2):
    ws.Range("A40:B41").Value = ((1, 2), (3, 4))
    ws.Rows("40:41").Insert()
    a = ws.Range("A42").Value
    ws.Rows("40:41").Delete()
    ws.Range(ws.Cells(1, 20), ws.Cells(1, 21)).EntireColumn.Insert()
    ws.Range(ws.Cells(1, 20), ws.Cells(1, 21)).EntireColumn.Delete()
    ws.Range("A40").ClearContents()
    b = (ws.Range("A40").Value, ws.Range("B40").Value)
    ws.Range("A40:B41").ClearFormats()
    ws.Range("A40:B41").Clear()
    ws.Calculate()
    app.Calculate()
    return a, b, ws.UsedRange.Address


def str_and_default_members(app, wb, ws, ws2):
    ws.Range("A1").Value = "first"
    return (str(ws.Range("A1")), ws.Range("A1")(), repr(ws.Range("A1")) == repr(ws.Range("A2")), bool(ws.Range("A1")),
            ws.Range("A1") == ws.Range("A1"), wb.Worksheets(ws.Name) == ws)


IDIOMS = {f.__name__: f for f in (ranges, used_and_end, values_scalar, values_grid, number_formats,
                                  fonts_fills_borders, widths_heights, merge, format_conditions, sheets_add_move_delete,
                                  sheets_copy_list, collections_and_iteration, tables, names, hyperlinks, autofilter,
                                  page_setup, page_breaks, chart, chart_lines, chart_title_missing, application, root_properties,
                                  vba_bridge, union_address,
                                  workbooks_add_open_close, activation_and_selection, special_and_errors,
                                  styles_and_xml, insert_delete_clear, str_and_default_members)}


# ------------------------------------------------------------------ comparing and timing
def norm(v):
    """A result as comparable data: types kept, COM objects by a projection, the side's sheet names replaced."""
    if isinstance(v, dynamic.CDispatch):
        return ("COM", _projection(v))
    if isinstance(v, (tuple, list)):
        return (type(v).__name__, tuple(norm(x) for x in v))
    if callable(v):
        return ("callable",)
    if isinstance(v, str):
        for names in SIDES.values():
            for n in sorted(names, key=len, reverse=True):
                v = v.replace(n, "SIDE")
        return ("str", v)
    return (type(v).__name__, v)


def _projection(ob):
    for name in ("Address", "Name", "Count"):
        try:
            x = getattr(ob, name)
        except (AttributeError, pythoncom.com_error):
            continue
        if not callable(x):
            return name, norm(x)
    return "?"


def _outcome(result) -> str:
    """'ran' when the idiom ran to its end, else the error that stopped it (the same on both sides to be equal)."""
    kind, value = result
    if kind == "tuple" and value and value[0] in (("str", "com_error"), ("str", "AttributeError"), ("str", "TypeError")):
        return "stopped by " + " ".join(str(v[1]) for v in value[:2])
    return "ran"


def run(fn, root, book: str, side: str):
    wb = root.Workbooks(book)
    s1, s2 = SIDES[side]
    try:
        return norm(fn(root, wb, wb.Worksheets(s1), wb.Worksheets(s2)))
    except Exception as e:   # noqa: BLE001 - the error is the result to compare
        return norm(_err(e))


def medians_ms(fn, std, fast, book: str, runs: int) -> tuple[float, float]:
    """Median milliseconds of fn through each wrapper, the runs interleaved so that a slower Excel weighs on both."""
    ts: dict[str, list[float]] = {"std": [], "fast": []}
    for _ in range(runs):
        for side, root in (("std", std), ("fast", fast)):
            t0 = time.perf_counter()
            run(fn, root, book, side)
            ts[side].append((time.perf_counter() - t0) * 1000)
    return statistics.median(ts["std"]), statistics.median(ts["fast"])


# ------------------------------------------------------------------ our own hidden Excel
def _excel_pids() -> set[int]:
    return {p.pid for p in psutil.process_iter(["name"]) if (p.info["name"] or "").lower() == "excel.exe"}


def start_excel():
    before = _excel_pids()
    app = wc.DispatchEx("Excel.Application")
    _, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
    if pid in before:
        raise SystemExit(f"Refusing: DispatchEx attached to an Excel that already existed (pid {pid}).")
    ctime = psutil.Process(pid).create_time()
    app.Visible = False
    app.IgnoreRemoteRequests = True            # files the user double-clicks open in their Excel, not in this one
    app.DisplayAlerts = False
    app.AutomationSecurity = 1                  # as host._start: the bridge idiom runs VBA (the opened file has none)
    app.AskToUpdateLinks = False
    return app, pid, ctime


def kill_if_ours(pid: int, ctime: float) -> bool:
    try:
        p = psutil.Process(pid)
        if abs(p.create_time() - ctime) < 1.0 and "-embedding" in " ".join(p.cmdline()).lower():
            p.kill()
            return True
    except psutil.Error:
        pass
    return False


def write_open_file():
    import openpyxl
    folder = config.HOME / "tmp"
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f"check_fastcom_{time.time_ns()}.xlsx"
    book = openpyxl.Workbook()
    book.active.title = "Opened"
    book.active.append(["name", 1.5])
    book.active.append(["other", 2])
    book.save(path)
    return path


OPEN_FILE = None


def main(argv=None) -> int:
    global OPEN_FILE
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--runs", type=int, default=7, help="timed runs per idiom and wrapper (median)")
    ap.add_argument("--only", help="run only the idioms whose name contains this")
    args = ap.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")
    idioms = {k: f for k, f in IDIOMS.items() if not args.only or args.only in k}
    OPEN_FILE = write_open_file()
    app, pid, ctime = start_excel()
    print(f"check_fastcom {dt.datetime.now():%Y-%m-%d %H:%M}: hidden Excel pid {pid}, {len(idioms)} idioms, "
          f"fastcom {'on' if config.FASTCOM else 'OFF (SV_FASTCOM=0)'}", flush=True)
    unequal: dict[str, list] = {}
    outcome: dict[str, str] = {}
    wb = fast = first = None
    try:
        wb = app.Workbooks.Add()
        book = str(wb.Name)
        first = wb.Worksheets(1)
        first.Name = SIDES["std"][0]
        for name in (SIDES["std"][1], *SIDES["fast"], NEUTRAL):
            wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count)).Name = name
        wb.Worksheets(NEUTRAL).Activate()
        order = list(idioms)
        for label, names in (("forward", order), ("reversed", order[::-1])):
            for temp in ("cold", "warm"):
                if temp == "cold":
                    fastcom._IFACE.clear()
                fast = fastcom.wrap(app)
                for k in names:
                    a, b = run(idioms[k], app, book, "std"), run(idioms[k], fast, book, "fast")
                    outcome.setdefault(k, _outcome(a))
                    if a != b:
                        unequal.setdefault(k, []).append((f"{label} {temp}", a, b))
        fast = fastcom.wrap(app)
        tot_s = tot_f = 0.0
        for k, fn in idioms.items():
            s, f = medians_ms(fn, app, fast, book, args.runs)
            tot_s, tot_f = tot_s + s, tot_f + f
            print(f"{k:<28} equal {'yes' if k not in unequal else 'NO '}  standard {s:8.1f} ms  fast {f:8.1f} ms  "
                  f"x{s / f if f else 0:4.1f}  ({outcome[k]})", flush=True)
        print(f"{'total':<28} {'':9}  standard {tot_s:8.1f} ms  fast {tot_f:8.1f} ms  x{tot_s / tot_f:4.1f}")
        print(f"interfaces learnt: {len(fastcom._IFACE)}")
        for k, rounds in unequal.items():
            for when, a, b in rounds:
                print(f"\nNOT EQUAL {k} ({when}):\n  standard {str(a)[:1500]}\n  fast     {str(b)[:1500]}")
    finally:
        _quit(app, wb, pid, ctime)
        app = wb = fast = first = None              # Excel ends only when every reference to it is gone
        fastcom._IFACE.clear()
        _wait_gone(pid, ctime)
        try:
            OPEN_FILE.unlink()
        except OSError:
            pass
    if unequal:
        print(f"\n{len(unequal)} idiom(s) not equal")
        return 1
    print("\nall idioms equal")
    return 0


def _quit(app, wb, pid: int, ctime: float) -> None:
    """Close the scratch workbook unsaved and quit our Excel; kill only that process if Quit hangs."""
    watchdog = threading.Timer(QUIT_WAIT_S, lambda: kill_if_ours(pid, ctime))
    watchdog.start()
    try:
        if wb is not None:
            wb.Close(False)
        app.Quit()
    except pywintypes.com_error:
        pass
    finally:
        watchdog.cancel()


def _wait_gone(pid: int, ctime: float) -> None:
    import gc
    gc.collect()
    t0 = time.monotonic()
    while time.monotonic() - t0 < QUIT_WAIT_S and psutil.pid_exists(pid):
        time.sleep(0.5)
    if psutil.pid_exists(pid) and kill_if_ours(pid, ctime):
        print(f"killed our hidden Excel pid {pid} after Quit")


if __name__ == "__main__":
    sys.exit(main())
