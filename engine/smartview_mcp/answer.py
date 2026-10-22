"""Build a 'Q<n> Svar' answer sheet in the user's workbook from a declarative spec.

Every number is a live formula: pulled cells are XLOOKUPs into the Smart View data sheets (by the labels Smart View
shows), derived cells are readable formulas over other answer cells (SUM over ranges, skill tn-formelstil). Refreshing
a data sheet updates the answer.

spec = {
  "title": "Technology - P&L YTD Sep 26 vs Sep 25",
  "subtitle": "NOK · FinMax · Actual · un.SY.2TEC · Accumulated",
  "key_points": ["EBITDA ... ", ...],
  "columns": [{"header": "YTD Sep 26", "kind": "heltall"}, {"header": "YTD Sep 25"}, {"header": "Variance"},
              {"header": "Variance %", "kind": "prosent1"}],
  "sections": [{"title": "Profit and loss", "lines": [
      {"key": "rev", "label": "C30 REVENUES", "kind": "line" | "sum" | "grand", "num": "tall1" (optional row override),
       "cells": [ {"pull": "Q1 FinMax", "row": ["ac.C30"], "col": ["Sep 26"]},
                  {"pull": "Q1 FinMax", "row": ["ac.C30"], "col": ["Sep 25"]},
                  {"var": [0, 1]},            # = col0 - col1 on this line
                  {"varpct": [0, 1]},         # = (col0 - col1) / |col1|
                  {"expr": "rev+cogs"},       # line keys, same column
                  {"sum": ["rev", "cogs"]}, {"value": 123}, {"text": "n/a"}, null,
                  {"drill": "d", "where": {"JE_SOURCE": "Payables"}, "agg": "sum" | "count" | "debit" | "credit"}
                                              # sv_drill lines table Linjer_Q<n>[_alias], exact match per column
                ]}]}],
  "checks": [{"label": "C30 + C40S = C48", "expr": "rev+cogs-gm", "columns": [0, 1]},
             {"label": "Period is closed", "static": "OK" | "DIFF", "detail": "Sep 26 <= Current Reporting Month Sep 26"}],
  "notes": ["..."], "source": "Source: Oracle EPM Cloud (FinMax) via Smart View, retrieved 2026-10-13 22:50",
  "chart": {"type": "waterfall" | "column" | "bar" | "line", "lines": ["rev", ...], "column": 0, "title": "..."}
}
"""
from __future__ import annotations

import re

from openpyxl.utils import get_column_letter as L

from . import formulas as fx
from .drill_xl import drill_formula, missing_columns, table_name
from .drill_xl import table_columns as drill_table_columns
from .host import HostError, InputError
from .userexcel import Pull, Target, add_sheet, com, lookup_formula

FIRST_COL = 3            # column C: first value column; B = labels, A = margin
NUM_KINDS = {"heltall", "tall1", "prosent1", "prosent", "faktor"}
CHART_TYPES = {"column": 51, "bar": 57, "line": 4, "waterfall": 119}   # xlColumnClustered, xlBarClustered, xlLine, xlWaterfall


def answer_sheet(t: Target, q: int):
    """Create 'Q<n> Svar' before the question's data sheets, or clear it for a rebuild (follow-up)."""
    name = f"Q{q} Svar"
    for s in list(t.wb.Worksheets):
        if s.Name == name:
            com(lambda: s.Cells.Clear())
            for co in list(com(lambda: s.ChartObjects())):
                com(co.Delete)
            return s
    group = [s for s in t.wb.Worksheets if s.Name.startswith(f"Q{q} ")]
    if group:
        ws = com(lambda: t.wb.Worksheets.Add(group[0]))     # positional Before (see userexcel.add_sheet)
        t.sheets_added.append(name)
        com(lambda: setattr(ws, "Name", name))
    else:
        ws = add_sheet(t, name)
    com(lambda: setattr(ws.Tab, "Color", 0x64381F))     # TN navy (BGR of #1F3864)
    return ws


def _expr_to_formula(expr: str, col_letter: str, rows_by_key: dict) -> str:
    def rep(m):
        k = m.group(0)
        if k not in rows_by_key:
            raise InputError(f"Unknown line key '{k}' in expression '{expr}'. Known: {sorted(rows_by_key)}")
        return f"{col_letter}{rows_by_key[k]}"
    return "=" + re.sub(r"[A-Za-z_][A-Za-z0-9_]*", rep, expr)


def build(t: Target, q: int, spec: dict, pulls: dict[str, Pull]) -> dict:
    from .userexcel import fast_excel
    with fast_excel(t.app):
        res = _build(t, q, spec, pulls)
    from .userexcel import wait_calculated
    from .timing import stage
    with stage("calculate + read back"):
        ws = t.wb.Worksheets(res.pop("_ws_name"))
        com(lambda: ws.Calculate())
        wait_calculated(t.app)
        grid = com(lambda: ws.Range(ws.Cells(1, 1), ws.Cells(max(res["_r"], 2), FIRST_COL + res["_ncol"])).Value)
    cell = lambda row, col: grid[row - 1][col - 1] if row - 1 < len(grid) and col - 1 < len(grid[row - 1]) else None
    values = {k: [cell(row, FIRST_COL + j) for j in range(res["_ncol"])] for k, row in res["_rows_by_key"].items()}
    errors = [k for k, vs in values.items() if any(isinstance(v, int) and v < -2146826000 for v in vs)]
    status = {}
    from openpyxl.utils import column_index_from_string as CI
    for m in res["_rows_meta"]:
        if m["kind"] == "check":
            status[cell(m["row"], 2)] = cell(m["row"], CI(m["status_col"]))
    with stage("formula check (tn-formelstil)"):
        formula_style = fx.check(fx.sheet_formulas(ws, com))
    return {"sheet": ws.Name, "values": values, "checks": status, "warnings": res["warnings"], "lint": res["lint"],
            "formula_errors": errors, "formula_style": formula_style}


def _build(t: Target, q: int, spec: dict, pulls: dict) -> dict:
    """Write, style and chart the answer sheet (inside fast_excel: no redraw, manual calculation)."""
    from . import styling
    from .timing import stage
    with stage("prepare sheet"):
        ws = answer_sheet(t, q)
        com(ws.Activate)
    with stage("write cells"):
        out = _write(t, q, spec, pulls, ws)
    rows_meta, plan, rows_by_key, checks, cols, ncol, last_col, header_row, r, warnings = out
    longest = max([len(str(ln.get("label", ""))) for _, ln, _ in plan if isinstance(ln, dict) and "cells" in ln] +
                  [len(f"{c['label']} — {cols[0].get('header', '')}") for c in checks if "static" not in c] + [40])
    roles = {"last_col": last_col, "label_cols": ["B"], "freeze_rows": header_row,
             "col_widths": {"A": 2.5, "B": min(80, max(46, longest + 4)), f"C:{last_col}": 17}, "rows": rows_meta}
    t.status(f"formaterer 'Q{q} Svar' i TN-rapportstil ...")
    with stage("style (tn-rapportstil)"):
        with stage("ensure styles (1st)"):
            styling.ensure_styles(t.wb)
        lint = styling.apply_report(ws, roles) or []
    if spec.get("chart"):
        with stage("chart"):
            try:
                _chart(ws, spec["chart"], rows_by_key, header_row, ncol)
            except Exception as e:   # noqa: BLE001 - a chart must never break the answer
                warnings.append(f"chart not created: {e}")
    return {"_ws_name": ws.Name, "_r": r, "_ncol": ncol, "_rows_by_key": rows_by_key, "_rows_meta": rows_meta,
            "warnings": warnings, "lint": lint}


def _write(t: Target, q: int, spec: dict, pulls: dict, ws):
    """All answer cells (title, key points, lines, checks, notes); returns the layout for styling."""
    t.status(f"bygger svar-arket 'Q{q} Svar' ...")
    cols = spec.get("columns") or []
    if not cols:
        raise InputError("answer spec needs 'columns'")
    ncol = len(cols)
    last_col = L(FIRST_COL + ncol - 1)
    rows_meta: list[dict] = []
    r = 1

    def put(row, col, value):
        # text is stored as text: Excel would otherwise turn 'Jan 26' into a date, '1/2' into a fraction ...
        if isinstance(value, str) and not value.startswith("="):
            value = "'" + value
        com(lambda: setattr(ws.Cells(row, col), "Value", value))

    put(r, 2, spec.get("title", f"Q{q}")); rows_meta.append({"row": r, "kind": "title"}); r += 1
    put(r, 2, spec.get("subtitle", "NOK")); rows_meta.append({"row": r, "kind": "subtitle"}); r += 2

    if spec.get("key_points"):
        put(r, 2, "Key points"); rows_meta.append({"row": r, "kind": "section"}); r += 1
        for kp in spec["key_points"]:
            put(r, 2, f"•  {kp}"); rows_meta.append({"row": r, "kind": "text"}); r += 1
        r += 1

    header_row = r
    for j, c in enumerate(cols):
        put(r, FIRST_COL + j, c.get("header", ""))
    rows_meta.append({"row": r, "kind": "header", "num_cols": [L(FIRST_COL + j) for j in range(ncol)]})
    r += 1

    # first pass: assign rows to line keys (expressions may reference lines further down)
    plan, rows_by_key, rr = [], {}, r
    for sec in spec.get("sections", []):
        plan.append(("section", sec, rr)); rr += 1
        for ln in sec.get("lines", []):
            if ln.get("key"):
                rows_by_key[ln["key"]] = rr
            plan.append(("line", ln, rr)); rr += 1
        rr += 1
    warnings = []
    table_cols: dict[str, list[str] | None] = {}          # drill tables: name -> header (None = not found)
    for kind, item, row in plan:
        if kind == "section":
            put(row, 2, item.get("title", "")); rows_meta.append({"row": row, "kind": "section"})
            continue
        ln = item
        put(row, 2, ln.get("label", ln.get("key", "")))
        num = {}
        cells = ln.get("cells", [])
        for j in range(ncol):
            c = cells[j] if j < len(cells) else None
            col = L(FIRST_COL + j)
            kind_num = ln.get("num") or cols[j].get("kind") or "heltall"
            if kind_num not in NUM_KINDS:
                kind_num = "heltall"
            if isinstance(c, dict) and "varpct" in c:
                kind_num = cols[j].get("kind") if cols[j].get("kind", "").startswith("prosent") else "prosent1"
            num[col] = kind_num
            if c is None:
                continue
            try:
                if "drill" in c:                                   # drill lines table: {"drill": alias, "where", "agg"}
                    tn_ = table_name(q, c["drill"])
                    if tn_ not in table_cols:
                        try:
                            table_cols[tn_] = drill_table_columns(t.wb, tn_)
                        except Exception:   # noqa: BLE001 - unknown table: write the formula, no column check
                            table_cols[tn_] = None
                    miss = missing_columns(table_cols[tn_], c.get("where") or {})
                    if miss:                                       # e.g. a detail column left out above 50 000 lines
                        raise InputError(f"the lines table {tn_} has no column {', '.join(miss)} (cell left empty)")
                    f = drill_formula(tn_, c.get("where") or {}, c.get("agg", "sum"))
                elif "pull" in c:
                    p = pulls.get(c["pull"])
                    if p is None:
                        raise InputError(f"Unknown pull sheet '{c['pull']}'. Pulls in this answer: {sorted(pulls)}")
                    f = lookup_formula(p, list(c["row"]), list(c["col"]))
                elif "var" in c:
                    a, b = (L(FIRST_COL + k) for k in c["var"])
                    f = fx.variance(f"{a}{row}", f"{b}{row}")
                elif "varpct" in c:
                    a, b = (L(FIRST_COL + k) for k in c["varpct"])
                    f = fx.variance_pct(f"{a}{row}", f"{b}{row}")
                elif "expr" in c:
                    keys = fx.expr_keys_if_plain_sum(c["expr"])        # 'a+b+c' -> SUM, never a + chain
                    f = (fx.sum_rows(col, [rows_by_key[k] for k in keys]) if keys
                         else _expr_to_formula(c["expr"], col, rows_by_key))
                elif "sum" in c:
                    f = fx.sum_rows(col, [rows_by_key[k] for k in c["sum"]])
                elif "ref" in c:                                   # another answer cell: {"ref": {"key":.., "col":..}}
                    f = f"={L(FIRST_COL + int(c['ref'].get('col', j)))}{rows_by_key[c['ref']['key']]}"
                elif "rowsum" in c:                                # sum of other columns on this line
                    f = fx.sum_cols(row, [FIRST_COL + k for k in c["rowsum"]])
                elif "value" in c:
                    f = c["value"]
                elif "text" in c:
                    f = "'" + str(c["text"])
                else:
                    continue
                cell = ws.Cells(row, FIRST_COL + j)
                if isinstance(f, str) and f.startswith("="):
                    com(lambda: setattr(cell, "Formula2", f))
                else:
                    com(lambda: setattr(cell, "Value", f))
            except Exception as e:   # noqa: BLE001 - one bad cell spec must not abort the whole answer
                if isinstance(e, KeyError):
                    e = InputError(f"unknown line key {e} (known: {sorted(rows_by_key)})")
                warnings.append(f"{ln.get('label')}: {e}")
        rows_meta.append({"row": row, "kind": {"sum": "sum", "grand": "grand"}.get(ln.get("kind"), "line"), "num": num})
    r = rr

    checks = spec.get("checks") or []
    if checks:
        put(r, 2, "Checks"); rows_meta.append({"row": r, "kind": "section"}); r += 1
        for ck in checks:
            if "static" in ck:
                put(r, 2, ck["label"] + (f" ({ck['detail']})" if ck.get("detail") else ""))
                put(r, FIRST_COL + 1, ck["static"])
                rows_meta.append({"row": r, "kind": "check", "num": {}, "status_col": L(FIRST_COL + 1)}); r += 1
                continue
            for j in ck.get("columns", [0]):
                col = L(FIRST_COL + j)
                put(r, 2, f"{ck['label']} — {cols[j].get('header', '')}")
                try:
                    f = _expr_to_formula(ck["expr"], col, rows_by_key)
                    com(lambda: setattr(ws.Cells(r, FIRST_COL), "Formula2", f))
                    com(lambda: setattr(ws.Cells(r, FIRST_COL + 1), "Formula2",
                                        f'=IF(ISNUMBER({L(FIRST_COL)}{r}),IF(ABS({L(FIRST_COL)}{r})<1,"OK","DIFF"),"DIFF")'))
                except (HostError, InputError) as e:
                    warnings.append(f"check {ck['label']}: {e}")
                rows_meta.append({"row": r, "kind": "check", "num": {L(FIRST_COL): "heltall"},
                                  "status_col": L(FIRST_COL + 1)})
                r += 1
        r += 1

    for note in spec.get("notes") or []:
        put(r, 2, note); rows_meta.append({"row": r, "kind": "note"}); r += 1
    if spec.get("source"):
        put(r, 2, spec["source"]); rows_meta.append({"row": r, "kind": "note"}); r += 1

    used = {m["row"] for m in rows_meta}
    for blank in range(1, r):
        if blank not in used:
            rows_meta.append({"row": blank, "kind": "blank"})
    rows_meta.sort(key=lambda m: m["row"])
    return rows_meta, plan, rows_by_key, checks, cols, ncol, last_col, header_row, r, warnings


def _chart(ws, ch: dict, rows_by_key: dict, header_row: int, ncol: int):
    """Either one column across several lines (ch['lines'] + ch['column']; e.g. a waterfall bridge) or one line
    across several columns (ch['row'] + ch['columns']; e.g. a monthly trend)."""
    xl_type = CHART_TYPES.get(ch.get("type", "column"), 51)
    if not ch.get("row") and len(ch.get("columns") or []) > 1:     # several columns across lines: one series each
        return _multi_chart(ws, ch, rows_by_key, header_row, ncol, xl_type)
    if ch.get("row"):
        if ch["row"] not in rows_by_key:
            return
        r = rows_by_key[ch["row"]]
        cs = ch.get("columns") or list(range(ncol))
        if cs == list(range(cs[0], cs[0] + len(cs))):
            a, b = L(FIRST_COL + cs[0]), L(FIRST_COL + cs[-1])
            values, xvalues = f"{a}{r}:{b}{r}", f"{a}{header_row}:{b}{header_row}"
        else:
            values = ",".join(f"{L(FIRST_COL + k)}{r}" for k in cs)
            xvalues = ",".join(f"{L(FIRST_COL + k)}{header_row}" for k in cs)
    else:
        keys = [k for k in ch.get("lines", []) if k in rows_by_key]
        if not keys:
            return
        col = L(FIRST_COL + int(ch.get("column", 0)))
        rws = [rows_by_key[k] for k in keys]
        if rws == list(range(rws[0], rws[0] + len(rws))):          # contiguous: plain ranges (waterfall needs it)
            values, xvalues = f"{col}{rws[0]}:{col}{rws[-1]}", f"B{rws[0]}:B{rws[-1]}"
        else:
            values = ",".join(f"{col}{r}" for r in rws)
            xvalues = ",".join(f"B{r}" for r in rws)
    left = ws.Cells(header_row, FIRST_COL + ncol + 1).Left
    top = ws.Cells(header_row, 2).Top
    try:
        shape = ws.Shapes.AddChart2(-1, xl_type, left, top, 480, 300)
        if xl_type == 119:
            shape.Chart.SetSourceData(ws.Range(f"{xvalues},{values}") if ":" in values else ws.Range(values))
            chart = shape.Chart
            chart.HasTitle = True
            chart.ChartTitle.Text = ch.get("title", "")
            chart.HasLegend = False
            n = chart.FullSeriesCollection(1).Points().Count
            for idx in ch.get("totals", [1, n]):        # start and end bars are totals, not deltas
                try:
                    chart.FullSeriesCollection(1).Points(idx).IsTotal = True
                except Exception:   # noqa: BLE001
                    pass
            _millions_axis(chart)
            return
    except Exception:   # noqa: BLE001 - waterfall not available for this range: plain column chart instead
        try:
            shape.Delete()
        except Exception:   # noqa: BLE001
            pass
        shape = ws.Shapes.AddChart2(-1, 51, left, top, 480, 300)
    chart = shape.Chart
    while chart.SeriesCollection().Count:
        chart.SeriesCollection(1).Delete()
    s = chart.SeriesCollection().NewSeries()
    s.Values = ws.Range(values)
    s.XValues = ws.Range(xvalues)
    chart.HasTitle = True
    chart.ChartTitle.Text = ch.get("title", "")
    chart.HasLegend = False
    try:
        chart.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Arial"
    except Exception:   # noqa: BLE001
        pass
    _millions_axis(chart)


def _multi_chart(ws, ch: dict, rows_by_key: dict, header_row: int, ncol: int, xl_type: int):
    """ch['lines'] x ch['columns']: one series per answer column (named by its header), e.g. this year vs last year
    per month. Waterfall is single-series only and falls back to clustered columns."""
    keys = [k for k in ch.get("lines", []) if k in rows_by_key]
    if not keys:
        return
    rws = [rows_by_key[k] for k in keys]
    contiguous = rws == list(range(rws[0], rws[0] + len(rws)))
    sheet_ref = "'" + ws.Name.replace("'", "''") + "'!"

    def ref(col: str):
        """Series source: a plain range, or (for lines in several sections) a union reference string - Excel rejects
        a multi-area Range object as Values/XValues (0x800A03EC) but accepts '=(Sheet!$E$5,Sheet!$E$9)'."""
        if contiguous:
            return ws.Range(f"{col}{rws[0]}:{col}{rws[-1]}")
        return "=(" + ",".join(f"{sheet_ref}${col}${r}" for r in rws) + ")"
    xvalues = ref("B")
    left = ws.Cells(header_row, FIRST_COL + ncol + 1).Left
    top = ws.Cells(header_row, 2).Top
    shape = ws.Shapes.AddChart2(-1, 51 if xl_type == 119 else xl_type, left, top, 520, 300)
    chart = shape.Chart
    while chart.SeriesCollection().Count:
        chart.SeriesCollection(1).Delete()
    for k in ch["columns"]:
        col = L(FIRST_COL + int(k))
        s = chart.SeriesCollection().NewSeries()
        s.Name = str(ws.Cells(header_row, FIRST_COL + int(k)).Value or "")
        s.Values = ref(col)
        s.XValues = xvalues
    chart.HasTitle = True
    chart.ChartTitle.Text = ch.get("title", "")
    chart.HasLegend = True
    try:
        chart.Legend.Position = -4107        # xlLegendPositionBottom
        chart.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Arial"
    except Exception:   # noqa: BLE001
        pass
    _millions_axis(chart)


def _millions_axis(chart):
    """Value axis in MNOK ('250') instead of raw NOK ('250000000,0'). Uses the axis display unit, because scaling
    format codes ('#,##0,,') are read in the user's locale on charts and come out garbled in nb-NO."""
    try:
        ax = chart.Axes(2)
        ax.DisplayUnit = -6                  # xlMillions (-5 is xlHundredThousands)
        ax.HasDisplayUnitLabel = True
        ax.DisplayUnitLabel.Text = "MNOK"
        ax.TickLabels.NumberFormat = "0"
    except Exception:   # noqa: BLE001
        pass
