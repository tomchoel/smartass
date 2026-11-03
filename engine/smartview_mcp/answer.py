"""Build a 'Q<n> Answer' answer sheet in the user's workbook from a declarative spec.

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
                  {"share": {"col": 0, "of": "tot"}},   # = col0 / col0 of line 'tot' (a share of the total)
                  {"expr": "rev+cogs"},       # line keys, same column
                  {"sum": ["rev", "cogs"]}, {"avg": ["f0", "f1"]}, {"value": 123}, {"text": "n/a"}, null,
                  {"drill": "d", "where": {"JE_SOURCE": "Payables"}, "agg": "sum" | "count" | "debit" | "credit"}
                                              # sv_drill lines table Lines_Q<n>[_alias], exact match per column
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
from .answer_write import FIRST_COL, _chart, _write
from .excel.jobs import com
from .excel.sheets import add_sheet, Pull
from .excel.target import Target


def answer_sheet(t: Target, q: int, suffix: str = "Answer", owned: dict | None = None):
    """Create 'Q<n> <suffix>' before the question's data sheets, or clear it for a rebuild (follow-up): only the
    tool-owned block `owned` ({"rows", "cols"}) and smartass charts, or everything when no block is stored."""
    from .core import owned as owned_mod
    name = f"Q{q} {suffix}"
    for s in list(t.wb.Worksheets):
        if s.Name == name:
            owned_mod.clear(s, owned)
            return s
    group = [s for s in t.wb.Worksheets if s.Name.startswith(f"Q{q} ")]
    if group:
        ws = com(lambda: t.wb.Worksheets.Add(group[0]))     # positional Before (see sheets.add_sheet)
        t.sheets_added.append(name)
        com(lambda: setattr(ws, "Name", name))
    else:
        ws = add_sheet(t, name)
    com(lambda: setattr(ws.Tab, "Color", 0x64381F))     # TN navy (BGR of #1F3864)
    return ws


def build(t: Target, q: int, spec: dict, pulls: dict[str, Pull], suffix: str = "Answer",
          owned: dict | None = None, reserve: int = 0) -> dict:
    from .excel.jobs import fast_excel
    with fast_excel(t.app):
        res = _build(t, q, spec, pulls, suffix, owned, reserve)
    from .excel.sheets import wait_calculated
    from .timing import stage
    with stage("calculate + read back"):
        ws = t.wb.Worksheets(res.pop("_ws_name"))
        com(lambda: ws.Calculate())
        wait_calculated(t.app)
        grid = com(lambda: ws.Range(ws.Cells(1, 1), ws.Cells(max(res["_r"], 2), FIRST_COL + res["_ncol"])).Value)
    cell = lambda row, col: grid[row - 1][col - 1] if row - 1 < len(grid) and col - 1 < len(grid[row - 1]) else None
    values = {k: [cell(row, FIRST_COL + j) for j in range(res["_ncol"])] for k, row in res["_rows_by_key"].items()}
    errors = [k for k, vs in values.items() if any(error_code(v) for v in vs)]
    with stage("scan error cells"):
        error_warns = error_warnings(ws.Name, grid, res["_trace"], _pull_grids(t, pulls))
    status = {}
    from openpyxl.utils import column_index_from_string as CI
    for m in res["_rows_meta"]:
        if m["kind"] == "check":
            status[cell(m["row"], 2)] = cell(m["row"], CI(m["status_col"]))
    with stage("formula check (tn-formelstil)"):
        formula_style = fx.check(fx.sheet_formulas(ws, com))
    return {"sheet": ws.Name, "values": values, "checks": status, "warnings": res["warnings"], "lint": res["lint"],
            "formula_errors": errors, "error_warnings": error_warns, "formula_style": formula_style,
            "extent": res["_extent"]}


# ------------------------------------------------------------------ error cells (answer + its data sheets)
XL_ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
             -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826238: "#SPILL!",
             -2146826229: "#CALC!"}             # CVErr codes as COM returns them (ints)
ERROR_TEXTS = {"#NoAccess", "#No Access", "#Invalid", "#VALUE!", "#REF!", "#N/A", "#DIV/0!", "#NAME?"}
_PULL_CAUSE = {"#NoAccess": "no read access or member not valid in this cube",
               "#No Access": "no read access or member not valid in this cube",
               "#Invalid": "member combination not valid in this cube"}


def error_code(v) -> str | None:
    """'#VALUE!' etc. for an Excel error (COM int) or a Smart View / Excel error label; None for anything else."""
    if isinstance(v, int) and not isinstance(v, bool) and v in XL_ERRORS:
        return XL_ERRORS[v]
    if isinstance(v, str) and v.strip() in ERROR_TEXTS:
        return v.strip()
    return None


def _addr(sheet: str, row: int, col: int) -> str:
    return "'" + sheet.replace("'", "''") + f"'!{L(col)}{row}"


def _error_cells(grid) -> list[tuple[int, int, str]]:
    out = []
    for i, r in enumerate(grid or ()):
        for j, v in enumerate(r or ()):
            code = error_code(v)
            if code:
                out.append((i + 1, j + 1, code))
    return out


def cell_trace(plan: list, rows_by_key: dict, pulls: dict, ncol: int) -> dict:
    """(row, col) of every answer line cell -> {"key", "pull": "'sheet'!B3"} for pulled cells, or {"key", "deps":
    [(row, col), ...]} for cells computed from other answer cells (to say what an error depends on)."""
    out = {}
    for kind, ln, row in plan:
        if kind != "line" or not ln.get("key"):
            continue
        cells = ln.get("cells", [])
        for j in range(min(ncol, len(cells))):
            c, col = cells[j], FIRST_COL + j
            if not isinstance(c, dict):
                continue
            ent: dict = {"key": ln["key"]}
            try:
                if "pull" in c and c["pull"] in pulls:
                    p = pulls[c["pull"]]
                    i = p.layout.row_tuples.index(tuple(c["row"]))
                    k = p.layout.col_tuples.index(tuple(c["col"]))
                    ent["pull"] = _addr(p.sheet, 2 + p.layout.nc + i, p.layout.nr + 1 + k)
                elif "var" in c or "varpct" in c:
                    ent["deps"] = [(row, FIRST_COL + k) for k in (c.get("var") or c.get("varpct"))]
                elif "expr" in c:
                    keys = re.findall(r"[A-Za-z_][A-Za-z0-9_]*", c["expr"])
                    ent["deps"] = [(rows_by_key[k], col) for k in keys if k in rows_by_key]
                elif "sum" in c and isinstance(c["sum"], list):
                    ent["deps"] = [(rows_by_key[k], col) for k in c["sum"] if k in rows_by_key]
                elif "avg" in c and isinstance(c["avg"], list):
                    ent["deps"] = [(rows_by_key[k], col) for k in c["avg"] if k in rows_by_key]
                elif "ref" in c and c["ref"].get("key") in rows_by_key:
                    ent["deps"] = [(rows_by_key[c["ref"]["key"]], FIRST_COL + int(c["ref"].get("col", j)))]
                elif "share" in c and c["share"].get("of") in rows_by_key:
                    k = FIRST_COL + int(c["share"]["col"])
                    ent["deps"] = [(row, k), (rows_by_key[c["share"]["of"]], k)]
                elif "rowsum" in c:
                    ent["deps"] = [(row, FIRST_COL + k) for k in c["rowsum"]]
            except (ValueError, KeyError, TypeError):
                pass                                    # untraceable: the error is still named, just without a cause
            out[(row, col)] = ent
    return out


MAX_LINE_WARNINGS = 10


def error_warnings(sheet: str, grid, trace: dict, pull_grids: dict) -> list[str]:
    """Warnings for every error cell in the answer sheet and its data sheets, with the likely cause: on a data sheet
    '#NoAccess' = no read access / member not in this cube; on an answer line, the data-sheet cell it depends on."""
    out = []
    for name, g in pull_grids.items():
        by_code: dict[str, list] = {}
        for r, c, code in _error_cells(g):
            by_code.setdefault(code, []).append((r, c))
        for code, cells in by_code.items():
            cause = f": {_PULL_CAUSE[code]}" if code in _PULL_CAUSE else ""
            if len(cells) == 1:
                out.append(f"error cell {_addr(name, *cells[0])} is {code}{cause}")
                continue
            first = ", ".join(f"{L(c)}{r}" for r, c in cells[:3]) + (", ..." if len(cells) > 3 else "")
            out.append(f"error cells on '{name}': {len(cells)} x {code} (first: {first}){cause}")
    vals = {(r, c): code for r, c, code in _error_cells(grid)}

    def sources(rc, seen) -> list[str]:
        if rc in seen:
            return []
        seen.add(rc)
        ent = trace.get(rc) or {}
        if "pull" in ent:
            return [ent["pull"]]
        found = []
        for d in ent.get("deps", []):
            if d in vals:                               # follow only inputs that are errors themselves
                found += [s for s in sources(d, seen) if s not in found]
        return found

    lines: dict = {}                                    # one warning per answer line (or untraced cell)
    for (r, c), code in sorted(vals.items()):
        ent = trace.get((r, c))
        k = ent["key"] if ent else (r, c)
        if k in lines:
            lines[k][1] += 1
            continue
        where = f"{_addr(sheet, r, c)} is {code}"
        if ent is None:
            lines[k] = [f"error cell {where}", 0]
            continue
        src = sources((r, c), set())
        lines[k] = [f"formula error on line '{ent['key']}': {where}" + (f" (depends on {', '.join(src)})" if src else ""), 0]
    texts = [t + (f" (+{n} more cells)" if n else "") for t, n in lines.values()]
    out += texts[:MAX_LINE_WARNINGS]
    if len(texts) > MAX_LINE_WARNINGS:
        out.append(f"... and {len(texts) - MAX_LINE_WARNINGS} more answer lines with error cells on '{sheet}'")
    return out


def _pull_grids(t: Target, pulls: dict) -> dict:
    """Values of the question's data sheets (same read as a pull: ops._used_values); unreadable sheets are skipped."""
    from .ops import _used_values
    out = {}
    for name in pulls:
        try:
            out[name] = com(lambda: _used_values(t.wb.Worksheets(name)))
        except Exception:   # noqa: BLE001 - the scan is advisory
            pass
    return out


CHART_COLS = 10         # a chart sits one column right of the table and is 480-520 pt (~8 standard columns) wide


def _extent(r: int, ncol: int, chart: bool) -> dict:
    """The tool-owned block: A1 to the last written row (r - 1) and the last value column, plus the chart's columns
    when there is one, so user content is not put under a chart that a redraw deletes."""
    return {"rows": r - 1, "cols": FIRST_COL + ncol - 1 + (CHART_COLS if chart else 0)}


def _build(t: Target, q: int, spec: dict, pulls: dict, suffix: str = "Answer", owned: dict | None = None,
           reserve: int = 0) -> dict:
    """Write, style and chart the answer sheet (inside fast_excel: no redraw, manual calculation). reserve: rows
    below the answer to make room for (the findings table, written after the build)."""
    from . import styling
    from .core import owned as owned_mod
    from .timing import stage
    chart = bool(spec.get("chart"))
    with stage("prepare sheet"):
        ws = answer_sheet(t, q, suffix, owned)
        com(ws.Activate)
    with stage("write cells"):
        out = _write(t, q, spec, pulls, ws, suffix, make_room=lambda new: owned_mod.make_room(
            ws, owned, {"rows": new["rows"] + reserve, "cols": new["cols"] + (CHART_COLS if chart else 0)}))
    rows_meta, plan, rows_by_key, checks, cols, ncol, last_col, header_row, r, warnings = out
    extent = _extent(r, ncol, chart)
    longest = max([len(str(ln.get("label", ""))) for _, ln, _ in plan if isinstance(ln, dict) and "cells" in ln] +
                  [len(f"{c['label']} — {cols[0].get('header', '')}") for c in checks if "static" not in c] + [40])
    roles = {"last_col": last_col, "label_cols": ["B"], "freeze_rows": header_row, "last_row": extent["rows"],
             "col_widths": {"A": 2.5, "B": min(80, max(46, longest + 4)), f"C:{last_col}": 17}, "rows": rows_meta}
    t.status(f"formatting 'Q{q} {suffix}' in TN report style ...")
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
            "_trace": cell_trace(plan, rows_by_key, pulls, ncol), "warnings": warnings, "lint": lint,
            "_extent": extent}
