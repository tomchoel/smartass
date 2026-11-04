"""Write an answer sheet: cell rectangles in one COM call, line formulas, formats, named ranges and the chart."""
from __future__ import annotations

import re

from openpyxl.utils import get_column_letter as L

from smartass.xl import formulas as fx
from smartass.xl.drill_xl import drill_formula, missing_columns, table_columns as drill_table_columns, table_name
from smartass.xl.jobs import com
from smartass.xl.sheets import lookup_formula
from smartass.xl.target import Target
from smartass.xl.host import HostError, InputError

FIRST_COL = 3             # column C: first value column; B = labels, A = margin
NUM_KINDS = {"heltall", "tall1", "prosent1", "prosent", "faktor"}
CHART_TYPES = {"column": 51, "bar": 57, "line": 4, "waterfall": 119}   # xlColumnClustered, xlBarClustered, xlLine, xlWaterfall


_LEAD_SUM = re.compile(r"\s*([A-Za-z_][A-Za-z0-9_]*(?:\s*\+\s*[A-Za-z_][A-Za-z0-9_]*){2,})\s*(?=-|$)")


def _expr_to_formula(expr: str, col_letter: str, rows_by_key: dict) -> str:
    """Line keys -> cells. A leading run of 3+ keys added with + becomes SUM(...) (tn-formelstil blocks + chains):
    'a+b+c-d' -> '=SUM(C1,C5,C9)-C12'."""
    def rep(m):
        k = m.group(0)
        if k not in rows_by_key:
            raise InputError(f"Unknown line key '{k}' in expression '{expr}'. Known: {sorted(rows_by_key)}")
        return f"{col_letter}{rows_by_key[k]}"
    head, rest = "", expr
    m = _LEAD_SUM.match(expr)
    if m:
        keys = re.split(r"\s*\+\s*", m.group(1))
        cells = [rep(re.match(r".+", k)) for k in keys]
        head, rest = "SUM(" + ",".join(cells) + ")", expr[m.end():]
    return "=" + head + re.sub(r"[A-Za-z_][A-Za-z0-9_]*", rep, rest)


class _Cells:
    """The answer cells, buffered and written as one rectangle: one COM call instead of one per cell (a fin_pnl
    answer took ~10 s cell by cell). Should Excel refuse the batch (one bad formula), the cells are written one by
    one, so a bad cell costs only its own warning, as before."""

    def __init__(self):
        self.cells: dict[tuple[int, int], tuple[object, str]] = {}

    def put(self, row: int, col: int, value, label: str = "") -> None:
        self.cells[(row, col)] = (value, label)

    def flush(self, ws, warnings: list[str]) -> None:
        if not self.cells:
            return
        rows, cols = [r for r, _ in self.cells], [c for _, c in self.cells]
        r0, c0 = min(rows), min(cols)
        grid = [[""] * (max(cols) - c0 + 1) for _ in range(max(rows) - r0 + 1)]
        for (r, c), (v, _) in self.cells.items():
            grid[r - r0][c - c0] = v
        rng = ws.Range(ws.Cells(r0, c0), ws.Cells(max(rows), max(cols)))
        try:
            com(lambda: setattr(rng, "Formula2", tuple(tuple(x) for x in grid)))
            return
        except HostError:
            raise                                          # Excel busy: the whole answer has to wait, not one cell
        except Exception:   # noqa: BLE001 - one bad formula refuses the batch: fall back to cell by cell
            pass
        for (r, c), (v, label) in sorted(self.cells.items()):
            cell = ws.Cells(r, c)
            try:
                com(lambda: setattr(cell, "Formula2" if isinstance(v, str) and v.startswith("=") else "Value", v))
            except HostError:
                raise
            except Exception as e:   # noqa: BLE001
                warnings.append(f"{label or f'row {r}'}: {e}")


def _write(t: Target, q: int, spec: dict, pulls: dict, ws, suffix: str = "Answer", make_room=None):
    """All answer cells (title, key points, lines, checks, notes); returns the layout for styling. make_room(extent)
    is called with the block about to be written ({"rows", "cols"}) just before the cells go to the sheet."""
    t.status(f"building the answer sheet 'Q{q} {suffix}' ...")
    cols = spec.get("columns") or []
    if not cols:
        raise InputError("answer spec needs 'columns'")
    ncol = len(cols)
    last_col = L(FIRST_COL + ncol - 1)
    rows_meta: list[dict] = []
    r = 1
    buf = _Cells()

    def put(row, col, value):
        # text is stored as text: Excel would otherwise turn 'Jan 26' into a date, '1/2' into a fraction ...
        if isinstance(value, str) and not value.startswith("="):
            value = "'" + value
        buf.put(row, col, value)

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
    table_cols: dict[str, list[str] | None | bool] = {}   # drill tables: name -> header (None = not found,
    #                                                       False = unreadable: write the formula, no column check)
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
                        except Exception:   # noqa: BLE001 - tables unreadable: write the formula, no column check
                            table_cols[tn_] = False
                    if table_cols[tn_] is None:                    # Excel rejects a reference to a missing table
                        buf.put(row, FIRST_COL + j, f"'#n/a: {tn_} not found", str(ln.get("label")))
                        raise InputError(f"the lines table {tn_} not found: drill it first (sv_drill q={q})")
                    miss = missing_columns(table_cols[tn_] or None, c.get("where") or {})
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
                elif "avg" in c:                                   # {"avg": [keys]}: AVERAGE over those lines
                    f = fx.avg_rows(col, [rows_by_key[k] for k in c["avg"]])
                elif "ref" in c:                                   # another answer cell: {"ref": {"key":.., "col":..}}
                    f = f"={L(FIRST_COL + int(c['ref'].get('col', j)))}{rows_by_key[c['ref']['key']]}"
                elif "share" in c:                                 # {"share": {"col", "of"}}: this line / line 'of'
                    a = L(FIRST_COL + int(c["share"]["col"]))
                    f = fx.share(f"{a}{row}", f"{a}{rows_by_key[c['share']['of']]}")
                elif "rowsum" in c:                                # sum of other columns on this line
                    f = fx.sum_cols(row, [FIRST_COL + k for k in c["rowsum"]])
                elif "value" in c:
                    f = c["value"]
                elif "text" in c:
                    f = "'" + str(c["text"])
                else:
                    continue
                buf.put(row, FIRST_COL + j, f, str(ln.get("label")))
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
                    buf.put(r, FIRST_COL, f, f"check {ck['label']}")
                    buf.put(r, FIRST_COL + 1,
                            f'=IF(ISNUMBER({L(FIRST_COL)}{r}),IF(ABS({L(FIRST_COL)}{r})<1,"OK","DIFF"),"DIFF")',
                            f"check {ck['label']}")
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

    if make_room:
        make_room({"rows": r - 1, "cols": FIRST_COL + ncol - 1})
    buf.flush(ws, warnings)
    used = {m["row"] for m in rows_meta}
    for blank in range(1, r):
        if blank not in used:
            rows_meta.append({"row": blank, "kind": "blank"})
    rows_meta.sort(key=lambda m: m["row"])
    return rows_meta, plan, rows_by_key, checks, cols, ncol, last_col, header_row, r, warnings


def _named(shape):
    """Name the chart so a redraw deletes it and only it (owned.drop_charts); user charts keep their names."""
    from smartass.xl.owned import CHART_NAME
    try:
        shape.Name = CHART_NAME
    except Exception:   # noqa: BLE001 - an unnamed chart is only deleted by a full clear
        pass
    return shape


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
        shape = _named(ws.Shapes.AddChart2(-1, xl_type, left, top, 480, 300))
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
        shape = _named(ws.Shapes.AddChart2(-1, 51, left, top, 480, 300))
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
    shape = _named(ws.Shapes.AddChart2(-1, 51 if xl_type == 119 else xl_type, left, top, 520, 300))
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
