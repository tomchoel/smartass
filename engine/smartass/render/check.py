# engine/smartass/render/check.py
"""The fast layout check after every draw (sub-spec 5 §5, D6): the SheetPlan plus a few COM reads (the chart count,
one Value2 read of the drawn block, one read per Data header status cell). Each finding is one line with its fix,
returned in the reply as `layout`; a finding is a warning, never a stop, and the check never raises.

Rules:
  1. zones in the fixed order (§1.1), no Plan zone;
  2. at least one chart block, and at least one chart drawn on the sheet;
  3. every numeric cell of a table line has a number format (p.rows num);
  4. no number too wide for its column (shown as ###), no General number shown as E+ (estimated from the
     formatted text against plan.widths: Range.Text of a multi-cell range is Null in Excel, so it cannot be read
     in one call);
  5. no Key point cell holds an Excel error;
  6. every Data header control total reads OK;
  7. the Sources rows of this Q were written."""
from __future__ import annotations

from .layout import DEFAULT_WIDTH, FIRST_COL, TEXT_COL, col_letter, col_number

ZONE_ORDER = ("title", "subtitle", "key_points", "charts", "tables", "checks", "notes", "chart_data", "source")
NUMERIC_KINDS = ("line", "sum", "grand")
XL_ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
             -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!"}
GENERAL_DIGITS = 11                      # a General cell switches to E+ notation from about 11 digits
PAD = 0.0                                # tn.text_units (x 1.08) already over-estimates; the margin fits in it
BOLD_KINDS = ("sum", "grand")


def _number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool) and v not in XL_ERRORS


def _error(v) -> str | None:
    if isinstance(v, int) and not isinstance(v, bool) and v in XL_ERRORS:
        return XL_ERRORS[v]
    if isinstance(v, str) and v.startswith("#") and (v.endswith(("!", "?")) or v == "#N/A"):
        return v
    return None


def shown_text(v: float, num: str) -> str:
    """The text Excel shows for v in a TN number kind (styling.NUM_FORMATS): thousands separators, negatives in
    parentheses (or '-' for percentages), the ▲ / ▼ ( ) of a variance, the trailing _) space of a positive."""
    if num == "markor":
        return "●"
    pct = num in ("prosent1", "avvikpst", "prosent")
    body = f"{abs(v) * 100:,.1f}%" if pct else f"{abs(v):,.1f}" if num == "tall1" else f"{abs(v):,.0f}"
    if num in ("avvik", "avvikpst"):
        return f"▲ {body}" if v >= 0 else f"▼ ({body})"
    if pct:
        return f"-{body}" if v < 0 else body
    return f"({body})" if v < 0 else f"{body})"            # '_)' keeps a parenthesis-wide space


def shown_width(v: float, num: str, bold: bool = False) -> float:
    """The shown text's width in column units (Arial 10, tn.text_units)."""
    from ..xl.vendor import tn_rapportstil as tn
    return tn.text_units(shown_text(v, num), 10, bold)


def _grid(plan, ws):
    """{(row, col): value} of A1:<last_col><last_row> in one Value2 read."""
    from ..xl.jobs import com
    rng = ws.Range(f"A1:{plan.last_col}{plan.last_row}")
    v = com(lambda: rng.Value2)
    if not isinstance(v, tuple):
        v = ((v,),)
    return {(r + 1, c + 1): x for r, row in enumerate(v) for c, x in enumerate(row) if x is not None}


def _zones(plan) -> list[str]:
    out = []
    names = [z for z, _ in plan.zones]
    if "plan" in names:
        out.append("layout: the sheet has a Plan zone; fix: the plan is not drawn (remove it from the layout)")
    known = [(z, r) for z, r in plan.zones if z in ZONE_ORDER]
    rank = [ZONE_ORDER.index(z) for z, _ in known]
    rows = [r for _, r in known]
    if rank != sorted(rank) or rows != sorted(rows):
        out.append(f"layout: zones out of order ({', '.join(names)}); fix: keep {' > '.join(ZONE_ORDER)}")
    return out


def _charts(plan, ws) -> list[str]:
    from ..xl.jobs import com
    if not plan.charts:
        return ["layout: no chart block; fix: add a chart block chosen by the message (skill reference charts.md)"]
    if not com(lambda: ws.ChartObjects().Count):
        return ["layout: no chart was drawn; fix: see the chart warnings, patch the chart block (kind, source)"]
    return []


def _formats(plan, grid) -> list[str]:
    by_row: dict[int, list[tuple[int, object]]] = {}
    for (r, c), v in plan.cells.items():
        if c >= FIRST_COL:
            by_row.setdefault(r, []).append((c, v))
    out = []
    for d in plan.rows:
        if d.get("kind") not in NUMERIC_KINDS or "status_col" in d:
            continue
        row, num = d["row"], d.get("num") or {}
        missing = []
        for c, v in sorted(by_row.get(row, ()), key=lambda x: x[0]):
            r = row
            numeric = (isinstance(v, str) and v.startswith("=")) or _number(v)
            if numeric and not num.get(col_letter(c)):
                missing.append(col_letter(c))
                shown = grid.get((r, c))
                if _number(shown) and abs(shown) >= 10 ** GENERAL_DIGITS:
                    out.append(f"layout: {col_letter(c)}{r} shows a number as E+; fix: give the row a number format")
        if missing:
            out.append(f"layout: row {row} has no number format in {', '.join(missing)}; fix: set num for those "
                       "columns in the layout")
    return out


def _narrow(plan, grid) -> list[str]:
    worst: dict[str, tuple[float, int]] = {}          # column -> (shown width, row) of its widest number
    for d in plan.rows:
        for letter, num in (d.get("num") or {}).items():
            v = grid.get((d["row"], col_number(letter)))
            if _number(v):
                n = shown_width(v, num, d.get("kind") in BOLD_KINDS)
                if n > worst.get(letter, (0.0, 0))[0]:
                    worst[letter] = (n, d["row"])
    out = []
    for letter, (n, row) in sorted(worst.items(), key=lambda x: col_number(x[0])):
        width = plan.widths.get(letter, DEFAULT_WIDTH)
        if n + PAD > width:
            out.append(f"layout: column {letter} too narrow for {letter}{row} (needs {n:.1f}, width {width:g}: "
                       f"shown as ###); fix: widen column {letter} or show the block in MNOK")
    return out


def _key_points(plan, grid) -> list[str]:
    rows = [r for z, r in plan.zones]
    start = next((r for z, r in plan.zones if z == "key_points"), None)
    if start is None:
        return []
    end = min([r for r in rows if r > start] or [plan.last_row + 1])
    out = []
    for r in range(start + 1, end):
        err = _error(grid.get((r, TEXT_COL)))
        if err:
            out.append(f"layout: key point {col_letter(TEXT_COL)}{r} shows {err}; fix: the line it points at is "
                       "missing or empty - patch the key point or its table")
    return out


def _data(data_ws, status_cells) -> list[str]:
    from ..xl.jobs import com
    out = []
    for qid, cell in (status_cells or {}).items():
        v = com(lambda cell=cell: data_ws.Range(cell).Value)
        if v != "OK":
            out.append(f"layout: the Data header control total of query {qid} says {v or 'nothing'} ({cell}); "
                       "fix: fin_refresh, and never edit the Data table by hand")
    return out


def check(plan, ws, data_ws, sources_ok: bool, status_cells: dict | None = None) -> list[str]:
    """One line per finding with its fix (module doc). status_cells: query id -> the Data header status cell
    (datasheet.Header.status). Never raises: a rule that cannot run says so."""
    out: list[str] = []

    def run(name, fn, *a):
        try:
            out.extend(fn(*a))
        except Exception as e:   # noqa: BLE001 - the check never fails an answer
            out.append(f"layout: check '{name}' could not run ({e})")
    run("zones", _zones, plan)
    run("charts", _charts, plan, ws)
    try:
        grid = _grid(plan, ws)
    except Exception as e:   # noqa: BLE001
        out.append(f"layout: check could not read the sheet ({e})")
        grid = None
    if grid is not None:
        run("formats", _formats, plan, grid)
        run("widths", _narrow, plan, grid)
        run("key points", _key_points, plan, grid)
    run("data header", _data, data_ws, status_cells)
    if not sources_ok:
        out.append("layout: the Sources rows of this answer are missing; fix: see the Sources warning and ask again")
    return out
