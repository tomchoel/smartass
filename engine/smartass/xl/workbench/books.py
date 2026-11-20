"""Find the user's open workbooks and read them: workbooks, overview, read."""
from __future__ import annotations

import logging
import re

import pywintypes
from openpyxl.utils import get_column_letter as L

from ..instances import busy_error, busy_holder, excel_instances, interactive_excels, is_archive, local_path
from ..jobs import com, hold_turn, start_wait, starting
from ..host import HostError, InputError

log = logging.getLogger("smartview.workbench")
MAX_READ_CELLS = 5000


# ------------------------------------------------------------------ targets
class Book:
    def __init__(self, app, wb, pid: int):
        self.app, self.wb, self.pid = app, wb, pid
        self.notes: list[str] = []

    @property
    def name(self) -> str:
        return com(lambda: self.wb.Name)

    @property
    def full_name(self) -> str:
        return str(com(lambda: self.wb.FullName))

    def sheet(self, name: str | None):
        if not name:
            return com(lambda: self.wb.ActiveSheet)
        try:
            return com(lambda: self.wb.Worksheets(name))
        except pywintypes.com_error:
            names = [s.Name for s in self.wb.Worksheets]
            raise InputError(f"No sheet '{name}' in '{self.name}'. Sheets: {names}") from None


def _books(app) -> list[tuple]:
    """(workbook, its Name, its FullName) of every workbook but add-ins open in app's Excel: read as one, again while
    Excel is busy - it rejects calls while it saves, also a save after the last reply (live, 20.11.2026), and
    find_book reads them before the job's turn there - for 60 s or what is left of the job's start (starting); then
    the busy message."""
    def read():
        out = []
        for wb in app.Workbooks:
            n = str(wb.Name)
            if not n.lower().endswith((".xlam", ".xla")):
                out.append((wb, n, str(wb.FullName)))
        return out
    with start_wait(60.0) as limit:
        return com(read, limit=limit)


def find_book(workbook: str | None, write: bool = False, allow_archive: bool = False) -> Book:
    """The user's active workbook (workbook=None) or an open one by name / file name / full path, in any Excel. Its
    reads before the turn and the first contact share the job's start (jobs.starting)."""
    with starting():
        return _find_book(workbook, write, allow_archive)


def _find_book(workbook: str | None, write: bool, allow_archive: bool) -> Book:
    from ..target import active_book
    instances = excel_instances()
    excels = [(i.pid, i.hwnd, i.app) for i in instances if i.app is not None]
    if not excels:
        holder = busy_holder(instances, workbook.strip()) if workbook else None
        if holder is not None:
            raise busy_error(holder.pid, workbook.strip())
        if instances:
            raise busy_error(instances[0].pid, "your workbook")
        raise HostError("No open Excel window found. Open the workbook in Excel first.")
    hit = None
    if not workbook:
        for pid, _h, app in excels:                      # front-most Excel first
            wb, name = active_book(app)
            if wb is not None and not name.lower().endswith(".xlam"):
                hit = (pid, app, wb)
                break
        if hit is None:
            raise HostError("No active workbook in Excel. Open or click into the workbook in Excel, or pass workbook=.")
    else:
        want = workbook.strip().lower()
        stem = lambda s: re.sub(r"\.(xlsx|xlsm|xlsb|xls)$", "", s)
        cands = []
        books = [(pid, app, wb, name, full) for pid, _h, app in excels for wb, name, full in _books(app)]
        for pid, app, wb, name, full in books:
            n, fn = name.lower(), full.lower()
            if want in (n, fn, local_path(full).lower()) or stem(want) == stem(n):
                cands.append((pid, app, wb, name))
        if not cands:
            for pid, app, wb, name, _full in books:      # unique partial name as a convenience
                if stem(want) in name.lower():
                    cands.append((pid, app, wb, name))
        if not cands:
            holder = busy_holder(instances, workbook.strip())
            if holder is not None:
                raise busy_error(holder.pid, workbook.strip())
            raise InputError(f"No open workbook matches '{workbook}' (open now: {list_names(excels)}). Next step: pass one of those names, or open the workbook in Excel first.")
        if len(cands) > 1:
            raise InputError(f"'{workbook}' matches several open workbooks: {[c[3] for c in cands]}. "
                             "Use the full name.")
        hit = cands[0][:3]
    pid, app, wb = hit
    hold_turn(pid, what="this Excel change", app=app)
    b = Book(app, wb, pid)
    if write and is_archive(local_path(b.full_name)) and not allow_archive:
        raise InputError(f"'{b.name}' is in the Produktregnskapet archive (audit record). Not changed - pass "
                         "allow_archive=true only if the user explicitly asked to edit it.")
    if write:
        _autosave_off(b)
    return b


def _autosave_off(b: Book) -> None:
    """AutoSave (OneDrive/SharePoint files) would save every change to the cloud at once - the user decided that
    Claude never saves. Switch it off for this workbook (that itself saves nothing) and tell the user."""
    try:
        if com(lambda: b.wb.AutoSaveOn, limit=10):
            com(lambda: setattr(b.wb, "AutoSaveOn", False), limit=10)
            b.notes.append("AutoSave was ON for this workbook and has been switched OFF so nothing is saved "
                           "automatically. Save with Ctrl+S when you are happy (or switch AutoSave back on).")
    except Exception:   # noqa: BLE001 - property missing on local files / old builds
        pass


def list_names(excels=None) -> list[str]:
    return [name for _p, _h, app in (excels or interactive_excels()) for _wb, name, _full in _books(app)]


def workbooks() -> list[dict]:
    out = []
    for pid, _h, app in interactive_excels():
        active = com(lambda: app.ActiveWorkbook, limit=20)
        aname = str(active.Name) if active is not None else None
        for wb, _name, _full in _books(app):
            try:
                sel = _addr(com(lambda: app.Selection)) if str(wb.Name) == aname else None
            except Exception:   # noqa: BLE001 - selection may be a chart/shape
                sel = None
            out.append({"name": str(wb.Name), "path": local_path(str(wb.FullName)), "excel_pid": pid,
                        "active": str(wb.Name) == aname, "sheets": [s.Name for s in wb.Worksheets],
                        "active_sheet": str(wb.ActiveSheet.Name), "selection": sel,
                        "unsaved_changes": not bool(wb.Saved), "read_only": bool(wb.ReadOnly),
                        "archive": is_archive(local_path(str(wb.FullName)))})
    return out


# ------------------------------------------------------------------ helpers
def _grid(v) -> list[list]:
    if v is None or not isinstance(v, tuple):
        return [[v]]
    return [list(r) for r in v]


def _clean(v):
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return int(v)
    if isinstance(v, int) and v < -2146826000:                    # Excel error codes come back as big negatives
        return {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
                -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!"}.get(v, "#ERROR")
    if isinstance(v, pywintypes.TimeType):
        return v.isoformat()[:10]
    return v


def _rng(ws, address: str | None):
    if not address or address.lower() == "selection":
        sel = com(lambda: ws.Application.Selection)
        try:
            if com(lambda: sel.Worksheet.Name) == com(lambda: ws.Name):
                return _one_area(ws, sel)
        except Exception:   # noqa: BLE001
            pass
        raise InputError("Give a range address (the current selection is not a range on that sheet).")
    try:
        rng = com(lambda: ws.Range(address))
    except pywintypes.com_error:
        raise InputError(f"'{address}' is not a valid range address.") from None
    return _one_area(ws, rng)


def _one_area(ws, rng):
    """Snapshots and reverts cover one rectangular block on one sheet: refuse multi-area ranges and names that
    point at another sheet (they would be recorded under the wrong sheet)."""
    if int(com(lambda: rng.Areas.Count)) > 1:
        raise InputError("That range has several areas; do one rectangular block per call.")
    if com(lambda: rng.Worksheet.Name) != com(lambda: ws.Name):
        raise InputError(f"That range is on sheet '{com(lambda: rng.Worksheet.Name)}'; pass sheet= that sheet.")
    return rng


def _addr(rng) -> str:
    """'B2:F7' (relative). pywin32 late binding exposes parameterised properties as Get<Name>(...)."""
    return str(com(lambda: rng.GetAddress(False, False)))


def _ncells(rng) -> int:
    return int(com(lambda: rng.Rows.Count)) * int(com(lambda: rng.Columns.Count))


def _label_context(ws, row: int, col: int) -> dict:
    """The row label (first text in columns A-C on that row) and the column header (nearest text above in the
    column, within 40 rows) - so a reference can be explained in words."""
    out = {}
    try:
        left = _grid(com(lambda: ws.Range(ws.Cells(row, 1), ws.Cells(row, min(max(col - 1, 1), 3))).Value))[0]
        txt = [str(x).strip() for x in left if isinstance(x, str) and x.strip()]
        if txt:
            out["row_label"] = txt[0][:80]
        top = max(1, row - 40)
        above = [r[0] for r in _grid(com(lambda: ws.Range(ws.Cells(top, col), ws.Cells(max(row - 1, top), col)).Value))]
        for v in reversed(above):
            if isinstance(v, str) and v.strip():
                out["col_header"] = v.strip()[:60]
                break
    except Exception:   # noqa: BLE001 - context is a convenience
        pass
    return out


# ------------------------------------------------------------------ read
def read(workbook, sheet, address, what: str = "values", max_cells: int = MAX_READ_CELLS) -> dict:
    b = find_book(workbook)
    ws = b.sheet(sheet)
    rng = _rng(ws, address)
    n = _ncells(rng)
    if n > max_cells:
        r1 = com(lambda: rng.Row); c1 = com(lambda: rng.Column)
        cols = int(com(lambda: rng.Columns.Count))
        keep = max(1, max_cells // cols)
        rng = com(lambda: ws.Range(ws.Cells(r1, c1), ws.Cells(r1 + keep - 1, c1 + cols - 1)))
        truncated = f"{n} cells requested; first {keep} rows returned (raise max_cells to read more)"
    else:
        truncated = None
    out = {"workbook": b.name, "sheet": com(lambda: ws.Name), "address": _addr(rng)}
    if what in ("values", "both"):
        out["values"] = [[_clean(v) for v in row] for row in _grid(com(lambda: rng.Value2))]
    if what in ("formulas", "both"):
        r0, c0 = com(lambda: rng.Row), com(lambda: rng.Column)
        f = _grid(com(lambda: rng.Formula2))
        out["formulas"] = {f"{L(c0 + j)}{r0 + i}": v for i, row in enumerate(f) for j, v in enumerate(row)
                           if isinstance(v, str) and v.startswith("=")}
    if what == "formats":
        r0, c0 = com(lambda: rng.Row), com(lambda: rng.Column)
        nr, nc = int(com(lambda: rng.Rows.Count)), int(com(lambda: rng.Columns.Count))
        fm = _format_grid(com(lambda: rng.NumberFormat), nr, nc,
                          lambda i, j: com(lambda: ws.Cells(r0 + i, c0 + j).NumberFormat))
        out["number_formats"] = fm
        styles = {}
        for i in range(len(fm)):
            for j in range(len(fm[0])):
                c = ws.Cells(r0 + i, c0 + j)
                styles.setdefault(str(com(lambda: c.Style.Name)), []).append(f"{L(c0 + j)}{r0 + i}")
        out["styles"] = {k: (v if len(v) <= 12 else v[:12] + [f"... +{len(v) - 12}"]) for k, v in styles.items()}
    if truncated:
        out["truncated"] = truncated
    return out


def _format_grid(whole, nr: int, nc: int, cell) -> list[list]:
    """Number formats in the range's shape: Range.NumberFormat is ONE value for a uniform range and None when the
    cells differ - then each cell is read (cell(i, j))."""
    if whole is not None:
        return [[whole] * nc for _ in range(nr)]
    return [[cell(i, j) for j in range(nc)] for i in range(nr)]


def overview(workbook, sheet, preview_rows: int = 12) -> dict:
    b = find_book(workbook)
    sheets = [b.sheet(sheet)] if sheet else list(com(lambda: b.wb.Worksheets))
    act = com(lambda: b.wb.ActiveSheet.Name)
    out = {"workbook": b.name, "path": local_path(b.full_name), "active_sheet": act,
           "unsaved_changes": not bool(com(lambda: b.wb.Saved)), "sheets": []}
    for ws in sheets:
        ur = com(lambda: ws.UsedRange)
        info = {"name": com(lambda: ws.Name), "used": _addr(ur), "visible": com(lambda: ws.Visible) == -1}
        counts = {}
        for kind, label in ((-4123, "formulas"), (2, "constants")):
            try:
                counts[label] = int(com(lambda: ur.SpecialCells(kind).Count, limit=10))
            except pywintypes.com_error:
                counts[label] = 0
        info.update(counts)
        try:
            info["tables"] = [{"name": t.Name, "range": _addr(t.Range)} for t in ws.ListObjects]
        except Exception:   # noqa: BLE001
            pass
        rows = preview_rows if (sheet or len(sheets) == 1) else min(preview_rows, 5)
        if rows:
            r0, c0 = com(lambda: ur.Row), com(lambda: ur.Column)
            nr = min(rows, int(com(lambda: ur.Rows.Count)))
            nc = min(26, int(com(lambda: ur.Columns.Count)))
            block = com(lambda: ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + nr - 1, c0 + nc - 1)).Value2)
            info["preview_from"] = f"{L(c0)}{r0}"
            info["preview"] = [[(str(v)[:40] if isinstance(v, str) else _clean(v)) for v in row] for row in _grid(block)]
        out["sheets"].append(info)
    try:
        out["names"] = [{"name": n.Name, "refers_to": str(n.RefersTo)[:80]} for n in list(b.wb.Names)[:40]
                        if not str(n.Name).startswith(("_xlfn", "_xlpm"))]
    except Exception:   # noqa: BLE001
        pass
    return out
