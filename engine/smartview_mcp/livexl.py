"""xl_* tools: work LIVE in the user's open workbooks (any Excel window), without Smart View.

Agreed with the user 2026-10-15:
- the ACTIVE workbook by default, another open one by name;
- a snapshot before EVERY change (values, formulas, formats via Excel's XML-spreadsheet export) + xl_revert,
  because Ctrl+Z cannot undo COM writes. Snapshots: %LOCALAPPDATA%\\smartass\\snapshots (config.HOME), kept 7 days, never OneDrive;
- the screen is frozen while writing, then Excel jumps to the changed area;
- the user's workbook is NEVER saved here (no Save / SaveAs anywhere in this module);
- every written formula goes through the tn-formelstil checker; 'block' findings must be fixed.
All functions run on a user-Excel job thread (host.call(user_excel=True)) and take that Excel's turn.
"""
from __future__ import annotations

import datetime as dt
import os
import hashlib
import json
import logging
import re
import threading
import time
import uuid

import pywintypes
from openpyxl.utils import get_column_letter as L
from openpyxl.utils.cell import range_boundaries

from . import config
from . import formulas as fx
from .host import HostError, InputError
from .userexcel import (
    _wait_ready,
    busy_error,
    busy_holder,
    com,
    excel_instances,
    fast_excel,
    hold_turn,
    interactive_excels,
    is_archive,
    local_path,
)

log = logging.getLogger("smartview.livexl")

SNAP_DIR = config.SVKIT / "snapshots"
KEEP_DAYS = 7
MAX_READ_CELLS = 5000
MAX_SNAPSHOT_CELLS = 250_000
JUMP = os.environ.get("SV_XL_JUMP", "1") == "1"     # tests set 0: do not move the user's view
TN_EMPH = {"": "", "sum": "Sum", "hovedsum": "Hovedsum", "grand": "Hovedsum", "input": "Input", "nokkel": "Nøkkel",
           "nøkkel": "Nøkkel", "key": "Nøkkel"}


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


def _books(app):
    out = []
    for wb in com(lambda: app.Workbooks):
        n = str(wb.Name)
        if not n.lower().endswith((".xlam", ".xla")):
            out.append(wb)
    return out


def find_book(workbook: str | None, write: bool = False, allow_archive: bool = False) -> Book:
    """The user's active workbook (workbook=None) or an open one by name / file name / full path, in any Excel."""
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
            wb = com(lambda: app.ActiveWorkbook, limit=20)
            if wb is not None and not str(wb.Name).lower().endswith(".xlam"):
                hit = (pid, app, wb)
                break
        if hit is None:
            raise HostError("No active workbook in Excel. Open or click into the workbook in Excel, or pass workbook=.")
    else:
        want = workbook.strip().lower()
        stem = lambda s: re.sub(r"\.(xlsx|xlsm|xlsb|xls)$", "", s)
        cands = []
        for pid, _h, app in excels:
            for wb in _books(app):
                n, fn = str(wb.Name).lower(), str(wb.FullName).lower()
                if want in (n, fn, local_path(str(wb.FullName)).lower()) or stem(want) == stem(n):
                    cands.append((pid, app, wb))
        if not cands:
            for pid, _h, app in excels:                  # unique partial name as a convenience
                for wb in _books(app):
                    if stem(want) in str(wb.Name).lower():
                        cands.append((pid, app, wb))
        if not cands:
            holder = busy_holder(instances, workbook.strip())
            if holder is not None:
                raise busy_error(holder.pid, workbook.strip())
            raise InputError(f"No open workbook matches '{workbook}' (open now: {list_names(excels)}). Next step: pass one of those names, or open the workbook in Excel first.")
        if len(cands) > 1:
            raise InputError(f"'{workbook}' matches several open workbooks: {[str(c[2].Name) for c in cands]}. "
                             "Use the full name.")
        hit = cands[0]
    pid, app, wb = hit
    hold_turn(pid, what="this Excel change")
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
    return [str(wb.Name) for _p, _h, app in (excels or interactive_excels()) for wb in _books(app)]


def workbooks() -> list[dict]:
    out = []
    for pid, _h, app in interactive_excels():
        active = com(lambda: app.ActiveWorkbook, limit=20)
        aname = str(active.Name) if active is not None else None
        for wb in _books(app):
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


# ------------------------------------------------------------------ snapshots
def _purge_old():
    cut = time.time() - KEEP_DAYS * 86400
    for p in list(SNAP_DIR.glob("*.json")) + list(SNAP_DIR.glob("*.meta")):
        try:
            if p.stat().st_mtime < cut:
                p.unlink()
        except OSError:
            pass


def _capture(ws, rng) -> dict:
    """XML-spreadsheet export (values with types, formats, merges) + the real Formula2 of formula cells (the XML
    turns dynamic arrays into literal per-cell formulas) + column widths."""
    n = _ncells(rng)
    if n > MAX_SNAPSHOT_CELLS:
        raise InputError(f"The change covers {n} cells - too big to snapshot safely (limit {MAX_SNAPSHOT_CELLS}). "
                         "Work on a smaller range.")
    r0, c0 = com(lambda: rng.Row), com(lambda: rng.Column)
    f2 = _grid(com(lambda: rng.Formula2))
    formulas = {f"{L(c0 + j)}{r0 + i}": v for i, row in enumerate(f2) for j, v in enumerate(row)
                if isinstance(v, str) and v.startswith("=")}
    widths = {}
    ncol = int(com(lambda: rng.Columns.Count))
    if ncol <= 200:
        for j in range(ncol):
            widths[L(c0 + j)] = com(lambda: ws.Columns(c0 + j).ColumnWidth)
    consts = _constants(f2, _grid(com(lambda: rng.Value2)), r0, c0) if n <= MAX_STYLE_CELLS else None
    return {"xml": com(lambda: rng.GetValue(11)), "formulas": formulas, "widths": widths,
            "styles": _style_runs(ws, rng, r0, c0), "constants": consts}


def _constants(f2: list[list], v2: list[list], r0: int, c0: int) -> dict:
    """{address: value} of every non-formula cell, typed (text stays text, empty = None). A cell with no formula but
    a value is a spill child of a dynamic array: not a constant (writing it would break the spill)."""
    out = {}
    for i, row in enumerate(f2):
        for j, f in enumerate(row):
            v = v2[i][j]
            if isinstance(f, str) and f.startswith("=") or (f in ("", None) and v not in ("", None)):
                continue
            out[f"{L(c0 + j)}{r0 + i}"] = None if v == "" else v
    return out


def _same(a, b) -> bool:
    num = (int, float)
    if isinstance(a, num) and isinstance(b, num) and not isinstance(a, bool) and not isinstance(b, bool):
        return a == b
    return type(a) is type(b) and a == b


def _coerced(consts: dict, now: list[list], r0: int, c0: int) -> list[tuple[str, object]]:
    """(address, snapshot value) of the constants the XML import did not bring back as they were: it turns text
    like '2026' into numbers and can leave stray values in cells that were empty."""
    out = []
    for i, row in enumerate(now):
        for j, v in enumerate(row):
            a = f"{L(c0 + j)}{r0 + i}"
            if a in consts and not _same(None if v == "" else v, consts[a]):
                out.append((a, consts[a]))
    return out


MAX_STYLE_CELLS = 20_000


def _style_runs(ws, rng, r0: int, c0: int) -> dict | None:
    """{style name: ['C3:E3', 'B7', ...]} - the XML import keeps the look but resets every cell to 'Normal', which
    would lose the TN named styles. One COM read per cell, so only up to MAX_STYLE_CELLS."""
    nr, nc = int(com(lambda: rng.Rows.Count)), int(com(lambda: rng.Columns.Count))
    if nr * nc > MAX_STYLE_CELLS:
        return None
    runs: dict[str, list[str]] = {}
    for i in range(nr):
        prev, start = None, 0
        for j in range(nc + 1):
            if j < nc:
                cell = ws.Cells(r0 + i, c0 + j)
                name = str(com(lambda: cell.Style.Name)) + "\x01" + str(com(lambda: cell.NumberFormat))
            else:
                name = None
            if name != prev:
                if prev is not None:
                    a, z = f"{L(c0 + start)}{r0 + i}", f"{L(c0 + j - 1)}{r0 + i}"
                    runs.setdefault(prev, []).append(a if a == z else f"{a}:{z}")
                prev, start = name, j
    return runs


def _state_hash(rng) -> str:
    return hashlib.sha1(json.dumps(_grid(com(lambda: rng.Formula2)), default=str).encode()).hexdigest()


_seq = 0
_seq_lock = threading.Lock()


def _new_change(b: Book, ws, op: str, address: str, before: dict | None, extra: dict | None = None,
                label: str | None = None) -> dict:
    SNAP_DIR.mkdir(parents=True, exist_ok=True)
    _purge_old()
    global _seq
    with _seq_lock:                              # ids sort in creation order: UTC (no DST jumps) + counter
        _seq = (_seq + 1) % 1000
        now = dt.datetime.now(dt.timezone.utc)
        cid = now.strftime("%Y%m%d-%H%M%S-") + f"{now.microsecond:06d}{_seq:03d}-" + uuid.uuid4().hex[:3]
    return {"id": cid, "ts": dt.datetime.now().isoformat(timespec="seconds"), "workbook": b.name,
            "full_name": b.full_name, "pid": b.pid, "sheet": com(lambda: ws.Name), "op": op, "address": address,
            "label": label, "before": before, "extra": extra or {}, "reverted": False, "state": "pending"}


META_KEYS = ("id", "ts", "workbook", "full_name", "pid", "sheet", "op", "address", "label", "reverted", "state")


def _save_change(ch: dict) -> None:
    """Full snapshot + a small meta file (listing changes must not parse every snapshot's XML)."""
    (SNAP_DIR / f"{ch['id']}.json").write_text(json.dumps(ch, ensure_ascii=False, default=str), encoding="utf-8")
    (SNAP_DIR / f"{ch['id']}.meta").write_text(json.dumps({k: ch.get(k) for k in META_KEYS}, ensure_ascii=False,
                                                          default=str), encoding="utf-8")


def _apply(b: Book, ws, ch: dict, mutate, undo=None) -> None:
    """The snapshot is on disk BEFORE anything changes; if the change fails halfway it is rolled back from the
    snapshot (undo) and recorded as such, so no edit is ever left without a way back."""
    _save_change(ch)
    try:
        with fast_excel(b.app):
            mutate()
    except Exception as e:
        try:
            if undo:
                with fast_excel(b.app):
                    undo()
            ch["reverted"] = f"auto: the change failed ({str(e)[:120]})"
        except Exception:   # noqa: BLE001
            log.exception("auto-rollback of %s failed", ch["id"])
            ch["state"] = "failed - partly applied, revert with force=true"
        _save_change(ch)
        raise
    ch["state"] = "applied"


def _restore(ws, address: str, snap: dict) -> None:
    rng = com(lambda: ws.Range(address))
    com(lambda: rng.SetValue(11, snap["xml"]))
    r0, c0 = com(lambda: rng.Row), com(lambda: rng.Column)
    now = _grid(com(lambda: rng.Formula2))
    stray = []                        # spill children the XML wrote as literal formulas: clear before re-spilling
    for i, row in enumerate(now):
        for j, v in enumerate(row):
            a = f"{L(c0 + j)}{r0 + i}"
            if isinstance(v, str) and v.startswith("=") and a not in snap["formulas"]:
                stray.append(a)
    sep = com(lambda: ws.Application.International)[4]   # union addresses use the LOCALE list separator (nb-NO ';')
    for k in range(0, len(stray), 20):                    # Range() takes at most 255 characters
        com(lambda: ws.Range(sep.join(stray[k:k + 20])).ClearContents())
    for key, areas in (snap.get("styles") or {}).items():
        name, _, nf = key.partition("\x01")              # style name + the cell's own number format on top
        for k in range(0, len(areas), 15):
            part = com(lambda: ws.Range(sep.join(areas[k:k + 15])))
            try:
                if name != "Normal":
                    com(lambda: setattr(part, "Style", name))
                if nf:
                    com(lambda: setattr(part, "NumberFormat", nf))
            except pywintypes.com_error:                  # style removed from the workbook meanwhile
                log.warning("style %s not restored on %s", name, areas[k:k + 15])
    for a, f in snap["formulas"].items():
        cell = com(lambda: ws.Range(a))
        if com(lambda: cell.Formula2) != f:
            com(lambda: setattr(cell, "Formula2", f))
    if snap.get("constants") is not None:
        for a, v in _coerced(snap["constants"], _grid(com(lambda: rng.Value2)), r0, c0):
            cell = com(lambda: ws.Range(a))
            if v is None:
                com(lambda: cell.ClearContents())
            elif isinstance(v, str):
                com(lambda: setattr(cell, "Formula2", "'" + v))    # as text, like xl_write stores it
            else:
                com(lambda: setattr(cell, "Value2", v))
    for col, w in (snap.get("widths") or {}).items():
        com(lambda: setattr(ws.Columns(col), "ColumnWidth", w))


def _jump(b: Book, ws, address: str) -> None:
    """After the frozen write: show the user what changed (the agreed 'frozen, then jump')."""
    if not JUMP:
        return
    try:
        com(lambda: b.wb.Activate(), limit=10)
        com(lambda: ws.Activate(), limit=10)
        com(lambda: b.app.Goto(ws.Range(address)), limit=10)
    except Exception:   # noqa: BLE001 - cosmetic
        pass


def _check_written(ws, rng) -> dict:
    """tn-formelstil on the formulas in the changed rows (whole used width, so row consistency is judged too)."""
    ur = com(lambda: ws.UsedRange)
    r1, r2 = com(lambda: rng.Row), com(lambda: rng.Row) + int(com(lambda: rng.Rows.Count)) - 1
    c1, c2 = com(lambda: ur.Column), com(lambda: ur.Column) + int(com(lambda: ur.Columns.Count)) - 1
    if (r2 - r1 + 1) * (c2 - c1 + 1) > 60_000:
        c1, c2 = com(lambda: rng.Column), com(lambda: rng.Column) + int(com(lambda: rng.Columns.Count)) - 1
    block = com(lambda: ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)))
    name = com(lambda: ws.Name)
    cells = [(name, f"{L(c1 + j)}{r1 + i}", v) for i, row in enumerate(_grid(com(lambda: block.Formula2)))
             for j, v in enumerate(row) if isinstance(v, str) and v.startswith("=")]
    return fx.check(cells) if cells else {"ok": True, "checked": 0, "block": [], "warn": []}


def _text_safe(v):
    """Values are typed like the user would: numbers stay numbers, '=...' is a formula, other text is stored AS TEXT
    (Excel would otherwise turn 'Sep 26' into a date and '1/2' into a fraction)."""
    if isinstance(v, str) and not v.startswith("=") and v != "":
        return "'" + v
    return v


# ------------------------------------------------------------------ write
def write(workbook, sheet, address, values=None, formulas=None, formula=None, number_format=None,
          allow_archive=False, label=None) -> dict:
    given = [x is not None for x in (values, formulas, formula)]
    if sum(given) > 1:
        raise InputError("Give ONE of values (2D), formulas (2D) or formula (one formula filled over the range).")
    if not any(given) and number_format is None:
        raise InputError("Nothing to write: give values, formulas, formula or number_format.")
    b = find_book(workbook, write=True, allow_archive=allow_archive)
    ws = b.sheet(sheet)
    _wait_ready(b.app)
    grid = values if values is not None else formulas
    if grid is not None:
        if not grid or not isinstance(grid, list) or not isinstance(grid[0], list):
            raise InputError("values/formulas must be a 2D list: [[row1...], [row2...]].")
        width = max(len(r) for r in grid)
        grid = [list(r) + [None] * (width - len(r)) for r in grid]
        anchor = _rng(ws, address)
        rng = com(lambda: anchor.Cells(1, 1).GetResize(len(grid), width))   # late binding: Get<Prop>(args)
    else:
        rng = _rng(ws, address)
    addr = _addr(rng)
    ch = _new_change(b, ws, "write", addr, _capture(ws, rng), label=label)

    def mutate():
        if values is not None:
            com(lambda: setattr(rng, "Formula2", tuple(tuple(_text_safe(v) for v in r) for r in grid)))
        elif formulas is not None:
            com(lambda: setattr(rng, "Formula2", tuple(tuple(r) for r in grid)))
        elif formula is not None:
            com(lambda: setattr(rng, "Formula2", formula))     # relative refs adjust per cell, like a fill
        if number_format is not None:
            com(lambda: setattr(rng, "NumberFormat", number_format))
    _apply(b, ws, ch, mutate, undo=lambda: _restore(ws, addr, ch["before"]))
    ch["after"] = _state_hash(rng)
    _save_change(ch)
    out = {"change_id": ch["id"], "workbook": b.name, "sheet": ch["sheet"], "address": addr,
           "note": "Not saved. xl_revert(change_id) undoes this change (Ctrl+Z cannot).", **_notes(b)}
    try:                                   # the change is applied and recorded: reporting must not hide that
        com(lambda: ws.Calculate())
        out["formula_style"] = _check_written(ws, rng)
        out["result_preview"] = [[_clean(v) for v in row] for row in _grid(com(lambda: rng.Value2))][:20]
        _jump(b, ws, addr)
    except Exception as e:   # noqa: BLE001
        out["warning"] = f"written, but the check/preview failed: {e}"
    return out


def _notes(b: Book) -> dict:
    return {"notes": b.notes} if b.notes else {}


def rows_cols(workbook, sheet, action: str, at: str, allow_archive=False, force=False, label=None) -> dict:
    """insert_rows / delete_rows ('5:7') or insert_cols / delete_cols ('C:D')."""
    b = find_book(workbook, write=True, allow_archive=allow_archive)
    ws = b.sheet(sheet)
    _wait_ready(b.app)
    is_row = action.endswith("rows")
    spec = at.replace("$", "").strip()
    if ":" not in spec:
        spec = f"{spec}:{spec}"
    target = com(lambda: (ws.Rows if is_row else ws.Columns)(spec))
    ur = com(lambda: ws.UsedRange)
    if is_row:
        r1 = com(lambda: target.Row); n = int(com(lambda: target.Rows.Count))
        c2 = com(lambda: ur.Column) + int(com(lambda: ur.Columns.Count)) - 1
        area = f"A{r1}:{L(max(c2, 1))}{r1 + n - 1}"
    else:
        c1 = com(lambda: target.Column); n = int(com(lambda: target.Columns.Count))
        r2 = com(lambda: ur.Row) + int(com(lambda: ur.Rows.Count)) - 1
        area = f"{L(c1)}1:{L(c1 + n - 1)}{max(r2, 1)}"
    first = com(lambda: target.Row) if is_row else com(lambda: target.Column)
    if action.startswith("delete"):
        deps = _dependents_outside(b, ws, area, is_row)
        if deps and not force:
            raise InputError(f"{len(deps)} formula cell(s) depend on {spec} (e.g. {deps[:5]}); deleting turns them "
                             "into #REF! and a revert cannot repair that. Pass force=true if that is intended.")
        before = _capture(ws, com(lambda: ws.Range(area)))
        follow = _fingerprint(ws, is_row, first + n)       # the line that moves up/left into the gap
    else:
        before = None
        follow = _fingerprint(ws, is_row, first)           # the line that moves down/right
    ch = _new_change(b, ws, action, spec, before, extra={"area": area, "count": n, "first": first,
                                                         "is_row": is_row, "follow": follow}, label=label)
    _apply(b, ws, ch, lambda: com(lambda: target.Insert() if action.startswith("insert") else target.Delete()))
    ch["after"] = None
    _save_change(ch)
    _jump(b, ws, spec)
    return {"change_id": ch["id"], "workbook": b.name, "sheet": ch["sheet"], "action": action, "at": spec,
            "note": "Not saved. xl_revert(change_id) undoes this.", **_notes(b)}


def _fingerprint(ws, is_row: bool, idx: int) -> str:
    """Hash of the first cells of one row/column: proves at revert time that rows/columns did not shift since."""
    line = com(lambda: ws.Range(ws.Cells(idx, 1), ws.Cells(idx, 40)) if is_row
               else ws.Range(ws.Cells(1, idx), ws.Cells(200, idx)))
    # values, not formulas: a moved row's formulas renumber (=C4-D4 -> =C5-D5) but its values stay the same
    return hashlib.sha1(json.dumps(_grid(com(lambda: line.Value2)), default=str).encode()).hexdigest()


def _inside(area: str, sheet_cell: tuple[int, int]) -> bool:
    c1, r1, c2, r2 = range_boundaries(area.replace("$", ""))
    c, r = sheet_cell
    return (c1 or 1) <= c <= (c2 or 16384) and (r1 or 1) <= r <= (r2 or 1048576)


def _dependents_outside(b: Book, ws, area: str, rows_mode: bool) -> list[str]:
    """Formula cells outside the area that would become #REF! if it is deleted. Excel only breaks a reference that
    lies ENTIRELY inside the deleted rows/columns; whole columns (for a row delete) and partly overlapping ranges
    (SUM(C3:C5) when row 3 goes) just shrink. Scans every sheet's formulas, own sheet included."""
    me = com(lambda: ws.Name)
    ac1, ar1, ac2, ar2 = range_boundaries(area.replace("$", ""))
    out = []
    for wso in com(lambda: b.wb.Worksheets):
        name = com(lambda: wso.Name)
        for _s, a, fo in fx.sheet_formulas(wso, com):
            if name == me:
                c, r = range_boundaries(a)[:2]
                if _inside(area, (c, r)):
                    continue
            elif me.lower() not in fo.lower():
                continue
            for ref in _refs_in(fo):
                sh, ra = _split(ref, name)
                if sh.lower() != me.lower():
                    continue
                try:
                    c1, r1, c2, r2 = range_boundaries(ra.replace("$", "").replace(".", ""))
                except ValueError:
                    continue
                if rows_mode:
                    broken = r1 is not None and r2 is not None and ar1 <= r1 and r2 <= ar2
                else:
                    broken = c1 is not None and c2 is not None and ac1 <= c1 and c2 <= ac2
                if broken:
                    out.append(f"{name}!{a}")
                    break
    return out


def add_sheet(workbook, name: str, after: str | None = None, allow_archive=False) -> dict:
    name = (name or "").strip()
    if not name or len(name) > 31 or re.search(r"[:\\/?*\[\]]", name) or name.startswith("'") or name.endswith("'"):
        raise InputError(f"'{name}' is not a valid sheet name (1-31 characters, none of : \\ / ? * [ ], no ' at the "
                         "ends).")
    b = find_book(workbook, write=True, allow_archive=allow_archive)
    _wait_ready(b.app)
    if any(str(s.Name).lower() == name.lower() for s in b.wb.Worksheets):
        raise InputError(f"A sheet named '{name}' already exists.")
    ref = b.sheet(after) if after else com(lambda: b.wb.Worksheets(b.wb.Worksheets.Count))
    ws = com(lambda: b.wb.Worksheets.Add(None, ref))   # positional: keyword After= is ignored via late-bound COM
    try:
        com(lambda: setattr(ws, "Name", name))
    except Exception:
        alerts = com(lambda: b.app.DisplayAlerts)
        com(lambda: setattr(b.app, "DisplayAlerts", False))
        try:
            com(lambda: ws.Delete())
        finally:
            com(lambda: setattr(b.app, "DisplayAlerts", alerts))
        raise
    ch = _new_change(b, ws, "add_sheet", "A1", None, extra={"name": com(lambda: ws.Name)})
    ch["state"] = "applied"
    _save_change(ch)
    _jump(b, ws, "A1")
    return {"change_id": ch["id"], "workbook": b.name, "sheet": ch["sheet"], "note": "Not saved.", **_notes(b)}


# ------------------------------------------------------------------ TN formatting
def _tn_style(style: str) -> str:
    from . import styling
    names = set(styling.tn.style_defs())
    if style in names:
        return style
    kind, _, emph = style.lower().partition(".")
    if kind in styling.tn.KIND_NAME and emph in TN_EMPH:
        return styling.tn.style_name(kind, TN_EMPH[emph])
    for n in names:                                    # 'tittel' -> 'TN Tittel'
        if n.lower() == f"tn {style.lower()}":
            return n
    raise InputError(f"Unknown TN style '{style}'. Use a kind[.emphasis] ({sorted(styling.tn.KIND_NAME)} x "
                     f"{sorted(k for k in TN_EMPH if k)}) or a style name: {sorted(names)}")


def format_range(workbook, sheet, address, style=None, number_format=None, column_width=None,
                 allow_archive=False, label=None) -> dict:
    from . import styling
    b = find_book(workbook, write=True, allow_archive=allow_archive)
    ws = b.sheet(sheet)
    _wait_ready(b.app)
    rng = _rng(ws, address)
    addr = _addr(rng)
    st = _tn_style(style) if style else None
    ch = _new_change(b, ws, "format", addr, _capture(ws, rng), label=label)

    def mutate():
        if st:
            styling.ensure_styles(b.wb)
            com(lambda: setattr(rng, "Style", st))
        if number_format is not None:
            com(lambda: setattr(rng, "NumberFormat", number_format))
        if column_width is not None:
            com(lambda: setattr(rng.EntireColumn, "ColumnWidth", float(column_width)))
    _apply(b, ws, ch, mutate, undo=lambda: _restore(ws, addr, ch["before"]))
    ch["after"] = _state_hash(rng)
    _save_change(ch)
    _jump(b, ws, addr)
    return {"change_id": ch["id"], "workbook": b.name, "sheet": ch["sheet"], "address": addr, "style": st,
            "note": "Not saved. xl_revert(change_id) undoes this."}


def _expand_rows(spec: dict) -> dict[int, str]:
    out = {}
    for k, kind in (spec or {}).items():
        a, _, z = str(k).partition("-")
        for r in range(int(a), int(z or a) + 1):
            out[r] = kind
    return out


def format_sheet(workbook, sheet, last_col: str, rows: dict, label_cols=None, num: dict | None = None,
                 col_widths: dict | None = None, freeze_rows: int | None = None, allow_archive=False) -> dict:
    """Whole sheet in TN report style (tn-rapportstil, same engine as the smartview answers). rows: {"1": "title",
    "2": "subtitle", "4": "header", "5-12": "line", "13": "sum", "15": "grand", "16": "note", "3": "section" ...};
    rows with content that are not listed become 'line', empty ones 'blank'. num: {"C": "heltall", "F": "prosent1"}
    number kind per value column (default: numbers -> heltall)."""
    from . import styling
    b = find_book(workbook, write=True, allow_archive=allow_archive)
    ws = b.sheet(sheet)
    _wait_ready(b.app)
    label_cols = label_cols or ["B"]
    kinds = _expand_rows(rows)
    bad = {k for k in kinds.values() if k not in ("title", "subtitle", "section", "text", "header", "line", "sum",
                                                  "grand", "note", "blank")}
    if bad:
        raise InputError(f"Unknown row kinds {sorted(bad)}; use title, subtitle, section, text, header, line, sum, "
                         "grand, note, blank.")
    ur = com(lambda: ws.UsedRange)
    last_row = max([com(lambda: ur.Row) + int(com(lambda: ur.Rows.Count)) - 1] + list(kinds))
    area = f"A1:{last_col}{last_row}"
    ch = _new_change(b, ws, "format_sheet", area, _capture(ws, com(lambda: ws.Range(area))))
    vals = _grid(com(lambda: ws.Range(area).Value2))
    first_val_col = max(range_boundaries(f"{c}1")[0] for c in label_cols) + 1
    value_cols = [L(c) for c in range(first_val_col, range_boundaries(f"{last_col}1")[0] + 1)]
    meta = []
    for r in range(1, last_row + 1):
        kind = kinds.get(r) or ("line" if any(v not in (None, "") for v in vals[r - 1]) else "blank")
        m = {"row": r, "kind": kind}
        if kind == "header":
            m["num_cols"] = value_cols
        if kind in ("line", "sum", "grand") and num:
            m["num"] = dict(num)
        meta.append(m)
    roles = {"last_col": last_col, "label_cols": label_cols, "rows": meta,
             "freeze_rows": freeze_rows if freeze_rows is not None else next((m["row"] for m in meta
                                                                             if m["kind"] == "header"), 0),
             "col_widths": col_widths or {"A": 2.5}}
    lint: list = []

    def mutate():
        lint.extend(styling.apply_report(ws, roles) or [])
    _apply(b, ws, ch, mutate, undo=lambda: _restore(ws, area, ch["before"]))
    ch["after"] = _state_hash(com(lambda: ws.Range(area)))
    _save_change(ch)
    _jump(b, ws, "A1")
    return {"change_id": ch["id"], "workbook": b.name, "sheet": ch["sheet"], "area": area, "style_lint": lint,
            "note": "Not saved. xl_revert(change_id) restores the previous look (freeze panes / page setup stay)."}


# ------------------------------------------------------------------ audit
def _refs_in(formula: str) -> list[str]:
    mod = fx._load_checker()
    return [m.group(0) for m in mod.REF_RE.finditer(_strip_strings(formula))]


def _strip_strings(f: str) -> str:
    return re.sub(r'"(?:[^"]|"")*"', '""', f)


def _split(ref: str, default_sheet: str) -> tuple[str, str]:
    if "!" in ref:
        sh, a = ref.rsplit("!", 1)
        sh = sh.strip("'").replace("''", "'")
        sh = re.sub(r"^\[[^\]]+\]", "", sh)
        return sh, a
    return default_sheet, ref


def trace(workbook, sheet, cell: str, direction: str = "precedents", max_items: int = 40) -> dict:
    b = find_book(workbook)
    ws = b.sheet(sheet)
    c = _rng(ws, cell)
    me = com(lambda: ws.Name)
    f = com(lambda: c.Formula2)
    out = {"workbook": b.name, "sheet": me, "cell": _addr(c), "formula": f if str(f).startswith("=") else None,
           "value": _clean(com(lambda: c.Value2)), **_label_context(ws, com(lambda: c.Row), com(lambda: c.Column))}
    if direction == "precedents":
        if not out["formula"]:
            out["precedents"] = []
            return out
        items = []
        for ref in dict.fromkeys(_refs_in(f)):
            if "[" in ref.split("!")[0]:
                items.append({"ref": ref, "external": True, "note": "reference into another workbook"})
                continue
            sh, a = _split(ref, me)
            try:
                wsr = b.sheet(sh)
                a_clean = a.replace(".", "")              # trim refs (A:.A) -> plain for reading
                rr = com(lambda: wsr.Range(a_clean))
                n = _ncells(rr)
                item = {"ref": ref, "sheet": sh, "address": _addr(rr), "cells": n}
                if n == 1:
                    item["value"] = _clean(com(lambda: rr.Value2))
                    rf = com(lambda: rr.Formula2)
                    if isinstance(rf, str) and rf.startswith("="):
                        item["formula"] = rf
                    item.update(_label_context(wsr, com(lambda: rr.Row), com(lambda: rr.Column)))
                elif n <= 20_000:
                    item["sum"] = _clean(com(lambda: b.app.WorksheetFunction.Sum(rr)))
                items.append(item)
            except Exception as e:   # noqa: BLE001
                items.append({"ref": ref, "error": str(e)[:120]})
            if len(items) >= max_items:
                break
        out["precedents"] = items
        return out
    # dependents: same sheet via Excel, other sheets by scanning their formulas for a reference covering the cell
    deps = []
    try:
        dd = com(lambda: c.DirectDependents, limit=10)
        for a in re.split(r"[,;]", _addr(dd)):
            deps.append({"sheet": me, "address": a})
    except pywintypes.com_error:
        pass
    row, col = com(lambda: c.Row), com(lambda: c.Column)
    for wso in com(lambda: b.wb.Worksheets):
        name = com(lambda: wso.Name)
        if name == me:
            continue
        for _s, a, fo in fx.sheet_formulas(wso, com):
            if me.lower() not in fo.lower():
                continue
            for ref in _refs_in(fo):
                sh, ra = _split(ref, name)
                if sh.lower() != me.lower():
                    continue
                try:
                    c1, r1, c2, r2 = range_boundaries(ra.replace("$", "").replace(".", ""))
                except ValueError:
                    continue
                if (c1 or 1) <= col <= (c2 or 16384) and (r1 or 1) <= row <= (r2 or 1048576):
                    deps.append({"sheet": name, "address": a, "formula": fo[:160]})
                    break
            if len(deps) >= max_items:
                break
    out["dependents"] = deps[:max_items]
    return out


def check(workbook, sheet=None) -> dict:
    b = find_book(workbook)
    sheets = [b.sheet(sheet)] if sheet else list(com(lambda: b.wb.Worksheets))
    cells = []
    for ws in sheets:
        cells += fx.sheet_formulas(ws, com)
    res = fx.check(cells)
    res["workbook"] = b.name
    res["sheets"] = [com(lambda: s.Name) for s in sheets]
    return res


# ------------------------------------------------------------------ changes / revert
def _load(cid: str) -> dict:
    p = SNAP_DIR / f"{cid}.json"
    if not p.exists():
        raise InputError(f"No change '{cid}' (snapshots are kept {KEEP_DAYS} days). See xl_changes.")
    return json.loads(p.read_text(encoding="utf-8"))


def _same_book(workbook: str, ch: dict) -> bool:
    """workbook (a name, with or without extension, or a full path) names the change's workbook. A local path matches
    the OneDrive/SharePoint URL Excel reports as FullName for a synced file."""
    g, full = workbook.strip(), str(ch["full_name"])
    if g.lower() == full.lower():
        return True
    if re.search(r"[\\/]", g):
        def norm(x: str) -> str:
            return os.path.normcase(os.path.abspath(local_path(x)))
        return re.search(r"[\\/]", full) is not None and norm(g) == norm(full)

    def stem(x: str) -> str:
        return re.sub(r"\.(xlsx|xlsm|xlsb|xls)$", "", x.lower())
    return stem(g) == stem(str(ch["workbook"]))


def changes(workbook: str | None = None, limit: int = 20) -> list[dict]:
    if not SNAP_DIR.exists():
        return []
    out = []
    metas = {p.stem: p for p in SNAP_DIR.glob("*.meta")}
    files = {p.stem: metas.get(p.stem, p) for p in SNAP_DIR.glob("*.json")}
    for _k, p in sorted(files.items(), reverse=True):
        try:
            ch = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if workbook and not _same_book(workbook, ch):
            continue
        out.append({k: ch.get(k) for k in ("id", "ts", "workbook", "sheet", "op", "address", "label", "reverted")})
        if len(out) >= limit:
            break
    return out


def _has_content(app, ws) -> bool:
    """Values, formulas or shapes (charts, pictures) anywhere on the sheet. Not UsedRange: formats left by reverted
    changes keep it large."""
    return int(com(lambda: app.WorksheetFunction.CountA(ws.Cells))) > 0 or int(com(lambda: ws.Shapes.Count)) > 0


def revert(change_id: str | None, workbook: str | None, force: bool = False) -> dict:
    if not change_id:
        b0 = find_book(workbook)
        mine = [c for c in changes(b0.full_name, limit=200) if not c["reverted"]]
        if not mine:
            raise InputError(f"No change to revert for '{b0.name}'.")
        change_id = mine[0]["id"]
    ch = _load(change_id)
    if ch["reverted"]:
        raise InputError(f"Change {change_id} was already reverted.")
    b = find_book(ch["full_name"], write=True, allow_archive=True)    # reverting is always allowed
    _wait_ready(b.app)
    if not re.search(r"[\\/]", ch["full_name"]) and ch.get("pid") and ch["pid"] != b.pid and not force:
        raise InputError(f"'{ch['workbook']}' was never saved and this is another Excel than the one the change was "
                         "made in - probably a different workbook with the same name. Not reverted. Check the "
                         "workbook (xl_changes lists the changes per workbook).")
    later = [c for c in changes(ch["full_name"], limit=200) if c["id"] > change_id and not c["reverted"]
             and c["sheet"] == ch["sheet"]]
    if later and not force:
        raise InputError(f"Later changes on '{ch['sheet']}' are not reverted yet: {[c['id'] for c in later]}. "
                         "Revert them first (newest first), or pass force=true.")
    op = ch["op"]
    if op == "add_sheet":
        ws = b.sheet(ch["extra"]["name"])
        if _has_content(b.app, ws) and not force:
            raise InputError(f"Sheet '{ch['extra']['name']}' has content now; pass force=true to delete it anyway.")
        alerts = com(lambda: b.app.DisplayAlerts)
        com(lambda: setattr(b.app, "DisplayAlerts", False))
        try:
            com(lambda: ws.Delete())
        finally:
            com(lambda: setattr(b.app, "DisplayAlerts", alerts))
    else:
        if not any(str(x.Name) == ch["sheet"] for x in com(lambda: b.wb.Worksheets)):
            ch["reverted"] = "obsolete: sheet no longer exists"
            _save_change(ch)
            return {"reverted": change_id, "obsolete": True, "note": f"Sheet '{ch['sheet']}' no longer exists; "
                    "nothing to restore (marked obsolete)."}
        ws = b.sheet(ch["sheet"])
        with fast_excel(b.app):
            if op in ("write", "format", "format_sheet"):
                rng = com(lambda: ws.Range(ch["address"]))
                if ch.get("after") and _state_hash(rng) != ch["after"] and not force:
                    raise InputError(f"{ch['sheet']}!{ch['address']} was edited after change {change_id}; a revert "
                                     "would overwrite those edits. Pass force=true to revert anyway.")
                _restore(ws, ch["address"], ch["before"])
            elif op.startswith("insert"):
                spec, ex = ch["address"], ch["extra"]
                lines = com(lambda: (ws.Rows if op.endswith("rows") else ws.Columns)(spec))
                if not force:
                    if com(lambda: b.app.WorksheetFunction.CountA(lines)):
                        raise InputError(f"The inserted {spec} now has content; reverting would delete it. Pass "
                                         "force=true to delete anyway.")
                    if ex.get("follow") and _fingerprint(ws, ex["is_row"], ex["first"] + ex["count"]) != ex["follow"]:
                        raise InputError(f"Rows/columns around {spec} moved since the insert; reverting could delete "
                                         "the wrong lines. Pass force=true to revert anyway.")
                com(lambda: lines.Delete())
            elif op.startswith("delete"):
                spec, ex = ch["address"], ch["extra"]
                if not force and ex.get("follow") and _fingerprint(ws, ex["is_row"], ex["first"]) != ex["follow"]:
                    raise InputError(f"Rows/columns around {spec} moved since the delete; restoring could put the "
                                     "content in the wrong place. Pass force=true to revert anyway.")
                com(lambda: (ws.Rows if op.endswith("rows") else ws.Columns)(spec).Insert())
                _restore(ws, ex["area"], ch["before"])
            else:
                raise InputError(f"Cannot revert operation '{op}'.")
        com(lambda: ws.Calculate())
        _jump(b, ws, ch["address"] if op in ("write", "format") else "A1")
    ch["reverted"] = dt.datetime.now().isoformat(timespec="seconds")
    _save_change(ch)
    return {"reverted": change_id, "workbook": b.name, "sheet": ch["sheet"], "op": op, "address": ch["address"],
            "note": "Not saved."}
