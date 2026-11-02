"""Snapshots before every change, the change log, and revert (Ctrl+Z cannot undo COM writes)."""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import re
import threading
import time
import uuid

import pywintypes
from openpyxl.utils import get_column_letter as L

from . import books
from .. import config
from ..excel.instances import local_path
from ..excel.jobs import com, fast_excel
from ..excel.target import _wait_ready
from ..host import InputError

SNAP_DIR = config.SVKIT / "snapshots"
KEEP_DAYS = 7
MAX_SNAPSHOT_CELLS = 250_000
JUMP = os.environ.get("SV_XL_JUMP", "1") == "1"     # tests set 0: do not move the user's view


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
    n = books._ncells(rng)
    if n > MAX_SNAPSHOT_CELLS:
        raise InputError(f"The change covers {n} cells - too big to snapshot safely (limit {MAX_SNAPSHOT_CELLS}). "
                         "Work on a smaller range.")
    r0, c0 = com(lambda: rng.Row), com(lambda: rng.Column)
    f2 = books._grid(com(lambda: rng.Formula2))
    formulas = {f"{L(c0 + j)}{r0 + i}": v for i, row in enumerate(f2) for j, v in enumerate(row)
                if isinstance(v, str) and v.startswith("=")}
    widths = {}
    ncol = int(com(lambda: rng.Columns.Count))
    if ncol <= 200:
        for j in range(ncol):
            widths[L(c0 + j)] = com(lambda: ws.Columns(c0 + j).ColumnWidth)
    consts = _constants(f2, books._grid(com(lambda: rng.Value2)), r0, c0) if n <= MAX_STYLE_CELLS else None
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
    return hashlib.sha1(json.dumps(books._grid(com(lambda: rng.Formula2)), default=str).encode()).hexdigest()


_seq = 0
_seq_lock = threading.Lock()


def _new_change(b: books.Book, ws, op: str, address: str, before: dict | None, extra: dict | None = None,
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


def _apply(b: books.Book, ws, ch: dict, mutate, undo=None) -> None:
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
            books.log.exception("auto-rollback of %s failed", ch["id"])
            ch["state"] = "failed - partly applied, revert with force=true"
        _save_change(ch)
        raise
    ch["state"] = "applied"


def _restore(ws, address: str, snap: dict) -> None:
    rng = com(lambda: ws.Range(address))
    com(lambda: rng.SetValue(11, snap["xml"]))
    r0, c0 = com(lambda: rng.Row), com(lambda: rng.Column)
    now = books._grid(com(lambda: rng.Formula2))
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
                books.log.warning("style %s not restored on %s", name, areas[k:k + 15])
    for a, f in snap["formulas"].items():
        cell = com(lambda: ws.Range(a))
        if com(lambda: cell.Formula2) != f:
            com(lambda: setattr(cell, "Formula2", f))
    if snap.get("constants") is not None:
        for a, v in _coerced(snap["constants"], books._grid(com(lambda: rng.Value2)), r0, c0):
            cell = com(lambda: ws.Range(a))
            if v is None:
                com(lambda: cell.ClearContents())
            elif isinstance(v, str):
                com(lambda: setattr(cell, "Formula2", "'" + v))    # as text, like xl_write stores it
            else:
                com(lambda: setattr(cell, "Value2", v))
    for col, w in (snap.get("widths") or {}).items():
        com(lambda: setattr(ws.Columns(col), "ColumnWidth", w))


def _jump(b: books.Book, ws, address: str) -> None:
    """After the frozen write: show the user what changed (the agreed 'frozen, then jump')."""
    if not JUMP:
        return
    try:
        com(lambda: b.wb.Activate(), limit=10)
        com(lambda: ws.Activate(), limit=10)
        com(lambda: b.app.Goto(ws.Range(address)), limit=10)
    except Exception:   # noqa: BLE001 - cosmetic
        pass


def _fingerprint(ws, is_row: bool, idx: int) -> str:
    """Hash of the first cells of one row/column: proves at revert time that rows/columns did not shift since."""
    line = com(lambda: ws.Range(ws.Cells(idx, 1), ws.Cells(idx, 40)) if is_row
               else ws.Range(ws.Cells(1, idx), ws.Cells(200, idx)))
    # values, not formulas: a moved row's formulas renumber (=C4-D4 -> =C5-D5) but its values stay the same
    return hashlib.sha1(json.dumps(books._grid(com(lambda: line.Value2)), default=str).encode()).hexdigest()


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


def delete_sheet(app, ws) -> None:
    alerts = com(lambda: app.DisplayAlerts)
    com(lambda: setattr(app, "DisplayAlerts", False))
    try:
        com(lambda: ws.Delete())
    finally:
        com(lambda: setattr(app, "DisplayAlerts", alerts))


def revert(change_id: str | None, workbook: str | None, force: bool = False) -> dict:
    if not change_id:
        b0 = books.find_book(workbook)
        mine = [c for c in changes(b0.full_name, limit=200) if not c["reverted"]]
        if not mine:
            raise InputError(f"No change to revert for '{b0.name}'.")
        change_id = mine[0]["id"]
    ch = _load(change_id)
    if ch["reverted"]:
        raise InputError(f"Change {change_id} was already reverted.")
    b = books.find_book(ch["full_name"], write=True, allow_archive=True)    # reverting is always allowed
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
        delete_sheet(b.app, ws)
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
