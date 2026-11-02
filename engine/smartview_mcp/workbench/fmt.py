"""Format a range or sheet in TN style, trace precedents/dependents, run the formula checker."""
from __future__ import annotations

import re

import pywintypes
from openpyxl.utils import get_column_letter as L
from openpyxl.utils.cell import range_boundaries

from . import books, history
from .. import formulas as fx
from ..excel.jobs import com
from ..excel.target import _wait_ready
from ..host import InputError

TN_EMPH = {"": "", "sum": "Sum", "hovedsum": "Hovedsum", "grand": "Hovedsum", "input": "Input", "nokkel": "Nøkkel",
           "nøkkel": "Nøkkel", "key": "Nøkkel"}


# ------------------------------------------------------------------ TN formatting
def _tn_style(style: str) -> str:
    from .. import styling
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
    from .. import styling
    b = books.find_book(workbook, write=True, allow_archive=allow_archive)
    ws = b.sheet(sheet)
    _wait_ready(b.app)
    rng = books._rng(ws, address)
    addr = books._addr(rng)
    st = _tn_style(style) if style else None
    ch = history._new_change(b, ws, "format", addr, history._capture(ws, rng), label=label)

    def mutate():
        if st:
            styling.ensure_styles(b.wb)
            com(lambda: setattr(rng, "Style", st))
        if number_format is not None:
            com(lambda: setattr(rng, "NumberFormat", number_format))
        if column_width is not None:
            com(lambda: setattr(rng.EntireColumn, "ColumnWidth", float(column_width)))
    history._apply(b, ws, ch, mutate, undo=lambda: history._restore(ws, addr, ch["before"]))
    ch["after"] = history._state_hash(rng)
    history._save_change(ch)
    history._jump(b, ws, addr)
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
    from .. import styling
    b = books.find_book(workbook, write=True, allow_archive=allow_archive)
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
    ch = history._new_change(b, ws, "format_sheet", area, history._capture(ws, com(lambda: ws.Range(area))))
    vals = books._grid(com(lambda: ws.Range(area).Value2))
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
    history._apply(b, ws, ch, mutate, undo=lambda: history._restore(ws, area, ch["before"]))
    ch["after"] = history._state_hash(com(lambda: ws.Range(area)))
    history._save_change(ch)
    history._jump(b, ws, "A1")
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
    b = books.find_book(workbook)
    ws = b.sheet(sheet)
    c = books._rng(ws, cell)
    me = com(lambda: ws.Name)
    f = com(lambda: c.Formula2)
    out = {"workbook": b.name, "sheet": me, "cell": books._addr(c), "formula": f if str(f).startswith("=") else None,
           "value": books._clean(com(lambda: c.Value2)), **books._label_context(ws, com(lambda: c.Row), com(lambda: c.Column))}
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
                n = books._ncells(rr)
                item = {"ref": ref, "sheet": sh, "address": books._addr(rr), "cells": n}
                if n == 1:
                    item["value"] = books._clean(com(lambda: rr.Value2))
                    rf = com(lambda: rr.Formula2)
                    if isinstance(rf, str) and rf.startswith("="):
                        item["formula"] = rf
                    item.update(books._label_context(wsr, com(lambda: rr.Row), com(lambda: rr.Column)))
                elif n <= 20_000:
                    item["sum"] = books._clean(com(lambda: b.app.WorksheetFunction.Sum(rr)))
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
        for a in re.split(r"[,;]", books._addr(dd)):
            deps.append({"sheet": me, "address": a})
    except pywintypes.com_error:
        pass
    row, col = com(lambda: c.Row), com(lambda: c.Column)
    for name, a, fo, (c1, r1, c2, r2) in refs_into(b, me, own=False):
        if len(deps) >= max_items:
            break
        if (c1 or 1) <= col <= (c2 or 16384) and (r1 or 1) <= row <= (r2 or 1048576) \
                and not any(d["sheet"] == name and d["address"] == a for d in deps):
            deps.append({"sheet": name, "address": a, "formula": fo[:160]})
    out["dependents"] = deps[:max_items]
    return out


def refs_into(b: books.Book, me: str, own: bool):
    """(sheet, cell, formula, (c1, r1, c2, r2)) per reference to sheet `me` in the workbook's formulas; own: also
    the formulas on `me` (whose local references carry no sheet name)."""
    for wso in com(lambda: b.wb.Worksheets):
        name = com(lambda: wso.Name)
        if name == me and not own:
            continue
        for _s, a, fo in fx.sheet_formulas(wso, com):
            if name != me and me.lower() not in fo.lower():
                continue
            for ref in _refs_in(fo):
                sh, ra = _split(ref, name)
                if sh.lower() == me.lower():
                    try:
                        yield name, a, fo, range_boundaries(ra.replace("$", "").replace(".", ""))
                    except ValueError:
                        continue


def check(workbook, sheet=None) -> dict:
    b = books.find_book(workbook)
    sheets = [b.sheet(sheet)] if sheet else list(com(lambda: b.wb.Worksheets))
    cells = []
    for ws in sheets:
        cells += fx.sheet_formulas(ws, com)
    res = fx.check(cells)
    res["workbook"] = b.name
    res["sheets"] = [com(lambda: s.Name) for s in sheets]
    return res
