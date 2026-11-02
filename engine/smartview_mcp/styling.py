"""Styling adapter: TN report style on workbooks that are ALREADY OPEN in Excel (live pywin32 COM objects).

The style definitions (fonts, colours, named cell styles ``TN ...``, number formats, locale translation and the
Arial text-width / wrap calibration) are reused from the ``tn-rapportstil`` skill, through the pinned copy in
``smartview_mcp/vendor/tn_rapportstil.py``; this
module only adds the live-workbook plumbing: COM retry, row-role based application, a COM-based lint, and a
non-destructive look for Smart View ad hoc grids.

Public API
    ensure_styles(wb)                                   create/refresh the TN named styles (idempotent)
    apply_report(ws, roles) -> list[str]                style an answer sheet (contents untouched), returns lint findings
    apply_data_sheet(ws, header_rows, label_cols, fte_rows=None)   financial look for a live Smart View grid

Nothing here saves a workbook, attaches to a running Excel or clears cell contents.
"""
from __future__ import annotations

import time
from contextlib import contextmanager
from pathlib import Path

# ---------------------------------------------------------------- reuse the skill's definitions (pinned copy)
from .vendor import tn_rapportstil as tn

# ---------------------------------------------------------------- COM retry
_BUSY = {-2147418111, -2147417846}      # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER
RETRY_LIMIT_S = 30.0


VBA_E_IGNORE = -2146777998               # 0x800AC472: Excel busy (often while recalculating), inside DISP_E_EXCEPTION


def _hresult(exc) -> int | None:
    hr = getattr(exc, "hresult", None)
    if hr is None and getattr(exc, "args", None) and isinstance(exc.args[0], int):
        hr = exc.args[0]
    return hr


def _is_busy(exc) -> bool:
    hr = _hresult(exc)
    info = getattr(exc, "excepinfo", None)
    inner = info[5] if info and len(info) > 5 else None
    return hr in _BUSY or VBA_E_IGNORE in (hr, inner)


def _retry(fn, limit_s: float = RETRY_LIMIT_S):
    """Call fn(); when Excel rejects the call because it is busy, back off (0.2 s doubling to 3 s) up to limit_s."""
    t0, delay = time.monotonic(), 0.2
    while True:
        try:
            return fn()
        except Exception as exc:   # noqa: BLE001
            if not _is_busy(exc):
                raise
            if time.monotonic() - t0 > limit_s:   # same readable message as userexcel.com, not a raw COM code
                from .host import HostError
                raise HostError("Excel is busy - probably a cell is being edited or a dialog is open. Press Enter "
                                "or Esc in Excel and ask again.") from None
            time.sleep(delay)
            delay = min(delay * 2, 3.0)


def _get(obj, *path):
    for p in path:
        obj = _retry(lambda o=obj, p=p: getattr(o, p))
    return obj


def _set(obj, name, value):
    _retry(lambda: setattr(obj, name, value))


def _call(obj, name, *args):
    return _retry(lambda: getattr(obj, name)(*args))


# ---------------------------------------------------------------- small helpers
def _ci(letter: str) -> int:
    n = 0
    for ch in letter.upper():
        n = n * 26 + ord(ch) - 64
    return n


def _cl(i: int) -> str:
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _rng(ws, r1, c1, r2=None, c2=None):
    r2 = r1 if r2 is None else r2
    c2 = c1 if c2 is None else c2
    return _retry(lambda: ws.Range(f"{_cl(c1)}{r1}:{_cl(c2)}{r2}"))


def _rgb(c):
    return tn.rgb(c)


def _check_ranges(cells: list[tuple[int, int]]) -> list[str]:
    """The check status cells as runs of consecutive rows per column ('D10:D12'), so the check colours are set once
    per run instead of once per cell."""
    out: list[list[int]] = []                       # [col, first row, last row]
    for r, c in sorted(set(cells), key=lambda x: (x[1], x[0])):
        if out and out[-1][0] == c and out[-1][2] == r - 1:
            out[-1][2] = r
        else:
            out.append([c, r, r])
    return [f"{_cl(c)}{a}" if a == b else f"{_cl(c)}{a}:{_cl(c)}{b}" for c, a, b in out]


GREEN = (0, 112, 60)
GREY_TAB = (166, 166, 166)
NUM_ALIAS = {"heltall": "heltall", "tall1": "tall1", "prosent1": "prosent1", "prosent": "prosent",
             "faktor": "faktor", "belop": "belop"}
TALL1_THOUSANDS = '#,##0.0_);(#,##0.0);"–"_)'    # spec: 'tall1' is one decimal WITH thousands separator
HELTALL = tn.NUM["heltall"]


def _seps(xl):
    tup = _retry(lambda: xl.International)
    return tup[2], tup[3], tup[4]     # decimal, thousands, list separator


def _style_for(kind: str, emph: str) -> str:
    if kind not in NUM_ALIAS:
        raise ValueError(f"unknown number kind {kind!r}; use one of {sorted(NUM_ALIAS)}")
    return tn.style_name(kind, emph)


# ---------------------------------------------------------------- ensure_styles
def _styles_version() -> str:
    import hashlib
    import json
    return hashlib.sha1(json.dumps(tn.style_defs(), sort_keys=True, default=str).encode()).hexdigest()[:10]


def ensure_styles(wb) -> None:
    """Install the TN named styles once per workbook. Rewriting them (~25 styles x ~25 properties) makes Excel restyle
    every cell using them, in every open workbook - measured 25-160 s per answer. A hidden workbook name records the
    installed version; the styles are rewritten only when missing or when tn_rapportstil changes them."""
    ver = _styles_version()
    if has_styles(wb, ver):
        return
    tpl = template_path(ver)
    if tpl.exists():                              # one Styles.Merge from the template instead of ~600 property sets
        try:
            xl = _get(wb, "Application")
            # Merge asks "Merge styles that have the same names?" in the user's (interactive) Excel and blocks until
            # someone clicks - alerts off answers Yes (= take the TN definitions), then the user's setting returns.
            alerts = _get(xl, "DisplayAlerts")
            _set(xl, "DisplayAlerts", False)
            try:
                src = _retry(lambda: xl.Workbooks.Open(str(tpl), UpdateLinks=0, ReadOnly=True, AddToMru=False))
                try:
                    _retry(lambda: wb.Styles.Merge(src))
                finally:
                    _retry(lambda: src.Close(False))
            finally:
                _set(xl, "DisplayAlerts", alerts if alerts is not None else True)
            _mark(wb, ver)
            return
        except Exception:   # noqa: BLE001 - fall back to writing the styles one by one
            pass
    _ensure_styles(wb)
    _mark(wb, ver)


def has_styles(wb, ver: str | None = None) -> bool:
    try:
        return str(_retry(lambda: wb.Names("_tnStylesVer").RefersTo)).strip('="') == (ver or _styles_version())
    except Exception:   # noqa: BLE001 - name missing
        return False


def _mark(wb, ver: str) -> None:
    try:
        _retry(lambda: wb.Names.Add(Name="_tnStylesVer", RefersTo=f'="{ver}"', Visible=False))
    except Exception:   # noqa: BLE001 - only an optimisation
        pass


def template_path(ver: str | None = None) -> Path:
    from . import config
    return config.SVKIT / f"tn_template_{ver or _styles_version()}.xltx"


def ensure_template(xl) -> Path:
    """An Excel template (.xltx) holding the TN styles + version marker: new analysis workbooks are created from it
    (styles for free) and existing workbooks merge their styles from it. Built once per style version (~10 s)."""
    from . import config
    ver = _styles_version()
    path = template_path(ver)
    if path.exists():
        return path
    for old in config.SVKIT.glob("tn_template_*.xltx"):
        try:
            old.unlink()
        except OSError:
            pass
    wb = _retry(lambda: xl.Workbooks.Add())
    try:
        _ensure_styles(wb)
        _mark(wb, ver)
        from .ops import _label
        _retry(lambda: _label(wb))
        _retry(lambda: wb.SaveAs(str(path), 54))           # xlOpenXMLTemplate
    finally:
        _retry(lambda: wb.Close(False))
    return path


def _ensure_styles(wb) -> None:
    """Create/refresh all TN named styles in this workbook (idempotent). Reuses tn_rapportstil.ensure_styles."""
    xl = _get(wb, "Application")
    _retry(lambda: tn.ensure_styles(xl, wb))


# ---------------------------------------------------------------- view helpers
@contextmanager
def _active(ws):
    """Make ws the active sheet (needed for window-level settings), restore the previous workbook/sheet after."""
    wb = _get(ws, "Parent")
    xl = _get(wb, "Application")
    prev_wb = prev_ws = None
    try:
        prev_wb, prev_ws = xl.ActiveWorkbook, xl.ActiveSheet
    except Exception:   # noqa: BLE001
        pass
    ok = True
    try:
        if prev_wb is None or _get(prev_wb, "FullName") != _get(wb, "FullName"):
            _call(wb, "Activate")
        _call(ws, "Activate")
    except Exception:   # noqa: BLE001
        ok = False
    try:
        yield ok
    finally:
        try:
            if prev_wb is not None and _get(prev_wb, "FullName") != _get(wb, "FullName"):
                _call(prev_wb, "Activate")
            if prev_ws is not None:
                _call(prev_ws, "Activate")
        except Exception:   # noqa: BLE001
            pass


def _visible_window(wb):
    try:
        win = _get(wb, "Windows")(1)
        return win if _get(win, "Visible") else None
    except Exception:   # noqa: BLE001
        return None


def _freeze(ws, rows: int, cols: int, gridlines_off: bool = True) -> bool:
    """Freeze panes below `rows` and right of `cols` and hide gridlines; only if the workbook window is visible."""
    wb = _get(ws, "Parent")
    win = _visible_window(wb)
    if win is None:
        return False
    with _active(ws) as ok:
        if not ok:
            return False
        try:
            _set(win, "FreezePanes", False)
            _set(win, "ScrollRow", 1)
            _set(win, "ScrollColumn", 1)
            if gridlines_off:
                _set(win, "DisplayGridlines", False)
            if rows or cols:
                _set(win, "SplitColumn", cols)
                _set(win, "SplitRow", rows)
                _set(win, "FreezePanes", True)
            return True
        except Exception:   # noqa: BLE001   (e.g. window has no usable client area) - leave no half-split behind
            try:
                _set(win, "FreezePanes", False); _set(win, "SplitColumn", 0); _set(win, "SplitRow", 0)
            except Exception:   # noqa: BLE001
                pass
            return False


# ---------------------------------------------------------------- apply_report
def _runs(styles: dict[int, str]):
    """{col: style} -> [(style, c1, c2)] merging adjacent equal styles."""
    out = []
    for c in sorted(styles):
        if out and out[-1][0] == styles[c] and out[-1][2] == c - 1:
            out[-1] = (out[-1][0], out[-1][1], c)
        else:
            out.append((styles[c], c, c))
    return out


def _parse_widths(spec: dict) -> dict[int, float]:
    w = {}
    for key, val in spec.items():
        a, _, b = key.partition(":")
        for c in range(_ci(a), _ci(b or a) + 1):
            w[c] = float(val)
    return w


def apply_report(ws, roles: dict) -> list[str]:
    """Style an answer sheet whose content is already written (never clears contents). Returns lint findings."""
    from .timing import stage
    with stage("A ensure styles"):
        wb = _get(ws, "Parent")
        xl = _get(wb, "Application")
        ensure_styles(wb)
    dec, thou, sep = _seps(xl)
    last = _ci(roles["last_col"])
    labels = sorted(_ci(c) for c in roles.get("label_cols", ["B"]))
    first_label = labels[0]
    widths = {c: 8.43 for c in range(1, last + 1)}
    widths.update({c: w for c, w in _parse_widths(roles.get("col_widths", {})).items() if c <= last})
    rows = {int(r["row"]): r for r in roles["rows"]}
    ur = _get(ws, "UsedRange")
    if roles.get("last_row"):              # the tool-owned block: never re-style the user's rows below it
        last_row = max(max(rows), int(roles["last_row"]))
    else:
        last_row = max(max(rows), _get(ur, "Row") + _get(ur, "Rows", "Count") - 1)

    raw = _get(_rng(ws, 1, 1, last_row, last), "Value2")
    if not isinstance(raw, tuple):
        raw = ((raw,),)
    def val(r, c):
        return raw[r - 1][c - 1]

    # clear stale formats (keeps contents), then base style
    with stage("A2 clear + base"):
        area = _rng(ws, 1, 1, last_row, last)
        _call(area, "ClearFormats")
        _set(area, "Style", "TN Tom")

    by_style: dict[str, list[str]] = {}
    tall1_cells: list[str] = []
    heights: dict[int, float] = {}
    merges: list[tuple[int, int]] = []
    checks: list[tuple[int, int]] = []
    assigned: dict[tuple[int, int], str] = {}             # (row, col) -> TN style set below (for the fast lint)

    for r in range(1, last_row + 1):
        spec = rows.get(r, {"kind": "blank"})
        kind = spec["kind"]
        cols: dict[int, str] = {}
        num = {_ci(k): v for k, v in (spec.get("num") or {}).items()}
        if kind in ("title", "subtitle", "section", "note", "blank"):
            name = {"title": "TN Tittel", "subtitle": "TN Undertittel", "section": "TN Seksjon",
                    "note": "TN Note", "blank": "TN Tom"}[kind]
            cols = {c: name for c in range(1, last + 1)}
            heights[r] = {"title": 28, "subtitle": 16, "section": 22}.get(kind, 15)
        elif kind == "text":
            cols = {c: "TN Tekst" if c >= first_label else "TN Tom" for c in range(1, last + 1)}
            width = sum(widths[c] for c in range(first_label, last + 1))
            txt = val(r, first_label)
            n = (tn.wrap_lines(str(txt), width) or 4) if txt else 1
            heights[r] = max(15, n * 13 + 2)
            if all(val(r, c) in (None, "") for c in range(first_label + 1, last + 1)):
                merges.append((r, first_label))
        elif kind == "header":
            nc = {_ci(c) for c in spec.get("num_cols", [])} if "num_cols" in spec else \
                set(range(labels[-1] + 1, last + 1))
            cols = {c: "TN Kolonne tall" if c in nc else "TN Kolonne" for c in range(1, last + 1)}
            lines = 1
            for c in range(1, last + 1):
                t = val(r, c)
                if t not in (None, ""):
                    lines = max(lines, tn.wrap_lines(str(t), widths[c], 9, True) or 4)
            heights[r] = max(26, lines * 12 + 3)
        elif kind in ("line", "sum", "grand", "check"):
            emph = {"line": "", "check": "", "sum": "Sum", "grand": "Hovedsum"}[kind]
            status = _ci(spec["status_col"]) if spec.get("status_col") else None
            for c in range(1, last + 1):
                if c < first_label:
                    cols[c] = "TN Tom"; continue
                v = val(r, c)
                if c in num:
                    k = num[c]
                    cols[c] = _style_for(k, emph)
                    if k == "tall1":
                        tall1_cells.append(f"{_cl(c)}{r}")
                elif c == status:
                    cols[c] = "TN Tekst tallkolonne" if not emph else tn.style_name("etikett", emph)
                elif c in labels:
                    cols[c] = tn.style_name("etikett", emph)
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    cols[c] = _style_for("heltall", emph)
                else:
                    cols[c] = tn.style_name("etikett", emph)
            heights[r] = 15
            if status:
                checks.append((r, status))
        else:
            raise ValueError(f"row {r}: unknown kind {kind!r}")
        for st, c1, c2 in _runs(cols):
            by_style.setdefault(st, []).append(f"{_cl(c1)}{r}:{_cl(c2)}{r}")
        for c, st in cols.items():
            assigned[(r, c)] = st

    with stage("B cell styles"):
        for st, addrs in by_style.items():
            for chunk in tn._chunks(addrs, sep):
                _set(_retry(lambda: ws.Range(chunk)), "Style", st)
        if tall1_cells:
            fmt = tn.to_local(TALL1_THOUSANDS, dec, thou)
            for chunk in tn._chunks(tall1_cells, sep):
                _set(_retry(lambda: ws.Range(chunk)), "NumberFormatLocal", fmt)

    with stage("C widths + heights"):
        for c, w in widths.items():
            if c in _parse_widths(roles.get("col_widths", {})):
                _set(_retry(lambda: ws.Columns(c)), "ColumnWidth", w)
        by_height: dict[float, list[str]] = {}            # one call per distinct height, not one per row
        for r, h in sorted(heights.items()):
            runs = by_height.setdefault(h, [])
            if runs and runs[-1][1] == r - 1:
                runs[-1] = (runs[-1][0], r)
            else:
                runs.append((r, r))
        for h, runs in by_height.items():
            for chunk in tn._chunks([f"{a}:{b}" for a, b in runs], sep):
                _set(_retry(lambda: ws.Range(chunk)), "RowHeight", h)
    with stage("D merges"):
        for r, c in merges:
            _retry(lambda: _set(_rng(ws, r, c, r, last), "WrapText", True))
            _set(_rng(ws, r, c, r, last), "MergeCells", True)
            _set(_rng(ws, r, c, r, last), "Style", "TN Tekst")      # merge resets the hidden cells to Normal
            _set(_rng(ws, r, c), "VerticalAlignment", tn.V_TOP)
    with stage("E check colours"):
        for chunk in tn._chunks(_check_ranges(checks), sep):          # one rule pair per chunk (was per cell: 7 s)
            cell = _retry(lambda: ws.Range(chunk))
            _call(_get(cell, "FormatConditions"), "Delete")
            fc = _call(_get(cell, "FormatConditions"), "Add", 1, 3, '="OK"')      # cell value = "OK"
            _set(_get(fc, "Font"), "Color", _rgb(GREEN)); _set(_get(fc, "Font"), "Bold", True)
            fc = _call(_get(cell, "FormatConditions"), "Add", 1, 4, '="OK"')      # not equal -> breach
            _set(_get(fc, "Font"), "Color", _rgb(tn.BREACH_FONT)); _set(_get(fc, "Font"), "Bold", True)
            _set(_get(fc, "Interior"), "Color", _rgb(tn.BREACH_FILL))

    with stage("F tab + freeze"):
        _set(_get(ws, "Tab"), "Color", _rgb(tn.NAVY))
        _freeze(ws, int(roles.get("freeze_rows") or 0), 0)
    with stage("G page setup"):
        _page_setup(ws, xl, last_row, last, roles)
    with stage("H lint"):
        return _lint_fast(ws, roles, raw, last, last_row, widths, assigned, merges, heights)


def _lint_fast(ws, roles, raw, last, last_row, widths, assigned, merges, heights) -> list[str]:
    """The tn-rapportstil checks that can fail here, without reading every cell back from Excel (the per-cell COM
    lint took 3-4 s): sheet-wide font and vertical borders from Excel, and text fit from the styles this module
    assigned itself (size/bold/wrap from tn.style_defs). Styles and number formats are set by our own code, so the
    per-cell 'has a TN style / has a number format' checks cannot fail and are not repeated."""
    errs: list[str] = []
    ur = _get(ws, "UsedRange")
    ur_last_c = _get(ur, "Column") + _get(ur, "Columns", "Count") - 1
    ur_last_r = _get(ur, "Row") + _get(ur, "Rows", "Count") - 1
    area = _rng(ws, 1, 1, max(ur_last_r, last_row), max(ur_last_c, last))
    fname = _get(_get(area, "Font"), "Name")
    if fname != tn.FONT:
        errs.append(f"more than one font, or not Arial (Font.Name={fname!r})")
    for edge, label in ((7, "left"), (10, "right"), (11, "inside")):
        try:
            if _get(_call(area, "Borders", edge), "LineStyle") != tn.NONE:
                errs.append(f"vertical border ({label}) on the sheet (not allowed in the report style)")
        except Exception:   # noqa: BLE001
            continue
    if ur_last_c > last:
        errs.append(f"content outside the table width (up to {roles['last_col']})")
    defs = tn.style_defs()
    merged = {(r, c) for r, c in merges}
    for r in range(1, min(len(raw), last_row) + 1):
        for c in range(1, min(len(raw[0]), last) + 1):
            v = raw[r - 1][c - 1]
            if not isinstance(v, str) or not v or v.startswith("="):
                continue
            p = defs.get(assigned.get((r, c), ""), {})
            size, bold = p.get("size", 10), bool(p.get("bold"))
            a = f"{_cl(c)}{r}"
            if (r, c) in merged:
                w = sum(widths.get(k, 8.43) for k in range(c, last + 1))
                lines = tn.wrap_lines(v, w, size, bold)
                h = heights.get(r, 15)
                if lines is None or lines * (size + 3) > h + 1:
                    errs.append(f"{a}: text does not fit ({lines or 'word too long'} lines, {h:.0f} pt)")
            elif p.get("h", -4131) != tn.H_RIGHT and not p.get("wrap"):
                avail = widths.get(c, 8.43)
                for k in range(c + 1, last + 1):
                    if raw[r - 1][k - 1] not in (None, ""):
                        break
                    avail += widths.get(k, 8.43)
                need = tn.text_units(v, size, bold)
                if need > avail:
                    errs.append(f"{a}: text is cut off or spills over (~{need:.0f} > {avail:.0f} chars)")
    return errs


def _page_setup(ws, xl, last_row, last, roles):
    # every PageSetup property normally round-trips to the printer driver (measured 14 s for 12 settings);
    # PrintCommunication=False batches them into one exchange when it is switched back on
    try:
        _set(xl, "PrintCommunication", False)
    except Exception:   # noqa: BLE001 - not available: just slower
        pass
    try:
        _page_setup_props(ws, xl, last_row, last, roles)
    finally:
        try:
            _set(xl, "PrintCommunication", True)
        except Exception:   # noqa: BLE001
            pass


def _page_setup_props(ws, xl, last_row, last, roles):
    try:
        ps = _get(ws, "PageSetup")
        _set(ps, "PrintArea", f"$A$1:${_cl(last)}${last_row}")
        _set(ps, "PrintTitleRows", "$1:$2")
        _set(ps, "Orientation", 2)
        _set(ps, "Zoom", False)
        _set(ps, "FitToPagesWide", 1)
        _set(ps, "FitToPagesTall", False)
        _set(ps, "CenterHorizontally", True)
        cm = _call(xl, "CentimetersToPoints", 1.5)
        _set(ps, "LeftMargin", cm); _set(ps, "RightMargin", cm)
        _set(ps, "CenterHeader", '&"Arial"&8' + roles.get("print_header", "Intern"))
        _set(ps, "RightFooter", '&"Arial"&8Side &P av &N')
        _set(ps, "LeftFooter", '&"Arial"&8&F')
    except Exception:   # noqa: BLE001  (no printer installed etc.)
        pass


# ---------------------------------------------------------------- apply_data_sheet
def _numeric_cells(rng):
    """Union of numeric constants and numeric formula results in rng, or None."""
    parts = []
    for kind in (2, -4123):          # xlCellTypeConstants, xlCellTypeFormulas
        try:
            parts.append(_retry(lambda k=kind: rng.SpecialCells(k, 1)))   # xlNumbers
        except Exception as exc:   # noqa: BLE001
            if _hresult(exc) in _BUSY:
                raise
    return parts


def apply_data_sheet(ws, header_rows: int, label_cols: int, fte_rows: list[int] | None = None) -> None:
    """Financial look for a LIVE Smart View ad hoc grid. Never clears/moves cells, never ClearFormats, never
    inserts rows/columns. Rows/cols are absolute sheet indexes (1-based)."""
    wb = _get(ws, "Parent")
    xl = _get(wb, "Application")
    dec, thou, _ = _seps(xl)
    ur = _get(ws, "UsedRange")
    r0, c0 = _get(ur, "Row"), _get(ur, "Column")
    last_r = r0 + _get(ur, "Rows", "Count") - 1
    last_c = c0 + _get(ur, "Columns", "Count") - 1
    last_r, last_c = max(last_r, header_rows), max(last_c, label_cols)

    _set(_get(_rng(ws, 1, 1, last_r, last_c), "Font"), "Name", tn.FONT)
    _set(_get(_rng(ws, 1, 1, last_r, last_c), "Font"), "Size", 10)

    if header_rows > 0:
        hdr = _rng(ws, 1, 1, header_rows, last_c)
        f = _get(hdr, "Font")
        _set(f, "Bold", True); _set(f, "Color", _rgb(tn.NAVY))
        if last_c > label_cols:
            _set(_rng(ws, 1, label_cols + 1, header_rows, last_c), "HorizontalAlignment", tn.H_RIGHT)
        if label_cols:
            _set(_rng(ws, 1, 1, header_rows, label_cols), "HorizontalAlignment", tn.H_LEFT)
        bd = _call(_rng(ws, header_rows, 1, header_rows, last_c), "Borders", tn.B)
        _set(bd, "LineStyle", tn.CONT); _set(bd, "Weight", tn.THIN); _set(bd, "Color", _rgb(tn.NAVY))

    if last_r > header_rows:
        if label_cols:
            _set(_rng(ws, header_rows + 1, 1, last_r, label_cols), "HorizontalAlignment", tn.H_LEFT)
        if last_c > label_cols:
            data = _rng(ws, header_rows + 1, label_cols + 1, last_r, last_c)
            f0 = tn.to_local(HELTALL, dec, thou)
            f1 = tn.to_local(TALL1_THOUSANDS, dec, thou)
            for part in _numeric_cells(data):
                _set(part, "NumberFormatLocal", f0)
                _set(part, "HorizontalAlignment", tn.H_RIGHT)
            for fr in fte_rows or []:
                if header_rows < fr <= last_r:
                    for part in _numeric_cells(_rng(ws, fr, label_cols + 1, fr, last_c)):
                        _set(part, "NumberFormatLocal", f1)

    # widths: ~46 shared by the label columns, 16 for data columns
    per_label = max(20.0, 46.0 / label_cols) if label_cols > 1 else 46.0
    for c in range(1, label_cols + 1):
        _set(_retry(lambda: ws.Columns(c)), "ColumnWidth", per_label)
    if last_c > label_cols:
        _set(_rng(ws, 1, label_cols + 1, 1, last_c).EntireColumn, "ColumnWidth", 16)

    _set(_get(ws, "Tab"), "Color", _rgb(GREY_TAB))
    _freeze(ws, header_rows, label_cols)
