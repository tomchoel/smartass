"""Styling adapter: TN report style on workbooks that are ALREADY OPEN in Excel (live pywin32 COM objects).

The style definitions (fonts, colours, named cell styles ``TN ...``, number formats, locale translation and the
Arial text-width / wrap calibration) are reused from the ``tn-rapportstil`` skill, through the pinned copy in
``smartass/xl/vendor/tn_rapportstil.py``; this
module only adds the live-workbook plumbing: COM retry, row-role based application, a COM-based lint, and a
non-destructive look for Smart View ad hoc grids.

Public API
    ensure_styles(wb)                                   create/refresh the TN named styles (idempotent)
    apply_report(ws, roles) -> list[str]                style an answer sheet (contents untouched), returns lint findings
    apply_data_sheet(ws, header_rows, label_cols, fte_rows=None)   financial look for a live Smart View grid

Nothing here saves a workbook, attaches to a running Excel or clears cell contents.
"""
from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path

# ---------------------------------------------------------------- reuse the skill's definitions (pinned copy)
from .vendor import tn_rapportstil as tn

# ---------------------------------------------------------------- COM retry
RETRY_LIMIT_S = 30.0


def _retry(fn, limit_s: float = RETRY_LIMIT_S):
    """Call fn(); while Excel is busy it is retried up to limit_s, then a readable HostError (as jobs.com)."""
    from . import jobs
    return jobs.com(fn, limit_s, check=False)


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
             "faktor": "faktor", "belop": "belop",
             "avvik": "heltall", "avvikpst": "prosent1", "markor": "heltall"}    # new kinds: base TN style
# the scenario kinds (sub-spec 5 §1.5) on top of their base style, as en-US codes set through NumberFormat (not
# Local, so no locale translation): Δ NOK, Δ % and the Material marker (1 -> ●, 0 -> blank)
NUM_FORMATS = {"avvik": '[Color10]"▲ "#,##0;[Red]"▼ ("#,##0")";"–"',
               "avvikpst": '[Color10]"▲ "0.0 %;[Red]"▼ ("0.0 %")";"–"',
               "markor": '"●";"●";""'}
TALL1_THOUSANDS = '#,##0.0_);(#,##0.0);"–"_)'    # spec: 'tall1' is one decimal WITH thousands separator
HELTALL = tn.NUM["heltall"]


def _style_for(kind: str, emph: str) -> str:
    if kind not in NUM_ALIAS:
        raise ValueError(f"unknown number kind {kind!r}; use one of {sorted(NUM_ALIAS)}")
    return tn.style_name(NUM_ALIAS[kind], emph)


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
    from ..core import config
    return config.SVKIT / f"tn_template_{ver or _styles_version()}.xltx"


def ensure_template(xl) -> Path:
    """An Excel template (.xltx) holding the TN styles + version marker: new analysis workbooks are created from it
    (styles for free) and existing workbooks merge their styles from it. Built once per style version (~10 s)."""
    from ..core import config
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
        from .instances import _label
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
def _busy(e: BaseException) -> bool:
    from . import jobs
    return jobs.is_busy(e)


@contextmanager
def _active(ws):
    """Make ws the active sheet (needed for window-level settings), restore the previous workbook/sheet after.
    Three COM calls: read the previous workbook/sheet, activate, restore."""
    def read():
        wb = ws.Parent
        try:
            prev_wb, prev_ws = wb.Application.ActiveWorkbook, wb.Application.ActiveSheet
        except Exception as e:   # noqa: BLE001
            if _busy(e):
                raise
            prev_wb = prev_ws = None
        other = prev_wb is not None and prev_wb.FullName != wb.FullName
        return wb, prev_wb, prev_ws, other

    try:
        wb, prev_wb, prev_ws, other = _retry(read)
    except Exception:   # noqa: BLE001 - cannot tell which workbook is active: no window settings
        yield False
        return

    def enter():
        if prev_wb is None or other:
            wb.Activate()
        ws.Activate()
    ok = True
    try:
        _retry(enter)
    except Exception:   # noqa: BLE001
        ok = False
    try:
        yield ok
    finally:
        def leave():
            if other:
                prev_wb.Activate()
            if prev_ws is not None:
                prev_ws.Activate()
        try:
            _retry(leave)
        except Exception:   # noqa: BLE001
            pass


def _freeze(ws, rows: int, cols: int) -> bool:
    """Freeze panes below `rows` and right of `cols` and hide gridlines; only if the workbook window is visible.
    The window settings go in one COM call (rerun as a whole when Excel is busy: every setting is absolute)."""
    def visible():
        win = ws.Parent.Windows(1)
        return win if win.Visible else None
    try:
        win = _retry(visible)
    except Exception:   # noqa: BLE001
        win = None
    if win is None:
        return False

    def settings():
        win.FreezePanes = False
        win.ScrollRow = 1
        win.ScrollColumn = 1
        win.DisplayGridlines = False
        if rows or cols:
            win.SplitColumn = cols
            win.SplitRow = rows
            win.FreezePanes = True

    def undo():
        win.FreezePanes = False
        win.SplitColumn = 0
        win.SplitRow = 0
    with _active(ws) as ok:
        if not ok:
            return False
        try:
            _retry(settings)
            return True
        except Exception:   # noqa: BLE001   (e.g. window has no usable client area) - leave no half-split behind
            try:
                _retry(undo)
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


def _blocks(cells: list[tuple[int, int, int]]) -> list[str]:
    """[(row, c1, c2)] (one style's column runs) -> addresses, equal runs on consecutive rows joined into one
    rectangle ('A5:G7' instead of 'A5:G5', 'A6:G6', 'A7:G7')."""
    out: list[list[int]] = []                       # [r1, r2, c1, c2]
    open_: dict[tuple[int, int], list[int]] = {}
    for r, c1, c2 in sorted(cells):
        blk = open_.get((c1, c2))
        if blk is not None and blk[1] == r - 1:
            blk[1] = r
        else:
            blk = [r, r, c1, c2]
            open_[(c1, c2)] = blk
            out.append(blk)
    return [f"{_cl(c1)}{r1}:{_cl(c2)}{r2}" for r1, r2, c1, c2 in out]


def _merge_text(row, first, style: str = "TN Tekst") -> None:
    """A text row merged over the table width, in one COM call: wrap, merge, the TN Tekst style again (a merge
    resets the hidden cells to Normal), top alignment of the first cell."""
    row.WrapText = True
    row.MergeCells = True
    row.Style = style
    if style != "TN Tekst":            # TN Tekst wraps by itself; another style (a title) resets WrapText
        row.WrapText = True
    first.VerticalAlignment = tn.V_TOP


def _code_look(rng) -> None:
    """The code column (FinMax/FinPlan member codes): grey and left-aligned, in one COM call."""
    rng.Font.Color = _rgb(tn.GREY)
    rng.HorizontalAlignment = tn.H_LEFT


def _check_rules(cell) -> None:
    """The check colours of one range in one COM call: green bold for "OK", red on red otherwise. Starts with
    Delete, so a rerun leaves the same two rules."""
    fcs = cell.FormatConditions
    fcs.Delete()
    fc = fcs.Add(1, 3, '="OK"')                            # cell value = "OK"
    fc.Font.Color = _rgb(GREEN)
    fc.Font.Bold = True
    fc = fcs.Add(1, 4, '="OK"')                            # not equal -> breach
    fc.Font.Color = _rgb(tn.BREACH_FONT)
    fc.Font.Bold = True
    fc.Interior.Color = _rgb(tn.BREACH_FILL)


def apply_report(ws, roles: dict) -> list[str]:
    """Style an answer sheet whose content is already written (never clears contents). Returns lint findings."""
    from ..core.timing import stage
    own = bool(roles.get("last_row"))     # the tool-owned block: never re-style the user's rows below it

    def info():
        wb = ws.Parent
        xl = wb.Application
        tup = xl.International
        ur_last = None if own else ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        return wb, xl, tup[2], tup[3], tup[4], ur_last
    with stage("A ensure styles"):
        wb, xl, dec, thou, sep, ur_last = _retry(info)
        ensure_styles(wb)
    last = _ci(roles["last_col"])
    text_last = _ci(roles.get("print_last_col") or roles["last_col"])   # text rows end where the print area does
    labels = sorted(_ci(c) for c in roles.get("label_cols", ["B"]))
    first_label = labels[0]
    widths = {c: 8.43 for c in range(1, last + 1)}
    widths.update({c: w for c, w in _parse_widths(roles.get("col_widths", {})).items() if c <= last})
    rows = {int(r["row"]): r for r in roles["rows"]}
    if own:
        last_row = max(max(rows), int(roles["last_row"]))
    else:
        last_row = max(max(rows), ur_last)
    block = f"A1:{_cl(last)}{last_row}"

    raw = _retry(lambda: ws.Range(block).Value2)
    if not isinstance(raw, tuple):
        raw = ((raw,),)
    def val(r, c):
        return raw[r - 1][c - 1]

    # clear stale formats (keeps contents), then base style
    def clear_and_base():
        area = ws.Range(block)
        area.ClearFormats()
        area.Style = "TN Tom"
    with stage("A2 clear + base"):
        _retry(clear_and_base)

    by_style: dict[str, list[str]] = {}
    tall1_cells: list[str] = []
    formats: dict[str, list[str]] = {}                    # NUM_FORMATS kind -> cells (set after the styles)
    codes = {_ci(c) for c in roles.get("code_cols", [])}
    code_cells: list[tuple[int, int, int]] = []           # (row, col, col): code cells of table rows (grey, left)
    heights: dict[int, float] = {}
    merges: list[tuple[int, int]] = []
    title_merges: list[tuple[int, int]] = []          # a wrapped title: merged over the width in its own style
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
            txt = val(r, first_label) if kind == "title" else None
            if txt:                     # a long message title wraps inside the table width, never off the page
                width = sum(widths[c] for c in range(first_label, text_last + 1))
                n = tn.wrap_lines(str(txt), width, 16, True) or 1
                if n > 1:
                    heights[r] = n * 21 + 4
                    title_merges.append((r, first_label))
        elif kind == "subhead":         # a sub-heading inside a table (concept sections): label only, no numbers
            cols = {c: "TN Underseksjon" if c >= first_label else "TN Tom" for c in range(1, last + 1)}
            heights[r] = 15
        elif kind == "text":
            cols = {c: "TN Tekst" if c >= first_label else "TN Tom" for c in range(1, last + 1)}
            width = sum(widths[c] for c in range(first_label, text_last + 1))
            txt = val(r, first_label)
            n = (tn.wrap_lines(str(txt), width) or 4) if txt else 1
            heights[r] = max(15, n * 13 + 2)
            if all(val(r, c) in (None, "") for c in range(first_label + 1, text_last + 1)):
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
                    elif k in NUM_FORMATS:
                        formats.setdefault(k, []).append(f"{_cl(c)}{r}")
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
            code_cells += [(r, c, c) for c in codes if c not in num]
        else:
            raise ValueError(f"row {r}: unknown kind {kind!r}")
        for st, c1, c2 in _runs(cols):
            by_style.setdefault(st, []).append((r, c1, c2))
        for c, st in cols.items():
            assigned[(r, c)] = st

    # every call below sets absolute values, so one that Excel rejects while busy is rerun as a whole
    with stage("B cell styles"):
        for st, cells in by_style.items():                # equal column runs on consecutive rows: one rectangle
            for chunk in tn._chunks(_blocks(cells), sep):
                _retry(lambda: setattr(ws.Range(chunk), "Style", st))
        if tall1_cells:
            fmt = tn.to_local(TALL1_THOUSANDS, dec, thou)
            for chunk in tn._chunks(tall1_cells, sep):
                _retry(lambda: setattr(ws.Range(chunk), "NumberFormatLocal", fmt))
        for k, cells in formats.items():
            for chunk in tn._chunks(cells, sep):
                _retry(lambda: setattr(ws.Range(chunk), "NumberFormatLocal", tn.to_local(NUM_FORMATS[k], dec, thou)))
        if code_cells:
            for chunk in tn._chunks(_blocks(code_cells), sep):
                _retry(lambda: _code_look(ws.Range(chunk)))
        helpers = sorted(_ci(c) for c in roles.get("helper_cols", []))
        if helpers:                                       # chart helper columns: grey, they are not the answer
            hrng = f"{_cl(helpers[0])}1:{_cl(helpers[-1])}{last_row}"
            _retry(lambda: setattr(ws.Range(hrng).Font, "Color", _rgb(tn.GREY)))

    with stage("C widths + heights"):
        given = _parse_widths(roles.get("col_widths", {}))
        by_width: dict[float, list[tuple[int, int, int]]] = {}    # one call per run of equal widths, not per column
        for c, w in widths.items():
            if c in given:
                runs = by_width.setdefault(w, [])
                if runs and runs[-1][2] == c - 1:
                    runs[-1] = (1, runs[-1][1], c)
                else:
                    runs.append((1, c, c))
        for w, runs in by_width.items():
            for _, c1, c2 in runs:
                _retry(lambda: setattr(ws.Range(f"{_cl(c1)}1:{_cl(c2)}1").EntireColumn, "ColumnWidth", w))
        by_height: dict[float, list[str]] = {}            # one call per distinct height, not one per row
        for r, h in sorted(heights.items()):
            runs = by_height.setdefault(h, [])
            if runs and runs[-1][1] == r - 1:
                runs[-1] = (runs[-1][0], r)
            else:
                runs.append((r, r))
        for h, runs in by_height.items():
            for chunk in tn._chunks([f"{a}:{b}" for a, b in runs], sep):
                _retry(lambda: setattr(ws.Range(chunk), "RowHeight", h))
    with stage("D merges"):
        for r, c in merges:
            _retry(lambda: _merge_text(ws.Range(f"{_cl(c)}{r}:{_cl(text_last)}{r}"), ws.Range(f"{_cl(c)}{r}")))
        for r, c in title_merges:
            _retry(lambda: _merge_text(ws.Range(f"{_cl(c)}{r}:{_cl(text_last)}{r}"), ws.Range(f"{_cl(c)}{r}"),
                                       "TN Tittel"))
    with stage("E check colours"):
        for chunk in tn._chunks(_check_ranges(checks), sep):          # one rule pair per chunk (was per cell: 7 s)
            _retry(lambda: _check_rules(ws.Range(chunk)))

    with stage("F tab + freeze"):
        _retry(lambda: setattr(ws.Tab, "Color", _rgb(tn.NAVY)))
        _freeze(ws, int(roles.get("freeze_rows") or 0), 0)
    with stage("G page setup"):
        if roles.get("print_footer"):         # a v2 answer: its own footer, no repeated title rows
            area_col = roles.get("print_area_col") or _cl(text_last)
            area_w = sum((widths.get(c, 8.43) * 7 + 5) * 0.75 for c in range(1, _ci(area_col) + 1))
            breaks = keep_breaks(heights, last_row, [tuple(s) for s in roles.get("keep", [])], area_w)
            print_setup(ws, area_col, last_row, roles["print_footer"], header=roles.get("print_header"),
                        breaks=breaks)
        else:
            _page_setup(ws, xl, last_row, text_last, roles)
    with stage("H lint"):
        return _lint_fast(ws, roles, raw, last, last_row, widths, assigned, merges + title_merges, heights)


def _lint_fast(ws, roles, raw, last, last_row, widths, assigned, merges, heights) -> list[str]:
    """The tn-rapportstil checks that can fail here, without reading every cell back from Excel (the per-cell COM
    lint took 3-4 s): sheet-wide font and vertical borders from Excel, and text fit from the styles this module
    assigned itself (size/bold/wrap from tn.style_defs). Styles and number formats are set by our own code, so the
    per-cell 'has a TN style / has a number format' checks cannot fail and are not repeated."""
    errs: list[str] = []
    text_last = _ci(roles.get("print_last_col") or roles["last_col"])   # merged text rows end at the print area

    def read():                                       # used range, font and vertical borders: one COM call
        ur = ws.UsedRange
        ur_last_c = ur.Column + ur.Columns.Count - 1
        ur_last_r = ur.Row + ur.Rows.Count - 1
        area = ws.Range(f"A1:{_cl(max(ur_last_c, last))}{max(ur_last_r, last_row)}")
        lines = []
        for edge in (7, 10, 11):
            try:
                lines.append(area.Borders(edge).LineStyle)
            except Exception as e:   # noqa: BLE001
                if _busy(e):
                    raise
                lines.append(tn.NONE)                 # unreadable edge: not reported (as before)
        # the font of the drawn block only: cells a redraw cleared outside it (a wider earlier answer) fall back to
        # the Normal style's font, are empty and print nothing
        return ur_last_c, ws.Range(f"A1:{_cl(last)}{last_row}").Font.Name, lines
    ur_last_c, fname, lines = _retry(read)
    if fname != tn.FONT:
        errs.append(f"more than one font, or not Arial (Font.Name={fname!r})")
    for style, label in zip(lines, ("left", "right", "inside")):
        if style != tn.NONE:
            errs.append(f"vertical border ({label}) on the sheet (not allowed in the report style)")
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
                w = sum(widths.get(k, 8.43) for k in range(c, text_last + 1))
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


PAGE_SETUP_NAME = "smartass_pagesetup"
PAGE_SETUP_VERSION = "v1"          # bump whenever _page_setup_props changes what it sets


def _page_setup_key(last_row, last, roles) -> str:
    """The sheet-level name's value: the setup's version + what varies per draw (print area, header text), so a
    redraw with another size or header is set up again."""
    header = roles.get("print_header", "Intern")
    return f'="{PAGE_SETUP_VERSION}|$A$1:${_cl(last)}${last_row}|' + header.replace('"', '""') + '"'


def _old_setup_key(ws) -> str:
    try:
        return str(_retry(lambda: ws.Names(PAGE_SETUP_NAME).RefersTo))
    except Exception:   # noqa: BLE001 - name missing: not set up by us yet
        return ""


def _has_page_setup(ws, key: str) -> bool:
    return _old_setup_key(ws) == key


# A4 landscape (842 x 595 pt) with the setup's margins: 1,5 cm left/right, Excel's default 0,75" top/bottom
PRINT_WIDTH_PT, PRINT_HEIGHT_PT = 841.9 - 2 * 42.52, 595.3 - 2 * 54.0


def keep_breaks(heights: dict[int, float], last_row: int, keep: list[tuple[int, int]], area_width_pt: float,
                default_height: float = 15.0) -> list[int]:
    """The rows to start a new page at so that no keep span (a section from its heading to its last line, the
    charts zone) is cut by a page break, when the span fits on one page. Simulates the pages of a sheet printed one
    page wide (Excel's zoom: the whole percent that fits the print area's width). A span taller than a page flows."""
    zoom = min(100, int(PRINT_WIDTH_PT / area_width_pt * 100)) / 100 if area_width_pt > 0 else 1.0
    page = PRINT_HEIGHT_PT / zoom * 0.97                  # a small margin for Excel's own rounding
    starts = {a: b for a, b in keep if b >= a}
    used, out = 0.0, []
    for r in range(1, last_row + 1):
        h = heights.get(r, default_height)
        if r in starts and used > 0:
            span = sum(heights.get(k, default_height) for k in range(r, starts[r] + 1))
            if used + span > page and span <= page:
                out.append(r)
                used = 0.0
        if used + h > page:                               # Excel's automatic break
            used = 0.0
        used += h
    return out


def _page_setup(ws, xl, last_row, last, roles):
    """The answer's page setup, once per sheet: skipped (one COM call) when the sheet-level name smartass_pagesetup
    says this exact setup is already there (follow-up redraws); the first draw is unchanged and ends by adding it."""
    key = _page_setup_key(last_row, last, roles)
    if _has_page_setup(ws, key):
        return
    keep_footer = _has_center_footer(ws)        # read before PrintCommunication goes off
    # every PageSetup property normally round-trips to the printer driver (measured 14 s for 12 settings);
    # PrintCommunication=False batches them into one exchange when it is switched back on
    try:
        _set(xl, "PrintCommunication", False)
    except Exception:   # noqa: BLE001 - not available: just slower
        pass
    restored = False
    try:
        done = _page_setup_props(ws, xl, last_row, last, roles, keep_footer)
    finally:
        try:
            _set(xl, "PrintCommunication", True)
            restored = True
        except Exception:   # noqa: BLE001
            pass
    if done and restored:           # batched settings not flushed: do not mark the sheet as set up
        try:
            _retry(lambda: ws.Names.Add(Name=PAGE_SETUP_NAME, RefersTo=key, Visible=False))
        except Exception:   # noqa: BLE001 - only an optimisation
            pass


def _has_center_footer(ws) -> bool:
    """A v2 answer's print_setup footer (source, date, page &P of &N) is there: the v1 setup (xl_tidy) then leaves
    the footers alone, so the page never shows two page numbers."""
    try:
        return bool(str(_retry(lambda: ws.PageSetup.CenterFooter) or ""))   # one COM call
    except Exception:   # noqa: BLE001 - no printer / not readable: v1 footers as before
        return False


def _page_setup_props(ws, xl, last_row, last, roles, keep_footer: bool = False) -> bool:
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
        if not keep_footer:                     # a v2 CenterFooter already carries the page number
            _set(ps, "RightFooter", '&"Arial"&8Side &P av &N')
            _set(ps, "LeftFooter", '&"Arial"&8&F')
        return True
    except Exception:   # noqa: BLE001  (no printer installed etc.)
        return False


PRINT_SETUP_VERSION = "a1"        # bump whenever print_setup changes what it sets


def print_setup(ws, last_col: str, last_row: int, footer: str, header: str | None = None, *,
                fit_wide: bool = True, title_rows: str = "", breaks: list[int] | tuple = ()) -> None:
    """A v2 answer's page setup: print area A1:<last_col><last_row> (the tables' last column, chart helpers
    outside), landscape, one page wide and as many tall as needed, no repeated title rows (the zones vary), no
    gridlines, `footer` centred (it carries the page number; the left and right footers are cleared). The display
    gridlines are apply_report's; ActiveWindow is never touched. Batched (PrintCommunication off) and skipped (one
    COM call) when the sheet-level name says this exact setup is already there; a new footer (date) sets it up
    again. fit_wide False (a wide lines table): 100 % zoom, as many pages wide as needed, title_rows ('$4:$4')
    repeated on every page."""
    header = "Intern" if header is None else header
    area = f"$A$1:${last_col}${last_row}"
    parts = (PRINT_SETUP_VERSION, area, header, footer) + (() if fit_wide else ("zoom100", title_rows)) \
        + ((f"breaks {','.join(map(str, breaks))}",) if breaks else ())
    key = '="' + "|".join(parts).replace('"', '""') + '"'
    old = _old_setup_key(ws)
    if old == key:
        return
    if breaks or "breaks " in old:      # manual breaks (keep a section on one page): the old ones go first
        try:
            _retry(lambda: ws.ResetAllPageBreaks())
            for r in breaks:
                _retry(lambda r=r: ws.HPageBreaks.Add(ws.Range(f"A{r}")))
        except Exception:   # noqa: BLE001 - no printer: the automatic breaks stay
            pass
    xl = _retry(lambda: ws.Parent.Application)
    try:
        _set(xl, "PrintCommunication", False)
    except Exception:   # noqa: BLE001 - not available: just slower
        pass
    done = restored = False
    try:
        ps = _get(ws, "PageSetup")
        scale = (("Zoom", False), ("FitToPagesWide", 1), ("FitToPagesTall", False)) if fit_wide else             (("Zoom", 100),)
        for name, value in (("PrintArea", area), ("PrintTitleRows", title_rows), ("Orientation", 2), *scale,
                            ("PrintGridlines", False),
                            ("CenterHorizontally", True), ("CenterHeader", '&"Arial"&8' + header),
                            ("LeftFooter", ""), ("RightFooter", ""), ("CenterFooter", footer)):
            _set(ps, name, value)
        cm = _call(xl, "CentimetersToPoints", 1.5)
        _set(ps, "LeftMargin", cm)
        _set(ps, "RightMargin", cm)
        done = True
    except Exception:   # noqa: BLE001  (no printer installed etc.)
        pass
    finally:
        try:
            _set(xl, "PrintCommunication", True)
            restored = True
        except Exception:   # noqa: BLE001
            pass
    if done and restored:
        try:
            _retry(lambda: ws.Names.Add(Name=PAGE_SETUP_NAME, RefersTo=key, Visible=False))
        except Exception:   # noqa: BLE001 - only an optimisation
            pass


# ---------------------------------------------------------------- apply_data_sheet
def _format_numbers(rng, fmt: str, align: bool) -> None:
    """Number format (and right alignment) on the numeric constants and numeric formula results in rng, inside ONE
    COM call; no such cells is not an error, a busy Excel is raised (the call is rerun as a whole)."""
    for kind in (2, -4123):          # xlCellTypeConstants, xlCellTypeFormulas
        try:
            part = rng.SpecialCells(kind, 1)   # xlNumbers
        except Exception as e:   # noqa: BLE001 - no such cells
            if _busy(e):
                raise
            continue
        part.NumberFormatLocal = fmt
        if align:
            part.HorizontalAlignment = tn.H_RIGHT


def local_format(obj, fmt: str) -> str:
    """fmt (en-US codes: ',' thousands, '.' decimal) in the user's locale. Through late-bound COM (pywin32 passes the
    user's LCID) Excel reads Range.NumberFormat and DataLabels.NumberFormat in the user's locale too: '#,##0.0' was
    refused under nb-NO (live 05.11.2026). obj: any Excel object with .Application. Without one (fakes): fmt."""
    try:
        tup = obj.Application.International
        return tn.to_local(fmt, tup[2], tup[3])
    except Exception:   # noqa: BLE001 - a fake or a detached object: the code as given
        return fmt


CHART_THOUSANDS = "\xa0"


def chart_format(obj, fmt: str) -> str:
    """fmt (en-US codes) for a chart's DataLabels.NumberFormat. Excel parses that one differently from a cell's:
    under nb-NO the thousands and scaling character must be U+00A0, not the plain space International returns
    ('#\\xa0##0,0\\xa0\\xa0' shows 3 811,4 for 3 811 374 124.3; '# ##0,0  ' showed '3811374 124,3', live 05.11.2026).
    So every ',' (outside quotes) becomes U+00A0 and '.' the locale decimal; spaces stay plain (a U+00A0 after a
    digit placeholder would scale). An en-US Excel (',' thousands) keeps the code. Without an app (fakes): fmt."""
    try:
        tup = obj.Application.International
        dec, thou = tup[2], tup[3]
    except Exception:   # noqa: BLE001 - a fake or a detached object: the code as given
        return fmt
    return tn.to_local(fmt, dec, "," if thou == "," else CHART_THOUSANDS)


def _sheet_info(ws):
    """(decimal sep, thousands sep, first row, first col, row count, col count) of ws's UsedRange: one COM call."""
    tup = ws.Parent.Application.International
    ur = ws.UsedRange
    return tup[2], tup[3], ur.Row, ur.Column, ur.Rows.Count, ur.Columns.Count


def apply_data_sheet(ws, header_rows: int, label_cols: int, fte_rows: list[int] | None = None) -> None:
    """Financial look for a LIVE Smart View ad hoc grid. Never clears/moves cells, never ClearFormats, never
    inserts rows/columns. Rows/cols are absolute sheet indexes (1-based). The fixed look (fonts, header, alignment,
    border) is one COM call, the number formats one per range, widths + tab colour one: every setting is absolute,
    so a call that Excel rejects while busy is rerun as a whole."""
    dec, thou, r0, c0, nrow, ncol = _retry(lambda: _sheet_info(ws))
    last_r, last_c = max(r0 + nrow - 1, header_rows), max(c0 + ncol - 1, label_cols)

    def a(r1, c1, r2, c2):
        return f"{_cl(c1)}{r1}:{_cl(c2)}{r2}"

    def look():
        f = ws.Range(a(1, 1, last_r, last_c)).Font
        f.Name = tn.FONT
        f.Size = 10
        if header_rows > 0:
            f = ws.Range(a(1, 1, header_rows, last_c)).Font
            f.Bold = True
            f.Color = _rgb(tn.NAVY)
            if last_c > label_cols:
                ws.Range(a(1, label_cols + 1, header_rows, last_c)).HorizontalAlignment = tn.H_RIGHT
            if label_cols:
                ws.Range(a(1, 1, header_rows, label_cols)).HorizontalAlignment = tn.H_LEFT
            bd = ws.Range(a(header_rows, 1, header_rows, last_c)).Borders(tn.B)
            bd.LineStyle = tn.CONT
            bd.Weight = tn.THIN
            bd.Color = _rgb(tn.NAVY)
        if last_r > header_rows and label_cols:
            ws.Range(a(header_rows + 1, 1, last_r, label_cols)).HorizontalAlignment = tn.H_LEFT
    _retry(look)

    if last_r > header_rows and last_c > label_cols:
        f0 = tn.to_local(HELTALL, dec, thou)
        f1 = tn.to_local(TALL1_THOUSANDS, dec, thou)
        _retry(lambda: _format_numbers(ws.Range(a(header_rows + 1, label_cols + 1, last_r, last_c)), f0, True))
        for fr in fte_rows or []:
            if header_rows < fr <= last_r:
                _retry(lambda fr=fr: _format_numbers(ws.Range(a(fr, label_cols + 1, fr, last_c)), f1, False))

    # widths: ~46 shared by the label columns, 16 for data columns
    per_label = max(20.0, 46.0 / label_cols) if label_cols > 1 else 46.0

    def widths_and_tab():
        if label_cols:
            ws.Range(a(1, 1, 1, label_cols)).EntireColumn.ColumnWidth = per_label
        if last_c > label_cols:
            ws.Range(a(1, label_cols + 1, 1, last_c)).EntireColumn.ColumnWidth = 16
        ws.Tab.Color = _rgb(GREY_TAB)
    _retry(widths_and_tab)
    _freeze(ws, header_rows, label_cols)
