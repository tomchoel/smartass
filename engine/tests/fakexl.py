"""A small in-memory Excel object model for the COM writers (styling, source sheets, lines sheets, sheet order).

Every property a writer sets lands on the cells it covers (per cell: value, number format, font/fill/border
attributes, style, alignment, conditional formats), on columns (width), rows (height), the sheet (tab colour, page
setup, names, window state) or the workbook (sheet order, names). snapshot(wb) returns the whole state in a compact,
comparable form (cells with equal settings collapsed into rectangles), so a test can pin a writer's exact effect and
prove that a batched version leaves the same workbook. count_com(monkeypatch) counts jobs.com calls.

For the move-in of a built file: wb.Sheets.Add(Before, After, Count, Type:=path) inserts every sheet of an .xlsx
(values, sheet visibility, tab colours) as Excel does it (live, 20.11.2026: a taken name becomes 'name (2)', the
first new sheet is active, the active workbook stays; a sheet with a table is refused - the fake's stand-in, as
movein's guard always refuses such a file first and Excel's own behaviour there was not measured);
Application.Workbooks (Workbooks(name); Open reads an .xlsx the same way, plus its tables; Close),
wb.Worksheets([names]).Copy / .Delete as Excel does them (an omitted argument, pythoncom.ArgNotFound, reads as
None), LinkSources / BreakLink, and Application.events: every step asked of the fake Excel, in order across sheets
and workbooks (not part of snapshot).

What Excel does to formulas when sheets and tables change (live, 20.11.2026, the probe of Task 3): renaming a sheet or
a table rewrites every formula that names it - cells, defined names, chart series - outside string literals
(_renamed); deleting a sheet turns a reference to it, or to a table on it, into #REF! (_deleted). A protected sheet's
hidden formula (Worksheet.hidden, with ProtectContents) refuses Formula2, alone or in a range, as Excel's does.
"""
from __future__ import annotations

import copy
import datetime
import hashlib
import re
from pathlib import Path
from types import SimpleNamespace

MAXR, MAXC = 1048576, 16384
REF = -2146826265                        # #REF! as Value2 reads it (pywin32 and the add-in: an int)
CLIP_R, CLIP_C = 400, 80                 # cell-level settings on whole rows/columns: only this block is tracked
SEP = ";"                                # nb-NO list separator (Range unions)
class _International(tuple):
    """Application.International: read whole (a 0-based tuple) or as International(i), 1-based as Excel numbers it."""

    def __call__(self, i: int):
        return self[i - 1]


def _international() -> tuple:
    """nb-NO's Application.International (0-based): separators, the date codes and its short date dd.mm.yyyy."""
    v: list = [None] * 51
    v[2], v[3], v[4] = ",", "\xa0", SEP           # xlDecimalSeparator, xlThousandsSeparator, xlListSeparator
    v[16] = "."                                    # xlDateSeparator
    v[18:21] = ["å", "M", "D"]                     # xlYearCode, xlMonthCode, xlDayCode
    v[31] = 1.0                                    # xlDateOrder: day, month, year
    v[40] = v[41] = v[42] = True                   # xlMonthLeadingZero, xlDayLeadingZero, xl4DigitYears
    return _International(v)


INTERNATIONAL = _international()


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def col_letter(i: int) -> str:
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_COLS = re.compile(r"^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$")
_ROWS = re.compile(r"^\$?(\d+):\$?(\d+)$")


def parse_area(a: str) -> tuple[int, int, int, int]:
    a = a.strip()
    if m := _COLS.match(a):
        c1, c2 = col_index(m.group(1)), col_index(m.group(2))
        return 1, min(c1, c2), MAXR, max(c1, c2)
    if m := _ROWS.match(a):
        r1, r2 = int(m.group(1)), int(m.group(2))
        return min(r1, r2), 1, max(r1, r2), MAXC
    parts = a.split(":")
    if len(parts) == 1:
        parts = parts * 2
    (m1, m2) = (_CELL.match(parts[0]), _CELL.match(parts[1]))
    if not (m1 and m2):
        raise ValueError(f"bad address {a!r}")
    r1, c1, r2, c2 = int(m1.group(2)), col_index(m1.group(1)), int(m2.group(2)), col_index(m2.group(1))
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def parse_address(addr: str) -> list[tuple[int, int, int, int]]:
    if "," in addr and SEP != ",":
        raise ValueError(f"union with ',' in a {SEP!r} locale: {addr!r}")
    return [parse_area(a) for a in addr.split(SEP)]


def rect_text(r1, c1, r2, c2) -> str:
    if r1 == 1 and r2 == MAXR:
        return f"{col_letter(c1)}:{col_letter(c2)}"
    if c1 == 1 and c2 == MAXC:
        return f"{r1}:{r2}"
    a, b = f"{col_letter(c1)}{r1}", f"{col_letter(c2)}{r2}"
    return a if a == b else f"{a}:{b}"


def compress(cells) -> str:
    """A set of (row, col) as a canonical list of rectangles: column runs per row, merged over consecutive rows."""
    by_row: dict[int, list[int]] = {}
    for r, c in cells:
        by_row.setdefault(r, []).append(c)
    runs_by_row: dict[int, list[tuple[int, int]]] = {}
    for r, cs in by_row.items():
        cs.sort()
        runs = []
        for c in cs:
            if runs and runs[-1][1] == c - 1:
                runs[-1] = (runs[-1][0], c)
            else:
                runs.append((c, c))
        runs_by_row[r] = runs
    open_: dict[tuple[int, int], int] = {}            # (c1, c2) -> first row
    rects, prev_row = [], None
    for r in sorted(runs_by_row) + [None]:
        now = set(runs_by_row.get(r, [])) if r is not None else set()
        nxt = {}
        for span, top in open_.items():
            if r is not None and span in now and prev_row == r - 1:
                nxt[span] = top
            else:
                rects.append((top, span[0], prev_row, span[1]))
        for span in now:
            if span not in nxt:
                nxt[span] = r
        open_ = nxt
        prev_row = r
    return ",".join(rect_text(*x) for x in sorted(rects))


_HYPERLINK = re.compile(r'^=HYPERLINK\("(#(?:[^"]|"")*)","((?:[^"]|"")*)"\)$')


def hyperlink(v) -> tuple[str, str] | None:
    """(target, text) of a cell value that is a =HYPERLINK("#<target>","<text>") formula (quotes undoubled)."""
    m = _HYPERLINK.match(v) if isinstance(v, str) else None
    return (m.group(1).replace('""', '"'), m.group(2).replace('""', '"')) if m else None


def shown(v):
    """What Excel shows for a value: a HYPERLINK formula's text, anything else as it is."""
    link = hyperlink(v)
    return link[1] if link else v


class ComError(Exception):
    """Not a 'busy' error: jobs.com passes it straight through."""


def _given(args: tuple) -> tuple:
    """Positional arguments with an omitted one (pythoncom.ArgNotFound, sent as DISP_E_PARAMNOTFOUND) read as None,
    as Excel reads both."""
    import pythoncom
    return tuple(None if a is pythoncom.ArgNotFound else a for a in args)


def _event(ws, *what) -> None:
    """One step in the Application's event list (everything asked of the fake Excel, in order, across workbooks)."""
    events = getattr(ws.wb.Application, "events", None)
    if events is not None:
        events.append((ws.Name, *what))


# ------------------------------------------------------------------ formulas following renames and deletes
_LITERAL = re.compile(r'("(?:[^"]|"")*")')
_QUOTED_REF = re.compile(r"'((?:[^']|'')*)'!")
_PLAIN = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")


def _quoted(name: str) -> str:
    return name if _PLAIN.match(name) else "'" + name.replace("'", "''") + "'"


def _outside_literals(text: str, fn) -> str:
    """fn applied to the parts of a formula outside its string literals."""
    return "".join(part if part.startswith('"') else fn(part) for part in _LITERAL.split(text))


def _formulas(wb):
    """(get, set) of every formula text of the workbook: cells, defined names (the workbook's and the sheets'), chart
    series."""
    for ws in wb.order:
        for props in ws.cells.values():
            v = props.get("Value")
            if isinstance(v, str) and v.startswith("="):
                yield (lambda p=props: p["Value"]), (lambda x, p=props: p.__setitem__("Value", x))
        for item in ws.Names.items.values():
            yield (lambda n=item: n.RefersTo), (lambda x, n=item: setattr(n, "RefersTo", x))
        for co in ws.charts:
            for s in co.Chart.series:
                yield (lambda s=s: s.Formula), (lambda x, s=s: setattr(s, "Formula", x))
    for item in wb.Names.items.values():
        yield (lambda n=item: n.RefersTo), (lambda x, n=item: setattr(n, "RefersTo", x))
    for ch in wb.Charts:
        for s in getattr(ch, "series", []):
            yield (lambda s=s: s.Formula), (lambda x, s=s: setattr(s, "Formula", x))


def _rewrite(wb, fn) -> None:
    for get, put in list(_formulas(wb)):
        v = get()
        if isinstance(v, str) and v.startswith("="):
            new = _outside_literals(v, fn)
            if new != v:
                put(new)


def _shown(v):
    """Value2 as Excel shows it where the fake can tell without calculating: a formula that names #REF! shows #REF!
    (its error code), everything else as stored."""
    if isinstance(v, tuple):
        return tuple(tuple(_shown(x) for x in row) for row in v)
    return REF if isinstance(v, str) and v.startswith("=") and "#REF!" in v else v


def _renamed_sheet(wb, old: str, new: str) -> None:
    """As Excel: every reference to the sheet names it by its new name (quoted when it needs quotes)."""
    def sheet(m):
        parts = m.group(1).replace("''", "'").split(":")
        if "[" in m.group(1) or not any(p.casefold() == old.casefold() for p in parts):
            return m.group(0)
        parts = [new if p.casefold() == old.casefold() else p for p in parts]
        return "'" + ":".join(parts).replace("'", "''") + "'!"

    def fn(text):
        text = _QUOTED_REF.sub(sheet, text)
        if _PLAIN.match(old):
            text = re.sub(rf"(?<![\w.'!])({re.escape(old)})!", lambda m: _quoted(new) + "!", text, flags=re.I)
        return text
    _rewrite(wb, fn)


def _table_word(name: str) -> str:
    return rf"(?<![\w.\\?!\[]){re.escape(name)}(?![\w.\\?])"


def _renamed_table(wb, old: str, new: str) -> None:
    """As Excel: every reference to the table names it by its new name."""
    _rewrite(wb, lambda text: re.sub(_table_word(old), new, text, flags=re.I))


def _deleted(wb, sheets: list) -> None:
    """As Excel: a reference to a deleted sheet becomes #REF! ('X'!E12 -> #REF!E12), a table on it #REF! with its
    columns."""
    names = {s.Name.casefold() for s in sheets}
    tables = [lo.Name for s in sheets for lo in s.ListObjects]

    def sheet(m):
        parts = m.group(1).replace("''", "'").split(":")
        return "#REF!" if "[" not in m.group(1) and any(p.casefold() in names for p in parts) else m.group(0)

    def fn(text):
        text = _QUOTED_REF.sub(sheet, text)
        for s in sheets:
            if _PLAIN.match(s.Name):
                text = re.sub(rf"(?<![\w.'!]){re.escape(s.Name)}!", "#REF!", text, flags=re.I)
        for t in tables:
            text = re.sub(_table_word(t) + r"(\[(?:[^\[\]]|\[[^\]]*\])*\])?", "#REF!", text, flags=re.I)
        return text
    _rewrite(wb, fn)


# ------------------------------------------------------------------ ranges
class _Sub:
    """Range.Font / Range.Interior / Range.Borders(i) / FormatCondition.Font: attribute sets land per cell."""

    def __init__(self, rng, prefix, cells=None):
        object.__setattr__(self, "_rng", rng)
        object.__setattr__(self, "_prefix", prefix)
        object.__setattr__(self, "_cells", cells)

    def _targets(self):
        return self._cells if self._cells is not None else self._rng.cells()

    def __setattr__(self, k, v):
        for rc in self._targets():
            self._rng.ws.cell(rc)[f"{self._prefix}.{k}"] = v

    def __getattr__(self, k):
        default = {"Font.Name": "Arial", "Font.Size": 10}.get(f"{self._prefix}.{k}")
        if self._prefix.startswith("Borders") and k == "LineStyle":
            default = -4142
        vals = {self._rng.ws.cells.get(rc, {}).get(f"{self._prefix}.{k}", default) for rc in self._targets()}
        return vals.pop() if len(vals) == 1 else None


class _Rule:
    def __init__(self, rule: dict):
        self.rule = rule
        self.Font = _RuleSub(rule, "Font")
        self.Interior = _RuleSub(rule, "Interior")


class _RuleSub:
    def __init__(self, rule, prefix):
        object.__setattr__(self, "_rule", rule)
        object.__setattr__(self, "_prefix", prefix)

    def __setattr__(self, k, v):
        self._rule[f"{self._prefix}.{k}"] = v


class _FormatConditions:
    def __init__(self, rng):
        self.rng = rng

    def Delete(self):
        for rc in self.rng.cells():
            self.rng.ws.cell(rc).pop("FC", None)

    def Add(self, kind, op, formula):
        rule = {"type": kind, "op": op, "f": formula}
        for rc in self.rng.cells():
            self.rng.ws.cell(rc).setdefault("FC", []).append(rule)
        return _Rule(rule)


_FORMAT_PROPS = {"NumberFormatLocal", "NumberFormat", "HorizontalAlignment", "VerticalAlignment", "Style",
                 "WrapText", "MergeCells", "IndentLevel"}
_STYLE_PARTS = ("NumberFormat", "HorizontalAlignment", "VerticalAlignment", "WrapText", "IndentLevel", "Font.",
                "Interior.", "Borders")


class Range:
    def __init__(self, ws, rects):
        object.__setattr__(self, "ws", ws)
        object.__setattr__(self, "rects", list(rects))

    # -- geometry
    def cells(self):
        out = []
        for r1, c1, r2, c2 in self.rects:
            for r in range(r1, min(r2, CLIP_R) + 1):
                for c in range(c1, min(c2, CLIP_C) + 1):
                    out.append((r, c))
        return out

    @property
    def Row(self):
        return self.rects[0][0]

    @property
    def Column(self):
        return self.rects[0][1]

    @property
    def Rows(self):
        r1, _, r2, _ = self.rects[0]
        return SimpleNamespace(Count=r2 - r1 + 1)

    @property
    def Columns(self):
        """Range.Columns: Count, AutoFit, and Columns(i) - the range's i-th column."""
        r1, c1, r2, c2 = self.rects[0]
        rng = self

        class Columns:
            Count = c2 - c1 + 1

            def AutoFit(self):
                rng.ws.log.append(("AutoFit", ";".join(rect_text(*x) for x in rng.rects)))

            def __call__(self, i):
                return Range(rng.ws, [(r1, c1 + i - 1, r2, c1 + i - 1)])
        return Columns()

    @property
    def Address(self):
        return ",".join(_abs(x) for x in self.rects)

    @property
    def EntireColumn(self):
        return Range(self.ws, [(1, c1, MAXR, c2) for _, c1, _, c2 in self.rects])

    @property
    def EntireRow(self):
        return Range(self.ws, [(r1, 1, r2, MAXC) for r1, _, r2, _ in self.rects])

    def End(self, direction):
        assert direction == -4162                       # xlUp
        r1, c1, _, _ = self.rects[0]
        used = [r for (r, c), p in self.ws.cells.items() if c == c1 and r <= r1 and "Value" in p]
        return Range(self.ws, [(max(used) if used else 1, c1, max(used) if used else 1, c1)])

    # -- sub objects
    @property
    def Font(self):
        return _Sub(self, "Font")

    @property
    def Interior(self):
        return _Sub(self, "Interior")

    def Borders(self, idx):
        cells = []
        for r1, c1, r2, c2 in self.rects:
            r2c, c2c = min(r2, CLIP_R), min(c2, CLIP_C)
            for r in range(r1, r2c + 1):
                for c in range(c1, c2c + 1):
                    if (idx == 7 and c == c1) or (idx == 8 and r == r1) or (idx == 9 and r == r2c) or \
                            (idx == 10 and c == c2c) or (idx == 11 and c > c1) or (idx == 12 and r > r1):
                        cells.append((r, c))
        return _Sub(self, f"Borders{idx}", cells)

    @property
    def FormatConditions(self):
        return _FormatConditions(self)

    @property
    def Hyperlinks(self):
        """Range.Hyperlinks: the sheet's links on these cells; Delete removes them and, as Excel, the cells' formats."""
        inside = set(self.cells())
        mine = [h for h in self.ws.Hyperlinks if h.cell in inside]

        def delete():
            self.ws.Hyperlinks[:] = [h for h in self.ws.Hyperlinks if h.cell not in inside]
            self.ClearFormats()
        return SimpleNamespace(Count=len(mine), Delete=delete)

    # -- values
    def _get_value(self):
        r1, c1, r2, c2 = self.rects[0]
        if (r1, c1) == (r2, c2):
            return self.ws.cells.get((r1, c1), {}).get("Value")
        return tuple(tuple(self.ws.cells.get((r, c), {}).get("Value") for c in range(c1, c2 + 1))
                     for r in range(r1, r2 + 1))

    def _set_value(self, v):
        r1, c1, r2, c2 = self.rects[0]
        if isinstance(v, (tuple, list)):
            assert len(v) == r2 - r1 + 1 and all(len(row) == c2 - c1 + 1 for row in v), "shape mismatch"
            for i, row in enumerate(v):
                for j, x in enumerate(row):
                    self.ws.put(r1 + i, c1 + j, x)
        else:
            for r, c in self.cells():
                self.ws.put(r, c, v)

    def __getattr__(self, k):
        if k in ("Value", "Value2", "Formula", "Formula2"):
            if k.startswith("Formula") and self.ws.ProtectContents and self.ws.hidden & set(self.cells()):
                raise ComError("a hidden formula of a protected sheet cannot be read")
            v = self._get_value()
            return _shown(v) if k == "Value2" else v
        if k == "PrefixCharacter":
            return "'" if self.rects[0][:2] in self.ws.prefixed else ""
        if k == "HasFormula":                            # one cell: a formula, not text written with an apostrophe
            v = self._get_value()
            return isinstance(v, str) and v.startswith("=") and self.rects[0][:2] not in self.ws.prefixed
        if k in ("HasArray", "CurrentArray", "FormulaArray"):
            r, c = self.rects[0][:2]
            area = next((a for a in self.ws.arrays if a[0] <= r <= a[2] and a[1] <= c <= a[3]), None)
            if k == "HasArray":
                return area is not None
            if area is None:
                raise ComError("no array formula here")
            return Range(self.ws, [area]) if k == "CurrentArray" else self.ws.cells[area[:2]]["Value"]
        if k == "ColumnWidth":
            vals = {self.ws.colw.get(c, self.ws.StandardWidth) for _, c1, _, c2 in self.rects
                    for c in range(c1, max(c1, min(c2, CLIP_C)) + 1)}
            return vals.pop() if len(vals) == 1 else None
        if k == "Width":                                 # points: whole pixels of the sheet's digit width + 5
            return sum((int(self.ws.colw.get(c, self.ws.StandardWidth) * self.ws.digit_px + 0.5) + 5) * 0.75
                       for _, c1, _, c2 in self.rects for c in range(c1, max(c1, min(c2, CLIP_C)) + 1))
        if k == "RowHeight":
            vals = {self.ws.rowh.get(r, 15) for r1, _, r2, _ in self.rects for r in range(r1, min(r2, CLIP_R) + 1)}
            return vals.pop() if len(vals) == 1 else None
        if k in _FORMAT_PROPS:
            vals = {self.ws.cells.get(rc, {}).get(k) for rc in self.cells()}
            return vals.pop() if len(vals) == 1 else None
        raise AttributeError(k)

    def __setattr__(self, k, v):
        self.ws.log.append((k, ";".join(rect_text(*x) for x in self.rects)))
        _event(self.ws, k, ";".join(rect_text(*x) for x in self.rects))
        if k == "FormulaArray":                          # a legacy array: every cell holds the formula
            self._check_tables(v)
            area = self.rects[0]
            self.ws.arrays = [a for a in self.ws.arrays if a != area] + [area]
            for r, c in self.cells():
                self.ws.put(r, c, v)
        elif k in ("Value", "Value2", "Formula", "Formula2"):
            self._check_tables(v)
            self._set_value(v)
        elif k == "ColumnWidth":
            for _, c1, _, c2 in self.rects:
                for c in range(c1, min(c2, CLIP_C) + 1):
                    self.ws.colw[c] = v
        elif k == "RowHeight":
            for r1, _, r2, _ in self.rects:
                for r in range(r1, min(r2, CLIP_R) + 1):
                    self.ws.rowh[r] = v
        elif k in _FORMAT_PROPS:
            for rc in self.cells():
                p = self.ws.cell(rc)
                if k == "Style":                       # as Excel: a style replaces the formats it includes
                    for x in [x for x in p if x.startswith(_STYLE_PARTS)]:
                        del p[x]
                p[k] = v
        else:
            raise AttributeError(f"fake Range cannot set {k}")

    # -- methods
    def ClearFormats(self):
        for rc in self.cells():
            p = self.ws.cells.get(rc)
            if p:
                for k in [k for k in p if k != "Value"]:
                    del p[k]

    def ClearContents(self):
        for rc in self.cells():
            self.ws.cells.get(rc, {}).pop("Value", None)

    def Clear(self):
        for rc in self.cells():
            self.ws.cells.pop(rc, None)
        inside = set(self.cells())
        self.ws.Hyperlinks[:] = [h for h in self.ws.Hyperlinks if h.cell not in inside]   # as Excel: Clear drops links

    def AutoFilter(self, *a):
        """As Excel: without arguments it TOGGLES the sheet's AutoFilter."""
        assert not a
        self.ws.log.append(("AutoFilter", ";".join(rect_text(*x) for x in self.rects)))
        if self.ws.AutoFilterMode:
            self.ws.AutoFilterMode, self.ws.autofilter = False, None
        else:
            self.ws.AutoFilterMode, self.ws.autofilter = True, self.rects[0]

    def SpecialCells(self, kind, value_type=None):
        if kind == 11:                                  # xlCellTypeLastCell: the used range's last cell
            r1, c1, r2, c2 = self.ws.UsedRange.rects[0]
            return Range(self.ws, [(r2, c2, r2, c2)])
        hits = []
        for rc in self.cells():
            v = self.ws.cells.get(rc, {}).get("Value")
            is_formula = isinstance(v, str) and v.startswith("=")
            numeric = isinstance(v, (int, float)) and not isinstance(v, bool)
            if kind == 2 and numeric and not is_formula:
                hits.append(rc)
            elif kind == -4123 and is_formula:
                hits.append(rc)
        if not hits:
            raise ComError("No cells were found.")
        return Range(self.ws, [(r, c, r, c) for r, c in hits])

    def Select(self):
        self.ws.wb.selection[self.ws.Name] = ";".join(rect_text(*x) for x in self.rects)

    def AutoFit(self):
        pass

    def _check_tables(self, v) -> None:
        """As Excel: a formula naming a table that does not exist (yet) in the workbook is refused (0x800A03EC)."""
        flat = [x for row in v for x in row] if isinstance(v, (tuple, list)) else [v]
        names = {lo.Name for ws in self.ws.wb.order for lo in ws.ListObjects}
        for x in flat:
            if isinstance(x, str) and x.startswith("="):
                for name in re.findall(r"((?:Data|Lines)_Q\d+(?:_\w+)?)\[", x):
                    if name not in names:
                        raise ComError(f"formula names a missing table {name}: {x}")

    def _shift_rows(self, at: int, n: int) -> None:
        """Move every cell, row height and table at or below row `at` by n rows (n < 0: up)."""
        ws = self.ws
        ws.cells = {((r + n) if r >= at else r, c): v for (r, c), v in ws.cells.items()}
        ws.rowh = {((r + n) if r >= at else r): h for r, h in ws.rowh.items()}
        for h in ws.Hyperlinks:
            if h.cell[0] >= at:
                h.cell = (h.cell[0] + n, h.cell[1])
        for lo in ws.ListObjects:
            r1, c1, r2, c2 = lo.rect
            if r1 >= at:
                lo.rect = (r1 + n, c1, r2 + n, c2)

    def Insert(self, shift=None):
        """Whole rows only (as ws.Rows('1:k').Insert()): the rows below move down, tables keep name and identity."""
        r1, c1, r2, c2 = self.rects[0]
        assert (c1, c2) == (1, MAXC), "fake Insert: whole rows only"
        self.ws.log.append(("Insert", rect_text(r1, c1, r2, c2)))
        self._shift_rows(r1, r2 - r1 + 1)

    def Delete(self, shift=None):
        r1, c1, r2, c2 = self.rects[0]
        assert (c1, c2) == (1, MAXC), "fake Delete: whole rows only"
        self.ws.log.append(("Delete", rect_text(r1, c1, r2, c2)))
        for lo in self.ws.ListObjects:
            assert not (lo.rect[0] <= r2 and lo.rect[2] >= r1), "fake Delete: rows inside a table"
        for rc in [rc for rc in self.ws.cells if r1 <= rc[0] <= r2]:
            del self.ws.cells[rc]
        self.ws.Hyperlinks[:] = [h for h in self.ws.Hyperlinks if not r1 <= h.cell[0] <= r2]
        self._shift_rows(r2 + 1, -(r2 - r1 + 1))


def _abs(rect) -> str:
    r1, c1, r2, c2 = rect
    a, b = f"${col_letter(c1)}${r1}", f"${col_letter(c2)}${r2}"
    return a if a == b else f"{a}:{b}"


class _Lines:
    """ws.Rows / ws.Columns: callable (one row/column or 'a:b') with .Count."""

    def __init__(self, ws, rows: bool):
        self.ws, self.rows = ws, rows
        self.Count = MAXR if rows else MAXC

    def __call__(self, key):
        if self.rows:
            r1, _, r2, _ = parse_area(f"{key}:{key}" if isinstance(key, int) else str(key))
            return Range(self.ws, [(r1, 1, r2, MAXC)])
        if isinstance(key, int):
            return Range(self.ws, [(1, key, MAXR, key)])
        key = str(key)
        return Range(self.ws, [parse_area(key if ":" in key else f"{key}:{key}")])


class ListObject:
    def __init__(self, ws, rng):
        self.ws, self.rect, self.Name, self.TableStyle = ws, rng.rects[0], "Table1", None
        self.ShowTableStyleRowStripes = True

    def __setattr__(self, k, v):
        old = self.__dict__.get("Name")
        super().__setattr__(k, v)
        if k == "Name" and old is not None and old != v and self in self.ws.ListObjects:
            _renamed_table(self.ws.wb, old, v)               # as Excel: the formulas follow the table

    @property
    def Range(self):
        return Range(self.ws, [self.rect])

    @property
    def HeaderRowRange(self):
        r1, c1, _, c2 = self.rect
        return Range(self.ws, [(r1, c1, r1, c2)])

    @property
    def DataBodyRange(self):
        r1, c1, r2, c2 = self.rect
        return Range(self.ws, [(r1 + 1, c1, r2, c2)]) if r2 > r1 else None

    def Resize(self, rng):
        self.rect = rng.rects[0]

    def Delete(self):
        self.ws.ListObjects.remove(self)
        self.Range.Clear()


class _ListObjects(list):
    def __init__(self, ws):
        super().__init__()
        self.ws = ws

    @property
    def Count(self):
        return len(self)

    def __call__(self, key):
        if isinstance(key, int):
            return self[key - 1]
        for lo in self:
            if lo.Name == key:
                return lo
        raise ComError(f"no table {key}")

    def Add(self, source_type, rng, link_source=None, has_headers=None):
        self.ws.log.append(("ListObjects.Add", ";".join(rect_text(*x) for x in rng.rects)))
        _event(self.ws, "ListObjects.Add", ";".join(rect_text(*x) for x in rng.rects))
        lo = ListObject(self.ws, rng)
        self.append(lo)
        return lo


class _Name(SimpleNamespace):
    """A defined name; RefersToLocal is RefersTo (the fake has no locale)."""

    @property
    def RefersToLocal(self):
        return self.RefersTo

    @RefersToLocal.setter
    def RefersToLocal(self, v):
        self.RefersTo = v


class _Names:
    def __init__(self, owner):
        self.owner, self.items = owner, {}

    def Add(self, Name=None, RefersTo=None, Visible=True, **kw):
        self.items[Name] = _Name(Name=Name, RefersTo=RefersTo, Visible=Visible)
        return self.items[Name]

    def __call__(self, name):
        if name not in self.items:
            raise ComError(f"no name {name}")
        return self.items[name]

    def Item(self, name):
        """As Excel: a name is found without case, or by its index."""
        if isinstance(name, int):
            return list(self.items.values())[name - 1]
        for key, item in self.items.items():
            if key.lower() == str(name).lower():
                return item
        raise ComError(f"no name {name}")

    def __iter__(self):
        return iter(list(self.items.values()))

    @property
    def Count(self):
        return len(self.items)


class _Series(SimpleNamespace):
    """A chart series: its formula (=SERIES(...))."""


class _Chart:
    def __init__(self, formulas):
        self.series = [_Series(Formula=f) for f in formulas]

    def SeriesCollection(self, i=None):
        return list(self.series) if i is None else self.series[i - 1]


class _ChartObjects(list):
    """ws.ChartObjects(): chart objects (.Chart) a test adds with add(series formulas)."""

    @property
    def Count(self):
        return len(self)

    def add(self, *formulas):
        co = SimpleNamespace(Name=f"Chart {len(self) + 1}", Chart=_Chart(formulas))
        self.append(co)
        return co


class ChartSheet(_Chart):
    """A chart sheet: its name and its series' formulas."""

    def __init__(self, name, *formulas):
        super().__init__(formulas)
        self.Name = name


class _Charts(list):
    """wb.Charts: the chart sheets, Charts(name) as Excel finds one."""

    def __call__(self, key):
        for ch in self:
            if str(ch.Name).lower() == str(key).lower():
                return ch
        raise ComError(f"no chart sheet {key}")

    @property
    def Count(self):
        return len(self)


class _Hyperlinks(list):
    """ws.Hyperlinks: Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay) POSITIONAL only (late-bound COM
    ignored the keyword TextToDisplay in v1); one link per cell, the text becomes the cell value."""

    def __init__(self, ws):
        super().__init__()
        self.ws = ws

    def Add(self, anchor, address, sub=None, tip=None, text=None, **kw):
        assert not kw, "Hyperlinks.Add: positional arguments only"
        r, c = anchor.rects[0][:2]
        self[:] = [h for h in self if h.cell != (r, c)]
        link = SimpleNamespace(cell=(r, c), Address=address, SubAddress=sub, ScreenTip=tip, TextToDisplay=text)
        self.append(link)
        if text:
            self.ws.put(r, c, text)
        return link

    def at(self, r, c):
        return next((h for h in self if h.cell == (r, c)), None)


class _Recorder:
    """PageSetup / Tab: plain attribute bag."""

    def __init__(self, **kw):
        self.__dict__.update(kw)


class Worksheet:
    def __init__(self, wb, name):
        self.wb, self.Name, self.Visible = wb, name, -1
        self.cells: dict[tuple[int, int], dict] = {}
        self.colw: dict[int, float] = {}
        self.rowh: dict[int, float] = {}
        self.Tab = _Recorder(Color=False)
        self.PageSetup = _Recorder()
        self.Names = _Names(self)
        self.ListObjects = _ListObjects(self)
        self.Rows = _Lines(self, True)
        self.Columns = _Lines(self, False)
        self.Hyperlinks = _Hyperlinks(self)
        self.AutoFilterMode, self.autofilter = False, None
        self.log: list = []
        self.prefixed: set[tuple[int, int]] = set()     # cells written with a leading apostrophe (text as text)
        self.page_breaks: list[int] | None = None       # manual page breaks: None until reset
        self.StandardHeight, self.StandardWidth, self.digit_px = 15.0, 8.43, 7     # Calibri 11, as a new workbook
        self.charts = _ChartObjects()                   # ChartObjects(): charts with series formulas
        self.ProtectContents = False
        self.hidden: set[tuple[int, int]] = set()        # a protected sheet's hidden formulas: Formula2 refuses them
        self.arrays: list[tuple[int, int, int, int]] = []    # legacy array formulas (FormulaArray), by their area

    def __setattr__(self, k, v):
        old = self.__dict__.get("Name")
        super().__setattr__(k, v)
        if k == "Name" and old is not None and old != v and self in self.wb.order:
            _renamed_sheet(self.wb, old, v)               # as Excel: the formulas follow the sheet

    @property
    def Application(self):
        return self.wb.Application

    def ChartObjects(self, i=None):
        return self.charts if i is None else self.charts[i - 1]

    @property
    def Parent(self):
        return self.wb

    def cell(self, rc):
        return self.cells.setdefault(rc, {})

    def put(self, r, c, x):
        if x is None or x == "":
            self.cells.get((r, c), {}).pop("Value", None)
            return
        if isinstance(x, str) and x.startswith("'"):
            x = x[1:]
            self.prefixed.add((r, c))
        else:
            self.prefixed.discard((r, c))
        self.cell((r, c))["Value"] = x

    def Range(self, a, b=None):
        if b is not None:
            r1 = min(a.rects[0][0], b.rects[0][0]); c1 = min(a.rects[0][1], b.rects[0][1])
            r2 = max(a.rects[0][2], b.rects[0][2]); c2 = max(a.rects[0][3], b.rects[0][3])
            return Range(self, [(r1, c1, r2, c2)])
        return Range(self, parse_address(a))

    def Cells(self, r, c):
        return Range(self, [(r, c, r, c)])

    @property
    def UsedRange(self):
        keys = [rc for rc, p in self.cells.items() if p]
        if not keys:
            return Range(self, [(1, 1, 1, 1)])
        rs, cs = [r for r, _ in keys], [c for _, c in keys]
        return Range(self, [(min(rs), min(cs), max(rs), max(cs))])

    def Activate(self):
        if self not in self.wb.order:
            raise ComError("the sheet was deleted")
        self.wb.Activate()
        self.wb.active = self

    def Calculate(self):
        _event(self, "Calculate")

    def ResetAllPageBreaks(self):
        self.page_breaks = []

    @property
    def HPageBreaks(self):
        return SimpleNamespace(Add=lambda rng: self.page_breaks.append(rng.Row))

    def Delete(self):
        _SheetGroup(self.wb, [self]).Delete()

    def Move(self, *args, **kw):
        assert not kw
        args = _given(args)
        before = args[0] if args else None
        after = args[1] if len(args) > 1 else None
        self.wb.order.remove(self)
        k = self.wb.order.index(before if before is not None else after)
        self.wb.order.insert(k if before is not None else k + 1, self)


class _Window:
    """One window per workbook; split/freeze settings are kept per active sheet, as Excel does."""

    def __init__(self, wb):
        object.__setattr__(self, "wb", wb)
        object.__setattr__(self, "Visible", True)

    def __setattr__(self, k, v):
        if k == "Visible":
            object.__setattr__(self, k, v)
            _event(self.wb.order[0], "Window.Visible", v)
            return
        self.wb.window_state.setdefault(self.wb.active.Name, {})[k] = v

    def __getattr__(self, k):
        return self.wb.window_state.get(self.wb.active.Name, {}).get(k)


class _Sheets:
    def __init__(self, wb):
        self.wb = wb

    def __call__(self, key):
        if isinstance(key, (list, tuple)):
            return _SheetGroup(self.wb, [self(k) for k in key])
        if isinstance(key, int):
            return self.wb.order[key - 1]
        for s in self.wb.order:
            if s.Name.lower() == str(key).lower():       # as Excel: a sheet's name is found without case
                return s
        raise ComError(f"no sheet {key}")

    def __iter__(self):
        return iter(list(self.wb.order))

    @property
    def Count(self):
        return len(self.wb.order)

    def Add(self, *args, **kw):
        """Positional (Before, After, Count, Type), as late-bound COM needs. With Type (a path): every sheet of that
        .xlsx inserted at the anchor, in the file's order (_insert); else one new empty sheet."""
        assert not kw, "Sheets.Add: positional arguments only"
        before, after, _count, path = (_given(args) + (None,) * 4)[:4]
        if path is not None:
            return _insert(self.wb, path, before, after)
        ws = Worksheet(self.wb, f"Sheet{len(self.wb.order) + 1}")
        if after is not None:
            self.wb.order.insert(self.wb.order.index(after) + 1, ws)
        elif before is not None:
            self.wb.order.insert(self.wb.order.index(before), ws)
        else:
            self.wb.order.append(ws)
        return ws


class _SheetGroup:
    """wb.Worksheets([name, ...]): several sheets in one call (Copy, Delete)."""

    def __init__(self, wb, sheets):
        self.wb, self.sheets = wb, list(sheets)

    def Copy(self, *args, **kw):
        """As Excel: positional (Before, After) only; into another workbook a sheet with a table is refused (Excel can
        crash), a taken name becomes 'name (2)' silently, a formula pointing at a sheet left behind leaves a link to
        the source file, and the destination workbook is activated on the first copy."""
        assert not kw, "Worksheets.Copy: positional (Before, After) only"
        args = _given(args)
        before = args[0] if args else None
        after = args[1] if len(args) > 1 else None
        anchor = before if before is not None else after
        dest = anchor.wb
        if dest is not self.wb and any(s.ListObjects for s in self.sheets):
            raise ComError("a sheet with a table copied between workbooks (Excel can crash)")
        taken = {s.Name.lower() for s in dest.order}
        k = dest.order.index(anchor) + (0 if before is not None else 1)
        copies = []
        for s in self.sheets:
            name, i = s.Name, 2
            while name.lower() in taken:
                name, i = f"{s.Name} ({i})", i + 1
            taken.add(name.lower())
            c = Worksheet(dest, name)
            c.cells, c.colw, c.rowh = copy.deepcopy(s.cells), dict(s.colw), dict(s.rowh)
            c.prefixed, c.Visible, c.Tab.Color = set(s.prefixed), s.Visible, s.Tab.Color
            dest.order.insert(k, c)
            k += 1
            copies.append(c)
        left = {s.Name for s in self.wb.order} - {s.Name for s in self.sheets}
        formulas = [p.get("Value") for s in self.sheets for p in s.cells.values()]
        if dest is not self.wb and any(isinstance(f, str) and f.startswith("=") and
                                       any(f"'{n}'!" in f or f"{n}!" in f for n in left) for f in formulas):
            dest.links.append(self.wb.FullName)
        dest.Activate()
        dest.active = copies[0]
        _event(copies[0], "Copy", [s.Name for s in self.sheets])

    def Delete(self):
        """As Excel: with DisplayAlerts on, a delete asks the user first (the fake refuses instead of waiting); what
        pointed at the sheets shows #REF!."""
        if self.wb.Application.DisplayAlerts:
            raise ComError("a sheet delete with DisplayAlerts on waits for the user")
        for s in self.sheets:
            self.wb.order.remove(s)
        _deleted(self.wb, self.sheets)
        if self.wb.active in self.sheets:
            self.wb.active = self.wb.order[0]
        _event(self.sheets[0], "Delete", [s.Name for s in self.sheets])


_VISIBLE = {"visible": -1, "hidden": 0, "veryHidden": 2}


def _read_file(wb, path) -> list:
    """The sheets of an .xlsx as fake Worksheets of wb (values, visibility, tab colour, tables), in the file's order."""
    import openpyxl
    book = openpyxl.load_workbook(path)
    out = []
    for src in book.worksheets:
        ws = Worksheet(wb, src.title)
        ws.Visible = _VISIBLE[src.sheet_state]
        tab = src.sheet_properties.tabColor
        if tab is not None and tab.rgb:                    # 'FF1F3864' -> BGR, as Excel reads Tab.Color
            r, g, b = (int(tab.rgb[i:i + 2], 16) for i in (2, 4, 6))
            ws.Tab.Color = r + (g << 8) + (b << 16)
        for row in src.iter_rows():
            for c in row:
                if c.value is not None:
                    ws.cell((c.row, c.column))["Value"] = c.value
        for name, ref in src.tables.items():
            lo = ListObject(ws, Range(ws, [parse_area(ref)]))
            lo.Name = name
            ws.ListObjects.append(lo)
        for name, dn in src.defined_names.items():           # the sheet's own names, as Excel reads them
            ws.Names.Add(Name=name, RefersTo="=" + dn.attr_text, Visible=not dn.hidden)
        if src.auto_filter.ref:
            ws.AutoFilterMode, ws.autofilter = True, parse_area(src.auto_filter.ref)
        out.append(ws)
    return out


DIGIT_PX = {("Telenor", 11.0): 8}                 # the Normal fonts whose digit width the fake knows (else 7 px)


def _normal_of(wb, path) -> None:
    """An opened file's Normal font and every sheet's standard row height and digit width, as Excel reads them."""
    import openpyxl
    book = openpyxl.load_workbook(path)
    font = book._fonts[0]
    wb.normal = (font.name, float(font.sz))
    for ws, src in zip(wb.order, book.worksheets):
        ws.StandardHeight = float(src.sheet_format.defaultRowHeight or 15)
        ws.digit_px = DIGIT_PX.get(wb.normal, 7)


def _insert(dest, path, before, after):
    """wb.Sheets.Add(Type:=path): the file's sheets at the anchor (Before, else After, else before the active
    sheet), a taken name as 'name (2)'. The first new sheet is the workbook's active one; the application's active
    workbook does not change (live, 20.11.2026). A sheet with a table is refused before anything goes in: the fake's
    stand-in (Excel can crash when such a sheet moves between workbooks; movein refuses such a file before the
    insert, so what Excel itself does here was never measured). Returns the first new sheet."""
    new = _read_file(dest, path)
    if any(s.ListObjects for s in new):
        raise ComError("a sheet with a table inserted from a file (Excel can crash)")
    taken = {s.Name.lower() for s in dest.order}
    anchor = before if before is not None else after if after is not None else dest.active
    k = dest.order.index(anchor) + (0 if after is None else 1)
    for s in new:
        name, i = s.Name, 2
        while name.lower() in taken:
            name, i = f"{s.Name} ({i})", i + 1
        taken.add(name.lower())
        s.Name = name
        dest.order.insert(k, s)
        k += 1
    dest.active = new[0]
    _event(new[0], "Sheets.Add", [s.Name for s in new], str(path))
    return new[0]


class _Workbooks(list):
    """Application.Workbooks: every open workbook; Open reads a file (values, sheet visibility, tables)."""

    def __init__(self, app):
        super().__init__()
        self.app = app

    @property
    def Count(self):
        return len(self)

    def __call__(self, key):
        """Workbooks(name) or Workbooks(index), as Excel: an unknown name is an error."""
        if isinstance(key, int):
            return self[key - 1]
        for wb in self:
            if wb.Name.lower() == str(key).lower():
                return wb
        raise ComError(f"no workbook {key}")

    def Open(self, path, *args, **kw):
        """Positional only, as late-bound COM needs: (Filename, UpdateLinks, ReadOnly, ...)."""
        assert not kw, "Workbooks.Open: positional arguments only"
        wb = Workbook(("the file's sheets",), app=self.app, styles_ver=False)
        wb.order = _read_file(wb, path)
        _normal_of(wb, path)
        wb.active = wb.order[0]
        wb.Name, wb.FullName, wb.opened_with = Path(path).name, str(path), args
        _event(wb.order[0], "Open", str(path))
        return wb


class Application:
    def __init__(self):
        self.Workbooks = _Workbooks(self)
        self.events: list = []
        self.International = INTERNATIONAL
        self.ScreenUpdating, self.Calculation, self.DisplayAlerts, self.EnableEvents = True, -4105, True, True
        self.PrintCommunication = True
        self.StatusBar = False
        self.ActiveWorkbook = None
        self.Visible = True

    @property
    def ActiveSheet(self):
        return self.ActiveWorkbook.active if self.ActiveWorkbook else None

    @property
    def ActiveWindow(self):
        return self.ActiveWorkbook.window

    def Intersect(self, a, b):
        """The cells two ranges of one sheet share, or None."""
        both = set(a.cells()) & set(b.cells()) if a.ws is b.ws else set()
        return Range(a.ws, [(r, c, r, c) for r, c in sorted(both)]) if both else None

    def CentimetersToPoints(self, cm):
        return round(cm * 28.3464567, 4)


class Workbook:
    def __init__(self, names=("Sheet1",), app=None, styles_ver=True):
        self.Application = app or Application()
        self.order = [Worksheet(self, n) for n in names]
        self.Worksheets = _Sheets(self)
        self.Sheets = self.Worksheets                   # the chart sheets are in Charts
        self.Names = _Names(self)
        self.Name, self.FullName = "book.xlsx", r"C:\fake\book.xlsx"
        self.links: list[str] = []                       # LinkSources(xlExcelLinks)
        self.Charts = _Charts()                          # chart sheets (objects with a Name; ChartSheet)
        self.window = _Window(self)
        self.window_state: dict[str, dict] = {}
        self.selection: dict[str, str] = {}
        self.normal = ("Calibri", 11.0)                  # the Normal style's font
        self.saved = 0                                   # Save() calls
        self.ReadOnly = False
        self.active = self.order[0]
        self.Application.ActiveWorkbook = self
        if isinstance(getattr(self.Application, "Workbooks", None), _Workbooks):
            self.Application.Workbooks.append(self)
        if styles_ver:
            from smartass.xl import styling
            self.Names.Add(Name="_tnStylesVer", RefersTo=f'="{styling._styles_version()}"', Visible=False)

    def Windows(self, i):
        return self.window

    def Styles(self, name):
        assert name == "Normal", name
        return SimpleNamespace(Font=SimpleNamespace(Name=self.normal[0], Size=self.normal[1]))

    @property
    def ActiveSheet(self):
        return self.active

    def Activate(self):
        self.Application.ActiveWorkbook = self

    def Save(self):
        self.saved += 1

    def LinkSources(self, kind):
        assert kind == 1                                   # xlExcelLinks
        return tuple(self.links) or None

    def BreakLink(self, name, kind):
        assert kind == 1
        self.links.remove(name)
        _event(self.order[0], "BreakLink", name)

    def Close(self, save=None):
        """As Excel: the workbook leaves Workbooks; the one opened before it becomes active."""
        assert save is False, "a fake workbook is only closed without saving"
        app = self.Application
        app.Workbooks.remove(self)
        if app.ActiveWorkbook is self:
            app.ActiveWorkbook = app.Workbooks[-1] if app.Workbooks else None
        _event(self.order[0], "Close")


# ------------------------------------------------------------------ state + counting
def _key(v):
    if isinstance(v, list):                              # conditional format rules
        return repr([sorted(r.items()) for r in v])
    return v


def sheet_state(ws) -> dict:
    groups: dict[tuple, list] = {}
    values = []
    for (r, c), props in ws.cells.items():
        for k, v in props.items():
            if k == "Value":
                values.append(((r, c), repr(v)))
                continue
            groups.setdefault((k, _key(v)), []).append((r, c))
    fmt = {f"{k}={v!r}": compress(cells) for (k, v), cells in sorted(groups.items(), key=lambda x: repr(x[0]))}
    return {
        "values": hashlib.sha1(repr(sorted(values)).encode()).hexdigest()[:12],
        "formats": fmt,
        "widths": {col_letter(c): w for c, w in sorted(ws.colw.items())},
        "heights": _runs(ws.rowh),
        "tab": ws.Tab.Color,
        "page_setup": dict(sorted(ws.PageSetup.__dict__.items())),
        "names": {n: (x.RefersTo, x.Visible) for n, x in sorted(ws.Names.items.items())},
        "window": dict(sorted(ws.wb.window_state.get(ws.Name, {}).items())),
        "selection": ws.wb.selection.get(ws.Name),
        "visible": ws.Visible,
        "tables": [(lo.Name, rect_text(*lo.rect), lo.TableStyle) for lo in ws.ListObjects],
    }


def _runs(d: dict) -> dict:
    out: dict = {}
    for k in sorted(d):
        out.setdefault(d[k], []).append((k, k))
    return {v: compress([(r, 1) for r, _ in ks]).replace("A", "") for v, ks in out.items()}


def snapshot(wb) -> dict:
    return {"order": [s.Name for s in wb.order],
            "names": {n: x.RefersTo for n, x in sorted(wb.Names.items.items())},
            "sheets": {s.Name: sheet_state(s) for s in wb.order}}


def count_com(monkeypatch) -> list:
    """Count jobs.com calls (the writers' unit of work, as timing's 'com calls.*' counter), also where a module
    imported com by name."""
    import sys

    from smartass.xl import jobs
    calls: list = []
    real = jobs.com

    def counting(call, *a, **k):
        calls.append(call)
        return real(call, *a, **k)
    monkeypatch.setattr(jobs, "com", counting)
    for name, mod in list(sys.modules.items()):
        if name.startswith("smartass") and getattr(mod, "com", None) is real:
            monkeypatch.setattr(mod, "com", counting)
    return calls


class FixedDateTime(datetime.datetime):
    @classmethod
    def now(cls, tz=None):
        return cls(2026, 11, 3, 9, 30)


def freeze_time(monkeypatch, *modules) -> None:
    """datetime.now() fixed in these modules (they import datetime as dt): timestamps in cells stay comparable."""
    ns = SimpleNamespace(datetime=FixedDateTime, timedelta=datetime.timedelta, date=datetime.date)
    for mod in modules:
        monkeypatch.setattr(mod, "dt", ns)
