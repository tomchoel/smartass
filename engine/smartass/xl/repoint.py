"""A follow-up's formulas re-pointed (light and fast phase 3b, Task 3): a reference elsewhere in the workbook to one of
the question's parked sheets or tables ('old Q3 Answer'!E12, Data_Q3_old[Value]) rewritten to name its namesake in
the new answer, so it survives the drop of the old sheets instead of turning #REF!. Pure text, no Excel:
xl/movein.record_mentions finds the formulas (the verb mentions, inside Excel) and record_repoint writes what plan
gives here.

Exactly those references are rewritten. A sheet reference is a quoted name before "!" (a parked name always holds a
space - "old ..." - so Excel always quotes it), also one end of a 3D reference ('Summary:old Q3 Answer'!E12); a
table reference is the table's name as a word of its own outside brackets (Data_Q3_old[Value], ROWS(Data_Q3_old)).
Never rewritten: text inside a string literal, a sheet whose name merely contains a parked one ('old Q3 Answer 2'),
another workbook's sheet or name ([Other.xlsx], Other.xlsx!x), a sheet's own defined name (Summary!x), a column of
another table that happens to be called like a parked table (T[Data_Q3_old]), a longer name that contains one
(Data_Q3_old2). Only names with a namesake in the new answer are in the maps: a reference to a parked sheet the new
answer has no sheet for stays, and shows #REF! once that sheet is dropped (the old answer's cell is gone). The text
is the notation it was read in (Formula2, FormulaArray and a series formula in en-US, a name's RefersToLocal in the
user's): separators and function names are never touched, only the names."""
from __future__ import annotations

import os
from dataclasses import dataclass

PREFIX_MIN = 4                          # a shorter common start of the parked sheets' names is no use as a gate


@dataclass(frozen=True)
class Repoint:
    """One formula to rewrite: kind - "cell" (Range.Formula2), "array" (a legacy array's CurrentArray.FormulaArray),
    "series" (a chart on a worksheet: ref "i/j" is ChartObjects(i)'s SeriesCollection(j)), "chart" (a chart sheet's
    SeriesCollection(ref)) or "name" (Names.Item(ref).RefersToLocal; where is the name); where - the sheet, chart
    sheet or name; old and new - the formula's text before and after."""
    kind: str
    where: str
    ref: str
    old: str
    new: str

    @property
    def where_text(self) -> str:
        if self.kind in ("cell", "array"):
            return f"{self.where}!{self.ref}"
        if self.kind in ("series", "chart"):
            return f"a chart series on {self.where}"
        return f"the name {self.where}"


def _word(ch: str) -> bool:
    """A character of a name in a formula (a table, a defined name, a function, a cell address)."""
    return ch.isalnum() or ch in "_.\\?"


def _closing(text: str, i: int, quote: str) -> int:
    """The index of the quote that closes the one at i (a doubled quote is one inside), or len(text)."""
    j, n = i + 1, len(text)
    while j < n:
        if text[j] == quote:
            if j + 1 < n and text[j + 1] == quote:
                j += 2
                continue
            return j
        j += 1
    return n


def _bracket_end(text: str, i: int) -> int:
    """The index of the "]" that closes the "[" at i (nested brackets counted; inside a structured reference "'"
    escapes the next character), or len(text)."""
    depth, j, n = 0, i, len(text)
    while j < n:
        ch = text[j]
        if ch == "'":
            j += 2
            continue
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                return j
        j += 1
    return n


def _sheet(quoted: str, sheets: dict[str, str]) -> str | None:
    """The quoted sheet reference's new text ('...' without the "!"), or None when it names no parked sheet: one
    sheet or a 3D range of two, never another workbook's ("[" in it)."""
    name = quoted[1:-1].replace("''", "'")
    parts = name.split(":")
    if "[" in name or len(parts) > 2:
        return None
    new = [sheets.get(p.casefold(), p) for p in parts]
    if new == parts:
        return None
    return "'" + ":".join(new).replace("'", "''") + "'"


def repoint(formula: str, sheets: dict[str, str], tables: dict[str, str]) -> str:
    """formula with each reference to a parked sheet (sheets: parked name casefolded -> the new name) or table
    (tables: the same) naming the new one; the same text when it has none (see the module docstring)."""
    out: list[str] = []
    i, n = 0, len(formula)
    while i < n:
        ch = formula[i]
        if ch == '"':                                       # a string literal: as it is
            j = _closing(formula, i, '"')
            out.append(formula[i:j + 1])
            i = j + 1
        elif ch == "'":                                     # a quoted sheet name
            j = _closing(formula, i, "'")
            quoted = formula[i:j + 1]
            new = _sheet(quoted, sheets) if formula[j + 1:j + 2] == "!" else None
            out.append(new or quoted)
            i = j + 1
        elif ch == "[":                                     # a structured reference's columns, a workbook's name
            j = _bracket_end(formula, i)
            out.append(formula[i:j + 1])
            i = j + 1
        elif _word(ch):
            j = i
            while j < n and _word(formula[j]):
                j += 1
            word = formula[i:j]
            qualified = formula[i - 1:i] == "!" or formula[j:j + 1] == "!"     # a sheet's or book's own, or a sheet
            out.append(word if qualified else tables.get(word.casefold(), word))
            i = j
        else:
            out.append(ch)
            i += 1
    return "".join(out)


def scan_texts(sheets: list[str], tables: list[str]) -> list[str]:
    """What the verb mentions looks for (case ignored): the parked sheets' names - as their common start when they
    share one of PREFIX_MIN characters or more (fewer texts to compare each formula with; the rewrite decides
    exactly) - and each parked table's name."""
    common = os.path.commonprefix([s.casefold() for s in sheets]) if len(sheets) > 1 else ""
    found = [sheets[0][:len(common)]] if len(common) >= PREFIX_MIN else list(sheets)
    return found + [t for t in tables if t not in found]


def _scope(name: str) -> str | None:
    """The sheet a defined name belongs to ("Summary!x", "'old Q3 Answer'!x"), None for the workbook's own."""
    sheet, bang, _ = name.rpartition("!")
    if not bang:
        return None
    return sheet[1:-1].replace("''", "'") if sheet.startswith("'") else sheet


def plan(rows, sheets: dict[str, str], tables: dict[str, str], skip: list[str]) -> tuple[list[Repoint], list]:
    """What the verb mentions found (rows of (kind, where, ref, formula)) as the rewrites to make, each formula whose
    text changes, and the sheets it could not read as (sheet, why). A name of a skipped sheet's own (the parked and
    the new sheets': their print areas) goes with its sheet and is left out."""
    skipped = {s.casefold() for s in skip}
    out, unread = [], []
    for kind, where, ref, formula in rows:
        if kind == "unread":
            unread.append((where, formula))
            continue
        if kind == "name" and (_scope(where) or "").casefold() in skipped:
            continue
        if kind in ("cell", "array", "series") and where.casefold() in skipped:
            continue
        new = repoint(formula, sheets, tables)
        if new != formula:
            out.append(Repoint(kind, where, ref, formula, new))
    return out, unread
