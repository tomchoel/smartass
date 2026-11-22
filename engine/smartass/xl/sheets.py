"""Question sheets (numbering, rollback, parking), data-sheet metadata, stored specs/pulls and lookup formulas."""
from __future__ import annotations

import json
import re
from collections.abc import Callable
from dataclasses import dataclass

from . import jobs, target
from ..epm.grid import Axis, Layout, parse
from ..core.errors import InputError

# ------------------------------------------------------------------ question numbering and sheets
_Q = re.compile(r"^Q(\d+) (Answer|FinMax|FinLite|FinPlan|Lines)\b")   # our sheets only, not the user's "Q3 Forecast"
_LINES = re.compile(r"^Q(\d+) Lines\b")
_TOPIC = re.compile(r"^Q(\d+) [A-Z][A-Za-z0-9 &/-]{1,24}$")             # an answer's topic sheet ('Q3 By project')
_DRAWN_LINES = re.compile(r"^Q(\d+) Lines b\d+$")                         # a v2 lines block's sheet ('Q3 Lines b0')


def answer_sheets(names, q: int) -> list[str]:
    """The sheets a v2 answer Q<q> drew, in workbook order: 'Q<q> Answer', 'Q<q> Data' and every 'Q<q> Lines b<i>'
    (a lines block's). A user's own 'Q<q> Notes' or 'Q<q> Lines notes' is not one of them."""
    return [n for n in names if n in (f"Q{q} Answer", f"Q{q} Data")
            or ((m := _DRAWN_LINES.match(n)) is not None and int(m.group(1)) == q)]


def answer_tables(q: int, tables) -> list[str]:
    """Of these table names (the tables on the answer's sheets), the ones the answer made: 'Data_Q<q>' and every
    'Lines_Q<q>_<alias>'."""
    return [n for n in tables if n == f"Data_Q{q}" or n.startswith(f"Lines_Q{q}_")]


def question_of(name: str, names) -> int | None:
    """The question number of one of our sheets: the fixed kinds, or a topic sheet next to its 'Q<n> Answer'."""
    m = _Q.match(name)
    if m is None:
        m = _TOPIC.match(name)
        m = m if m and f"Q{m.group(1)} Answer" in names else None
    return int(m.group(1)) if m else None


def next_question(t: target.Target) -> int:
    return next_question_in([s.Name for s in t.wb.Worksheets])


def next_question_in(names: list[str]) -> int:
    """The next question number of a workbook with these sheets."""
    return max((n for x in names if (n := question_of(x, names)) is not None), default=0) + 1


def pending_drill_question(t: target.Target) -> int | None:
    """The highest q that has drilled 'Q<q> Lines…' sheets but no 'Q<q> Answer' yet (fin_drill ran before fin_analyze),
    else None."""
    names = [s.Name for s in t.wb.Worksheets]
    drilled = {int(m.group(1)) for n in names if (m := _LINES.match(n))}
    pending = [q for q in drilled if f"Q{q} Answer" not in names]
    return max(pending, default=None)


def rollback(t: target.Target) -> list[str]:
    """Delete the sheets this request added (after a failure), so no half-built analysis is left behind: one quiet
    batch (DisplayAlerts off meanwhile and the user's setting put back by the batch itself), each Delete an operation
    that may fail (a sheet that never came in), the status bar cleared. Runs even for an abandoned job (check=False).
    Returns the names deleted."""
    from .inproc.batch import Batch
    from .inproc.runner import run
    from .jobs import book_name
    if not t.sheets_added:                              # nothing came in: nothing to run
        return []
    b = Batch(calculation=False)
    spans = []
    with b.attempt():
        wbn = b.workbook(book_name(t))
        for name in reversed(t.sheets_added):
            start = b.position
            wbn.Worksheets(name).Delete()
            spans.append((name, range(start, b.position)))
        b.app.StatusBar = False
    try:
        failed = {i for i, _, _ in run(t.app, b, limit=30, check=False).skipped}
    except Exception:   # noqa: BLE001 - nothing could be deleted
        failed = set(range(b.position))
    removed = [name for name, span in spans if not failed & set(span) and 0 not in failed]
    t.sheets_added.clear()
    return removed


def park_sheets(t: target.Target, names: list[str]) -> dict[str, str]:
    """Rename these sheets out of the way ('~<name>'). Returns {parked name: original name}."""
    parked = {}
    for name in names:
        if not _sheet_exists(t, name):
            continue
        s = jobs.com(lambda name=name: t.wb.Worksheets(name))
        tmp = unique_sheet_name(t, ("~" + name)[:31])
        jobs.com(lambda s=s, tmp=tmp: setattr(s, "Name", tmp))
        parked[tmp] = name
    return parked


def restore_parked(t: target.Target, parked: dict[str, str]) -> None:
    for tmp, orig in parked.items():
        if _sheet_exists(t, tmp) and not _sheet_exists(t, orig):
            jobs.com(lambda tmp=tmp, orig=orig: setattr(t.wb.Worksheets(tmp), "Name", orig), check=False)


def _sheet_exists(t: target.Target, name: str) -> bool:
    return any(s.Name == name for s in t.wb.Worksheets)


def unique_sheet_name(t: target.Target, base: str) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", base)[:31]
    name, i = base, 2
    while _sheet_exists(t, name):
        suffix = f" {i}"
        name = base[:31 - len(suffix)] + suffix
        i += 1
    return name


def place_after(t: target.Target, after_q: int | None = None):
    """The sheet a new sheet of question `after_q` goes after (place_after_in), or None."""
    sheets = list(t.wb.Worksheets)
    name = place_after_in([s.Name for s in sheets], after_q)
    return next(s for s in sheets if s.Name == name) if name is not None else None


def place_after_in(names: list[str], after_q: int | None = None) -> str | None:
    """Of these sheet names, the one a new sheet of question `after_q` goes after: the last sheet of that question,
    else the last sheet that is not Sources / Sources SV / _svmeta; None when there is none (it then goes first)."""
    if after_q is not None:
        group = [n for n in names if question_of(n, names) == after_q]
        if group:
            return group[-1]
    others = [n for n in names if n not in ("Sources", "Sources SV", META_SHEET)]
    return others[-1] if others else None


def add_sheet(t: target.Target, name: str, after_q: int | None = None, color: int | None = None):
    """Add a sheet after the last sheet of question `after_q` (or at the end, before 'Sources'): place_after."""
    sheets = list(t.wb.Worksheets)
    anchor = place_after(t, after_q)
    # positional (Before, After): the keyword After= is silently ignored through late-bound COM here (measured
    # 15.10.2026: the sheet landed before the ACTIVE sheet instead). Only 'Sources'/meta sheets yet: go first.
    ws = jobs.com(lambda: t.wb.Worksheets.Add(None, anchor) if anchor is not None else t.wb.Worksheets.Add(sheets[0]))
    t.sheets_added.append(name)

    def name_and_tab():
        ws.Name = name
        if color is not None:
            ws.Tab.Color = color
    jobs.com(name_and_tab)
    return ws


# ------------------------------------------------------------------ data pull into the user's workbook
@dataclass
class Pull:
    sheet: str
    connection: str
    layout: Layout
    parsed: object
    seconds: float
    row_labels: dict          # tuple(requested members) -> tuple(labels as shown in the grid)
    col_labels: dict
    positional: bool = False  # labels not unique -> formulas reference cells by position


META_SHEET = "_svmeta"


def _name_hidden(ws, name: str) -> None:
    ws.Name = name
    ws.Visible = 0                                      # xlSheetHidden


SPEC_CHUNK = 30000                                      # a cell holds max 32 767 characters


@dataclass
class SpecRows:
    """Where the stored specs are (record_spec_rows): sheet - the hidden sheet is there; top - its used range's first
    row; width - its used range's columns; keys - column A of its used range, from top; free - the first row below
    column A's last value."""
    sheet: bool = False
    top: int = 2
    width: int = 0
    keys: tuple = ()
    free: int = 2

    def row(self, key: str) -> tuple[int, int]:
        """(the row of key, else the first free row; how many columns to write over: the used range's width when the
        key is there, so an earlier longer spec leaves no chunk)."""
        for i, k in enumerate(self.keys):
            if k == key:
                return self.top + i, self.width
        return self.free, 0


def record_spec_rows(b, wbn, key: str = "spec") -> Callable:
    """The hidden sheet's SpecRows, read in b (only column A of its used range: the specs themselves are not read)."""
    with b.attempt():                                   # no hidden sheet: nothing to read
        ws = wbn.Worksheets(META_SHEET)
        used = ws.UsedRange
        b.ret(used.Row, f"{key}.top")
        b.ret(used.Columns.Count, f"{key}.width")
        b.ret(used.Columns(1).Value2, f"{key}.keys")
        b.ret(ws.Cells(ws.Rows.Count, 1).End(-4162).Row, f"{key}.last")

    def rows(res) -> SpecRows:
        v = res.values
        if f"{key}.last" not in v:
            return SpecRows()
        if f"{key}.keys" not in v:                      # the sheet is there, its rows cannot be told
            raise ValueError(f"the stored specs on {META_SHEET} could not be read")
        keys = v.get(f"{key}.keys")
        keys = tuple(r[0] for r in keys) if isinstance(keys, tuple) else (keys,)
        return SpecRows(True, int(v[f"{key}.top"]), int(v[f"{key}.width"]), keys, int(v[f"{key}.last"]) + 1)
    return rows


def save_spec(t: target.Target, q: int, spec: dict, ids: dict[str, str] | None = None, owned: dict | None = None,
              create: bool = False, b=None, rows: SpecRows | None = None) -> bool:
    """Keep the question's last answer spec (unexpanded) + pull ids + the tool-owned range of its answer sheet
    ({"rows", "cols"}) in the hidden sheet, for patch-only follow-ups: one write (A and every 30 000-character chunk
    in one rectangle, as wide as the row was, so an earlier longer spec leaves no chunk) at the row rows says
    (record_spec_rows; else read in a batch of its own). Without the hidden sheet: nothing (v1), or create it first,
    the workbook's last sheet, hidden (create: True is returned then). Recorded into b, else run in a batch of its
    own."""
    from .inproc.batch import MISSING, Batch
    from .inproc.runner import run
    from .jobs import book_name
    if rows is None:
        rb = Batch(quiet=False)
        get = record_spec_rows(rb, rb.workbook(book_name(t)))
        rows = get(run(t.app, rb))
    created = not rows.sheet
    if created and not create:
        return False
    own = b is None
    b = Batch(calculation=False) if own else b
    wbn = b.workbook(book_name(t))
    key = f"SPEC Q{q}"
    r, width = rows.row(key)
    if created:
        wss = wbn.Worksheets
        ws = wss.Add(MISSING, wbn.Worksheets(wss.Count))       # positional: after the last sheet
        ws.Name = META_SHEET
        ws.Visible = 0                                  # xlSheetHidden
    else:
        ws = wbn.Worksheets(META_SHEET)
    blob = json.dumps({"spec": spec, "ids": ids or {}, "owned": owned}, ensure_ascii=False)
    row = ["'" + key, *("'" + blob[k:k + SPEC_CHUNK] for k in range(0, len(blob), SPEC_CHUNK))]
    row += [""] * (width - len(row))
    ws.Range(ws.Cells(r, 1), ws.Cells(r, len(row))).Value = (tuple(row),)
    if own:
        run(t.app, b)
    return created


def _stored_specs(t: target.Target) -> dict[int, dict]:
    """Every SPEC Q<n> row of the hidden sheet: {q: {"spec", "ids", "owned"}}."""
    if not _sheet_exists(t, META_SHEET):
        return {}
    vals = jobs.com(lambda: t.wb.Worksheets(META_SHEET).UsedRange.Value)
    vals = vals if isinstance(vals, tuple) else ((vals,),)
    out = {}
    for row in vals:
        m = re.match(r"^SPEC Q(\d+)$", row[0]) if row and isinstance(row[0], str) else None
        if m:
            d = json.loads("".join(str(x) for x in row[1:] if x))
            out[int(m.group(1))] = {"spec": d.get("spec") or {}, "ids": d.get("ids") or {}, "owned": d.get("owned")}
    return out


def load_answer(t: target.Target, q: int) -> dict | None:
    """The stored answer of question q: {"spec", "ids", "owned"} (owned None for answers before phase 5)."""
    return _stored_specs(t).get(q)


def stored_questions(t: target.Target) -> list[int]:
    """The questions with a stored answer spec v1 (drawn by fin_analyze), ascending."""
    return sorted(q for q, a in _stored_specs(t).items() if a["spec"].get("version") == 1)


def load_pulls(t: target.Target, q: int) -> dict:
    """Rebuild Pull objects of question q from the hidden metadata + the live data sheets."""
    from ..epm.ops import _used_values
    if not _sheet_exists(t, META_SHEET):
        return {}
    ws = jobs.com(lambda: t.wb.Worksheets(META_SHEET))
    vals = jobs.com(lambda: ws.UsedRange.Value)
    vals = vals if isinstance(vals, tuple) else ((vals,),)
    out = {}
    for row in vals:
        sheet = row[0]
        if not isinstance(sheet, str) or not sheet.startswith(f"Q{q} ") or not _sheet_exists(t, sheet):
            continue
        meta = json.loads("".join(str(x) for x in row[1:] if x))
        lay = Layout([tuple(x) for x in meta["pov"]], [Axis(d, m) for d, m in meta["rows"]],
                     [Axis(d, m) for d, m in meta["cols"]])
        parsed = parse(lay, _used_values(jobs.com(lambda: t.wb.Worksheets(sheet))), False, False)
        parsed.pov_problems[:0] = meta.get("warnings", [])
        out[sheet] = _make_pull(sheet, meta["connection"], lay, parsed, 0.0)
    return out


def _make_pull(name, connection, lay, parsed, secs) -> "Pull":
    row_map = {tuple(combo): tuple(parsed.row_labels[i]) for i, combo in enumerate(lay.row_tuples)
               if i < len(parsed.row_labels)}
    col_map = {tuple(combo): tuple(str(x).strip() if x is not None else "" for x in parsed.col_labels[j])
               for j, combo in enumerate(lay.col_tuples)}
    positional = len(set(row_map.values())) < len(row_map) or len(set(col_map.values())) < len(col_map)
    return Pull(name, connection, lay, parsed, secs, row_map, col_map, positional)


def lookup_formula(p: Pull, row_members: list[str], col_members: list[str]) -> str:
    """Live label-keyed lookup into a pull sheet (skill tn-formelstil), robust to rows moving on refresh:
        =XLOOKUP("C30 REVENUES (ac)",'Q1 FinMax'!$A:$A,XLOOKUP("Sep 26",'Q1 FinMax'!$A$2:$M$2,'Q1 FinMax'!$A:$M))
    Rows are whole columns (never a guessed $A$5:$A$400); the header row and the block share the exact data width,
    so the arrays always match. Several column dimensions -> XLOOKUP(1,(row2="a")*(row3="b"),...). Several ROW
    dimensions need bounded criteria (whole-column array tests take seconds per sheet) -> INDEX/XMATCH over trim
    refs ($A:.$A), whose index is independent of the block size. A missing label shows #N/A, never a wrong number."""
    rlab = p.row_labels.get(tuple(row_members))
    clab = p.col_labels.get(tuple(col_members))
    if rlab is None or clab is None:
        raise InputError(f"No cell for rows {row_members} / columns {col_members} in '{p.sheet}'. "
                         f"Available rows: {list(p.row_labels)[:6]} columns: {list(p.col_labels)[:6]}")
    s = "'" + p.sheet.replace("'", "''") + "'!"
    nr, nc = p.layout.nr, p.layout.nc
    from openpyxl.utils import get_column_letter as L
    if p.positional:                                   # duplicate labels: a label lookup could hit the wrong row
        i = p.layout.row_tuples.index(tuple(row_members))
        j = p.layout.col_tuples.index(tuple(col_members))
        return f"={s}${L(nr + 1 + j)}${2 + nc + i}"
    W = L(nr + len(p.layout.col_tuples))                 # last data column (suppression is off: width is fixed)
    lab = lambda x: '"' + str(x).strip().replace('"', '""') + '"'   # exact label (no TRIM in these formulas)
    if nc == 1:
        col_lookup, col_arr = lab(clab[0]), f"{s}$A$2:${W}$2"
    else:
        col_lookup, col_arr = "1", "*".join(f"({s}$A${2 + d}:${W}${2 + d}={lab(clab[d])})" for d in range(nc))
    if nr == 1:
        return f"=XLOOKUP({lab(rlab[0])},{s}$A:$A,XLOOKUP({col_lookup},{col_arr},{s}$A:${W}))"
    row_match = "XMATCH(1," + "*".join(f"({s}${L(1 + d)}:.${L(1 + d)}={lab(rlab[d])})" for d in range(nr)) + ")"
    col_match = (f"XMATCH({col_lookup},{col_arr})" if nc == 1 else f"XMATCH(1,{col_arr})")
    return f"=INDEX({s}$A:${W},{row_match},{col_match})"
