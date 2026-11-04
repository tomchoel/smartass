"""Formelsjekk for TN-arbeidsbøker (regler i SKILL.md).

CLI:  python formelsjekk.py <fil.xlsx> [--sheet NAVN ...] [--json]
      Avslutter med kode 1 hvis det finnes BLOCK-funn.

Kjernen (check_formula / check_cells) er ren Python (kun stdlib) og kan importeres av andre programmer.
openpyxl importeres først inne i check_workbook.
"""
from __future__ import annotations

import re
import sys
from dataclasses import dataclass, asdict
from typing import Iterable, Optional


@dataclass
class Finding:
    sheet: str
    cell: str
    rule: str
    severity: str  # "block" | "warn"
    message: str
    formula: str
    suggestion: Optional[str] = None


# ---------------------------------------------------------------- tokenizer

_SHEET = r"(?:\[[^\]]+\])?(?:'(?:[^']|'')+'|[A-Za-z_][\w\.]*)!"
_CELL = r"\$?[A-Za-z]{1,3}\$?\d+"
_COL = r"\$?[A-Za-z]{1,3}"
_ROW = r"\$?\d+"
_SEP = r"\.?:\.?"
_BODY = rf"(?:{_CELL}(?:{_SEP}{_CELL})?|{_COL}{_SEP}{_COL}|{_ROW}{_SEP}{_ROW})"
REF_RE = re.compile(rf"(?:{_SHEET})?{_BODY}(?![\w(\[!])")
SHEET_RE = re.compile(_SHEET)
NUM_RE = re.compile(r"(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?")
NAME_RE = re.compile(r"[A-Za-z_\\][\w\.\\]*")
ERR_RE = re.compile(r"#[A-Za-z0-9/]+[!?]?")
SEP_RE = re.compile(_SEP)
CELL_PIECE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
COL_PIECE = re.compile(r"^(\$?)([A-Za-z]{1,3})$")
ROW_PIECE = re.compile(r"^(\$?)(\d+)$")


@dataclass
class Tok:
    kind: str  # str ref num func name struct err op
    text: str
    start: int
    end: int


def tokenize(s: str) -> list[Tok]:
    """Tokeniser formelteksten (uten ledende =)."""
    toks: list[Tok] = []
    i, n = 0, len(s)
    while i < n:
        ch = s[i]
        if ch.isspace():
            i += 1
            continue
        if ch == '"':
            j = i + 1
            while j < n:
                if s[j] == '"':
                    if j + 1 < n and s[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            toks.append(Tok("str", s[i:j + 1], i, j + 1))
            i = j + 1
            continue
        m = REF_RE.match(s, i)
        if m:
            toks.append(Tok("ref", m.group(0), i, m.end()))
            i = m.end()
            continue
        if ch == "'":
            j = i + 1
            while j < n:
                if s[j] == "'":
                    if j + 1 < n and s[j + 1] == "'":
                        j += 2
                        continue
                    break
                j += 1
            toks.append(Tok("name", s[i:j + 1], i, j + 1))
            i = j + 1
            continue
        m = NUM_RE.match(s, i)
        if m:
            toks.append(Tok("num", m.group(0), i, m.end()))
            i = m.end()
            continue
        m = NAME_RE.match(s, i)
        if m or ch == "[":
            j = m.end() if m else i
            if j < n and s[j] == "[":
                depth = 0
                while j < n:
                    if s[j] == "[":
                        depth += 1
                    elif s[j] == "]":
                        depth -= 1
                        if depth == 0:
                            j += 1
                            break
                    j += 1
                toks.append(Tok("struct", s[i:j], i, j))
            else:
                kind = "func" if j < n and s[j] == "(" else "name"
                toks.append(Tok(kind, s[i:j], i, j))
            i = j
            continue
        if ch == "#":
            m = ERR_RE.match(s, i)
            if m:
                toks.append(Tok("err", m.group(0), i, m.end()))
                i = m.end()
                continue
        two = s[i:i + 2]
        if two in ("<>", "<=", ">="):
            toks.append(Tok("op", two, i, i + 2))
            i += 2
            continue
        toks.append(Tok("op", ch, i, i + 1))
        i += 1
    return toks


def _fname(t: Tok) -> str:
    n = t.text.upper()
    for p in ("_XLFN.", "_XLWS.", "_XLPM."):
        n = n.replace(p, "")
    return n


# ---------------------------------------------------------------- ref helpers

def split_ref(text: str):
    """-> (arkprefiks, [deler]) for et ref-token."""
    m = SHEET_RE.match(text)
    sheet = m.group(0) if m else ""
    return sheet, SEP_RE.split(text[len(sheet):])


def col_to_num(col: str) -> int:
    v = 0
    for c in col.upper():
        v = v * 26 + ord(c) - 64
    return v


def _single_cell(text: str):
    """(ark, kolabs, kol, radabs, rad) hvis ref er én celle, ellers None."""
    sheet, pieces = split_ref(text)
    if len(pieces) != 1:
        return None
    m = CELL_PIECE.match(pieces[0])
    if not m:
        return None
    return sheet, m.group(1), m.group(2).upper(), m.group(3), int(m.group(4))


def _structure(toks: list[Tok]):
    """Parentes-par og argumentspenn per kall. -> (close, args)"""
    close: dict[int, int] = {}
    stack: list[int] = []
    for i, t in enumerate(toks):
        if t.kind == "op" and t.text == "(":
            stack.append(i)
        elif t.kind == "op" and t.text == ")" and stack:
            close[stack.pop()] = i
    args: dict[int, list[tuple[int, int]]] = {}
    for op, cl in close.items():
        spans, depth, st = [], 0, op + 1
        for k in range(op + 1, cl):
            t = toks[k]
            if t.kind == "op":
                if t.text in ("(", "{"):
                    depth += 1
                elif t.text in (")", "}"):
                    depth -= 1
                elif t.text in (",", ";") and depth == 0:
                    spans.append((st, k))
                    st = k + 1
        spans.append((st, cl))
        args[op] = spans
    return close, args


# ---------------------------------------------------------------- single rules

_MODE_ARGS = {
    "SUBTOTAL": {0}, "AGGREGATE": {0, 1}, "MATCH": {2}, "XMATCH": {2, 3},
    "XLOOKUP": {3, 4, 5}, "ROUND": {1}, "ROUNDUP": {1}, "ROUNDDOWN": {1}, "TRUNC": {1},
    "FIXED": {1}, "INDEX": {1, 2, 3}, "CHOOSE": {0}, "VLOOKUP": {2, 3}, "HLOOKUP": {2, 3},
    "LARGE": {1}, "SMALL": {1}, "RANK": {2}, "PERCENTILE": {1}, "QUARTILE": {1},
    "EDATE": {1}, "EOMONTH": {1}, "OFFSET": {1, 2, 3, 4},
}
_ALL_ARGS_OK = {
    "LEFT", "RIGHT", "MID", "FIND", "SEARCH", "REPT", "TEXTBEFORE", "TEXTAFTER", "SEQUENCE",
    "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS", "SORT", "SORTBY", "GROUPBY", "PIVOTBY",
    "WRAPCOLS", "WRAPROWS", "DATE", "TIME", "EXPAND", "TEXTSPLIT", "SUBSTITUTE", "TEXT",
    "REPLACE", "VSTACK", "HSTACK", "TOCOL", "TOROW", "ROWS", "COLUMNS",
}
_ALLOWED_NUMS = {0.0, 1.0, 100.0}


def _check_constants(toks, f, cell, sheet, out):
    stack: list[tuple[str, int]] = []
    brace = 0
    bad: list[str] = []
    pending_func = None
    for i, t in enumerate(toks):
        if t.kind == "func":
            pending_func = _fname(t)
            continue
        if t.kind == "op":
            if t.text == "(":
                stack.append((pending_func or "", 0))
            elif t.text == ")":
                if stack:
                    stack.pop()
            elif t.text in (",", ";") and stack and brace == 0:
                fn, k = stack[-1]
                stack[-1] = (fn, k + 1)
            elif t.text == "{":
                brace += 1
            elif t.text == "}":
                brace = max(0, brace - 1)
        pending_func = None
        if t.kind != "num" or brace:
            continue
        try:
            v = float(t.text)
        except ValueError:
            continue
        if v in _ALLOWED_NUMS:
            continue
        if i > 0 and toks[i - 1].kind == "op" and toks[i - 1].text == "^":
            continue
        if stack:
            fn, k = stack[-1]
            if fn in _ALL_ARGS_OK or k in _MODE_ARGS.get(fn, ()):
                continue
        bad.append(t.text)
    if bad:
        out.append(Finding(sheet, cell, "CONSTANT", "block",
                           "Hard-coded number(s) in formula: " + ", ".join(dict.fromkeys(bad)) +
                           ". Put inputs in a labelled cell and reference it.",
                           f, "Move the number to a labelled input cell and reference it."))


def _sum_suggestion(refs: list[str]) -> str:
    items: list[str] = []
    run: list = []
    direction = None

    def flush():
        nonlocal run, direction
        if not run:
            return
        if len(run) == 1:
            items.append(run[0][0])
        else:
            first, last = run[0][0], run[-1][0]
            items.append(first + ":" + last[len(split_ref(last)[0]):])
        run, direction = [], None

    for r in refs:
        p = _single_cell(r)
        if not run:
            run = [(r, p)]
            continue
        q = run[-1][1]
        d = None
        if p[0] == q[0] and p[2] == q[2] and p[4] == q[4] + 1:
            d = "v"
        elif p[0] == q[0] and p[4] == q[4] and col_to_num(p[2]) == col_to_num(q[2]) + 1:
            d = "h"
        if d and direction in (None, d):
            direction = d
            run.append((r, p))
        else:
            flush()
            run = [(r, p)]
    flush()
    return "SUM(" + ",".join(items) + ")"


def _is_cell_tok(t: Tok) -> bool:
    return t.kind == "ref" and _single_cell(t.text) is not None


def _check_plus_chain(toks, f, cell, sheet, out):
    n = len(toks)
    i = 0
    while i < n:
        if _is_cell_tok(toks[i]):
            prev = toks[i - 1] if i > 0 else None
            if prev is not None and (prev.kind != "op" or prev.text in ("*", "/", "^", "-", "&", "%", ")", "+")):
                i += 1
                continue
            ops = []
            j = i
            while j < n and _is_cell_tok(toks[j]):
                nxt = toks[j + 1] if j + 1 < n else None
                if nxt is not None and nxt.kind == "op" and nxt.text in ("*", "/", "^", "%"):
                    break
                ops.append(toks[j].text)
                if nxt is not None and nxt.kind == "op" and nxt.text == "+" and j + 2 < n \
                        and _is_cell_tok(toks[j + 2]):
                    j += 2
                else:
                    break
            if len(ops) >= 3:
                out.append(Finding(sheet, cell, "PLUS_CHAIN", "block",
                                   f"{len(ops)} cell references added with +. Use SUM.",
                                   f, _sum_suggestion(ops)))
                i = j + 1
                continue
        i += 1


def _check_fixed_ranges(toks, f, cell, sheet, out):
    seen = set()
    for t in toks:
        if t.kind != "ref" or t.text in seen:
            continue
        sh, pieces = split_ref(t.text)
        if len(pieces) != 2:
            continue
        a, b = CELL_PIECE.match(pieces[0]), CELL_PIECE.match(pieces[1])
        if not a or not b or not a.group(3) or not b.group(3):
            continue
        span = abs(int(b.group(4)) - int(a.group(4))) + 1
        if span < 50:
            continue
        seen.add(t.text)
        sug = f"{sh}{a.group(1)}{a.group(2)}:.{b.group(1)}{b.group(2)}"
        out.append(Finding(sheet, cell, "FIXED_RANGE", "block",
                           f"Oversized fixed range {t.text} ({span} rows). "
                           "Use trim refs or whole columns.", f, sug + " (trim ref) or whole column"))


def _check_legacy(toks, args, f, cell, sheet, out):
    names = []
    for i, t in enumerate(toks):
        if t.kind != "func":
            continue
        fn = _fname(t)
        if fn in ("VLOOKUP", "HLOOKUP"):
            names.append(fn)
        elif fn == "MATCH" and i + 1 in args and len(args[i + 1]) >= 2:
            (a0, a1), (b0, b1) = args[i + 1][0], args[i + 1][1]
            first = toks[a0:a1]
            if len(first) == 1 and first[0].kind == "num" and first[0].text == "1":
                second = toks[b0:b1]
                has_trim = any(x.kind == "func" and _fname(x) == "TRIM" for x in second)
                one_times = len(second) >= 2 and second[0].kind == "num" and second[0].text == "1" \
                    and second[1].text == "*"
                if has_trim or one_times:
                    names.append("INDEX/MATCH(1,1*...)")
    if names:
        out.append(Finding(sheet, cell, "LEGACY_LOOKUP", "block",
                           "Legacy lookup (" + ", ".join(dict.fromkeys(names)) + "). Use XLOOKUP "
                           "(two-way, whole columns, header row and block of equal width; keyed on labels).", f,
                           "XLOOKUP($B13,'Data'!$A:$A,XLOOKUP(C$4,'Data'!$A$2:$Z$2,'Data'!$A:$Z))"))


def _check_zero_varpct(toks, args, f, s, cell, sheet, out):
    for i, t in enumerate(toks):
        if t.kind != "func" or _fname(t) != "IF" or i + 1 not in args:
            continue
        a = args[i + 1]
        if len(a) < 3:
            continue
        c0, c1, c2 = (toks[x:y] for x, y in a[:3])
        if not (len(c1) == 1 and c1[0].kind == "num" and float(c1[0].text) == 0):
            continue
        if len(c0) < 3 or c0[-1].text != "0" or c0[-2].text != "=":
            continue
        starts_n = c0[0].kind == "func" and _fname(c0[0]) == "N"
        divides = any(x.kind == "op" and x.text == "/" for x in c2)
        if not (starts_n or divides):
            continue
        if starts_n and len(c0) > 5 and c0[1].text == "(" and c0[-3].text == ")":
            base = s[c0[2].start:c0[-4].end]
        else:
            base = s[c0[0].start:c0[-3].end]
        rest = s[c2[0].start:c2[-1].end] if c2 else ""
        out.append(Finding(sheet, cell, "ZERO_VARPCT", "block",
                           "Variance % returns 0 when base is zero. Return blank, no N().", f,
                           f'=IF({base}=0,"",{rest})'))
        return


def check_formula(formula: str, cell: str = "", sheet: str = "") -> list[Finding]:
    """Enkeltformel-regler. Rent: ingen I/O."""
    out: list[Finding] = []
    if not isinstance(formula, str):
        return out
    f = formula
    s = f.strip()
    if s.startswith("="):
        s = s[1:]
    toks = tokenize(s)
    _, args = _structure(toks)
    _check_plus_chain(toks, f, cell, sheet, out)
    _check_fixed_ranges(toks, f, cell, sheet, out)
    _check_legacy(toks, args, f, cell, sheet, out)
    _check_zero_varpct(toks, args, f, s, cell, sheet, out)
    _check_constants(toks, f, cell, sheet, out)
    fnames = {_fname(t) for t in toks if t.kind == "func"}
    vol = sorted(fnames & {"INDIRECT", "OFFSET"})
    if vol:
        out.append(Finding(sheet, cell, "VOLATILE", "warn",
                           "Volatile function(s): " + ", ".join(vol) + ".", f,
                           "Use INDEX/XLOOKUP or structured refs."))
    if len(f) > 200 and "LET" not in fnames:
        out.append(Finding(sheet, cell, "LONG_FORMULA", "warn",
                           f"Formula is {len(f)} chars without LET. Name repeated parts with LET.", f,
                           "Wrap repeated sub-expressions in LET."))
    return out


# ---------------------------------------------------------------- row consistency

_A1 = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_CONST = "\x00CONST"


def _norm_piece(p: str, col: int, row: int) -> str:
    m = CELL_PIECE.match(p)
    if m:
        c = f"${m.group(2).upper()}" if m.group(1) else f"C[{col_to_num(m.group(2)) - col}]"
        r = f"${m.group(4)}" if m.group(3) else f"R[{int(m.group(4)) - row}]"
        return c + r
    m = COL_PIECE.match(p)
    if m:
        return f"${m.group(2).upper()}" if m.group(1) else f"C[{col_to_num(m.group(2)) - col}]"
    m = ROW_PIECE.match(p)
    if m:
        return f"${m.group(2)}" if m.group(1) else f"R[{int(m.group(2)) - row}]"
    return p


def normalise(formula: str, col: int, row: int) -> str:
    """Relativ (R1C1-lignende) form: relative refs som offset, absolutte beholdes."""
    s = formula.strip()
    if s.startswith("="):
        s = s[1:]
    parts = []
    for t in tokenize(s):
        if t.kind == "ref":
            sh, pieces = split_ref(t.text)
            parts.append(sh + ":".join(_norm_piece(p, col, row) for p in pieces))
        elif t.kind in ("func", "name", "op"):
            parts.append(t.text.upper())
        else:
            parts.append(t.text)
    return "\x1f".join(parts)


def _is_number(s: str) -> bool:
    try:
        float(s)
        return True
    except ValueError:
        return False


def check_cells(cells: Iterable[tuple[str, str, str]]) -> list[Finding]:
    """(ark, A1-celle, formel) -> funn fra enkeltregler + radkonsistens."""
    out: list[Finding] = []
    rows: dict[tuple[str, int], dict[int, tuple[str, str, str]]] = {}
    for sheet, cell, formula in cells:
        if not isinstance(formula, str):
            continue
        m = _A1.match(cell or "")
        is_f = formula.lstrip().startswith("=")
        if is_f:
            out.extend(check_formula(formula, cell, sheet))
        if not m:
            continue
        col, row = col_to_num(m.group(1)), int(m.group(2))
        if is_f:
            rows.setdefault((sheet, row), {})[col] = (normalise(formula, col, row), cell, formula)
        elif _is_number(formula.strip()):
            rows.setdefault((sheet, row), {})[col] = (_CONST, cell, formula)
    for (sheet, row), d in rows.items():
        if sum(1 for v in d.values() if v[0] != _CONST) < 3:
            continue
        for c in sorted(d):
            if c - 1 not in d or c + 1 not in d:
                continue
            left, mid, right = d[c - 1], d[c], d[c + 1]
            if left[0] != right[0] or left[0] == _CONST or mid[0] == left[0]:
                continue
            if d.get(c - 2, ("",))[0] == mid[0] or d.get(c + 2, ("",))[0] == mid[0]:
                continue  # vekslende mønster, ikke tilfeldig avvik
            what = "Hard-coded value" if mid[0] == _CONST else "Different formula"
            out.append(Finding(sheet, mid[1], "INCONSISTENT_ROW", "warn",
                               f"{what} between identical formulas in {left[1]} and {right[1]}.",
                               mid[2], f"Same pattern as {left[1]}: {left[2]}"))
    return out


# ---------------------------------------------------------------- workbook / CLI

def check_workbook(path: str, sheets: Optional[list[str]] = None) -> list[Finding]:
    import warnings
    import openpyxl
    warnings.filterwarnings("ignore")
    wb = openpyxl.load_workbook(path, data_only=False)
    cells: list[tuple[str, str, str]] = []
    for ws in wb.worksheets:
        if sheets and ws.title not in sheets:
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None:
                    continue
                if hasattr(v, "text"):  # ArrayFormula (Formula2 / dynamiske matriser)
                    v = v.text
                if isinstance(v, str):
                    if v.startswith("="):
                        cells.append((ws.title, c.coordinate, v))
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    cells.append((ws.title, c.coordinate, repr(v)))
    return check_cells(cells)


def main(argv=None) -> int:
    import argparse
    import json
    ap = argparse.ArgumentParser(description="TN formelsjekk")
    ap.add_argument("file")
    ap.add_argument("--sheet", action="append", help="kun dette arket (kan gjentas)")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    fs = check_workbook(a.file, a.sheet)
    if a.json:
        print(json.dumps([asdict(x) for x in fs], ensure_ascii=False, indent=1))
    else:
        cur = None
        for x in sorted(fs, key=lambda x: (x.sheet, x.severity != "block", x.rule)):
            if x.sheet != cur:
                cur = x.sheet
                print(f"\n== {cur}")
            print(f"  [{x.severity.upper()}] {x.cell} {x.rule}: {x.message}")
            print(f"      {x.formula[:160]}")
            if x.suggestion:
                print(f"      -> {x.suggestion[:160]}")
        nb = sum(1 for x in fs if x.severity == "block")
        print(f"\n{nb} BLOCK, {len(fs) - nb} WARN")
    return 1 if any(x.severity == "block" for x in fs) else 0


if __name__ == "__main__":
    sys.exit(main())
