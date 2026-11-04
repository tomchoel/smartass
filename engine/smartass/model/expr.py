"""Expression trees: one tree per output cell, two interpreters.
evaluate() gives the number the model reports; to_excel() gives the live formula written on the sheet.
None means missing (no fact), never zero. Formulas follow tn-formelstil: SUMIFS on labels, SUM over ranges,
variance % empty when the base is zero."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable

from .frame import Frame

Addr = Callable[[str], str]


class Cells:
    """Lazily evaluated cell values by key; a Ref may point forwards or backwards."""

    def __init__(self, frame: Frame, exprs: dict[str, "Expr | None"]):
        self.frame, self.exprs = frame, exprs
        self.done: dict[str, float | None] = {}
        self._busy: set[str] = set()

    def get(self, key: str) -> float | None:
        if key in self.done:
            return self.done[key]
        if key not in self.exprs:
            raise KeyError(f"unknown cell {key}")
        if key in self._busy:
            raise ValueError(f"circular reference at {key}")
        self._busy.add(key)
        try:
            e = self.exprs[key]
            v = e.evaluate(self.frame, self) if e is not None else None
        finally:
            self._busy.discard(key)
        self.done[key] = v
        return v


class Expr:
    def evaluate(self, frame: Frame, cells: Cells) -> float | None:
        raise NotImplementedError

    def to_excel(self, table: str, addr: Addr, const: dict[str, str] | None = None) -> str:
        raise NotImplementedError


def _q(s: str) -> str:
    return '"' + s.replace('"', '""') + '"'


def _crit(s: str) -> str:
    """A SUMIFS criterion: * ? ~ are wildcards there, so escape them before quoting."""
    return _q(s.replace("~", "~~").replace("*", "~*").replace("?", "~?"))


@dataclass(frozen=True)
class Pick(Expr):
    filters: tuple[tuple[str, tuple[str, ...]], ...]

    def evaluate(self, frame, cells):
        return frame.total(dict(self.filters))

    def to_excel(self, table, addr, const=None):
        parts, multi = [f"{table}[Value]"], 0
        for col, vals in self.filters:
            if const and len(vals) == 1 and const.get(col) == vals[0]:
                continue                      # matches every row of the table
            if len(vals) == 1:
                crit = _crit(vals[0])
            else:
                multi += 1
                crit = "{" + ("," if multi == 1 else ";").join(_crit(v) for v in vals) + "}"
            parts += [f"{table}[{col}]", crit]
        if len(parts) == 1:
            return f"SUM({table}[Value])"
        f = "SUMIFS(" + ",".join(parts) + ")"
        return f"SUM({f})" if multi else f


@dataclass(frozen=True)
class Ref(Expr):
    key: str

    def evaluate(self, frame, cells):
        return cells.get(self.key)

    def to_excel(self, table, addr, const=None):
        return addr(self.key)


_A1 = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$")


def _ranges(addrs: list[str]) -> list[str]:
    runs: list[list[str]] = []
    last: tuple[str, int] | None = None
    for a in addrs:
        m = _A1.match(a)
        pos = (m.group(1), int(m.group(2))) if m else None
        if runs and pos and last and pos[0] == last[0] and pos[1] == last[1] + 1:
            runs[-1].append(a)
        else:
            runs.append([a])
        last = pos
    return [r[0] if len(r) == 1 else f"{r[0]}:{r[-1]}" for r in runs]


@dataclass(frozen=True)
class Sum(Expr):
    items: tuple[Expr, ...]

    def __post_init__(self):
        if not self.items:
            raise ValueError("Sum needs at least one item")

    def evaluate(self, frame, cells):
        vals = [e.evaluate(frame, cells) for e in self.items]
        return None if all(v is None for v in vals) else sum(v for v in vals if v is not None)

    def to_excel(self, table, addr, const=None):
        if self.items and all(isinstance(e, Ref) for e in self.items):
            args = _ranges([addr(e.key) for e in self.items])
        else:
            args = [e.to_excel(table, addr, const) for e in self.items]
        return "SUM(" + ",".join(args) + ")"


@dataclass(frozen=True)
class Sub(Expr):
    a: Expr
    b: Expr

    def evaluate(self, frame, cells):
        x, y = self.a.evaluate(frame, cells), self.b.evaluate(frame, cells)
        return None if x is None and y is None else (x or 0.0) - (y or 0.0)

    def to_excel(self, table, addr, const=None):
        def wrap(e: Expr) -> str:
            s = e.to_excel(table, addr, const)
            return f"({s})" if isinstance(e, Sub) else s
        return f"{wrap(self.a)}-{wrap(self.b)}"


@dataclass(frozen=True)
class Neg(Expr):
    a: Expr

    def evaluate(self, frame, cells):
        x = self.a.evaluate(frame, cells)
        return None if x is None else -x

    def to_excel(self, table, addr, const=None):
        s = self.a.to_excel(table, addr, const)
        return f"-({s})" if isinstance(self.a, Sub) else f"-{s}"


@dataclass(frozen=True)
class Ratio(Expr):
    """num / den, empty when the base is zero or missing; abs_base divides by |den| (variance %)."""
    num: Expr
    den: Expr
    abs_base: bool = False

    def evaluate(self, frame, cells):
        n, d = self.num.evaluate(frame, cells), self.den.evaluate(frame, cells)
        if n is None or d is None or d == 0:
            return None
        return n / (abs(d) if self.abs_base else d)

    def to_excel(self, table, addr, const=None):
        n, d = self.num.to_excel(table, addr, const), self.den.to_excel(table, addr, const)
        den = f"ABS({d})" if self.abs_base else f"({d})"
        return f'IF({d}=0,"",({n})/{den})'


@dataclass(frozen=True)
class Mul(Expr):
    """a x b (a concept's base x its key share, both cells); missing when either is missing."""
    a: Expr
    b: Expr

    def evaluate(self, frame, cells):
        x, y = self.a.evaluate(frame, cells), self.b.evaluate(frame, cells)
        return None if x is None or y is None else x * y

    def to_excel(self, table, addr, const=None):
        def wrap(e: Expr) -> str:
            s = e.to_excel(table, addr, const)
            return f"({s})" if isinstance(e, (Sub, Neg)) else s
        return f"{wrap(self.a)}*{wrap(self.b)}"


@dataclass(frozen=True)
class Const(Expr):
    """An input written as a number in its own cell (e.g. a 100 % key share), never inside a formula: the layout
    writes the value, not a formula."""
    value: float

    def evaluate(self, frame, cells):
        return self.value

    def to_excel(self, table, addr, const=None):
        raise TypeError("a Const is written as a value in its own cell, never inside a formula")


def pick(filters: dict[str, tuple[str, ...] | list[str]]) -> Expr:
    """A Pick; split into a Sum of Picks when more than two filters have several values (one SUMIFS can cross
    only a row array and a column array)."""
    items = []
    for c, v in filters.items():
        if isinstance(v, str):
            raise TypeError(f"filter {c}: give a tuple of values, not a string")
        items.append((c, tuple(dict.fromkeys(v))))
    for c, v in items:
        if not v:
            raise ValueError(f"filter {c} has no values")
    multis = [i for i, (_, v) in enumerate(items) if len(v) > 1]
    if len(multis) <= 2:
        return Pick(tuple(items))
    k = min(multis, key=lambda i: len(items[i][1]))
    col, vals = items[k]
    return Sum(tuple(pick({**dict(items), col: (v,)}) for v in vals))


def var_pct(cur: Expr, base: Expr) -> Expr:
    return Ratio(Sub(cur, base), base, abs_base=True)


def share(part: Expr, total: Expr) -> Expr:
    return Ratio(part, total)


@dataclass(frozen=True)
class LinePick(Expr):
    """Sum of AMOUNT over a line set's rows whose `column` (case-insensitive, trimmed) is one of `values`
    (column None: every row). Excel: SUMIFS on table Lines_Q<n>_<alias> (n from the Data table name)."""
    alias: str
    column: str | None
    values: tuple[str, ...] | None

    def evaluate(self, frame, cells):
        from .lines import AMOUNT, norm
        ls = frame.lines[self.alias]
        want = None if self.column is None else {norm(v) for v in self.values or ()}
        hit = [float(r.get(AMOUNT) or 0.0) for r in ls.rows
               if want is None or norm(r.get(self.column)) in want]
        return sum(hit) if hit else None

    def to_excel(self, table, addr, const=None):
        t = f"Lines_{table.split('_', 1)[1]}_{self.alias}"
        if self.column is None:
            return f"SUM({t}[AMOUNT])"
        vals = tuple(dict.fromkeys(self.values or ()))
        if len(vals) == 1:
            return f"SUMIFS({t}[AMOUNT],{t}[{self.column}],{_crit(vals[0])})"
        arr = "{" + ",".join(_crit(v) for v in vals) + "}"
        return f"SUM(SUMIFS({t}[AMOUNT],{t}[{self.column}],{arr}))"


@dataclass(frozen=True)
class CubeRef(Expr):
    """The cube value of the drilled cell, written next to the lines (sheet 'Q<n> Lines <alias>', D2)."""
    alias: str

    def evaluate(self, frame, cells):
        return frame.lines[self.alias].cube_value

    def to_excel(self, table, addr, const=None):
        from .lines import CUBE_CELL
        q = table.split("_", 1)[1]                      # 'Q3'
        return f"'{q} Lines {self.alias}'!{CUBE_CELL}"
