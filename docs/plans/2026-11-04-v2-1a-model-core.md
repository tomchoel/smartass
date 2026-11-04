# smartass v2 · Plan 1A: model core

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the pure-Python model layer of smartass v2. It turns an `AnalysisSpec` (typed blocks, relative
periods) plus fetched facts into expression trees that give both the reported numbers and the Excel formulas. It also
covers the data guards, findings and the grid requests the data layer will fetch. No Excel, no network.

**Architecture:** This is a new package `engine/smartass/` beside v1's `engine/smartview_mcp/`.

- **`model/spec.py`:** the pydantic contract.
- **`model/periods.py`:** period algebra.
- **`model/frame.py`:** facts in long format, the future `Data_Q<n>` table.
- **`model/expr.py`:** expression trees with two interpreters, `evaluate` (the number) and `to_excel` (the
  formula).
- **`model/blocks/`:** turns each block into lines of trees.
- **`model/validate.py`:** hard guards and soft findings.
- **`model/analyze.py`:** ties it together.
- **Knowledge:** reached only through the `Resolver` protocol. A `StaticResolver` serves tests now. Sub-spec 2
  supplies the real one.

**Tech Stack:** Python 3.12, pydantic 2.14 (already a dependency), pytest, ruff (rules E9, F).

**Spec:** `docs/specs/2026-11-04-v2-1-engine-design.md` §2–§3 (model layer). Plan 1B covers data provider, render,
tools and runtime. It starts when this plan is merged.

**Conventions for every task:**
- Work in `C:\dev\smartass` on branch `v2/design`.
- Run commands from `C:\dev\smartass\engine` in Git Bash (`cd /c/dev/smartass/engine`).
- Tests live in `engine/tests/v2/`. The directory has no `__init__.py`, so every test file name must be globally
  unique (`test_v2_*.py`) and helpers are imported as `from v2util import ...`.
- Never run python through a stdin heredoc. Always run tests with `uv run pytest`.
- Commit messages end with:

  ```
  Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
  ```

## File structure

| File | Responsibility |
|---|---|
| `engine/smartass/__init__.py` | Package marker |
| `engine/smartass/model/__init__.py` | Package marker |
| `engine/smartass/model/periods.py` | Month labels, shifts, relative period expressions |
| `engine/smartass/model/frame.py` | `Fact`, `Frame` (long table, totals by filter) |
| `engine/smartass/model/expr.py` | `Pick`, `Ref`, `Sum`, `Sub`, `Ratio`, `pick()`, `var_pct()`, `share()`, `Cells` |
| `engine/smartass/model/spec.py` | `AnalysisSpec`, `Plan`, `Query`, nine block classes, `block_metrics`, `block_series`, `json_schema` |
| `engine/smartass/model/resolver.py` | `Resolver` protocol, `UnknownMetric`, `MetricDef`, `StaticResolver` |
| `engine/smartass/model/series.py` | Series (value / ly / budget / forecast) → scenario, version, months |
| `engine/smartass/model/context.py` | `Ctx`: query lookup and `pick()` for a metric and series |
| `engine/smartass/model/blocks/common.py` | `Line`, `BlockOut`, `ck`, `HEAD`, `BASES`, `row_cells`, `derived`, `entity_lines` |
| `engine/smartass/model/blocks/table.py` | `build_table`, `build_kpi` |
| `engine/smartass/model/blocks/trend.py` | `build_trend`, `build_matrix` |
| `engine/smartass/model/blocks/bridge.py` | `build_bridge` |
| `engine/smartass/model/blocks/ranking.py` | `build_ranking` |
| `engine/smartass/model/blocks/drivers.py` | `build_drivers` |
| `engine/smartass/model/blocks/__init__.py` | `BUILDERS` registry |
| `engine/smartass/model/requests.py` | `GridRequest`, `requests()` |
| `engine/smartass/model/validate.py` | `GuardError`, `Finding`, `guard_spec`, `guard_result`, `sign_findings` |
| `engine/smartass/model/findings.py` | `material`, `key_values` |
| `engine/smartass/model/analyze.py` | `Result`, `plan_requests`, `compute` |
| `engine/tests/v2/v2util.py` | Shared fixture: resolver, facts, spec dict |
| `engine/tests/v2/test_v2_*.py` | One test file per module |

---

### Task 0: Scaffold the package

**Files:**
- Create: `engine/smartass/__init__.py`
- Create: `engine/smartass/model/__init__.py`
- Create: `engine/smartass/model/blocks/__init__.py` (empty for now; filled in Task 13)
- Modify: `engine/pyproject.toml` (`[tool.hatch.build.targets.wheel]`)
- Test: `engine/tests/v2/test_v2_scaffold.py`

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_scaffold.py
def test_package_imports():
    import smartass.model

    assert smartass.model.__doc__.startswith("Pure model layer")
```

- [ ] **Step 2: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_scaffold.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass'`

- [ ] **Step 3: Create the package**

```python
# engine/smartass/__init__.py
"""smartass v2 engine, built beside smartview_mcp until it passes the test set."""
```

```python
# engine/smartass/model/__init__.py
"""Pure model layer: spec -> requests -> facts -> expression trees -> values. No Excel, no network."""
```

```python
# engine/smartass/model/blocks/__init__.py
"""Block builders: each typed block becomes lines of expression trees."""
```

In `engine/pyproject.toml` change:

```toml
[tool.hatch.build.targets.wheel]
packages = ["smartview_mcp"]
```

to:

```toml
[tool.hatch.build.targets.wheel]
packages = ["smartview_mcp", "smartass"]
```

Then run `uv sync`.

- [ ] **Step 4: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_scaffold.py -q`
Expected: `1 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass engine/pyproject.toml engine/tests/v2/test_v2_scaffold.py
git commit -m "feat(v2): scaffold smartass.model package

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 1: Period algebra

**Files:**
- Create: `engine/smartass/model/periods.py`
- Test: `engine/tests/v2/test_v2_periods.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/v2/test_v2_periods.py
import pytest

from smartass.model.periods import label, parse, resolve, shift, span

CLOSED = (2026, 9)


def test_label_and_parse_round_trip():
    assert label((2026, 9)) == "Sep 26"
    assert parse("Sep 26") == (2026, 9)
    assert parse(" Jan 25 ") == (2025, 1)


def test_parse_rejects_other_text():
    with pytest.raises(ValueError, match="use 'Sep 26'"):
        parse("September 2026")


def test_shift_crosses_years():
    assert shift((2026, 1), -1) == (2025, 12)
    assert shift((2026, 9), -12) == (2025, 9)
    assert shift((2025, 12), 1) == (2026, 1)


def test_span_inclusive_and_ordered():
    assert span((2025, 11), (2026, 2)) == [(2025, 11), (2025, 12), (2026, 1), (2026, 2)]
    with pytest.raises(ValueError):
        span((2026, 2), (2026, 1))


@pytest.mark.parametrize("expr, first, last, n", [
    ("closed", (2026, 9), (2026, 9), 1),
    ("closed-ytd", (2026, 1), (2026, 9), 9),
    ("closed-12m", (2025, 10), (2026, 9), 12),
    ("fy", (2026, 1), (2026, 12), 12),
    ("Mar 26", (2026, 3), (2026, 3), 1),
    ("Jan 26..Sep 26", (2026, 1), (2026, 9), 9),
])
def test_resolve(expr, first, last, n):
    got = resolve(expr, CLOSED)
    assert (got[0], got[-1], len(got)) == (first, last, n)
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_periods.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.periods'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/periods.py
"""Months as (year, month) tuples, labelled like Smart View ('Sep 26'), and the relative period expressions a
saved spec uses so it can roll forward to the next closed month."""
from __future__ import annotations

import re

Month = tuple[int, int]
ABBR = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
_LABEL = re.compile(r"([A-Z][a-z]{2}) (\d{2})")


def label(m: Month) -> str:
    return f"{ABBR[m[1] - 1]} {m[0] % 100:02d}"


def parse(text: str) -> Month:
    hit = _LABEL.fullmatch(text.strip())
    if not hit or hit.group(1) not in ABBR:
        raise ValueError(f"not a month label: {text!r} (use 'Sep 26')")
    return (2000 + int(hit.group(2)), ABBR.index(hit.group(1)) + 1)


def shift(m: Month, n: int) -> Month:
    i = m[0] * 12 + (m[1] - 1) + n
    return (i // 12, i % 12 + 1)


def span(a: Month, b: Month) -> list[Month]:
    if b < a:
        raise ValueError(f"period ends before it starts: {label(a)}..{label(b)}")
    out, cur = [], a
    while cur <= b:
        out.append(cur)
        cur = shift(cur, 1)
    return out


def resolve(expr: str, closed: Month) -> list[Month]:
    """'closed', 'closed-ytd', 'closed-12m', 'fy' (the closed month's year), 'Sep 26' or 'Jan 26..Sep 26'."""
    e = expr.strip()
    if e == "closed":
        return [closed]
    if e == "closed-ytd":
        return span((closed[0], 1), closed)
    if e == "closed-12m":
        return span(shift(closed, -11), closed)
    if e == "fy":
        return span((closed[0], 1), (closed[0], 12))
    if ".." in e:
        a, b = e.split("..", 1)
        return span(parse(a), parse(b))
    return [parse(e)]
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_periods.py -q`
Expected: `10 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/periods.py engine/tests/v2/test_v2_periods.py
git commit -m "feat(v2): period algebra with relative expressions

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Frame of facts

**Files:**
- Create: `engine/smartass/model/frame.py`
- Test: `engine/tests/v2/test_v2_frame.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/v2/test_v2_frame.py
import pytest

from smartass.model.frame import COLUMNS, Fact, Frame


def f(acc, ent, per, v):
    return Fact("FinMax", "fm", acc, ent, per, "Actual", "", v)


FR = Frame([f("ac.C50S", "E1", "Sep 26", -40.0), f("ac.C60", "E1", "Sep 26", -20.0),
            f("ac.C50S", "E2", "Sep 26", -10.0)])


def test_total_sums_matching_facts():
    assert FR.total({"Account": ("ac.C50S", "ac.C60"), "Entity": ("E1",)}) == -60.0
    assert FR.total({"Account": ("ac.C50S",)}) == -50.0


def test_total_is_none_when_nothing_matches():
    assert FR.total({"Entity": ("E9",)}) is None


def test_unknown_filter_column_is_an_error():
    with pytest.raises(KeyError, match="Colour"):
        FR.total({"Colour": ("red",)})


def test_rows_follow_columns():
    assert COLUMNS[-1] == "Value"
    assert FR.rows()[0] == ("FinMax", "fm", "ac.C50S", "E1", "Sep 26", "Actual", "", -40.0)
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_frame.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.frame'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/frame.py
"""Fetched facts as one long table: the same rows become the Excel table Data_Q<n>.
Every cube's periods are normalised to month labels ('Sep 26') by the data layer; FinMax has no version ('')."""
from __future__ import annotations

from dataclasses import astuple, dataclass

COLUMNS = ("Cube", "Query", "Account", "Entity", "Period", "Scenario", "Version", "Value")
FILTER_COLUMNS = COLUMNS[:-1]


@dataclass(frozen=True)
class Fact:
    cube: str
    query: str
    account: str
    entity: str
    period: str
    scenario: str
    version: str
    value: float


class Frame:
    def __init__(self, facts=()):
        self.facts: list[Fact] = list(facts)

    def total(self, filters: dict[str, tuple[str, ...]]) -> float | None:
        """Sum of the facts matching every filter (column -> allowed values); None when no fact matches."""
        bad = set(filters) - set(FILTER_COLUMNS)
        if bad:
            raise KeyError(f"unknown filter column(s): {sorted(bad)}")
        hit = [x for x in self.facts
               if all(getattr(x, col.lower()) in allowed for col, allowed in filters.items())]
        return sum(x.value for x in hit) if hit else None

    def rows(self) -> list[tuple]:
        return [astuple(x) for x in self.facts]
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_frame.py -q`
Expected: `4 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/frame.py engine/tests/v2/test_v2_frame.py
git commit -m "feat(v2): long-format fact frame

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Expression trees (value and formula from one tree)

**Files:**
- Create: `engine/smartass/model/expr.py`
- Test: `engine/tests/v2/test_v2_expr.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/v2/test_v2_expr.py
import pytest

from smartass.model.expr import Cells, Pick, Ref, Sub, Sum, pick, share, var_pct
from smartass.model.frame import Fact, Frame

FR = Frame([Fact("FinMax", "fm", "ac.C50S", "E1", "Sep 26", "Actual", "", -40.0),
            Fact("FinMax", "fm", "ac.C60", "E1", "Sep 26", "Actual", "", -20.0),
            Fact("FinMax", "fm", "ac.C50S", "E1", "Sep 25", "Actual", "", -30.0)])
ADDR = {"a": "D4", "b": "D5", "c": "D6", "x": "F9"}.__getitem__


def ev(e, exprs=None):
    return e.evaluate(FR, Cells(FR, exprs or {}))


def test_pick_single_values():
    e = pick({"Account": ("ac.C50S",), "Period": ("Sep 26",)})
    assert ev(e) == -40.0
    assert e.to_excel("T", ADDR) == 'SUMIFS(T[Value],T[Account],"ac.C50S",T[Period],"Sep 26")'


def test_pick_one_and_two_multi_filters_use_row_then_column_arrays():
    one = pick({"Account": ("ac.C50S", "ac.C60"), "Period": ("Sep 26",)})
    assert ev(one) == -60.0
    assert one.to_excel("T", ADDR) == 'SUM(SUMIFS(T[Value],T[Account],{"ac.C50S","ac.C60"},T[Period],"Sep 26"))'
    two = pick({"Account": ("ac.C50S", "ac.C60"), "Period": ("Sep 26", "Sep 25")})
    assert ev(two) == -90.0
    assert two.to_excel("T", ADDR) == ('SUM(SUMIFS(T[Value],T[Account],{"ac.C50S","ac.C60"},'
                                       'T[Period],{"Sep 26";"Sep 25"}))')


def test_pick_splits_when_three_filters_have_several_values():
    e = pick({"Account": ("ac.C50S", "ac.C60"), "Entity": ("E1", "E2"), "Period": ("Sep 26", "Sep 25")})
    assert isinstance(e, Sum) and len(e.items) == 2
    assert ev(e) == -90.0


def test_pick_quotes_are_doubled_and_empty_filter_refused():
    assert '"a""b"' in pick({"Entity": ('a"b',)}).to_excel("T", ADDR)
    with pytest.raises(ValueError, match="no values"):
        pick({"Entity": ()})


def test_missing_is_none_not_zero():
    assert ev(pick({"Entity": ("E9",)})) is None


def test_refs_sum_and_sub():
    exprs = {"a": pick({"Account": ("ac.C50S",), "Period": ("Sep 26",)}),
             "b": pick({"Account": ("ac.C60",), "Period": ("Sep 26",)}),
             "c": Sum((Ref("a"), Ref("b")))}
    cells = Cells(FR, exprs)
    assert cells.get("c") == -60.0
    assert exprs["c"].to_excel("T", ADDR) == "SUM(D4:D5)"
    d = Sub(Ref("a"), Sub(Ref("b"), Ref("c")))
    assert d.evaluate(FR, cells) == -40.0 - (-20.0 - -60.0)
    assert d.to_excel("T", ADDR) == "D4-(D5-D6)"


def test_sum_of_non_adjacent_refs_lists_them():
    assert Sum((Ref("a"), Ref("c"), Ref("x"))).to_excel("T", ADDR) == "SUM(D4,D6,F9)"


def test_var_pct_and_share_blank_on_zero_base():
    exprs = {"a": pick({"Account": ("ac.C50S",), "Period": ("Sep 26",)}),
             "b": pick({"Account": ("ac.C50S",), "Period": ("Sep 25",)})}
    cells = Cells(FR, exprs)
    v = var_pct(Ref("a"), Ref("b"))
    assert v.evaluate(FR, cells) == pytest.approx((-40.0 - -30.0) / 30.0)
    assert v.to_excel("T", ADDR) == 'IF(D5=0,"",(D4-D5)/ABS(D5))'
    assert share(Ref("a"), Ref("b")).to_excel("T", ADDR) == 'IF(D5=0,"",(D4)/D5)'
    zero = Cells(FR, {"z": pick({"Entity": ("E9",)}), "a": exprs["a"]})
    assert var_pct(Ref("a"), Ref("z")).evaluate(FR, zero) is None


def test_cells_detect_cycles_and_unknown_keys():
    cells = Cells(FR, {"a": Ref("b"), "b": Ref("a")})
    with pytest.raises(ValueError, match="circular"):
        cells.get("a")
    with pytest.raises(KeyError, match="unknown cell"):
        Cells(FR, {}).get("nope")


def test_pick_is_hashable_tree():
    assert isinstance(pick({"Entity": ("E1",)}), Pick)
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_expr.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.expr'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/expr.py
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
        e = self.exprs[key]
        v = e.evaluate(self.frame, self) if e is not None else None
        self._busy.discard(key)
        self.done[key] = v
        return v


class Expr:
    def evaluate(self, frame: Frame, cells: Cells) -> float | None:
        raise NotImplementedError

    def to_excel(self, table: str, addr: Addr) -> str:
        raise NotImplementedError


def _q(s: str) -> str:
    return '"' + s.replace('"', '""') + '"'


@dataclass(frozen=True)
class Pick(Expr):
    filters: tuple[tuple[str, tuple[str, ...]], ...]

    def evaluate(self, frame, cells):
        return frame.total(dict(self.filters))

    def to_excel(self, table, addr):
        parts, multi = [f"{table}[Value]"], 0
        for col, vals in self.filters:
            if len(vals) == 1:
                crit = _q(vals[0])
            else:
                multi += 1
                crit = "{" + ("," if multi == 1 else ";").join(_q(v) for v in vals) + "}"
            parts += [f"{table}[{col}]", crit]
        f = "SUMIFS(" + ",".join(parts) + ")"
        return f"SUM({f})" if multi else f


@dataclass(frozen=True)
class Ref(Expr):
    key: str

    def evaluate(self, frame, cells):
        return cells.get(self.key)

    def to_excel(self, table, addr):
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

    def evaluate(self, frame, cells):
        vals = [e.evaluate(frame, cells) for e in self.items]
        return None if all(v is None for v in vals) else sum(v for v in vals if v is not None)

    def to_excel(self, table, addr):
        if self.items and all(isinstance(e, Ref) for e in self.items):
            args = _ranges([addr(e.key) for e in self.items])
        else:
            args = [e.to_excel(table, addr) for e in self.items]
        return "SUM(" + ",".join(args) + ")"


@dataclass(frozen=True)
class Sub(Expr):
    a: Expr
    b: Expr

    def evaluate(self, frame, cells):
        x, y = self.a.evaluate(frame, cells), self.b.evaluate(frame, cells)
        return None if x is None and y is None else (x or 0.0) - (y or 0.0)

    def to_excel(self, table, addr):
        def wrap(e: Expr) -> str:
            s = e.to_excel(table, addr)
            return f"({s})" if isinstance(e, Sub) else s
        return f"{wrap(self.a)}-{wrap(self.b)}"


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

    def to_excel(self, table, addr):
        n, d = self.num.to_excel(table, addr), self.den.to_excel(table, addr)
        den = f"ABS({d})" if self.abs_base else d
        return f'IF({d}=0,"",({n})/{den})'


def pick(filters: dict[str, tuple[str, ...] | list[str]]) -> Expr:
    """A Pick; split into a Sum of Picks when more than two filters have several values (one SUMIFS can cross
    only a row array and a column array)."""
    items = [(c, tuple(v)) for c, v in filters.items()]
    for c, v in items:
        if not v:
            raise ValueError(f"filter {c} has no values")
    multis = [i for i, (_, v) in enumerate(items) if len(v) > 1]
    if len(multis) <= 2:
        return Pick(tuple(items))
    col, vals = items[multis[0]]
    return Sum(tuple(pick({**dict(items), col: (v,)}) for v in vals))


def var_pct(cur: Expr, base: Expr) -> Expr:
    return Ratio(Sub(cur, base), base, abs_base=True)


def share(part: Expr, total: Expr) -> Expr:
    return Ratio(part, total)
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_expr.py -q`
Expected: `10 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/expr.py engine/tests/v2/test_v2_expr.py
git commit -m "feat(v2): expression trees - one tree gives the value and the Excel formula

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: The analysis spec contract

**Files:**
- Create: `engine/smartass/model/spec.py`
- Create: `engine/tests/v2/v2util.py` (the spec part; Task 5 adds the resolver and facts)
- Test: `engine/tests/v2/test_v2_spec.py`

- [ ] **Step 1: Write the shared fixture and the failing tests**

```python
# engine/tests/v2/v2util.py
"""Shared v2 fixtures. Every figure is invented."""
import copy

CLOSED = (2026, 9)

BASE_SPEC = {
    "version": 2,
    "question": "How is Business opex developing vs last year?",
    "plan": {"depth": "standard", "purpose": "Explain Business opex vs LY", "hypotheses": ["Consultants up"]},
    "queries": [{"id": "fm", "cube": "FinMax", "metrics": ["Revenue", "Opex", "EBITDA"],
                 "entity": "un.SY.2BUS", "periods": "closed", "compare": ["ly"]}],
    "blocks": [
        {"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA", "Revenue"]},
        {"type": "table", "query": "fm", "title": "P&L"},
        {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]},
        {"type": "ranking", "query": "fm", "metric": "Opex", "top": 1},
        {"type": "drivers", "query": "fm", "metric": "Opex", "depth": 2},
        {"type": "commentary"},
    ],
}


def spec_dict(**over) -> dict:
    d = copy.deepcopy(BASE_SPEC)
    d.update(over)
    return d
```

```python
# engine/tests/v2/test_v2_spec.py
import pytest
from pydantic import ValidationError
from v2util import spec_dict

from smartass.model.spec import AnalysisSpec, DriversBlock, json_schema


def test_base_spec_is_valid():
    s = AnalysisSpec.model_validate(spec_dict())
    assert isinstance(s.blocks[4], DriversBlock) and s.blocks[4].depth == 2


def test_unknown_fields_are_refused():
    with pytest.raises(ValidationError, match="extra"):
        AnalysisSpec.model_validate(spec_dict(colour="red"))


def test_purpose_required_unless_quick():
    with pytest.raises(ValidationError, match="purpose is required"):
        AnalysisSpec.model_validate(spec_dict(plan={"depth": "standard"}))
    AnalysisSpec.model_validate(spec_dict(plan={"depth": "quick"}))


def test_budget_only_in_finplan():
    d = spec_dict()
    d["queries"][0]["compare"] = ["ly", "budget"]
    with pytest.raises(ValidationError, match="only in FinPlan"):
        AnalysisSpec.model_validate(d)


def test_block_must_reference_a_known_query_and_its_metrics_and_series():
    d = spec_dict(blocks=[{"type": "ranking", "query": "xx", "metric": "Opex"}])
    with pytest.raises(ValidationError, match="unknown query 'xx'"):
        AnalysisSpec.model_validate(d)
    d = spec_dict(blocks=[{"type": "ranking", "query": "fm", "metric": "Capex"}])
    with pytest.raises(ValidationError, match="not in query 'fm'"):
        AnalysisSpec.model_validate(d)
    d = spec_dict()
    d["queries"][0]["compare"] = []
    with pytest.raises(ValidationError, match="compare"):
        AnalysisSpec.model_validate(d)


def test_table_shape_rules():
    bad = [{"type": "table", "query": "fm", "columns": ["value", "var"]},
           {"type": "table", "query": "fm", "columns": ["value", "share"]},
           {"type": "table", "query": "fm", "rows": "entities", "columns": ["value"]}]
    for b in bad:
        with pytest.raises(ValidationError):
            AnalysisSpec.model_validate(spec_dict(blocks=[b]))


def test_data_driver_needs_evidence():
    b = {"type": "commentary", "drivers": [{"line": "Opex", "text": "x", "basis": "data"}]}
    with pytest.raises(ValidationError, match="needs evidence"):
        AnalysisSpec.model_validate(spec_dict(blocks=[b]))


def test_duplicate_query_ids_refused():
    d = spec_dict()
    d["queries"].append(dict(d["queries"][0]))
    with pytest.raises(ValidationError, match="unique"):
        AnalysisSpec.model_validate(d)


def test_json_schema_lists_every_block_type():
    text = str(json_schema())
    for t in ("table", "kpi_tiles", "trend", "bridge", "ranking", "matrix", "drivers", "lines", "commentary"):
        assert f"'{t}'" in text
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_spec.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.spec'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/spec.py
"""Analysis spec v2: the contract the finance agent writes. Typed blocks, relative periods, strict (unknown
fields refused). Cross-references (block -> query, metrics, series) are checked here; members and metrics are
checked against knowledge in validate.py."""
from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, ConfigDict, Field, model_validator

Compare = Literal["ly", "budget", "forecast"]
Series = Literal["value", "ly", "budget", "forecast"]
Column = Literal["value", "ly", "budget", "forecast", "var", "var_pct", "share"]
BASE_SERIES = ("ly", "budget", "forecast")
PLAN_ONLY = ("budget", "forecast")


class _M(BaseModel):
    model_config = ConfigDict(extra="forbid")


class Plan(_M):
    depth: Literal["quick", "standard", "deep"]
    purpose: str = ""
    hypotheses: list[str] = []
    audience: Literal["controller", "management", "analyst"] = "controller"

    @model_validator(mode="after")
    def _purpose(self) -> "Plan":
        if self.depth != "quick" and not self.purpose.strip():
            raise ValueError("plan.purpose is required unless plan.depth is 'quick'")
        return self


class Query(_M):
    id: str = Field(pattern=r"^[a-z][a-z0-9_]{0,15}$")
    cube: Literal["FinMax", "FinLite", "FinPlan"]
    metrics: list[str] = Field(min_length=1)
    entity: str
    periods: str
    compare: list[Compare] = []
    preliminary: bool = False

    @model_validator(mode="after")
    def _plan_only(self) -> "Query":
        bad = [c for c in self.compare if c in PLAN_ONLY]
        if bad and self.cube != "FinPlan":
            raise ValueError(f"query {self.id}: {', '.join(bad)} exist only in FinPlan "
                             "(actual vs plan is done in FinPlan)")
        return self


class TableBlock(_M):
    type: Literal["table"]
    query: str
    title: str = ""
    rows: Literal["metrics", "entities"] = "metrics"
    metrics: list[str] = []
    metric: str | None = None
    columns: list[Column] = ["value", "ly", "var", "var_pct"]
    top: int | None = Field(None, ge=1)
    other: bool = True
    sort: Literal["none", "desc", "asc", "abs"] = "none"

    @model_validator(mode="after")
    def _shape(self) -> "TableBlock":
        cols = self.columns
        has_base = any(c in cols for c in BASE_SERIES)
        if ("var" in cols or "var_pct" in cols) and not ("value" in cols and has_base):
            raise ValueError("table: var and var_pct need 'value' and one of ly/budget/forecast in columns")
        if "share" in cols and (self.rows != "entities" or "value" not in cols):
            raise ValueError("table: share needs rows 'entities' and a 'value' column")
        if self.rows == "entities" and (not self.metric or "value" not in cols):
            raise ValueError("table: rows 'entities' needs 'metric' and a 'value' column")
        return self


class KpiBlock(_M):
    type: Literal["kpi_tiles"]
    query: str
    title: str = ""
    metrics: list[str] = Field(min_length=1, max_length=6)
    compare: Compare = "ly"


class TrendBlock(_M):
    type: Literal["trend"]
    query: str
    title: str = ""
    metrics: list[str] = Field(min_length=1)
    series: list[Series] = ["value", "ly"]
    cumulative: bool = False
    chart: bool = True


class BridgeBlock(_M):
    type: Literal["bridge"]
    query: str
    title: str = ""
    total: str
    steps: list[str] = Field(min_length=1)
    base: Compare = "ly"


class RankingBlock(_M):
    type: Literal["ranking"]
    query: str
    title: str = ""
    metric: str
    top: int = Field(10, ge=1)
    compare: Compare | None = "ly"


class MatrixBlock(_M):
    type: Literal["matrix"]
    query: str
    title: str = ""
    metric: str


class DriversBlock(_M):
    type: Literal["drivers"]
    query: str
    title: str = ""
    metric: str
    base: Compare = "ly"
    depth: int = Field(2, ge=1, le=4)
    drill: bool = False


class LinesBlock(_M):
    type: Literal["lines"]
    title: str = ""
    cell: dict[str, str]
    group_by: Literal["supplier", "project", "account"] = "supplier"


class Driver(_M):
    line: str
    text: str
    basis: Literal["data", "hypothesis"]
    evidence: str = ""

    @model_validator(mode="after")
    def _evidence(self) -> "Driver":
        if self.basis == "data" and not self.evidence.strip():
            raise ValueError(f"driver '{self.line}': basis 'data' needs evidence (a cell key)")
        return self


class CommentaryBlock(_M):
    type: Literal["commentary"]
    title: str = "Commentary"
    key_points: list[str] = []
    drivers: list[Driver] = []


Block = Annotated[Union[TableBlock, KpiBlock, TrendBlock, BridgeBlock, RankingBlock, MatrixBlock, DriversBlock,
                        LinesBlock, CommentaryBlock], Field(discriminator="type")]


def block_metrics(b) -> list[str]:
    if isinstance(b, TableBlock):
        return list(b.metrics) + ([b.metric] if b.metric else [])
    if isinstance(b, (KpiBlock, TrendBlock)):
        return list(b.metrics)
    if isinstance(b, BridgeBlock):
        return [b.total, *b.steps]
    if isinstance(b, (RankingBlock, MatrixBlock, DriversBlock)):
        return [b.metric]
    return []


def block_series(b) -> list[str]:
    """Comparison series a block needs from its query (besides 'value')."""
    if isinstance(b, TableBlock):
        return [c for c in b.columns if c in BASE_SERIES]
    if isinstance(b, KpiBlock):
        return [b.compare]
    if isinstance(b, TrendBlock):
        return [s for s in b.series if s != "value"]
    if isinstance(b, (BridgeBlock, DriversBlock)):
        return [b.base]
    if isinstance(b, RankingBlock):
        return [b.compare] if b.compare else []
    return []


class AnalysisSpec(_M):
    version: Literal[2]
    question: str = Field(min_length=1)
    plan: Plan
    queries: list[Query] = Field(min_length=1)
    blocks: list[Block] = Field(min_length=1)

    @model_validator(mode="after")
    def _refs(self) -> "AnalysisSpec":
        ids = [q.id for q in self.queries]
        if len(ids) != len(set(ids)):
            raise ValueError("query ids must be unique")
        by_id = {q.id: q for q in self.queries}
        for i, b in enumerate(self.blocks):
            qid = getattr(b, "query", None)
            if qid is None:
                continue
            q = by_id.get(qid)
            if q is None:
                raise ValueError(f"block {i} ({b.type}): unknown query '{qid}'")
            missing = [m for m in block_metrics(b) if m not in q.metrics]
            if missing:
                raise ValueError(f"block {i} ({b.type}): metrics {missing} are not in query '{qid}' "
                                 "(add them to its metrics)")
            absent = [s for s in block_series(b) if s not in q.compare]
            if absent:
                raise ValueError(f"block {i} ({b.type}): {absent} not in query '{qid}' compare")
        return self


def json_schema() -> dict:
    return AnalysisSpec.model_json_schema()
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_spec.py -q`
Expected: `9 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/spec.py engine/tests/v2/v2util.py engine/tests/v2/test_v2_spec.py
git commit -m "feat(v2): typed analysis spec contract with cross-reference checks

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: Resolver protocol and the test fixture

**Files:**
- Create: `engine/smartass/model/resolver.py`
- Modify: `engine/tests/v2/v2util.py` (append resolver and facts)
- Test: `engine/tests/v2/test_v2_resolver.py`

- [ ] **Step 1: Extend the fixture and write the failing tests**

Append to `engine/tests/v2/v2util.py`:

```python
from smartass.model.frame import Fact, Frame  # noqa: E402
from smartass.model.resolver import MetricDef, StaticResolver  # noqa: E402

TREE = {"un.SY.2BUS": ("un.SY.4BUS_B", "un.SY.G910"), "un.SY.4BUS_B": ("un.SY.B1", "un.SY.B2")}
LEAVES = {  # period -> entity -> account -> value
    "Sep 26": {"un.SY.B1": {"ac.C30": 100.0, "ac.C50S": -40.0, "ac.C60": -20.0},
               "un.SY.B2": {"ac.C30": 30.0, "ac.C50S": -10.0, "ac.C60": -5.0},
               "un.SY.G910": {"ac.C30": 20.0, "ac.C50S": -8.0, "ac.C60": -2.0}},
    "Sep 25": {"un.SY.B1": {"ac.C30": 90.0, "ac.C50S": -35.0, "ac.C60": -15.0},
               "un.SY.B2": {"ac.C30": 32.0, "ac.C50S": -12.0, "ac.C60": -6.0},
               "un.SY.G910": {"ac.C30": 18.0, "ac.C50S": -8.0, "ac.C60": -1.0}},
}


def resolver(**kw) -> StaticResolver:
    return StaticResolver(
        metrics=[MetricDef("Revenue", {"FinMax": ("ac.C30",), "FinPlan": ("Revenue",)}, sign="income"),
                 MetricDef("Opex", {"FinMax": ("ac.C50S", "ac.C60"), "FinPlan": ("OPEX",)}, sign="cost",
                           materiality=(5.0, 0.05)),
                 MetricDef("EBITDA", {"FinMax": ("ac.EBITDA",), "FinPlan": ("EBITDA",)}, sign="income"),
                 MetricDef("Draft thing", {"FinMax": ("ac.C99",)}, status="draft")],
        tree=TREE, labels={"un.SY.2BUS": "2BUS Business Norway"}, **kw)


def facts() -> Frame:
    """Leaves plus parents (EPM returns parent values), EBITDA = C30 + C50S + C60 per entity."""
    out = []
    for per, ents in LEAVES.items():
        vals = {e: dict(a) for e, a in ents.items()}
        for parent in ("un.SY.4BUS_B", "un.SY.2BUS"):
            vals[parent] = {acc: sum(vals[c][acc] for c in TREE[parent]) for acc in ("ac.C30", "ac.C50S", "ac.C60")}
        for ent, accs in vals.items():
            accs["ac.EBITDA"] = accs["ac.C30"] + accs["ac.C50S"] + accs["ac.C60"]
            out += [Fact("FinMax", "fm", acc, ent, per, "Actual", "", v) for acc, v in accs.items()]
    return Frame(out)
```

```python
# engine/tests/v2/test_v2_resolver.py
import pytest
from v2util import resolver

from smartass.model.resolver import UnknownMetric


def test_metric_members_per_cube_case_insensitive():
    r = resolver()
    assert r.members("opex", "FinMax") == ("ac.C50S", "ac.C60")
    assert r.members("Opex", "FinPlan") == ("OPEX",)


def test_raw_member_is_its_own_metric():
    r = resolver()
    assert r.members("ac.C30", "FinMax") == ("ac.C30",)
    assert r.sign("ac.C30") == "neutral" and r.status("ac.C30") == "approved"


def test_unknown_metric_and_missing_cube():
    r = resolver()
    with pytest.raises(UnknownMetric):
        r.members("Capex", "FinMax")
    with pytest.raises(UnknownMetric, match="FinPlan"):
        r.members("Draft thing", "FinPlan")


def test_tree_labels_and_status():
    r = resolver()
    assert r.children("un.SY.2BUS", "FinMax") == ("un.SY.4BUS_B", "un.SY.G910")
    assert r.children("un.SY.B1", "FinMax") == ()
    assert r.label("un.SY.2BUS", "FinMax") == "2BUS Business Norway"
    assert r.label("un.SY.B1", "FinMax") == "un.SY.B1"
    assert r.exists("un.SY.B2", "FinMax") and not r.exists("un.SY.ZZ", "FinMax")
    assert r.status("Draft thing") == "draft"
    assert r.materiality("Opex") == (5.0, 0.05)
    assert r.latest_forecast() == "Forecast Q326"
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_resolver.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.resolver'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/resolver.py
"""The model's only view of knowledge. Sub-spec 2 implements Resolver over knowledge.db; StaticResolver serves
tests and seeds."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Protocol

Sign = Literal["income", "cost", "neutral"]
Status = Literal["approved", "draft"]
DEFAULT_MATERIALITY = (5_000_000.0, 0.05)


class UnknownMetric(KeyError):
    pass


class Resolver(Protocol):
    def members(self, metric: str, cube: str) -> tuple[str, ...]: ...
    def sign(self, metric: str) -> Sign: ...
    def status(self, metric: str) -> Status: ...
    def materiality(self, metric: str) -> tuple[float, float]: ...
    def exists(self, member: str, cube: str) -> bool: ...
    def children(self, member: str, cube: str) -> tuple[str, ...]: ...
    def label(self, member: str, cube: str) -> str: ...
    def latest_forecast(self) -> str: ...


@dataclass
class MetricDef:
    name: str
    members: dict[str, tuple[str, ...]]
    sign: Sign = "neutral"
    status: Status = "approved"
    materiality: tuple[float, float] = DEFAULT_MATERIALITY


@dataclass
class StaticResolver:
    metrics: list[MetricDef]
    tree: dict[str, tuple[str, ...]] = field(default_factory=dict)
    labels: dict[str, str] = field(default_factory=dict)
    extra_members: set[str] = field(default_factory=set)
    forecast: str = "Forecast Q326"

    def __post_init__(self):
        self._m = {m.name.lower(): m for m in self.metrics}
        self._known = (set(self.extra_members) | set(self.tree)
                       | {c for cs in self.tree.values() for c in cs}
                       | {a for m in self.metrics for ms in m.members.values() for a in ms})

    def members(self, metric: str, cube: str) -> tuple[str, ...]:
        m = self._m.get(metric.lower())
        if m is None:
            if metric in self._known:
                return (metric,)
            raise UnknownMetric(metric)
        if cube not in m.members:
            raise UnknownMetric(f"{metric} has no {cube} definition")
        return m.members[cube]

    def sign(self, metric: str) -> Sign:
        m = self._m.get(metric.lower())
        return m.sign if m else "neutral"

    def status(self, metric: str) -> Status:
        m = self._m.get(metric.lower())
        return m.status if m else "approved"

    def materiality(self, metric: str) -> tuple[float, float]:
        m = self._m.get(metric.lower())
        return m.materiality if m else DEFAULT_MATERIALITY

    def exists(self, member: str, cube: str) -> bool:
        return member in self._known

    def children(self, member: str, cube: str) -> tuple[str, ...]:
        return self.tree.get(member, ())

    def label(self, member: str, cube: str) -> str:
        return self.labels.get(member, member)

    def latest_forecast(self) -> str:
        return self.forecast
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_resolver.py -q`
Expected: `4 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/resolver.py engine/tests/v2/v2util.py engine/tests/v2/test_v2_resolver.py
git commit -m "feat(v2): resolver protocol and static resolver

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: Series and the pick context

**Files:**
- Create: `engine/smartass/model/series.py`
- Create: `engine/smartass/model/context.py`
- Test: `engine/tests/v2/test_v2_context.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/v2/test_v2_context.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.context import Ctx
from smartass.model.series import months, scenario
from smartass.model.spec import AnalysisSpec, Query


def ctx():
    return Ctx(AnalysisSpec.model_validate(spec_dict()), resolver(), CLOSED, facts())


def test_scenarios_per_cube():
    r = resolver()
    assert scenario("FinMax", "ly", r) == ("Actual", "")
    assert scenario("FinPlan", "value", r) == ("Actual", "Working")
    assert scenario("FinPlan", "budget", r) == ("Target", "Final")
    assert scenario("FinPlan", "forecast", r) == ("Forecast Q326", "Final")
    with pytest.raises(ValueError, match="only in FinPlan"):
        scenario("FinMax", "budget", r)


def test_ly_shifts_months_back_a_year():
    q = Query(id="fm", cube="FinMax", metrics=["Opex"], entity="E", periods="closed-ytd")
    assert months(q, "ly", CLOSED)[0] == (2025, 1)
    assert months(q, "value", CLOSED, base=[(2026, 3)]) == [(2026, 3)]


def test_pick_value_and_ly_and_formula():
    c = ctx()
    q = c.query("fm")
    assert c.value(c.pick(q, "Opex", "value")) == -85.0
    assert c.value(c.pick(q, "Opex", "ly")) == -77.0
    assert c.value(c.pick(q, "Opex", "value", "un.SY.4BUS_B")) == -75.0
    assert c.value(c.pick(q, "Opex", "value", ("un.SY.4BUS_B", "un.SY.G910"))) == -85.0
    assert c.pick(q, "Opex", "value").to_excel("Data_Q1", str) == (
        'SUM(SUMIFS(Data_Q1[Value],Data_Q1[Cube],"FinMax",Data_Q1[Query],"fm",'
        'Data_Q1[Account],{"ac.C50S","ac.C60"},Data_Q1[Entity],"un.SY.2BUS",'
        'Data_Q1[Period],"Sep 26",Data_Q1[Scenario],"Actual"))')
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_context.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.context'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/series.py
"""A query's series (value, ly, budget, forecast) as scenario, version and months. Actual vs plan lives entirely in
FinPlan: budget and forecast exist only there (FinPlan actuals are version Working, plans version Final)."""
from __future__ import annotations

from .periods import Month, resolve, shift
from .resolver import Resolver
from .spec import Query


def scenario(cube: str, series: str, r: Resolver) -> tuple[str, str]:
    if cube == "FinPlan":
        return {"value": ("Actual", "Working"), "ly": ("Actual", "Working"), "budget": ("Target", "Final"),
                "forecast": (r.latest_forecast(), "Final")}[series]
    if series in ("budget", "forecast"):
        raise ValueError(f"{series} exists only in FinPlan")
    return ("Actual", "")


def months(q: Query, series: str, closed: Month, base: list[Month] | None = None) -> list[Month]:
    ms = list(base) if base is not None else resolve(q.periods, closed)
    return [shift(m, -12) for m in ms] if series == "ly" else ms
```

```python
# engine/smartass/model/context.py
"""Build context shared by the block builders: the spec, knowledge, the closed month and the fetched facts."""
from __future__ import annotations

from dataclasses import dataclass

from .expr import Cells, Expr, pick
from .frame import Frame
from .periods import Month, label
from .resolver import Resolver
from .series import months, scenario
from .spec import AnalysisSpec, Query


@dataclass
class Ctx:
    spec: AnalysisSpec
    resolver: Resolver
    closed: Month
    frame: Frame

    def query(self, qid: str) -> Query:
        return next(q for q in self.spec.queries if q.id == qid)

    def pick(self, q: Query, metric: str, series: str = "value", entity: str | tuple[str, ...] | None = None,
             base_months: list[Month] | None = None) -> Expr:
        scen, ver = scenario(q.cube, series, self.resolver)
        ents = (q.entity,) if entity is None else ((entity,) if isinstance(entity, str) else tuple(entity))
        f = {"Cube": (q.cube,), "Query": (q.id,), "Account": self.resolver.members(metric, q.cube),
             "Entity": ents, "Period": tuple(label(m) for m in months(q, series, self.closed, base_months)),
             "Scenario": (scen,)}
        if ver:
            f["Version"] = (ver,)
        return pick(f)

    def value(self, e: Expr) -> float | None:
        """Value of a self-contained tree (no Refs)."""
        return e.evaluate(self.frame, Cells(self.frame, {}))
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_context.py -q`
Expected: `3 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/series.py engine/smartass/model/context.py engine/tests/v2/test_v2_context.py
git commit -m "feat(v2): series and pick context

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: Block output and shared row builders

**Files:**
- Create: `engine/smartass/model/blocks/common.py`
- Test: `engine/tests/v2/test_v2_blocks_common.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/v2/test_v2_blocks_common.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.common import BlockOut, ck, entity_lines, evaluate, row_cells
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def ctx():
    return Ctx(AnalysisSpec.model_validate(spec_dict()), resolver(), CLOSED, facts())


def test_row_cells_value_ly_var_pct():
    c = ctx()
    q = c.query("fm")
    out = BlockOut("b0", "table", "t", ["Actual", "Last year", "Variance", "Variance %"])
    from smartass.model.blocks.common import Line
    out.lines.append(Line("m0", "Opex", row_cells(c, q, "Opex", ["value", "ly", "var", "var_pct"], "b0", "m0")))
    cells = evaluate([out], c.frame)
    assert [cells.get(ck("b0", "m0", j)) for j in range(3)] == [-85.0, -77.0, -8.0]
    assert cells.get(ck("b0", "m0", 3)) == pytest.approx(-8.0 / 77.0)


def test_entity_lines_top_other_total_and_check():
    c = ctx()
    q = c.query("fm")
    cols = ["value", "share"]
    out = BlockOut("b1", "table", "t", ["Actual", "Share %"], lines=entity_lines(c, q, "Opex", cols, "b1",
                                                                               top=1, sort="abs"))
    assert [ln.key for ln in out.lines] == ["e0", "other", "total", "check"]
    assert out.lines[0].label == "un.SY.4BUS_B"
    assert out.lines[2].label == "2BUS Business Norway" and out.lines[2].kind == "grand"
    cells = evaluate([out], c.frame)
    assert cells.get(ck("b1", "e0", 0)) == -75.0
    assert cells.get(ck("b1", "other", 0)) == -10.0
    assert cells.get(ck("b1", "total", 0)) == -85.0
    assert cells.get(ck("b1", "e0", 1)) == pytest.approx(75 / 85)
    assert cells.get(ck("b1", "check", 0)) == 0.0


def test_entity_lines_without_children_has_no_check():
    c = ctx()
    q = c.query("fm").model_copy(update={"entity": "un.SY.B1"})
    lines = entity_lines(c, q, "Opex", ["value"], "b2")
    assert [ln.key for ln in lines] == ["total"]
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_blocks_common.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.blocks.common'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/blocks/common.py
"""What every block builder produces (BlockOut: headers + lines of expression trees) and the shared builders for
metric rows and entity rows (children, top N, Other, parent total, a children = parent check)."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from ..context import Ctx
from ..expr import Cells, Expr, Ref, Sub, Sum, share, var_pct
from ..frame import Frame
from ..spec import Query

BASES = ("ly", "budget", "forecast")
HEAD = {"value": "Actual", "ly": "Last year", "budget": "Budget", "forecast": "Forecast", "var": "Variance",
        "var_pct": "Variance %", "share": "Share %"}


@dataclass
class Line:
    key: str
    label: str
    cells: list[Expr | None]
    kind: Literal["line", "sum", "grand", "check"] = "line"
    level: int = 0
    metric: str | None = None


@dataclass
class BlockOut:
    key: str
    type: str
    title: str
    headers: list[str]
    lines: list[Line] = field(default_factory=list)
    chart: str | None = None
    drills: list[dict] = field(default_factory=list)


def ck(block: str, line: str, col: int) -> str:
    return f"{block}.{line}.{col}"


def evaluate(outs: list[BlockOut], frame: Frame) -> Cells:
    exprs = {ck(o.key, ln.key, j): e for o in outs for ln in o.lines for j, e in enumerate(ln.cells)}
    cells = Cells(frame, exprs)
    for k in exprs:
        cells.get(k)
    return cells


def derived(c: str, cols: list[str], key: str, lkey: str, total_key: str | None = None) -> Expr | None:
    vi = cols.index("value")
    base = next(x for x in cols if x in BASES) if c in ("var", "var_pct") else None
    if c == "var":
        return Sub(Ref(ck(key, lkey, vi)), Ref(ck(key, lkey, cols.index(base))))
    if c == "var_pct":
        return var_pct(Ref(ck(key, lkey, vi)), Ref(ck(key, lkey, cols.index(base))))
    if c == "share":
        return share(Ref(ck(key, lkey, vi)), Ref(ck(key, total_key, vi))) if total_key else None
    raise ValueError(f"not a derived column: {c}")


def row_cells(ctx: Ctx, q: Query, metric: str, cols: list[str], key: str, lkey: str,
              entity: str | None = None, total_key: str | None = None) -> list[Expr | None]:
    return [ctx.pick(q, metric, c, entity) if c in ("value",) + BASES else derived(c, cols, key, lkey, total_key)
            for c in cols]


def entity_lines(ctx: Ctx, q: Query, metric: str, cols: list[str], key: str, *, top: int | None = None,
                 other: bool = True, sort: str = "none") -> list[Line]:
    r = ctx.resolver
    kids = list(r.children(q.entity, q.cube))

    def val(e: str) -> float:
        return ctx.value(ctx.pick(q, metric, "value", e)) or 0.0

    if sort != "none" or top:
        order = {"desc": lambda e: -val(e), "asc": val, "abs": lambda e: -abs(val(e)),
                 "none": lambda e: -abs(val(e))}[sort]
        kids.sort(key=order)
    shown, rest = (kids[:top], kids[top:]) if top else (kids, [])
    lines = [Line(f"e{i}", r.label(e, q.cube),
                  row_cells(ctx, q, metric, cols, key, f"e{i}", entity=e, total_key="total"), metric=metric)
             for i, e in enumerate(shown)]
    if rest and other:
        cells: list[Expr | None] = []
        for j, c in enumerate(cols):
            if c in ("value",) + BASES:
                cells.append(Sub(Ref(ck(key, "total", j)),
                                 Sum(tuple(Ref(ck(key, f"e{i}", j)) for i in range(len(shown))))))
            else:
                cells.append(derived(c, cols, key, "other", total_key="total"))
        lines.append(Line("other", "Other", cells, metric=metric))
    lines.append(Line("total", r.label(q.entity, q.cube),
                      row_cells(ctx, q, metric, cols, key, "total", entity=q.entity, total_key="total"),
                      kind="grand", metric=metric))
    if kids:
        vi = cols.index("value")
        check: list[Expr | None] = [None] * len(cols)
        check[vi] = Sub(ctx.pick(q, metric, "value", tuple(kids)), Ref(ck(key, "total", vi)))
        lines.append(Line("check", "Check: children = parent", check, kind="check"))
    return lines
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_blocks_common.py -q`
Expected: `3 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/blocks/common.py engine/tests/v2/test_v2_blocks_common.py
git commit -m "feat(v2): block output and shared metric/entity row builders

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 8: Table and KPI tiles

**Files:**
- Create: `engine/smartass/model/blocks/table.py`
- Test: `engine/tests/v2/test_v2_block_table.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/v2/test_v2_block_table.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.common import ck, evaluate
from smartass.model.blocks.table import build_kpi, build_table
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def run(blocks):
    s = AnalysisSpec.model_validate(spec_dict(blocks=blocks))
    c = Ctx(s, resolver(), CLOSED, facts())
    return c, s.blocks


def test_table_metric_rows_default_to_query_metrics():
    c, (b,) = run([{"type": "table", "query": "fm", "title": "P&L"}])
    out = build_table(b, c, "b0")
    assert out.headers == ["Actual", "Last year", "Variance", "Variance %"]
    assert [ln.label for ln in out.lines] == ["Revenue", "Opex", "EBITDA"]
    cells = evaluate([out], c.frame)
    assert [cells.get(ck("b0", "m2", j)) for j in range(3)] == [65.0, 63.0, 2.0]


def test_table_entity_rows():
    c, (b,) = run([{"type": "table", "query": "fm", "rows": "entities", "metric": "Opex",
                    "columns": ["value", "ly", "var", "share"], "sort": "abs"}])
    out = build_table(b, c, "b0")
    assert [ln.key for ln in out.lines] == ["e0", "e1", "total", "check"]
    cells = evaluate([out], c.frame)
    assert cells.get(ck("b0", "e0", 2)) == -7.0
    assert cells.get(ck("b0", "e1", 3)) == pytest.approx(10 / 85)


def test_kpi_tiles():
    c, (b,) = run([{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}])
    out = build_kpi(b, c, "b0")
    cells = evaluate([out], c.frame)
    assert out.type == "kpi_tiles"
    assert [cells.get(ck("b0", "k0", j)) for j in range(3)] == [65.0, 63.0, 2.0]
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_block_table.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.blocks.table'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/blocks/table.py
"""table: metric rows or entity rows (children of the query's entity) by columns; kpi_tiles: headline metrics
with one comparison."""
from __future__ import annotations

from ..context import Ctx
from ..spec import KpiBlock, TableBlock
from .common import HEAD, BlockOut, Line, entity_lines, row_cells


def build_table(b: TableBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cols = list(b.columns)
    out = BlockOut(key, "table", b.title or "Table", [HEAD[c] for c in cols])
    if b.rows == "entities":
        out.lines = entity_lines(ctx, q, b.metric, cols, key, top=b.top, other=b.other, sort=b.sort)
    else:
        for i, m in enumerate(b.metrics or q.metrics):
            out.lines.append(Line(f"m{i}", m, row_cells(ctx, q, m, cols, key, f"m{i}"), metric=m))
    return out


def build_kpi(b: KpiBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cols = ["value", b.compare, "var", "var_pct"]
    out = BlockOut(key, "kpi_tiles", b.title or "Key figures", [HEAD[c] for c in cols])
    for i, m in enumerate(b.metrics):
        out.lines.append(Line(f"k{i}", m, row_cells(ctx, q, m, cols, key, f"k{i}"), metric=m))
    return out
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_block_table.py -q`
Expected: `3 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/blocks/table.py engine/tests/v2/test_v2_block_table.py
git commit -m "feat(v2): table and kpi_tiles blocks

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 9: Trend and matrix

**Files:**
- Create: `engine/smartass/model/blocks/trend.py`
- Modify: `engine/tests/v2/v2util.py` (add a monthly frame)
- Test: `engine/tests/v2/test_v2_block_trend.py`

- [ ] **Step 1: Extend the fixture and write the failing tests**

Append to `engine/tests/v2/v2util.py`:

```python
def monthly_facts() -> Frame:
    """Opex for B1, B2, G910 and parents, Jul-Sep 26 and Jul-Sep 25: leaf value = -(month) per account."""
    out = []
    for yy in (26, 25):
        for mon, lab in ((7, "Jul"), (8, "Aug"), (9, "Sep")):
            per = f"{lab} {yy}"
            leaf = {e: {"ac.C50S": -float(mon), "ac.C60": -1.0} for e in ("un.SY.B1", "un.SY.B2", "un.SY.G910")}
            leaf["un.SY.4BUS_B"] = {a: leaf["un.SY.B1"][a] + leaf["un.SY.B2"][a] for a in ("ac.C50S", "ac.C60")}
            leaf["un.SY.2BUS"] = {a: leaf["un.SY.4BUS_B"][a] + leaf["un.SY.G910"][a] for a in ("ac.C50S", "ac.C60")}
            out += [Fact("FinMax", "fm", a, e, per, "Actual", "", v) for e, accs in leaf.items() for a, v in accs.items()]
    return Frame(out)
```

```python
# engine/tests/v2/test_v2_block_trend.py
from v2util import CLOSED, monthly_facts, resolver, spec_dict

from smartass.model.blocks.common import ck, evaluate
from smartass.model.blocks.trend import build_matrix, build_trend
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def run(block):
    d = spec_dict(blocks=[block])
    d["queries"][0]["periods"] = "Jul 26..Sep 26"
    s = AnalysisSpec.model_validate(d)
    return Ctx(s, resolver(), CLOSED, monthly_facts()), s.blocks[0]


def test_trend_months_series_and_total():
    c, b = run({"type": "trend", "query": "fm", "metrics": ["Opex"]})
    out = build_trend(b, c, "b0")
    assert out.headers == ["Opex Actual", "Opex Last year"] and out.chart == "line"
    assert [ln.label for ln in out.lines] == ["Jul 26", "Aug 26", "Sep 26", "Total"]
    cells = evaluate([out], c.frame)
    assert cells.get(ck("b0", "p0", 0)) == -24.0          # 3 entities x (-7 - 1)
    assert cells.get(ck("b0", "total", 0)) == -24.0 - 27.0 - 30.0
    assert cells.get(ck("b0", "p0", 1)) == -24.0          # Jul 25 has the same invented values


def test_trend_cumulative_has_no_total():
    c, b = run({"type": "trend", "query": "fm", "metrics": ["Opex"], "series": ["value"], "cumulative": True})
    out = build_trend(b, c, "b0")
    cells = evaluate([out], c.frame)
    assert [ln.key for ln in out.lines] == ["p0", "p1", "p2"]
    assert cells.get(ck("b0", "p2", 0)) == -81.0


def test_matrix_children_by_month_with_check():
    c, b = run({"type": "matrix", "query": "fm", "metric": "Opex"})
    out = build_matrix(b, c, "b0")
    assert out.headers == ["Jul 26", "Aug 26", "Sep 26"]
    assert [ln.key for ln in out.lines] == ["e0", "e1", "total", "check"]
    cells = evaluate([out], c.frame)
    assert cells.get(ck("b0", "e0", 2)) == -20.0          # 4BUS_B Sep: 2 x (-9 - 1)
    assert all(cells.get(ck("b0", "check", j)) == 0.0 for j in range(3))
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_block_trend.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.blocks.trend'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/blocks/trend.py
"""trend: one line per month of the query's periods, a column per metric x series (optionally cumulative);
matrix: children of the query's entity x months, with the parent total and a children = parent check."""
from __future__ import annotations

from ..context import Ctx
from ..expr import Expr, Ref, Sub, Sum
from ..periods import label, resolve
from ..spec import MatrixBlock, TrendBlock
from .common import HEAD, BlockOut, Line, ck


def build_trend(b: TrendBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    ms = resolve(q.periods, ctx.closed)
    combos = [(m, s) for m in b.metrics for s in b.series]
    out = BlockOut(key, "trend", b.title or "Trend", [f"{m} {HEAD[s]}" for m, s in combos],
                   chart="line" if b.chart else None)
    for i, mon in enumerate(ms):
        cells: list[Expr | None] = []
        for j, (m, s) in enumerate(combos):
            this = ctx.pick(q, m, s, base_months=[mon])
            cells.append(Sum((Ref(ck(key, f"p{i - 1}", j)), this)) if b.cumulative and i else this)
        out.lines.append(Line(f"p{i}", label(mon), cells))
    if not b.cumulative and len(ms) > 1:
        out.lines.append(Line("total", "Total", [Sum(tuple(Ref(ck(key, f"p{i}", j)) for i in range(len(ms))))
                                                 for j in range(len(combos))], kind="grand"))
    return out


def build_matrix(b: MatrixBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    ms = resolve(q.periods, ctx.closed)
    r = ctx.resolver
    kids = r.children(q.entity, q.cube)
    out = BlockOut(key, "matrix", b.title or b.metric, [label(m) for m in ms])
    for i, e in enumerate(kids):
        out.lines.append(Line(f"e{i}", r.label(e, q.cube), [ctx.pick(q, b.metric, "value", e, [m]) for m in ms],
                              metric=b.metric))
    out.lines.append(Line("total", r.label(q.entity, q.cube),
                          [ctx.pick(q, b.metric, "value", q.entity, [m]) for m in ms], kind="grand", metric=b.metric))
    if kids:
        out.lines.append(Line("check", "Check: children = parent",
                              [Sub(Sum(tuple(Ref(ck(key, f"e{i}", j)) for i in range(len(kids)))),
                                   Ref(ck(key, "total", j))) for j in range(len(ms))], kind="check"))
    return out
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_block_trend.py -q`
Expected: `3 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/blocks/trend.py engine/tests/v2/v2util.py engine/tests/v2/test_v2_block_trend.py
git commit -m "feat(v2): trend and matrix blocks

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 10: Bridge

**Files:**
- Create: `engine/smartass/model/blocks/bridge.py`
- Test: `engine/tests/v2/test_v2_block_bridge.py`

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_block_bridge.py
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.bridge import build_bridge
from smartass.model.blocks.common import ck, evaluate
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def test_bridge_from_ly_to_actual_with_other():
    s = AnalysisSpec.model_validate(spec_dict(blocks=[
        {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}]))
    c = Ctx(s, resolver(), CLOSED, facts())
    out = build_bridge(s.blocks[0], c, "b0")
    assert out.chart == "waterfall" and out.headers == ["Amount"]
    assert [ln.key for ln in out.lines] == ["start", "s0", "s1", "other", "end"]
    cells = evaluate([out], c.frame)
    got = {ln.key: cells.get(ck("b0", ln.key, 0)) for ln in out.lines}
    assert got == {"start": 63.0, "s0": 10.0, "s1": -8.0, "other": 0.0, "end": 65.0}
    addr = {ck("b0", k, 0): f"D{i + 4}" for i, k in enumerate(got)}.__getitem__
    assert out.lines[3].cells[0].to_excel("Data_Q1", addr) == "(D8-D4)-SUM(D5:D6)"
```

- [ ] **Step 2: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_block_bridge.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.blocks.bridge'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/blocks/bridge.py
"""bridge: waterfall from the base total (ly / budget / forecast) to the actual total, one step per metric's
variance, plus 'Other / unexplained' so the bridge always closes."""
from __future__ import annotations

from ..context import Ctx
from ..expr import Ref, Sub, Sum
from ..spec import BridgeBlock
from .common import HEAD, BlockOut, Line, ck


def build_bridge(b: BridgeBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    out = BlockOut(key, "bridge", b.title or f"{b.total} bridge", ["Amount"], chart="waterfall")
    out.lines.append(Line("start", f"{b.total} {HEAD[b.base]}", [ctx.pick(q, b.total, b.base)], kind="sum",
                          metric=b.total))
    for i, s in enumerate(b.steps):
        out.lines.append(Line(f"s{i}", s, [Sub(ctx.pick(q, s, "value"), ctx.pick(q, s, b.base))], metric=s))
    steps = Sum(tuple(Ref(ck(key, f"s{i}", 0)) for i in range(len(b.steps))))
    out.lines.append(Line("other", "Other / unexplained",
                          [Sub(Sub(Ref(ck(key, "end", 0)), Ref(ck(key, "start", 0))), steps)]))
    out.lines.append(Line("end", f"{b.total} {HEAD['value']}", [ctx.pick(q, b.total, "value")], kind="grand",
                          metric=b.total))
    return out
```

- [ ] **Step 4: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_block_bridge.py -q`
Expected: `1 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/blocks/bridge.py engine/tests/v2/test_v2_block_bridge.py
git commit -m "feat(v2): bridge block

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 11: Ranking

**Files:**
- Create: `engine/smartass/model/blocks/ranking.py`
- Test: `engine/tests/v2/test_v2_block_ranking.py`

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_block_ranking.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.common import ck, evaluate
from smartass.model.blocks.ranking import build_ranking
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def test_ranking_top_other_total_share_and_check():
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "ranking", "query": "fm", "metric": "Opex",
                                                       "top": 1}]))
    c = Ctx(s, resolver(), CLOSED, facts())
    out = build_ranking(s.blocks[0], c, "b0")
    assert out.headers == ["Actual", "Last year", "Variance", "Variance %", "Share %"]
    assert [ln.key for ln in out.lines] == ["e0", "other", "total", "check"]
    cells = evaluate([out], c.frame)
    assert cells.get(ck("b0", "e0", 0)) == -75.0
    assert cells.get(ck("b0", "other", 2)) == -1.0          # G910: -10 vs -9
    assert cells.get(ck("b0", "other", 4)) == pytest.approx(10 / 85)
    assert cells.get(ck("b0", "check", 0)) == 0.0


def test_ranking_without_compare():
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "ranking", "query": "fm", "metric": "Opex",
                                                       "compare": None}]))
    c = Ctx(s, resolver(), CLOSED, facts())
    assert build_ranking(s.blocks[0], c, "b0").headers == ["Actual", "Share %"]
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_block_ranking.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.blocks.ranking'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/blocks/ranking.py
"""ranking: the query entity's children ranked by size (absolute value), top N + Other + total, share %, and the
comparison when asked. Suppliers and projects come from drill lines (Plan 1B), not from the cube."""
from __future__ import annotations

from ..context import Ctx
from ..spec import RankingBlock
from .common import HEAD, BlockOut, entity_lines


def build_ranking(b: RankingBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    cols = ["value"] + ([b.compare, "var", "var_pct"] if b.compare else []) + ["share"]
    out = BlockOut(key, "ranking", b.title or f"Top {b.top} by {b.metric}", [HEAD[c] for c in cols])
    out.lines = entity_lines(ctx, q, b.metric, cols, key, top=b.top, other=True, sort="abs")
    return out
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_block_ranking.py -q`
Expected: `2 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/blocks/ranking.py engine/tests/v2/test_v2_block_ranking.py
git commit -m "feat(v2): ranking block

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 12: Drivers (automatic variance decomposition)

**Files:**
- Create: `engine/smartass/model/blocks/drivers.py`
- Test: `engine/tests/v2/test_v2_block_drivers.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/v2/test_v2_block_drivers.py
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.common import ck, evaluate
from smartass.model.blocks.drivers import build_drivers
from smartass.model.context import Ctx
from smartass.model.spec import AnalysisSpec


def run(**opts):
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "drivers", "query": "fm", "metric": "Opex", **opts}]))
    c = Ctx(s, resolver(), CLOSED, facts())
    return c, build_drivers(s.blocks[0], c, "b0")


def test_drivers_walk_the_biggest_contributors():
    c, out = run(depth=2)
    # 2BUS -8: 4BUS_B -7 covers 7/8 >= 80 % so G910 is not shown; 4BUS_B: B1 -10, then B2 +3 (10/13 < 80 %)
    assert [(ln.key, ln.label, ln.level) for ln in out.lines] == [
        ("top", "2BUS Business Norway", 0), ("d0", "un.SY.4BUS_B", 1),
        ("d0_0", "un.SY.B1", 2), ("d0_1", "un.SY.B2", 2)]
    cells = evaluate([out], c.frame)
    assert [cells.get(ck("b0", k, 2)) for k in ("top", "d0", "d0_0", "d0_1")] == [-8.0, -7.0, -10.0, 3.0]
    assert out.drills == []


def test_drivers_request_drills_at_the_leaves():
    _, out = run(depth=2, drill=True)
    assert out.drills == [{"entity": "un.SY.B1", "metric": "Opex", "query": "fm"},
                          {"entity": "un.SY.B2", "metric": "Opex", "query": "fm"}]


def test_depth_one_stops_at_children():
    _, out = run(depth=1)
    assert [ln.key for ln in out.lines] == ["top", "d0"]
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_block_drivers.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.blocks.drivers'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/blocks/drivers.py
"""drivers: decompose a metric's variance down the entity tree. At each level the children are ranked by absolute
variance and shown until they explain 80 % of the gross movement (at most 3); each shown child is expanded the same
way until `depth`. With drill, the leaves reached are returned as drill requests (Plan 1B turns them into ledger
lines), so commentary can rest on data."""
from __future__ import annotations

from ..context import Ctx
from ..expr import Sub
from ..spec import DriversBlock
from .common import HEAD, BlockOut, Line, row_cells

COVER = 0.8
MAX_PER_LEVEL = 3


def build_drivers(b: DriversBlock, ctx: Ctx, key: str) -> BlockOut:
    q = ctx.query(b.query)
    r = ctx.resolver
    cols = ["value", b.base, "var", "var_pct"]
    out = BlockOut(key, "drivers", b.title or f"Drivers of {b.metric}", [HEAD[c] for c in cols])

    def variance(entity: str) -> float:
        return ctx.value(Sub(ctx.pick(q, b.metric, "value", entity), ctx.pick(q, b.metric, b.base, entity))) or 0.0

    def add(entity: str, level: int, lkey: str) -> None:
        out.lines.append(Line(lkey, r.label(entity, q.cube), row_cells(ctx, q, b.metric, cols, key, lkey, entity),
                              level=level, metric=b.metric))

    def leaf(entity: str) -> None:
        if b.drill:
            out.drills.append({"entity": entity, "metric": b.metric, "query": q.id})

    def walk(entity: str, level: int, prefix: str) -> None:
        kids = r.children(entity, q.cube)
        if level >= b.depth or not kids:
            leaf(entity)
            return
        scored = sorted(((variance(k), k) for k in kids), key=lambda t: -abs(t[0]))
        gross = sum(abs(v) for v, _ in scored) or 1.0
        covered = 0.0
        for i, (v, k) in enumerate(scored):
            lkey = f"{prefix}{i}"
            add(k, level + 1, lkey)
            walk(k, level + 1, lkey + "_")
            covered += abs(v)
            if covered / gross >= COVER or i + 1 >= MAX_PER_LEVEL:
                break

    add(q.entity, 0, "top")
    walk(q.entity, 0, "d")
    return out
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_block_drivers.py -q`
Expected: `3 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/blocks/drivers.py engine/tests/v2/test_v2_block_drivers.py
git commit -m "feat(v2): drivers block - automatic variance decomposition with drill requests

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 13: Builder registry and grid requests

**Files:**
- Modify: `engine/smartass/model/blocks/__init__.py`
- Create: `engine/smartass/model/requests.py`
- Test: `engine/tests/v2/test_v2_requests.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/v2/test_v2_requests.py
from v2util import CLOSED, resolver, spec_dict

from smartass.model.blocks import BUILDERS
from smartass.model.requests import requests
from smartass.model.spec import AnalysisSpec


def test_registry_covers_the_computed_blocks():
    assert set(BUILDERS) == {"table", "kpi_tiles", "trend", "bridge", "ranking", "matrix", "drivers"}


def test_one_request_per_series_with_subtree_and_all_accounts():
    reqs = requests(AnalysisSpec.model_validate(spec_dict()), resolver(), CLOSED)
    assert [(x.query, x.periods, x.scenario) for x in reqs] == [("fm", ("Sep 26",), "Actual"),
                                                                ("fm", ("Sep 25",), "Actual")]
    assert reqs[0].accounts == ("ac.C30", "ac.C50S", "ac.C60", "ac.EBITDA")
    assert reqs[0].entities == ("un.SY.2BUS", "un.SY.4BUS_B", "un.SY.G910", "un.SY.B1", "un.SY.B2")


def test_no_entity_blocks_fetch_only_the_entity():
    d = spec_dict(blocks=[{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}])
    reqs = requests(AnalysisSpec.model_validate(d), resolver(), CLOSED)
    assert reqs[0].entities == ("un.SY.2BUS",)


def test_finplan_series_have_versions():
    d = spec_dict(blocks=[{"type": "kpi_tiles", "query": "fp", "metrics": ["EBITDA"], "compare": "budget"}])
    d["queries"] = [{"id": "fp", "cube": "FinPlan", "metrics": ["EBITDA"], "entity": "un.SY.2BUS",
                     "periods": "closed-ytd", "compare": ["budget", "forecast"]}]
    reqs = requests(AnalysisSpec.model_validate(d), resolver(), CLOSED)
    assert [(x.scenario, x.version) for x in reqs] == [("Actual", "Working"), ("Target", "Final"),
                                                       ("Forecast Q326", "Final")]
    assert len(reqs[0].periods) == 9
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_requests.py -q`
Expected: FAIL with `ImportError: cannot import name 'BUILDERS'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/blocks/__init__.py
"""Block builders: each typed block becomes lines of expression trees. 'lines' and 'commentary' are not computed
here: lines come from drill (Plan 1B), commentary is text the agent writes."""
from .bridge import build_bridge
from .drivers import build_drivers
from .ranking import build_ranking
from .table import build_kpi, build_table
from .trend import build_matrix, build_trend

BUILDERS = {"table": build_table, "kpi_tiles": build_kpi, "trend": build_trend, "bridge": build_bridge,
            "ranking": build_ranking, "matrix": build_matrix, "drivers": build_drivers}
```

```python
# engine/smartass/model/requests.py
"""What to fetch: one grid request per query and series (value + each comparison), with every account the
query's metrics need and the entity subtree as deep as its blocks go."""
from __future__ import annotations

from dataclasses import dataclass

from .periods import Month, label
from .resolver import Resolver
from .series import months, scenario
from .spec import AnalysisSpec, DriversBlock, MatrixBlock, RankingBlock, TableBlock


@dataclass(frozen=True)
class GridRequest:
    query: str
    cube: str
    accounts: tuple[str, ...]
    entities: tuple[str, ...]
    periods: tuple[str, ...]
    scenario: str
    version: str


def _depth(spec: AnalysisSpec, qid: str) -> int:
    d = 0
    for b in spec.blocks:
        if getattr(b, "query", None) != qid:
            continue
        if isinstance(b, DriversBlock):
            d = max(d, b.depth)
        elif isinstance(b, (RankingBlock, MatrixBlock)) or (isinstance(b, TableBlock) and b.rows == "entities"):
            d = max(d, 1)
    return d


def _subtree(r: Resolver, cube: str, root: str, depth: int) -> tuple[str, ...]:
    out, frontier = [root], [root]
    for _ in range(depth):
        frontier = [c for e in frontier for c in r.children(e, cube)]
        out += frontier
    return tuple(dict.fromkeys(out))


def requests(spec: AnalysisSpec, r: Resolver, closed: Month) -> list[GridRequest]:
    reqs = []
    for q in spec.queries:
        accounts = tuple(dict.fromkeys(a for m in q.metrics for a in r.members(m, q.cube)))
        ents = _subtree(r, q.cube, q.entity, _depth(spec, q.id))
        for s in ("value", *q.compare):
            scen, ver = scenario(q.cube, s, r)
            reqs.append(GridRequest(q.id, q.cube, accounts, ents, tuple(label(m) for m in months(q, s, closed)),
                                    scen, ver))
    return reqs
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_requests.py -q`
Expected: `4 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/blocks/__init__.py engine/smartass/model/requests.py engine/tests/v2/test_v2_requests.py
git commit -m "feat(v2): block registry and grid requests

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 14: Guards (hard on data, soft on method)

**Files:**
- Create: `engine/smartass/model/validate.py`
- Test: `engine/tests/v2/test_v2_validate.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/v2/test_v2_validate.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.common import BlockOut, Line, evaluate
from smartass.model.context import Ctx
from smartass.model.expr import pick
from smartass.model.frame import Fact, Frame
from smartass.model.spec import AnalysisSpec
from smartass.model.validate import GuardError, guard_result, guard_spec, sign_findings


def spec(**q_over):
    d = spec_dict(blocks=[{"type": "kpi_tiles", "query": "fm", "metrics": ["Opex"]}])
    d["queries"][0].update(q_over)
    return AnalysisSpec.model_validate(d)


def test_clean_spec_passes_without_findings():
    assert guard_spec(spec(), resolver(), CLOSED) == []


@pytest.mark.parametrize("over, match", [
    ({"entity": "un.SY.ZZ"}, "unknown entity"),
    ({"metrics": ["Opex", "Capex"]}, "unknown metric or member 'Capex'"),
    ({"metrics": ["Opex", "Draft thing"]}, "is a draft"),
    ({"periods": "Oct 26"}, "after the last closed month Sep 26"),
])
def test_hard_guards_stop_with_a_fix(over, match):
    with pytest.raises(GuardError, match=match) as e:
        guard_spec(spec(**over), resolver(), CLOSED)
    assert e.value.fix


def test_preliminary_month_is_allowed_with_a_warning():
    found = guard_spec(spec(periods="Oct 26", preliminary=True), resolver(), CLOSED)
    assert [f.severity for f in found] == ["warning"] and "preliminary" in found[0].text


def test_missing_last_year_is_a_soft_warning():
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "columns": ["value"]}])
    d["queries"][0]["compare"] = []
    found = guard_spec(AnalysisSpec.model_validate(d), resolver(), CLOSED)
    assert [f.text for f in found] == ["query fm has no last-year comparison"]


def test_check_lines_must_be_zero():
    fr = Frame([Fact("FinMax", "fm", "a", "e", "Sep 26", "Actual", "", 5.0)])
    out = BlockOut("b0", "ranking", "Top", ["Actual"],
                   lines=[Line("check", "Check: children = parent", [pick({"Account": ("a",)})], kind="check")])
    with pytest.raises(GuardError, match="off by 5"):
        guard_result([out], evaluate([out], fr))


def test_sign_findings_flag_positive_costs():
    s = spec()
    assert sign_findings(Ctx(s, resolver(), CLOSED, facts())) == []
    flipped = Frame([Fact(**{**f.__dict__, "value": -f.value}) for f in facts().facts])
    found = sign_findings(Ctx(s, resolver(), CLOSED, flipped))
    assert [f.severity for f in found] == ["critical"] and "Opex" in found[0].text
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_validate.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.validate'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/validate.py
"""Guards. Hard on data (stop, with the fix): unknown entity, unknown or draft metric, an open month without
'preliminary', a children = parent check that does not reconcile. Soft on method (findings for the critic): no
last-year comparison, preliminary months. A total whose sign contradicts the metric is Critical, not a stop
(FinPlan's sign convention is not verified yet, so FinPlan is not checked)."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from .blocks.common import BlockOut, ck
from .context import Ctx
from .expr import Cells
from .periods import Month, label, resolve
from .resolver import Resolver, UnknownMetric
from .spec import AnalysisSpec

TOLERANCE = 1.0  # NOK


class GuardError(Exception):
    def __init__(self, reason: str, fix: str):
        super().__init__(f"{reason} - {fix}")
        self.reason, self.fix = reason, fix


@dataclass
class Finding:
    severity: Literal["critical", "warning", "info"]
    text: str
    where: str = ""


def guard_spec(spec: AnalysisSpec, r: Resolver, closed: Month) -> list[Finding]:
    soft: list[Finding] = []
    for q in spec.queries:
        if not r.exists(q.entity, q.cube):
            raise GuardError(f"unknown entity {q.entity!r} in {q.cube}", "find the member with fin_lookup")
        for m in q.metrics:
            try:
                r.members(m, q.cube)
            except UnknownMetric:
                raise GuardError(f"unknown metric or member {m!r} in {q.cube}", "find it with fin_lookup") from None
            if r.status(m) != "approved":
                raise GuardError(f"metric {m!r} is a draft",
                                 "approve it (smartass knowledge approve) or use an approved metric")
        late = [label(x) for x in resolve(q.periods, closed) if x > closed]
        if late and not q.preliminary:
            raise GuardError(f"query {q.id}: {', '.join(late)} is after the last closed month {label(closed)}",
                             f"end the period at {label(closed)} or set preliminary: true")
        if late:
            soft.append(Finding("warning", f"query {q.id}: {', '.join(late)} is preliminary (month not closed)",
                                q.id))
        if spec.plan.depth != "quick" and q.cube != "FinPlan" and "ly" not in q.compare:
            soft.append(Finding("warning", f"query {q.id} has no last-year comparison", q.id))
    return soft


def guard_result(outs: list[BlockOut], cells: Cells) -> None:
    for o in outs:
        for ln in o.lines:
            if ln.kind != "check":
                continue
            for j, e in enumerate(ln.cells):
                v = cells.get(ck(o.key, ln.key, j)) if e is not None else None
                if v is not None and abs(v) > TOLERANCE:
                    raise GuardError(f"{o.title}: {ln.label} is off by {v:,.0f} NOK",
                                     "the children do not add up to the parent; check the entity tree with "
                                     "fin_lookup before using these numbers")


def sign_findings(ctx: Ctx) -> list[Finding]:
    out: list[Finding] = []
    for q in ctx.spec.queries:
        if q.cube == "FinPlan":
            continue
        for m in q.metrics:
            s = ctx.resolver.sign(m)
            if s == "neutral":
                continue
            v = ctx.value(ctx.pick(q, m, "value"))
            if v is None:
                continue
            if (s == "cost" and v > 0) or (s == "income" and v < 0):
                want = "negative" if s == "cost" else "positive"
                out.append(Finding("critical", f"{m} in query {q.id} is {v:,.0f}; a {s} metric should be {want} "
                                               f"in {q.cube}", q.id))
    return out
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_validate.py -q`
Expected: `9 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/validate.py engine/tests/v2/test_v2_validate.py
git commit -m "feat(v2): data guards and soft findings

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 15: Findings, `compute` and the end-to-end test

**Files:**
- Create: `engine/smartass/model/findings.py`
- Create: `engine/smartass/model/analyze.py`
- Test: `engine/tests/v2/test_v2_analyze.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/v2/test_v2_analyze.py
import pytest
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.analyze import compute, plan_requests
from smartass.model.spec import AnalysisSpec
from smartass.model.validate import GuardError


def test_end_to_end_base_spec():
    s = AnalysisSpec.model_validate(spec_dict())
    res = compute(s, resolver(), CLOSED, facts())
    assert [o.type for o in res.outs] == ["kpi_tiles", "table", "bridge", "ranking", "drivers"]
    assert res.key_values["Key figures · EBITDA"] == {"Actual": 65.0, "Last year": 63.0, "Variance": 2.0,
                                                     "Variance %": pytest.approx(2 / 63)}
    assert res.key_values["EBITDA bridge · Other / unexplained"] == {"Amount": 0.0}
    assert not any("Check" in k for k in res.key_values)
    keys = {m["key"] for m in res.material}
    assert {"b1.m1", "b4.top", "b4.d0", "b4.d0_0", "b4.d0_1"} <= keys      # Opex lines over 5 NOK / 5 %
    assert "b0.k0" not in keys                                             # EBITDA +2 is under 5 MNOK
    assert res.findings == []


def test_plan_requests_runs_the_guards_first():
    d = spec_dict()
    d["queries"][0]["entity"] = "un.SY.ZZ"
    with pytest.raises(GuardError, match="unknown entity"):
        plan_requests(AnalysisSpec.model_validate(d), resolver(), CLOSED)
    assert len(plan_requests(AnalysisSpec.model_validate(spec_dict()), resolver(), CLOSED)) == 2


def test_key_values_are_capped():
    s = AnalysisSpec.model_validate(spec_dict())
    res = compute(s, resolver(), CLOSED, facts(), kv_limit=3)
    assert len(res.key_values) == 3
```

- [ ] **Step 2: Run them and see them fail**

Run: `uv run pytest tests/v2/test_v2_analyze.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.model.analyze'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/model/findings.py
"""Material lines (they need a driver in the commentary) and key_values (the numbers the chat may quote)."""
from __future__ import annotations

from .blocks.common import BlockOut, ck
from .expr import Cells
from .resolver import Resolver


def material(outs: list[BlockOut], cells: Cells, r: Resolver) -> list[dict]:
    """A line is material when |variance| >= the metric's absolute threshold, or when |variance %| >= its
    percentage threshold and |variance| >= a tenth of the absolute threshold."""
    hits = []
    for o in outs:
        if "Variance" not in o.headers:
            continue
        vi = o.headers.index("Variance")
        pi = o.headers.index("Variance %") if "Variance %" in o.headers else None
        for ln in o.lines:
            if ln.kind == "check" or not ln.metric:
                continue
            v = cells.get(ck(o.key, ln.key, vi))
            if v is None:
                continue
            p = cells.get(ck(o.key, ln.key, pi)) if pi is not None else None
            a, pct = r.materiality(ln.metric)
            if abs(v) >= a or (p is not None and abs(p) >= pct and abs(v) >= a / 10):
                hits.append({"block": o.title, "line": ln.label, "key": f"{o.key}.{ln.key}", "variance": v,
                             "variance_pct": p})
    return hits


def key_values(outs: list[BlockOut], cells: Cells, limit: int = 40) -> dict[str, dict[str, float | None]]:
    kv: dict[str, dict[str, float | None]] = {}
    for o in outs:
        for ln in o.lines:
            if ln.kind == "check":
                continue
            if len(kv) >= limit:
                return kv
            kv[f"{o.title} · {ln.label}"] = {h: cells.get(ck(o.key, ln.key, j)) for j, h in enumerate(o.headers)}
    return kv
```

```python
# engine/smartass/model/analyze.py
"""The model's two entry points. plan_requests: guards, then what to fetch. compute: guards, blocks -> trees ->
values, the reconciliation guard, findings, material lines and key values. The data layer (Plan 1B) fetches between
the two; the render layer writes the same trees as formulas."""
from __future__ import annotations

from dataclasses import dataclass

from .blocks import BUILDERS
from .blocks.common import BlockOut, evaluate
from .context import Ctx
from .expr import Cells
from .findings import key_values, material
from .frame import Frame
from .periods import Month
from .requests import GridRequest, requests
from .resolver import Resolver
from .spec import AnalysisSpec
from .validate import Finding, guard_result, guard_spec, sign_findings


@dataclass
class Result:
    outs: list[BlockOut]
    cells: Cells
    findings: list[Finding]
    material: list[dict]
    key_values: dict[str, dict[str, float | None]]


def plan_requests(spec: AnalysisSpec, r: Resolver, closed: Month) -> list[GridRequest]:
    guard_spec(spec, r, closed)
    return requests(spec, r, closed)


def compute(spec: AnalysisSpec, r: Resolver, closed: Month, frame: Frame, kv_limit: int = 40) -> Result:
    soft = guard_spec(spec, r, closed)
    ctx = Ctx(spec, r, closed, frame)
    outs = [BUILDERS[b.type](b, ctx, f"b{i}") for i, b in enumerate(spec.blocks) if b.type in BUILDERS]
    cells = evaluate(outs, frame)
    guard_result(outs, cells)
    return Result(outs, cells, soft + sign_findings(ctx), material(outs, cells, r),
                  key_values(outs, cells, kv_limit))
```

- [ ] **Step 4: Run them and see them pass**

Run: `uv run pytest tests/v2/test_v2_analyze.py -q`
Expected: `3 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/findings.py engine/smartass/model/analyze.py engine/tests/v2/test_v2_analyze.py
git commit -m "feat(v2): compute entry point, material lines and key values

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 16: Full suite, lint and handover note

**Files:**
- Modify: `CHANGELOG.md` (under `Unreleased`)

- [ ] **Step 1: Run the whole v2 suite**

Run: `uv run pytest tests/v2 -q`
Expected: `72 passed` (1 + 10 + 4 + 10 + 9 + 4 + 3 + 3 + 3 + 3 + 1 + 2 + 3 + 4 + 9 + 3).

- [ ] **Step 2: Run the v1 suite to prove nothing broke**

Run: `uv run pytest -q`
Expected: every test passes (v1 count unchanged plus the v2 tests).

- [ ] **Step 3: Lint**

Run: `uv run ruff check .`
Expected: `All checks passed!`. Fix any unused import it names, then re-run.

- [ ] **Step 4: Add the changelog line**

Under `## Unreleased` in `CHANGELOG.md` add:

```markdown
- v2 model core (`engine/smartass/model`): typed analysis spec, period algebra, fact frame, expression trees
  (value and Excel formula from one tree), seven computed blocks (table, kpi_tiles, trend, matrix, bridge, ranking,
  drivers), data guards, material lines and key values. Not wired to tools yet (Plan 1B).
```

- [ ] **Step 5: Commit**

```bash
git add CHANGELOG.md
git commit -m "docs(v2): changelog for the model core

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## What Plan 1B covers (next)

Plan 1B is written after this plan is merged.
- `data/`: the `Provider` interface, wrapping v1 `svhttp`/`svxml`, ad hoc fallback, FinPlan period normalisation to
  `Sep 26` labels, drill lines.
- `render/`: `Data_Q<n>` table, layout of `BlockOut` lines, `to_excel` with an address map, TN style, charts, sheet
  order, readback "Sheet = model", packs via `fin_refresh`.
- `runtime/` fixes.
- The 8 MCP tools behind a second server entry `smartass2`.
- The test-set runner on recorded facts.
- Patches (follow-ups as JSON merge patches on the stored spec, with unchanged queries reused).
- Passing caveats from the resolver through to the reply.
- Readback rule (review 04.11): the model's `None` (missing) equals Excel's 0 / blank in "Sheet = model". SUMIFS
  over missing data returns 0, and a ratio with a missing numerator shows 0 in Excel.
- Write `Data_Q<n>[Period]` (and every label column) as text (`@` number format), or Excel turns "Sep 26" into a
  date and SUMIFS stops matching.
