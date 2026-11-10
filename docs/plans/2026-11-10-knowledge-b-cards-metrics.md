# Knowledge layer, phase B: behaviour cards and new metrics Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:**
- Every FinMax/FinLite/FinPlan account and entity has a behaviour card: its operator towards each parent, time balance, account type and shared flag.
- The model consolidates with the real operators.
- The semantic layer gains the missing metrics: EBIT, profit before tax, net income, capex, balance sheet, free cash flow, FinPlan volumes/ARPU/FTE, and statistics.

**Architecture:**
- Smart View cannot give member properties: probe 10.11, see "Probe results" below. So cards are **derived from data**.
  - One value per member is read in a few large ExecuteGrid retrieves, about 20 s for all of FinMax.
  - Every parent is then compared with the sum of its children offline.
  - Time balance is read from year vs months vs December.
- Derived cards live in `atlas.db` (table `card`), because they are rebuilt with the atlas.
- Manual overrides live in `knowledge/semantic/cards/*.toml` (git). They are merged into `knowledge.db` at build, and a manual field wins over a derived one.
- `weight(member, cube, parent)` reads the merged card. `fin_lookup` member hits carry the compact card.
- New metrics are verified live one by one before they are approved.

**Tech Stack:** Python 3.12, pydantic v2, sqlite3, pytest, ruff, uv; the engine's `SvHttp` (`engine/smartass/epm/svhttp.py`: `shared()`, `fetch(conn, pov, rows: list[Axis], cols: list[Axis])` → `Fetched(layout, values)`), `Axis` in `engine/smartass/epm/grid.py`.

**Spec:** `docs/specs/2026-11-10-knowledge-layer-design.md` §3 (cards) and §2 / Q7 (metrics).

**Refinement of the spec:**
- Derived cards are not written as TOML (26 000+ members). They are data in `atlas.db`.
- Only manual cards are TOML.
- Provenance per field is `derived` or `manual` plus a date. The `smartview` source is dropped, because the probe showed it does not exist.

**Probe results (10.11, scratchpad `probe_b/`, the facts this plan rests on):**
- **Smart View:** `GetMemberInformation` and `GetMembersInformation` fail on this provider (11.1.2.4.400). `EnumMembers getAtts=1` returns `0` for everything. Do **not** add them to `svxml.ALLOWED`.
- **Signs:** revenue is stored positive and costs negative in FinMax and FinPlan. Every P&L total is a plain sum, so `operator` is `+` throughout the P&L.
- **FinMax accounts:** slice un.SY, Actual, NOK, Periodic, Yr 25; 4 742 members, 1 226 parents.

  | Result | Parents |
  |---|---|
  | `+` | 654 |
  | `~` (ac.Statistics ignores ac.STA) | 1 |
  | unresolved (`All Accounts`, `ac.Export`) | 2 |
  | no data | 569 |

- **FinMax entities:** slice ac.C70, Yr 25; 21 597 members, 3 435 parents.

  | Result | Parents |
  |---|---|
  | `+` | 228 |
  | one `~` child (eliminations `…i`, un.NOR under All Entities, company roll-ups) | 90 |
  | unresolved | 108 |
  | no data | 3 009 |

  Of the 108 unresolved parents, 106 are 0 while their children have data (non-consolidating parents, so every child is `~`).
- **Time balance:**
  - FinMax stores every account, FTE included, as monthly movement: `Yr` = the sum of the months. A stock shows as Accumulated(Dec) / Accumulated(Jan) ≈ 1 (STA20Y 0.97, against 12.4 for C70).
  - FinPlan has real time balance: `Fixed Broadband subscriptions` YearTotal = December.
- **Cost:** a 2 000-row grid takes 0.3–0.5 s once the session is open; the first grid takes 8–10 s.

**Standing rules** (same as phase A, `docs/plans/2026-11-10-knowledge-a-one-source.md` header):
- **Running Python and tests:**
  - Run from `engine/` with `uv run ... < /dev/null`. Never run Python from stdin. Scratch files go in the session scratchpad.
  - Tests: `uv run python -B -m pytest -q tests < /dev/null`. Lint: `uv run ruff check . < /dev/null`.
  - Live EPM calls go through `svhttp.shared()` (headless). Call `drill_browser.shutdown_shared()` in a `finally`. If anything would open a visible window, stop.
- **Commits:** `git checkout engine/uv.lock` before every commit except the version bump, which commits it. Commit messages end with:
  ```
  Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01MW9ANq493fHqddMk5ipVTq
  ```
- **Goldens** (`uv run python -B -m smartass.testset run < /dev/null`) are Tom's: if one changes, stop and report.
- **Branch:** `v2/knowledge-b` from `main`. Merge, push and check that CI is green at the end.

---

## File structure

| File | Responsibility |
|---|---|
| `engine/smartass/knowledge/cards.py` (create) | Pure derivation: operators from values + tree; time balance from year/months; account type from branch + sign. No I/O. |
| `engine/smartass/knowledge/cards_fetch.py` (create) | Reads the value slices through `SvHttp` and writes table `card` in atlas.db. |
| `engine/smartass/knowledge/atlas.py` (modify) | `card` table in SCHEMA; `store_cards`, `cards_status` |
| `engine/smartass/ops/cli.py`, `ops/atlas_build.py` (modify) | `smartass atlas cards`; `atlas build` runs the cards at the end |
| `engine/smartass/knowledge/semantic.py` (modify) | `CardOverride` model; `load_all` reads `cards/*.toml` |
| `engine/smartass/knowledge/build.py` (modify) | `card` table in knowledge.db: derived merged with manual, plus the provenance; checks |
| `engine/smartass/knowledge/atlas_tree.py`, `resolver.py`, `model/resolver.py` (modify) | `weight(member, cube, parent=None)` from cards; `card()` for lookup |
| `engine/smartass/model/blocks/common.py` (modify) | `counted` / `children_sum` pass the parent |
| `knowledge/semantic/metrics/*.toml`, `kpis/*.toml`, `facts/*.toml` (create) | the new metrics, verified live |
| `engine/tests/v2/test_v2_knowledge_b.py` (create) | the tests of this plan |

---

### Task 1: Branch

- [ ] **Step 1:**

```bash
cd /c/dev/smartass && git checkout main && git pull -q && git checkout -b v2/knowledge-b
cd engine && uv run python -B -m pytest -q tests < /dev/null | tail -1
```

Expected: `1304 passed`.

---

### Task 2: Pure derivation (`cards.py`)

**Files:** Create `engine/smartass/knowledge/cards.py`, `engine/tests/v2/test_v2_knowledge_b.py`

- [ ] **Step 1: Write the failing tests**

```python
"""Knowledge layer phase B (docs/plans/2026-11-10-knowledge-b-cards-metrics.md)."""
import pytest

from smartass.knowledge.cards import account_type, operators, time_balance

KIDS = {"P": ["A", "B", "C"], "Q": ["D", "E"], "Z": ["F", "G"], "N": ["H"]}


def test_a_plain_sum_is_plus_for_every_child():
    got = operators(KIDS, {"P": 6.0, "A": 1.0, "B": 2.0, "C": 3.0})
    assert got[("P", "A")] == got[("P", "B")] == got[("P", "C")] == "+"


def test_one_child_left_out_is_tilde():
    got = operators(KIDS, {"Q": 5.0, "D": 5.0, "E": 7.0})
    assert got[("Q", "D")] == "+" and got[("Q", "E")] == "~"


def test_one_child_subtracted_is_minus():
    got = operators(KIDS, {"Q": -2.0, "D": 5.0, "E": 7.0})
    assert got[("Q", "E")] == "-" and got[("Q", "D")] == "+"


def test_a_zero_parent_over_children_with_data_ignores_them_all():
    got = operators(KIDS, {"Z": 0.0, "F": 4.0, "G": -1.0})
    assert got[("Z", "F")] == got[("Z", "G")] == "~"


def test_no_data_and_unresolved_are_unknown():
    got = operators(KIDS, {"P": 100.0, "A": 1.0, "B": 2.0, "C": 3.0})
    assert got[("P", "A")] is None and got[("N", "H")] is None


def test_a_second_slice_settles_what_the_first_could_not():
    first = operators(KIDS, {"Q": 5.0, "D": 5.0})                     # E has no data: + for D, E unknown
    both = operators(KIDS, {"Q": 5.0, "D": 5.0}, {"Q": 5.0, "D": 5.0, "E": 7.0}, prior=first)
    assert first[("Q", "E")] is None and both[("Q", "E")] == "~"


def test_time_balance():
    months = [100.0] + [1.0] * 11                                     # a level booked as movements (FinMax FTE)
    assert time_balance(year=sum(months), months=months, acc_jan=100.0, acc_dec=111.0) == "balance_last"
    flow = [10.0] * 12
    assert time_balance(year=120.0, months=flow, acc_jan=10.0, acc_dec=120.0) == "flow"
    level = [5.0, 6.0, 7.0, 8.0, 9.0, 9.0, 9.0, 9.0, 9.0, 9.0, 9.0, 10.0]  # FinPlan subscriptions: year = December
    assert time_balance(year=10.0, months=level) == "balance_last"
    assert time_balance(year=None, months=[None] * 12) is None


def test_account_type_from_branch_and_sign():
    assert account_type(["ac.C30", "ac.C48", "ac.EBITDA"], 5.0) == "revenue"
    assert account_type(["ac.C70", "ac.EBITDA_BOI"], -5.0) == "expense"
    assert account_type(["ac.STA20Y", "ac.STA", "ac.Statistics"], 3.0) == "statistic"
    assert account_type(["ac.C181", "ac.C1_C2"], 3.0) == "balance"
    assert account_type(["ac.X"], None) is None
```

- [ ] **Step 2: Run them to see them fail**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_b.py < /dev/null`
Expected: `ModuleNotFoundError: smartass.knowledge.cards`.

- [ ] **Step 3: Implement**

Create `engine/smartass/knowledge/cards.py`:

```python
"""Behaviour cards derived from data (spec 2026-11-10 §3; plan phase B). Pure functions: values in, card fields out.
Smart View gives no member properties on this provider (probe 10.11), so the operator of a child towards its parent
is read from the numbers: parent == sum of children -> '+'; one child that must be left out -> '~'; one that must be
subtracted -> '-'; a parent that is 0 while its children carry data does not consolidate -> every child '~'."""
from __future__ import annotations

Op = str | None
TOL_ABS, TOL_REL = 1.0, 1e-6


def close(a: float, b: float) -> bool:
    return abs(a - b) <= max(TOL_ABS, TOL_REL * max(abs(a), abs(b)))


def _one(kids: list[str], val: dict[str, float | None], parent: str) -> dict[str, Op]:
    pv = val.get(parent)
    cv = {c: val.get(c) for c in kids}
    if pv is None and all(x is None for x in cv.values()):
        return {}
    pv = pv or 0.0
    s = sum(x or 0.0 for x in cv.values())
    with_data = [c for c, x in cv.items() if x]
    if close(pv, s):
        return {c: "+" for c in with_data}
    if close(pv, 0.0) and with_data:
        return {c: "~" for c in with_data}
    diff = pv - s
    for c in with_data:
        if close(diff, -cv[c]):
            return {**{k: "+" for k in with_data}, c: "~"}
        if close(diff, -2 * cv[c]):
            return {**{k: "+" for k in with_data}, c: "-"}
    return {}


def operators(kids: dict[str, list[str]], *slices: dict[str, float | None],
              prior: dict[tuple[str, str], Op] | None = None) -> dict[tuple[str, str], Op]:
    """{(parent, child): '+' | '~' | '-' | None} over every parent in `kids`; later slices fill what earlier ones left
    unknown (a child without data in one slice), never overwrite a known operator."""
    out: dict[tuple[str, str], Op] = {(p, c): None for p, cs in kids.items() for c in cs}
    if prior:
        out.update({k: v for k, v in prior.items() if v is not None})
    for val in slices:
        for p, cs in kids.items():
            for c, op in _one(cs, val, p).items():
                if out.get((p, c)) is None:
                    out[(p, c)] = op
    return out


def time_balance(year: float | None, months: list[float | None], acc_jan: float | None = None,
                 acc_dec: float | None = None) -> str | None:
    """flow | balance_last | None. FinPlan: the year equals December (not the sum) -> balance_last. FinMax books
    levels as movements: Accumulated Dec / Accumulated Jan near 1 -> balance_last; near the month count -> flow."""
    ms = [m for m in months if m is not None]
    if year is None or not ms:
        return None
    last = next((m for m in reversed(months) if m is not None), None)
    total = sum(ms)
    if acc_jan and acc_dec is not None and abs(acc_jan) > TOL_ABS:
        ratio = acc_dec / acc_jan
        return "balance_last" if 0.5 <= ratio <= 2.0 and not close(total, last or 0.0) and close(year, total) \
            and abs(months[0] or 0.0) > 3 * max(abs(m) for m in ms[1:] or [0.0]) else "flow"
    if last is not None and close(year, last) and not close(year, total):
        return "balance_last"
    if close(year, total):
        return "flow"
    return None


REVENUE_ROOTS = {"ac.C30"}
EXPENSE_ROOTS = {"ac.C40S", "ac.C70", "ac.C50S", "ac.C60", "ac.C73", "ac.C76S", "ac.C75"}
STAT_ROOTS = {"ac.STA", "ac.Statistics"}
BALANCE_ROOTS = {"ac.C1_C2"}


def account_type(chain: list[str], value: float | None) -> str | None:
    """revenue | expense | statistic | balance | None from the member and its ancestors (chain[0] = the member) and,
    for P&L members outside the known roots, the stored sign (revenue +, costs -)."""
    names = set(chain)
    if names & STAT_ROOTS:
        return "statistic"
    if names & BALANCE_ROOTS:
        return "balance"
    if names & REVENUE_ROOTS:
        return "revenue"
    if names & EXPENSE_ROOTS:
        return "expense"
    if value is None:
        return None
    return "revenue" if value > 0 else "expense"
```

The FinMax rule in `time_balance` reads as follows:
- A level is booked in January and followed by small changes. So the first month dominates, and Accumulated Dec ≈ Accumulated Jan.
- If the test data shows the heuristic is wrong on real data in Task 3, adjust it there and pin the real case.

- [ ] **Step 4: Run tests, iterate until green**

Run: `uv run python -B -m pytest -q tests/v2/test_v2_knowledge_b.py < /dev/null`
Expected: `8 passed`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/knowledge/cards.py engine/tests/v2/test_v2_knowledge_b.py
git commit -m "feat(knowledge): behaviour cards derived from data - operators, time balance, account type (pure)"
```

---

### Task 3: Fetch the slices and store the cards in atlas.db

**Files:**
- Create: `engine/smartass/knowledge/cards_fetch.py`
- Modify: `engine/smartass/knowledge/atlas.py` (SCHEMA + `store_cards`, `card_rows`), `engine/smartass/ops/cli.py`, `engine/smartass/ops/atlas_build.py`
- Test: `engine/tests/v2/test_v2_knowledge_b.py`

- [ ] **Step 1: Schema and store** (in `atlas.py` SCHEMA)

```sql
CREATE TABLE IF NOT EXISTS card (
  connection TEXT, dimension TEXT, name TEXT, parent TEXT, operator TEXT, time_balance TEXT, account_type TEXT,
  shared INT, derived_at REAL, PRIMARY KEY (connection, dimension, name, parent)
);
```

```python
def store_cards(db, conn: str, dim: str, rows: list[tuple]) -> None:
    """rows: (name, parent, operator, time_balance, account_type, shared); replaces the dimension's cards."""
    now = time.time()
    with db:
        db.execute("DELETE FROM card WHERE connection=? AND dimension=?", (conn, dim))
        db.executemany("INSERT OR REPLACE INTO card VALUES (?,?,?,?,?,?,?,?,?)",
                       [(conn, dim, *r, now) for r in rows])
```

- [ ] **Step 2: Write the failing test for the fetch plan (offline, fake SvHttp)**

```python
from smartass.knowledge import atlas
from smartass.knowledge.cards_fetch import derive_dimension


class FakeSv:
    """fetch(conn, pov, rows, cols) -> Fetched-like with layout nr/nc and values[1+nc+j][nr]."""
    def __init__(self, values):
        self.values, self.calls = values, []

    def fetch(self, conn, pov, rows, cols):
        self.calls.append((conn, pov, rows[0].members, cols[0].members))
        names = rows[0].members

        class L:
            nr, nc = 1, 1
        grid = [[None, None], [None, None]] + [[n, self.values.get((cols[0].members[0], n))] for n in names]
        return type("F", (), {"layout": L, "values": grid})()


def test_derive_dimension_stores_operators_and_shared(tmp_path):
    db = atlas.connect(tmp_path / "a.db")
    db.executemany("INSERT INTO member VALUES (?,?,?,?,?,?,?)", [
        ("FinMax", "All Entities", "un.SY", "SY", "", 1, 1),
        ("FinMax", "All Entities", "A", "A", "un.SY", 2, 1),
        ("FinMax", "All Entities", "Ai", "Ai elim", "un.SY", 2, 2),
        ("FinMax", "All Entities", "A", "A", "OTHER", 2, 3)])
    db.commit()
    sv = FakeSv({("ac.C70", "un.SY"): 10.0, ("ac.C70", "A"): 10.0, ("ac.C70", "Ai"): 4.0})
    n = derive_dimension(sv, db, "FinMax", "All Entities", chunk=2)
    rows = {(r[0], r[1]): r for r in db.execute("SELECT name, parent, operator, shared FROM card")}
    assert rows[("A", "un.SY")][2] == "+" and rows[("Ai", "un.SY")][2] == "~"
    assert rows[("A", "un.SY")][3] == 1 and rows[("Ai", "un.SY")][3] == 0
    assert n == 3 and len(sv.calls) >= 2                                 # chunked
```

Adapt `FakeSv` to how `Fetched.values` is really indexed. Read `svhttp.Fetched` and the probe script `derive_all.py` (scratchpad `probe_b/`): it reads `values[1 + lay.nc + j][lay.nr]`. The test only needs to mirror that.

- [ ] **Step 3: Implement `cards_fetch.py`**

```python
"""Read value slices through SvHttp and store derived behaviour cards in atlas.db (plan phase B Task 3).
Read-only: ordinary retrieve grids only."""
from __future__ import annotations

import sqlite3
from collections import defaultdict

from ..epm.grid import Axis
from . import atlas
from .cards import account_type, operators

CHUNK = 2000
# (pov, column member) per cube and dimension: two slices each, the second fills children without data in the first.
SLICES = {
    ("FinMax", "All Accounts"): ("All Time Periods", [
        {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic", "All Entities": "un.SY"},
    ], ["Yr 25", "Yr 24"]),
    ("FinMax", "All Entities"): ("All Accounts", [
        {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic",
         "All Time Periods": "Yr 25"},
    ], ["ac.C70", "ac.C30"]),
}


def _values(sv, conn: str, pov: dict, dim: str, names: list[str], cdim: str, col: str) -> dict[str, float | None]:
    out: dict[str, float | None] = {}
    for i in range(0, len(names), CHUNK):
        part = names[i:i + CHUNK]
        f = sv.fetch(conn, pov, [Axis(dim, part)], [Axis(cdim, [col])])
        lay, v = f.layout, f.values
        for j, n in enumerate(part):
            x = v[1 + lay.nc + j][lay.nr]
            out[n] = x if isinstance(x, float) else None
    return out


def derive_dimension(sv, db: sqlite3.Connection, conn: str, dim: str, chunk: int = CHUNK) -> int:
    global CHUNK
    CHUNK = chunk
    pairs = db.execute("SELECT name, parent FROM member WHERE connection=? AND dimension=? ORDER BY ord",
                       (conn, dim)).fetchall()
    names = list(dict.fromkeys(n for n, _ in pairs))
    kids: dict[str, list[str]] = defaultdict(list)
    parents_of: dict[str, list[str]] = defaultdict(list)
    for n, p in pairs:
        if p:
            kids[p].append(n)
            parents_of[n].append(p)
    cdim, povs, cols = SLICES[(conn, dim)]
    slices = [_values(sv, conn, povs[0], dim, names, cdim, c) for c in cols]
    ops = operators(dict(kids), *slices)
    first = slices[0]

    def chain(n: str) -> list[str]:
        out, p, seen = [n], (parents_of.get(n) or [""])[0], {n}
        while p and p not in seen:
            out.append(p)
            seen.add(p)
            p = (parents_of.get(p) or [""])[0]
        return out

    rows = [(n, p, ops.get((p, n)), None,
             account_type(chain(n), first.get(n)) if dim == "All Accounts" else None,
             int(len(parents_of[n]) > 1)) for n, p in pairs if p]
    atlas.store_cards(db, conn, dim, rows)
    return len(rows)
```

Replace the `global CHUNK` hack with a `chunk` parameter passed down to `_values` (the code above shows the intent; do it cleanly). Time balance for accounts is Task 4.

- [ ] **Step 4: Test green, then a real run**

Run the unit test. Then write a scratch script that calls `derive_dimension(svhttp.shared(), atlas.connect(), "FinMax", d)` for `"All Accounts"` and `"All Entities"`, and prints counts per operator.

Expected (from the probe):
- About 650+ `+` parents in accounts.
- About 90 parents with a `~` child in entities.
- Run time under 1 minute.

Wrap the call in `try/finally: drill_browser.shutdown_shared()`.

Then add FinLite: it shares FinMax's members. Copy the FinMax cards for FinLite, using `INSERT ... SELECT 'FinLite', ...` where the FinLite atlas has the same (dimension, name, parent). Add FinPlan with slices `("FinPlan", "Account")` and `("FinPlan", "Entity")`:
- POV: `Scenario Actual`, `Version Working`, `Years FY25`, entity `BANORWAY`, column `YearTotal`.
- Entity slice: account `OPEX` and `Revenue`.
- Check the FinPlan dimension names and members in the atlas: `SELECT DISTINCT dimension FROM member WHERE connection='FinPlan'`. Read how `data/provider.py` builds a FinPlan POV and copy it.
- Record the real counts in the commit message.

- [ ] **Step 5: CLI + atlas build**

- `cli.py`: add `at_sub.add_parser("cards", help="derive behaviour cards from data (read-only, about 1 minute)")`. The handler runs every `SLICES` key plus the FinLite copy and prints counts.
- `atlas_build.run_build`: after a full build without `--dimension`, call the same function. If cards fail, print a warning and keep the atlas, with exit code 0 for the atlas part.

- [ ] **Step 6: Tests, commit**

```bash
git add -A engine/smartass engine/tests
git commit -m "feat(knowledge): derive behaviour cards from two data slices per dimension into atlas.db (smartass atlas cards)"
```

---

### Task 4: Time balance per account

**Files:** Modify `cards_fetch.py`; test in `test_v2_knowledge_b.py`

- [ ] **Step 1: Failing test**

A FakeSv returns 12 months, the year, and Accumulated Jan/Dec for two accounts. `derive_time_balance(sv, db, "FinMax")` sets `balance_last` on the stock and `flow` on the other, in table `card` (column `time_balance`, every row of that member).

- [ ] **Step 2: Implement `derive_time_balance(sv, db, conn)`:**
- **FinMax** (un.SY, Yr 25): read all accounts × the 12 months, Periodic; plus Accumulated `Jan 25` and `Dec 25`. That is three retrieves of about 4 742 rows × 12 or × 1 columns, chunked like `_values`, but with several columns. Check the real month member names in the atlas (`All Time Periods`).
- **FinPlan** (BANORWAY, Actual/Working, FY25): months `Jan`…`Dec` and `YearTotal`. Check the Period member names in the atlas.
- Apply `cards.time_balance`, then `UPDATE card SET time_balance=? WHERE connection=? AND dimension=? AND name=?`.
- **Pin real cases:** STA20Y gives `balance_last`, C70 gives `flow`, and FinPlan `Fixed Broadband subscriptions` gives `balance_last`. That last case may need Target/Final FY26 if Actual is empty. The probe found it on Target FY26, so add a Target slice for FinPlan accounts with no Actual data.
- If the heuristic misclassifies on real data, fix `time_balance` and add the real numbers as a unit test.

- [ ] **Step 3: Live run, counts, commit**

```bash
git commit -am "feat(knowledge): time balance per account from year vs months (FinPlan) and accumulated Jan/Dec (FinMax)"
```

---

### Task 5: Manual overrides and the merged card in knowledge.db

**Files:** Modify `semantic.py` (model + `load_all`), `build.py`; create `knowledge/semantic/cards/finmax.toml`; test.

- [ ] **Step 1: Failing tests**

```python
from smartass.knowledge.semantic import CardOverride


def test_card_override_model():
    CardOverride(cube="FinMax", member="un.SY.X", parent="un.SY", operator="~", note="elimination", source="s")
    with pytest.raises(ValueError):
        CardOverride(cube="FinMax", member="x", operator="*", source="s")
```

The second test builds a semantic folder with `cards/x.toml`:

```toml
[[card]]
cube = "FinMax"
member = "ac.STA"
parent = "ac.Statistics"
operator = "~"
note = "statistics never add up to the P&L"
source = "probe 10.11"
```

Over an atlas with a derived `+` for that pair, `KnowledgeResolver.card("ac.STA", "FinMax", "ac.Statistics")` gives `operator "~"` with provenance `manual`. Without the override it gives `+` with `derived`.

- [ ] **Step 2: Implement**
- `CardOverride(BaseModel, extra=forbid)`: `cube`, `member`, `parent: str = ""` (empty means every parent), `operator: Literal["+","-","~"] | None`, `time_balance: Literal["flow","balance_last","balance_first","average"] | None`, `account_type: Literal["revenue","expense","statistic","balance"] | None`, `storage: Literal["stored","dynamic_calc","shared","label_only"] | None`, `note: str = ""`, `source: str`.
- `load_all` reads `cards/*.toml` `[[card]]` lists, like the notes.
- In `build.py`, add the table `card (cube, dimension, member, parent, operator, time_balance, account_type, storage, shared, note, provenance TEXT)` to knowledge.db. Fill it by copying the atlas `card` rows, with provenance a JSON `{field: "derived@<date>"}`, then applying the overrides (field-wise; provenance `manual@<build date>`).
- **Check:** two overrides of the same (cube, member, parent) field with different values are an error. An override of a member not in the atlas is an error.
- `SCHEMA_VERSION` `"3b"`.

- [ ] **Step 3: First overrides from the probe**

In `knowledge/semantic/cards/finmax.toml`:
- `ac.STA` under `ac.Statistics`: `~`.
- `All Accounts`: unresolved in the probe. Leave it out of the override file, so it stays unknown.
- For each of the 106 non-consolidating entity parents the derivation already marks `~`, add nothing.

Only add overrides for the cases the derivation gets wrong. List them in the commit.

- [ ] **Step 4: Tests, commit**

```bash
git commit -am "feat(knowledge): manual card overrides in knowledge/semantic/cards merged over the derived cards"
```

---

### Task 6: The model consolidates with the real operators

**Files:**
- Modify: `atlas_tree.py` (`weight`), `knowledge/resolver.py` (`card`, `weight`), `model/resolver.py` (Protocol + StaticResolver `weight(member, cube, parent=None)`), `model/blocks/common.py` (`counted`, `children_sum` pass `q.entity` as parent)
- Test: `test_v2_knowledge_b.py`

- [ ] **Step 1: Failing test**

`KnowledgeResolver.weight("Ai", "FinMax", parent="un.SY")` must be `0.0` for a `~` card, `-1.0` for `-`, `1.0` for `+`, and `1.0` when there is no card. Also a pipeline-level test, using the existing fakes in `tests/v2/test_v2_pipeline.py` as the pattern: an entity whose children include a `~` elimination child reconciles (no guard stop) and lists only the counted children.

- [ ] **Step 2: Implement.**
- `weight(member, cube, parent=None)`: look up the merged card for (cube, member, parent). Without a parent, use the shallowest placement's parent, as `AtlasTree.parent` does. Map `+` → 1.0, `-` → -1.0, `~` → 0.0, and None → 1.0.
- `common.counted` / `children_sum`: call `weight(k, q.cube, q.entity)`. Check every other `weight(` call: `grep -rn "\.weight(" engine/smartass`.
- Remove the comment "The atlas has no consolidation operators" in `atlas_tree.py`.

- [ ] **Step 3: Full tests + goldens**

Run: pytest, ruff, then `uv run python -B -m smartass.testset run < /dev/null`.
Expected: 9/9. If a golden changes, a `~` child probably changed a ranking or drivers list. Stop and report the diff with the member names; Tom decides.

- [ ] **Step 4: Commit**

```bash
git commit -am "feat(model): consolidation weights from the behaviour cards (~ children no longer counted)"
```

---

### Task 7: Cards in `fin_lookup` and in caveats

**Files:** Modify `knowledge/resolver.py` (`search` member hits), `server.py` (`fin_lookup` docstring); test.

- [ ] **Step 1: Failing test**

A member hit carries `"card": {"time_balance": "balance_last", "account_type": "statistic", "children_ignored": ["X"]}`. The card is compact: fields that are None are omitted, and `children_ignored` / `children_subtracted` list that member's own children with `~` / `-`.

- [ ] **Step 2: Implement.**

In `search`, add `card` to member hits. The `fin_lookup` docstring gets one clause: "member hits carry a behaviour card (time balance, account type, children ignored/subtracted in the consolidation)".

- [ ] **Step 3: Answer-level guidance**

In `pipeline` caveat collection, for every touched account whose card is `balance_last`, add one info caveat `"<label> is a level (time balance last): read the period-end value, never sum months."`. It gets id `card_<member>`, and the same dedupe as facts. The FTE metric already behaves this way (`aggregation last`), so the caveat matters only for raw members. Test it.

- [ ] **Step 4: Commit**

```bash
git commit -am "feat(knowledge): fin_lookup member hits carry the behaviour card; levels get a time-balance caveat"
```

---

### Task 8: New metrics, verified live

**Files:** Create `knowledge/semantic/metrics/*.toml`, `kpis/*.toml`, `facts/*.toml`; scratch verification script.

Candidates from the atlas (10.11):

| id | FinMax / FinLite | FinPlan | Notes to verify |
|---|---|---|---|
| `ebit` | `ac.EBIT` (= EBITDA + C76S) | `EBIT` | sign; FinPlan Target has no D&A (fact `finplan_target_stops_at_ebitda`): budget EBIT may be empty |
| `depreciation_writedowns` | `ac.C76S` | (find) | C76S vs the existing `da` (C75): which is "D&A"? Fix `da.toml` if needed |
| `profit_before_tax` | `ac.C85` | (find) | = EBIT + C84S + C77 |
| `net_financial_items` | `ac.C84S` | (find) | |
| `net_income` | `ac.C87` (or `ac.C89T`) | `Net Income` | which one is "net income" for TN |
| `capex` | `ac.Capex` or `ac.P1A600` | `CAPEX` | balance type and scenario; does FinMax capex need `All Balance Types` other than Periodic? |
| `total_assets` | under `ac.C1_C2` | `Total Assets` | a balance: `balance = "Accumulated"`, `aggregation = "last"` |
| `free_cash_flow` | (none) | `Free Cash Flow` | FinPlan only |
| `subscriptions_*`, `arpu_*` | `ac.STF_SUBS`, `ac.STF_ARPU` (statistics) | `Volumes ARPU` children | which ones exist with Actual data; the memory says no actual subs in FinMax: check |
| `fte_finplan` | | `No of FTE Total Average`, `Total FTEs incl P&C` | which one matches FinMax STA20Y |
| `ebitda_margin_finplan` (KPI) | | from `ebitda` / `revenue` in FinPlan | KPI cube FinPlan |

- [ ] **Step 1: Verification script**

In the scratchpad, write a script that fetches each candidate through `svhttp.shared().fetch(...)` for:
- FinMax: un.SY, Actual, Yr 25 and Yr 26 YTD (closed month), Periodic and Accumulated.
- FinPlan: BANORWAY; Actual/Working FY25; Target/Final FY26.

It prints the values next to the card of each member (Task 4 time balance) and the sum checks (EBIT = EBITDA + C76S; PBT = EBIT + C84S + C77).

- [ ] **Step 2: Decide per metric from the numbers.**

For each metric, check:
- The sign is right (revenue +, costs −).
- The value has a plausible size (EBIT below EBITDA).
- The identity holds.
- The FinPlan counterpart is the same concept.

A candidate that fails gets no metric. Write a fact instead, saying what was found.

- [ ] **Step 3: Write each metric TOML**

Use the existing ones as models (`knowledge/semantic/metrics/ebitda.toml`) and fill in these fields:
- `status = "approved"`;
- `approved = { by = "auto", on = "<date>", via = "live check phase B" }`;
- `verified = { by = "auto", on = "<date>" }` (the `Metric` model has no `verified` field; add `verified: dict | None = None` to `Metric` if absent);
- `differs` saying how FinMax and FinPlan differ;
- materiality from the rule default, unless the unit is not NOK (volumes: count, ARPU: NOK per sub, so set `pct_floor`).
- Balance-sheet and level metrics: `balance = "Accumulated"`, `aggregation = "last"`.
- Add aliases (English and Norwegian: "driftsresultat" = EBIT, "resultat før skatt", "årsresultat"/"netto resultat", "investeringer"/"capex", "kontantstrøm", "abonnenter", "ARPU", "årsverk budsjett"). Make sure that no alias collides; the build checks this.

- [ ] **Step 4: Scenarios in FinMax/FinLite**

The atlas has 14 members in `All Scenarios`. Fetch C70 un.SY Yr 26 for each, and write a fact `finmax_scenarios.toml` (scope FinMax/FinLite `All Scenarios` members that have data) saying which scenarios carry data and what they are. Do not change `rules/scenarios.toml` (plan stays in FinPlan) unless a scenario is clearly a plan; if one is, report it and do not wire it.

- [ ] **Step 5: Build, tests, goldens, commit**

```bash
uv run smartass knowledge build < /dev/null
git add knowledge/semantic engine/smartass/knowledge/semantic.py
git commit -m "feat(knowledge): EBIT, PBT, net income, capex, balance, FCF, FinPlan volumes/ARPU/FTE metrics - verified live"
```

---

### Task 9: Docs, version, merge, CI

- [ ] **Step 1: Docs**
- `docs/DESIGN.md` knowledge section: cards (derived in atlas.db, manual overrides in `cards/`, weights from cards).
- `CHANGELOG.md` entry 0.2.23.
- Spec status line: "phase B done <date>".
- `skills/smartass/reference/cubes.md`: one line on cards in `fin_lookup`.

- [ ] **Step 2: Bump**

Run: `powershell -File scripts/bump.ps1` (0.2.22 → 0.2.23), then `git add engine/VERSION engine/pyproject.toml engine/uv.lock` and commit `chore(v2): version 0.2.23 (knowledge layer phase B)`.

- [ ] **Step 3: Merge, push, CI**

```bash
git checkout main && git merge --no-ff v2/knowledge-b -m "Merge v2/knowledge-b: behaviour cards and new metrics" && git push -q
gh run watch $(gh run list --limit 1 --json databaseId -q '.[0].databaseId') --exit-status; gh run list --limit 1
```

Expected: success.

- [ ] **Step 4: Smoke test through MCP**

Run these through `mcp__smartass__fin_lookup`:
- `fin_lookup("STA20Y", kind="member", cube="FinMax")`: the card says `balance_last` / `statistic`.
- `fin_lookup("un.SY", kind="member", cube="FinMax")`: `children_ignored` lists the elimination child.
- `fin_lookup("driftsresultat")`: the `ebit` metric.
