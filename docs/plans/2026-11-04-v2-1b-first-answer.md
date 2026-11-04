# smartass v2 · Plan 1B: first real answer

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Wire the v2 model core (`engine/smartass/model`, Plan 1A) to real data and to Excel, so that
`fin_analyze(spec)` on a second MCP server `smartass2` produces a real answer. The answer is a `Q<n> Answer` sheet of
live formulas over a `Q<n> Data` table, checked "Sheet = model".

**Architecture:**
- **New packages beside the model.**
  - `smartass/data`: a provider over v1's `SvHttp`.
  - `smartass/knowledge`: an interim resolver over v1's `atlas.db` and `catalog.toml`.
  - `smartass/render`: the sheet layout as a pure function, a thin COM writer, and the readback.
  - `smartass/pipeline.py`: fetch, then compute, then the Excel job.
  - `smartass/server.py`: three tools.
- **One seam to v1.** Everything v2 borrows from v1 (Excel host and jobs, targets, sheets, styling, `svhttp`,
  atlas) goes through `smartass/legacy.py`. Those modules are lifted into v2 when v1 is deleted.

**Tech Stack:** Python 3.12, pydantic 2.14, the `mcp` SDK 2.3 (`MCPServer`, as v1), pywin32 (COM via v1), pytest,
ruff.

**Decisions (Tom, 2026-11-04):**
- **Interim resolver:** built from atlas and catalogue now. Sub-spec 2 replaces it.
- **Plan split:** 1B delivers the first answer. Plan 1C adds follow-ups, `fin_refresh`, `fin_drill`, the `xl_*`
  tools, the ad hoc fallback, the runtime fixes and the test-set runner.
- **Answer files:** v2 writes them to `<analyser>\v2\`.
- **Live runs:** only when Tom says so. Everything here is tested offline; Task 12 is the one live step and waits
  for his go.

**Known limits in 1B, all stated in the reply:**
- No charts, and no ad hoc Excel fallback. When HTTP fails, the reply names the error and the next step.
- No `lines` block.
- The atlas has no consolidation operators, so every child counts with weight 1. A cube parent that is not the
  plain sum of its children stops at the reconciliation guard. That is the safe failure.

**Conventions for every task:**
- Work in `C:\dev\smartass` on a new branch `v2/1b`: `git switch -c v2/1b` from `main` in Task 0.
- Run commands from `/c/dev/smartass/engine` in Git Bash.
- Tests go in `engine/tests/v2/`. File names are globally unique (`test_v2_*.py`).
- Import shared fixtures with `from v2util import ...`. v1 test helpers come from `engine/tests/`, since pytest
  puts `tests/` on `sys.path` through its `conftest.py`, so `from svxml_util import ...` works.
- `tests/conftest.py` already points `SMARTASS_HOME` at a temp folder.
- Never run python through a stdin heredoc. Run bare python only with `< /dev/null`.
- Stage files explicitly, never with `git add -A`.
- Every commit message ends with:

  ```
  Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
  ```

## File structure

| File | Responsibility |
|---|---|
| `engine/smartass/legacy.py` | The seam to v1. `configure_v2()` sends this process's analysis files and session state to `<analyser>\v2` |
| `engine/smartass/knowledge/__init__.py` | Package marker |
| `engine/smartass/knowledge/interim.py` | `catalog_metrics()`: `MetricDef`s from `knowledge/catalog.toml`, joined across cubes on `concept` |
| `engine/smartass/knowledge/atlas_resolver.py` | `AtlasResolver`, which implements `model.resolver.Resolver` over `atlas.db` |
| `engine/smartass/data/__init__.py` | Package marker |
| `engine/smartass/data/provider.py` | `FetchResult`, `SvHttpProvider` (`GridRequest` → Smart View grids → `Fact`s; closed months) |
| `engine/smartass/render/__init__.py` | Package marker |
| `engine/smartass/render/layout.py` | `SheetPlan`, `layout()`: cells, formulas and style roles for `Q<n> Answer` (pure) |
| `engine/smartass/render/readback.py` | `compare()`: sheet values against model values (pure) |
| `engine/smartass/render/excel.py` | `write_data()`, `write_answer()`, `read_block()`: the COM writer |
| `engine/smartass/pipeline.py` | `prepare()` (pure + provider), `excel_job()` (COM), `reply()` |
| `engine/smartass/server.py` | `MCPServer("smartass2")` with `fin_analyze`, `fin_lookup`, `fin_status` |
| `engine/child_server_v2.py` | Starts `smartass.server.main` |
| `engine/run_server_v2.py` | The v1 relay, pointed at `child_server_v2.py` and also watching `smartass/` |

---

### Task 0: Branch and the v1 seam

**Files:**
- Create: `engine/smartass/legacy.py`
- Test: `engine/tests/v2/test_v2_legacy.py`

- [ ] **Step 1: Create the branch**

```bash
cd /c/dev/smartass && git switch main && git switch -c v2/1b
```

- [ ] **Step 2: Write the failing test**

```python
# engine/tests/v2/test_v2_legacy.py
from smartview_mcp import config as v1config
from smartview_mcp import session as v1session

from smartass import legacy


def test_configure_v2_moves_files_and_session(monkeypatch, tmp_path):
    monkeypatch.setattr(v1config, "ANALYSER_DIR", tmp_path / "analyser")
    monkeypatch.setattr(v1session, "_STATE", tmp_path / "sessions.json")
    monkeypatch.setattr(v1session, "_info", None)
    monkeypatch.setattr(legacy, "_configured", False)
    legacy.configure_v2()
    assert v1config.ANALYSER_DIR == tmp_path / "analyser" / "v2"
    assert v1session._STATE.name == "sessions_v2.json"
    assert v1session.info()["id"].startswith("v2-")
    legacy.configure_v2()                                   # idempotent
    assert v1config.ANALYSER_DIR == tmp_path / "analyser" / "v2"
```

- [ ] **Step 3: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_legacy.py -q`
Expected: FAIL with `ImportError: cannot import name 'legacy'`

- [ ] **Step 4: Implement**

```python
# engine/smartass/legacy.py
"""The one seam to v1 (smartview_mcp). Everything v2 borrows from v1 is imported here or from v1 modules named
in this docstring: excel.jobs, excel.target, excel.sheets, styling, sheetorder, answer_write._Cells, svhttp, grid,
atlas, server (host, guard, _call_user, lifespan). When v1 is deleted these are lifted into smartass/."""
from __future__ import annotations

import os
from pathlib import Path

import psutil

from smartview_mcp import config as v1config
from smartview_mcp import session as v1session

V2_DIR_NAME = "v2"
RELAY_SCRIPT = "run_server_v2.py"
_configured = False


def _session_info() -> dict:
    me = psutil.Process()
    relay = next((p for p in me.parents() if RELAY_SCRIPT in " ".join(v1session._cmdline(p))), None)
    p = relay or me
    cwd = v1session._cwd(p) or os.getcwd()
    return {"id": f"v2-{p.pid}-{int(p.create_time())}", "cwd": cwd, "project": v1session.safe_name(Path(cwd).name)}


def configure_v2() -> None:
    """In this process only: new analysis files go to <analyser>\\v2\\<project>, and the session workbook is
    remembered in sessions_v2.json under v2's own relay, so v2 never writes into a v1 session workbook."""
    global _configured
    if _configured:
        return
    if v1config.ANALYSER_DIR.name != V2_DIR_NAME:
        v1config.ANALYSER_DIR = v1config.ANALYSER_DIR / V2_DIR_NAME
    v1session._STATE = v1config.SVKIT / "sessions_v2.json"
    v1session._info = _session_info()
    _configured = True
```

- [ ] **Step 5: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_legacy.py -q`
Expected: `1 passed`

- [ ] **Step 6: Commit**

```bash
git add engine/smartass/legacy.py engine/tests/v2/test_v2_legacy.py
git commit -m "feat(v2): legacy seam - v2 answer files and session apart from v1

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 1: Interim metrics from the catalogue

**Files:**
- Create: `engine/smartass/knowledge/__init__.py`
- Create: `engine/smartass/knowledge/interim.py`
- Test: `engine/tests/v2/test_v2_interim.py`

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_interim.py
from smartass.knowledge.interim import catalog_metrics


def test_metrics_join_cubes_on_concept():
    by = {m.name: m for m in catalog_metrics()}
    assert by["Revenue"].members == {"FinMax": ("ac.C30",), "FinPlan": ("Revenue",)}
    assert by["Opex"].members == {"FinMax": ("ac.C70",), "FinPlan": ("OPEX",)}
    assert by["EBITDA"].members["FinPlan"] == ("EBITDA",)
    assert "FinPlan" not in by["Personnel cost"].members        # FinMax-only line


def test_signs_materiality_and_stock():
    by = {m.name: m for m in catalog_metrics()}
    assert by["Revenue"].sign == "income" and by["Opex"].sign == "cost" and by["EBITDA"].sign == "income"
    assert by["Other items"].sign == "neutral"
    assert by["Opex"].materiality == (5_000_000.0, 0.10)
    fte = by["FTE"]
    assert (fte.balance, fte.aggregation, fte.sign) == ("Accumulated", "last", "neutral")
    assert all(m.status == "approved" for m in by.values())


def test_finlite_shares_finmax_members():
    by = {m.name: m for m in catalog_metrics()}
    assert by["Revenue"].members.get("FinLite") is None      # FinLite is opted in per metric in sub-spec 2
```

- [ ] **Step 2: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_interim.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.knowledge'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/knowledge/__init__.py
"""Knowledge for v2. Interim: the v1 catalogue and atlas. Sub-spec 2 replaces this with the semantic layer."""
```

```python
# engine/smartass/knowledge/interim.py
"""Interim metrics until the semantic layer (sub-spec 2): v1's knowledge/catalog.toml P&L lines, joined across
cubes on `concept` (never on `key`). The catalogue is owner-approved, so every metric is 'approved'. Materiality is
the catalogue's owner rule: NOK 5 million, or 10 % from NOK 500k (= a tenth of the absolute threshold, which is how
model.findings.material reads it)."""
from __future__ import annotations

import tomllib

from smartview_mcp import config as v1config

from ..model.resolver import MetricDef

SIGN = {"revenue": "income", "gross_margin": "income", "ebitda_boi": "income", "ebitda": "income",
        "cogs": "cost", "personnel": "cost", "other_opex": "cost", "opex": "cost", "da": "cost"}


def catalog_metrics(path=None) -> list[MetricDef]:
    data = tomllib.loads((path or v1config.KNOWLEDGE_DIR / "catalog.toml").read_text(encoding="utf-8"))
    mat = data["materiality"]
    materiality = (float(mat["abs_nok"]), float(mat["pct"]))
    by_concept: dict[str, dict] = {}
    for cube in ("FinMax", "FinPlan"):
        for line in data["pnl"][cube]:
            concept = line.get("concept")
            if not concept:
                continue
            d = by_concept.setdefault(concept, {"label": line["label"], "members": {}, "stock": False})
            d["members"][cube] = (line["account"],)
            d["stock"] = d["stock"] or line.get("measure") == "stock"
            if cube == "FinMax":
                d["label"] = line["label"]
    out = []
    for concept, d in by_concept.items():
        stock = d["stock"]
        out.append(MetricDef(d["label"], d["members"], sign=SIGN.get(concept, "neutral"), status="approved",
                             materiality=materiality, balance="Accumulated" if stock else "Periodic",
                             aggregation="last" if stock else "sum"))
    return out
```

- [ ] **Step 4: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_interim.py -q`
Expected: `3 passed`. If a label differs, read `knowledge/catalog.toml` and fix the test's metric name to the
catalogue's FinMax label. The catalogue is the truth, so do not change the catalogue.

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/knowledge engine/tests/v2/test_v2_interim.py
git commit -m "feat(v2): interim metrics from the catalogue

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Atlas resolver

**Files:**
- Create: `engine/smartass/knowledge/atlas_resolver.py`
- Test: `engine/tests/v2/test_v2_atlas_resolver.py`

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_atlas_resolver.py
import pytest

from smartass.knowledge.atlas_resolver import AtlasResolver
from smartass.model.resolver import MetricDef, UnknownMetric
from smartview_mcp import atlas


@pytest.fixture
def db(tmp_path):
    con = atlas.connect(tmp_path / "atlas.db")
    rows = [("FinMax", "All Entities", "un.SY.2BUS", "2BUS Business Norway", "un.SY", 2, 1),
            ("FinMax", "All Entities", "un.SY.4BUS_B", "4BUS_B Business", "un.SY.2BUS", 3, 1),
            ("FinMax", "All Entities", "un.SY.G910", "G910 Talkmore Business", "un.SY.2BUS", 3, 2),
            ("FinMax", "All Accounts", "ac.C30", "C30 REVENUES", "ac.C48", 3, 1),
            ("FinMax", "All Accounts", "ac.X1", "X1 child of 2BUS name clash", "un.SY.2BUS", 4, 1),
            ("FinPlan", "Scenario", "Forecast Q226", "Forecast Q226", "Forecasts", 2, 1),
            ("FinPlan", "Scenario", "Forecast Q326", "Forecast Q326", "Forecasts", 2, 2)]
    con.executemany("INSERT INTO member (connection, dimension, name, label, parent, depth, ord) "
                    "VALUES (?,?,?,?,?,?,?)", rows)
    con.commit()
    return con


def res(db, yy=26):
    return AtlasResolver(db, [MetricDef("Revenue", {"FinMax": ("ac.C30",)}, sign="income")], forecast_year=yy)


def test_children_stay_in_the_members_dimension(db):
    assert res(db).children("un.SY.2BUS", "FinMax") == ("un.SY.4BUS_B", "un.SY.G910")
    assert res(db).children("un.SY.G910", "FinMax") == ()


def test_exists_label_and_metrics(db):
    r = res(db)
    assert r.exists("un.sy.2bus", "FinMax") and not r.exists("un.SY.2BUS", "FinPlan")
    assert r.label("un.SY.2BUS", "FinMax") == "2BUS Business Norway"
    assert r.label("un.SY.ZZ", "FinMax") == "un.SY.ZZ"
    assert r.members("revenue", "FinMax") == ("ac.C30",)
    assert r.members("ac.C30", "FinMax") == ("ac.C30",)
    with pytest.raises(UnknownMetric):
        r.members("ac.NOPE", "FinMax")
    assert r.weight("un.SY.G910", "FinMax") == 1.0


def test_latest_forecast_of_the_year(db):
    assert res(db).latest_forecast() == "Forecast Q326"
    with pytest.raises(UnknownMetric, match="Forecast Q<q>25"):
        res(db, 25).latest_forecast()
```

- [ ] **Step 2: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_atlas_resolver.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.knowledge.atlas_resolver'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/knowledge/atlas_resolver.py
"""Resolver over v1's atlas.db (every member of the three cubes) and the interim metrics. Limits until sub-spec 2:
no consolidation operators in the atlas, so every child weighs 1.0 (a parent that is not the plain sum of its
children stops at the reconciliation guard); no draft metrics."""
from __future__ import annotations

import re
import sqlite3

from ..model.resolver import DEFAULT_MATERIALITY, MetricDef, UnknownMetric


class AtlasResolver:
    def __init__(self, db: sqlite3.Connection, metrics: list[MetricDef], forecast_year: int):
        self.db = db
        self._m = {m.name.lower(): m for m in metrics}
        self.forecast_year = forecast_year

    def _metric(self, name: str) -> MetricDef | None:
        return self._m.get(name.lower())

    def metrics(self) -> list[MetricDef]:
        return list(self._m.values())

    def members(self, metric: str, cube: str) -> tuple[str, ...]:
        m = self._metric(metric)
        if m is None:
            if self.exists(metric, cube):
                return (self._spelling(metric, cube),)
            raise UnknownMetric(metric)
        if cube not in m.members:
            raise UnknownMetric(f"{metric} has no {cube} definition")
        return m.members[cube]

    def sign(self, metric):
        m = self._metric(metric)
        return m.sign if m else "neutral"

    def status(self, metric):
        m = self._metric(metric)
        return m.status if m else "approved"

    def materiality(self, metric):
        m = self._metric(metric)
        return m.materiality if m else DEFAULT_MATERIALITY

    def balance(self, metric):
        m = self._metric(metric)
        return m.balance if m else "Periodic"

    def aggregation(self, metric):
        m = self._metric(metric)
        return m.aggregation if m else "sum"

    def _row(self, member: str, cube: str):
        return self.db.execute("SELECT name, dimension, label FROM member WHERE connection=? AND name=? "
                               "COLLATE NOCASE ORDER BY depth LIMIT 1", (cube, member)).fetchone()

    def _spelling(self, member: str, cube: str) -> str:
        row = self._row(member, cube)
        return row[0] if row else member

    def exists(self, member: str, cube: str) -> bool:
        return self._row(member, cube) is not None

    def children(self, member: str, cube: str) -> tuple[str, ...]:
        row = self._row(member, cube)
        if row is None:
            return ()
        names = [r[0] for r in self.db.execute(
            "SELECT name FROM member WHERE connection=? AND dimension=? AND parent=? ORDER BY ord",
            (cube, row[1], row[0]))]
        return tuple(dict.fromkeys(names))

    def weight(self, member: str, cube: str) -> float:
        return 1.0

    def label(self, member: str, cube: str) -> str:
        row = self._row(member, cube)
        return (row[2] or row[0]) if row else member

    def latest_forecast(self) -> str:
        names = [r[0] for r in self.db.execute("SELECT DISTINCT name FROM member WHERE connection='FinPlan' AND "
                                               "dimension='Scenario' AND name LIKE 'Forecast Q%'")]
        cands = sorted(((int(m.group(1)), n) for n in names
                        if (m := re.fullmatch(r"Forecast Q(\d)(\d\d)", n)) and int(m.group(2)) == self.forecast_year),
                       reverse=True)
        if not cands:
            raise UnknownMetric(f"no Forecast Q<q>{self.forecast_year:02d} in the atlas")
        return cands[0][1]
```

Before running, check `engine/smartass/model/resolver.py`. If the `Resolver` protocol defines `balance`,
`aggregation` and `weight` under other names (they were added in the 1A review fixes), match those names
exactly.

- [ ] **Step 4: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_atlas_resolver.py -q`
Expected: `3 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/knowledge/atlas_resolver.py engine/tests/v2/test_v2_atlas_resolver.py
git commit -m "feat(v2): atlas resolver over atlas.db and the interim metrics

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Smart View HTTP provider

**Files:**
- Create: `engine/smartass/data/__init__.py`
- Create: `engine/smartass/data/provider.py`
- Test: `engine/tests/v2/test_v2_provider.py`

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_provider.py
from svxml_util import EchoTransport, atlas_with_dimensions

from smartass.data.provider import SvHttpProvider
from smartass.model.requests import GridRequest
from smartview_mcp import svhttp


def provider(tmp_path, value):
    sv = svhttp.SvHttp(EchoTransport(value), home="https://plan1.example",
                       db_path=atlas_with_dimensions(tmp_path / "a.db"))
    return SvHttpProvider(sv)


def test_finmax_grid_becomes_facts(tmp_path):
    seen = []

    def value(rows, cols, pov):
        seen.append((rows, cols, pov))
        return None if rows[1] == "E2" else 10.0

    p = provider(tmp_path, value)
    req = GridRequest("fm", "FinMax", ("ac.C30", "ac.C70"), ("E1", "E2"), ("Aug 26", "Sep 26"), "Actual", "",
                      "Periodic")
    got = p.fetch(req)
    assert got.route == "http" and got.warnings == []
    assert len(got.facts) == 4                                   # E2 is #Missing everywhere: no facts
    f = got.facts[0]
    assert (f.cube, f.query, f.entity, f.scenario, f.version, f.balance, f.value) == (
        "FinMax", "fm", "E1", "Actual", "", "Periodic", 10.0)
    assert {x.period for x in got.facts} == {"Aug 26", "Sep 26"}
    assert any("Actual" in pov and "NOK" in pov and "Periodic" in pov for _, _, pov in seen)


def test_finplan_splits_years_and_relabels_periods(tmp_path):
    p = provider(tmp_path, lambda rows, cols, pov: 1.0)
    req = GridRequest("fp", "FinPlan", ("EBITDA",), ("2BUS Business",), ("Dec 25", "Jan 26"), "Target", "Final",
                      "Periodic")
    got = p.fetch(req)
    assert sorted(f.period for f in got.facts) == ["Dec 25", "Jan 26"]
    assert {(f.scenario, f.version) for f in got.facts} == {("Target", "Final")}


def test_error_cells_are_warnings_not_numbers(tmp_path):
    p = provider(tmp_path, lambda rows, cols, pov: "#NoAccess" if cols[0] == "Sep 26" else 2.0)
    req = GridRequest("fm", "FinMax", ("ac.C30",), ("E1",), ("Aug 26", "Sep 26"), "Actual", "", "Periodic")
    got = p.fetch(req)
    assert [f.period for f in got.facts] == ["Aug 26"]
    assert got.warnings and "#NoAccess" in got.warnings[0]


def test_closed_months_per_cube(tmp_path):
    p = provider(tmp_path, lambda *a: 1.0)
    assert p.closed_months() == {"FinMax": (2026, 9), "FinLite": (2026, 9), "FinPlan": (2026, 9)}
```

- [ ] **Step 2: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_provider.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.data'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/data/__init__.py
"""Data access for v2: providers turn GridRequests into Facts."""
```

```python
# engine/smartass/data/provider.py
"""Provider over v1's Smart View HTTP route (svhttp.SvHttp, read-only, ~0.2 s per grid once signed in).
FinMax/FinLite: rows Accounts x Entities, columns Time Periods ('Sep 26'), POV scenario, NOK, balance.
FinPlan: rows Account x Entity, columns Period ('Sep'), POV Scenario, Version, Years ('FY26'), balance - one grid
per year, periods relabelled to 'Sep 26'. #Missing gives no fact (missing is never zero); #NoAccess / #Invalid
give a warning and no fact. The ad hoc Excel fallback arrives in Plan 1C."""
from __future__ import annotations

import time
from dataclasses import dataclass, field

from smartview_mcp.grid import Axis

from ..model.frame import Fact
from ..model.periods import Month, parse
from ..model.requests import GridRequest

MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


@dataclass
class FetchResult:
    request: GridRequest
    facts: list[Fact] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    route: str = "http"
    seconds: float = 0.0


def _grids(req: GridRequest) -> list[tuple[dict, list[Axis], list[Axis], dict[str, str]]]:
    """(pov, rows, cols, {column member: 'Sep 26' label}) per grid."""
    if req.cube == "FinPlan":
        by_year: dict[str, list[str]] = {}
        for p in req.periods:
            y, m = parse(p)
            by_year.setdefault(f"FY{y % 100:02d}", []).append(MONTHS[m - 1])
        out = []
        for fy, months in by_year.items():
            pov = {"Scenario": req.scenario, "Version": req.version, "Years": fy, "All Balance Types": req.balance}
            out.append((pov, [Axis("Account", list(req.accounts)), Axis("Entity", list(req.entities))],
                        [Axis("Period", months)], {mo: f"{mo} {fy[2:]}" for mo in months}))
        return out
    pov = {"All Scenarios": req.scenario, "All Currencies": "NOK", "All Balance Types": req.balance}
    return [(pov, [Axis("All Accounts", list(req.accounts)), Axis("All Entities", list(req.entities))],
             [Axis("All Time Periods", list(req.periods))], {p: p for p in req.periods})]


class SvHttpProvider:
    def __init__(self, sv):
        self.sv = sv

    def fetch(self, req: GridRequest) -> FetchResult:
        out = FetchResult(req)
        t0 = time.monotonic()
        for pov, rows, cols, labels in _grids(req):
            _, parsed, _ = self.sv.retrieve(req.cube, pov, rows, cols, False, False)
            if parsed.unknown_members:
                out.warnings.append(f"{req.cube}: Smart View did not know {', '.join(parsed.unknown_members)}")
            heads = parsed.columns[2:]
            for row in parsed.rows:
                account, entity, values = row[0], row[1], row[2:]
                for head, v in zip(heads, values):
                    if v is None:
                        continue
                    if isinstance(v, str):
                        out.warnings.append(f"{req.cube} {account} / {entity} / {head}: {v} (left out)")
                        continue
                    out.facts.append(Fact(req.cube, req.query, account, entity, labels[head], req.scenario,
                                          req.version, float(v), req.balance))
        out.seconds = time.monotonic() - t0
        return out

    def closed_months(self) -> dict[str, Month]:
        """FinMax's 'Current Reporting Month' for every cube (FinLite and FinPlan close with FinMax)."""
        m = parse(self.sv.closed_month("FinMax"))
        return {"FinMax": m, "FinLite": m, "FinPlan": m}
```

Before running, check how `Fact` is built in `engine/smartass/model/frame.py`. If `balance` is a keyword-only
field, or sits in another position, adapt the `Fact(...)` call.

- [ ] **Step 4: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_provider.py -q`
Expected: `4 passed`.

`EchoTransport`'s `value(rows, cols, pov)` receives member names. If FinPlan dimension names or the closed month
differ in the fake, read `tests/svxml_util.py` and adapt the test to the fake, not the provider.

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/data engine/tests/v2/test_v2_provider.py
git commit -m "feat(v2): Smart View HTTP provider - grids to facts, closed months

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: Sheet layout (pure)

**Files:**
- Create: `engine/smartass/render/__init__.py`
- Create: `engine/smartass/render/layout.py`
- Test: `engine/tests/v2/test_v2_layout.py`

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_layout.py
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.spec import AnalysisSpec
from smartass.render.layout import FIRST_COL, layout


def plan(blocks=None):
    d = spec_dict() if blocks is None else spec_dict(blocks=blocks)
    s = AnalysisSpec.model_validate(d)
    res = compute(s, resolver(), CLOSED, facts())
    return s, res, layout(3, s, res, stock=set(), subtitle="NOK · FinMax · closed Sep 26")


def test_title_subtitle_and_sections():
    s, _, p = plan()
    assert p.cells[(1, 2)] == "Explain Business opex vs LY"
    assert p.cells[(2, 2)] == "NOK · FinMax · closed Sep 26"
    sections = [p.cells[(r["row"], 2)] for r in p.rows if r["kind"] == "section"]
    assert sections[:2] == ["Plan", "Key figures"]
    assert "EBITDA bridge" in sections and "Commentary" in sections


def test_every_model_cell_is_a_formula_on_the_data_table():
    _, res, p = plan()
    for o in res.outs:
        for ln in o.lines:
            for j, e in enumerate(ln.cells):
                key = f"{o.key}.{ln.key}.{j}"
                if e is None:
                    assert key not in p.addr
                    continue
                r, c = p.at(key)
                assert c == FIRST_COL + j
                assert str(p.cells[(r, c)]).startswith("=")
    first = next(v for v in p.cells.values() if isinstance(v, str) and "SUMIFS" in v)
    assert "Data_Q3[Value]" in first


def test_refs_point_at_their_rows():
    _, res, p = plan([{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}])
    other = p.at("b0.other.0")
    end, start = p.at("b0.end.0"), p.at("b0.start.0")
    s0, s1 = p.at("b0.s0.0"), p.at("b0.s1.0")
    col = "C"
    assert p.cells[other] == f"=({col}{end[0]}-{col}{start[0]})-SUM({col}{s0[0]}:{col}{s1[0]})"


def test_number_formats_by_column_code():
    _, res, p = plan([{"type": "table", "query": "fm", "columns": ["value", "ly", "var", "var_pct"]}])
    line = next(r for r in p.rows if r["kind"] == "line")
    assert line["num"] == {"C": "heltall", "D": "heltall", "E": "heltall", "F": "prosent1"}
    header = next(r for r in p.rows if r["kind"] == "header")
    assert header["num_cols"] == ["C", "D", "E", "F"]
    assert p.last_col == "F"


def test_check_lines_and_findings_rows():
    _, res, p = plan([{"type": "ranking", "query": "fm", "metric": "Opex", "top": 1}])
    assert any(r["kind"] == "check" for r in p.rows)
    assert p.last_row == max(r["row"] for r in p.rows)
```

- [ ] **Step 2: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_layout.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.render'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/render/__init__.py
"""Rendering for v2: the layout is pure; excel.py writes it through COM."""
```

```python
# engine/smartass/render/layout.py
"""Q<n> Answer as a SheetPlan: every cell (text, or a formula from the model's expression trees) plus the style
roles v1's styling.apply_report understands. Column B holds labels, values start in C. Two passes: rows and
addresses first, then formulas, so a Ref may point forwards (bridge 'Other' -> 'end'). Pure: no Excel."""
from __future__ import annotations

from dataclasses import dataclass, field

from ..model.analyze import Result
from ..model.blocks.common import ck
from ..model.spec import AnalysisSpec, CommentaryBlock

FIRST_COL = 3
LABEL_COL = 2
PCT_CODES = {"var_pct", "share"}


def col_letter(c: int) -> str:
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


@dataclass
class SheetPlan:
    cells: dict[tuple[int, int], object] = field(default_factory=dict)
    rows: list[dict] = field(default_factory=list)
    addr: dict[str, str] = field(default_factory=dict)
    last_col: str = "C"
    last_row: int = 1

    def at(self, key: str) -> tuple[int, int]:
        a = self.addr[key]
        col = "".join(ch for ch in a if ch.isalpha())
        n = 0
        for ch in col:
            n = n * 26 + ord(ch) - 64
        return int(a[len(col):]), n


def _num(code: str, stock: bool) -> str:
    if code in PCT_CODES:
        return "prosent1"
    return "tall1" if stock else "heltall"


def layout(q: int, spec: AnalysisSpec, res: Result, stock: set[str], subtitle: str) -> SheetPlan:
    p = SheetPlan()
    table = f"Data_Q{q}"
    widest = max([len(o.headers) for o in res.outs] or [1])
    r = 1

    def text(row: int, value: str, kind: str) -> None:
        p.cells[(row, LABEL_COL)] = value
        p.rows.append({"row": row, "kind": kind})

    text(r, spec.plan.purpose or spec.question, "title")
    text(r + 1, subtitle, "subtitle")
    r += 3
    text(r, "Plan", "section")
    r += 1
    text(r, f"Question: {spec.question}", "text")
    r += 1
    for h in spec.plan.hypotheses:
        text(r, f"•  Hypothesis: {h}", "text")
        r += 1
    r += 1

    placed: list[tuple[int, object]] = []          # (row, Line) per model line, for pass 2
    outs = {o.key: o for o in res.outs}
    for i, b in enumerate(spec.blocks):
        o = outs.get(f"b{i}")
        if isinstance(b, CommentaryBlock):
            text(r, b.title, "section")
            r += 1
            for kp in b.key_points:
                text(r, f"•  {kp}", "text")
                r += 1
            for d in b.drivers:
                ev = f"; evidence {d.evidence}" if d.evidence else ""
                basis = "Hypothesis - not verified in data" if d.basis == "hypothesis" else "data"
                text(r, f"•  {d.line}: {d.text} ({basis}{ev})", "text")
                r += 1
            r += 1
            continue
        if o is None:
            continue
        text(r, o.title, "section")
        r += 1
        letters = [col_letter(FIRST_COL + j) for j in range(len(o.headers))]
        for j, h in enumerate(o.headers):
            p.cells[(r, FIRST_COL + j)] = h
        p.rows.append({"row": r, "kind": "header", "num_cols": letters})
        r += 1
        for ln in o.lines:
            p.cells[(r, LABEL_COL)] = ("   " * ln.level) + ln.label
            is_stock = ln.metric in stock if ln.metric else False
            num = {letters[j]: _num(o.cols[j], is_stock) for j in range(len(o.headers))}
            p.rows.append({"row": r, "kind": ln.kind, "num": num})
            for j, e in enumerate(ln.cells):
                if e is not None:
                    p.addr[ck(o.key, ln.key, j)] = f"{letters[j]}{r}"
            placed.append((r, o, ln))
            r += 1
        r += 1

    if res.findings:
        text(r, "Findings", "section")
        r += 1
        for f in res.findings:
            text(r, f"{f.severity.capitalize()}: {f.text}", "note")
            r += 1
        r += 1
    text(r, f"Source: Oracle EPM Cloud via Smart View. Every number is a formula on table {table} "
            f"(sheet Q{q} Data).", "note")

    def addr(key: str) -> str:
        return p.addr[key]

    for row, o, ln in placed:
        for j, e in enumerate(ln.cells):
            if e is not None:
                p.cells[(row, FIRST_COL + j)] = "=" + e.to_excel(table, addr)
    p.last_col = col_letter(FIRST_COL + widest - 1)
    p.last_row = r
    return p
```

Before running, confirm that `BlockOut` has `cols` (column codes, added in the 1A review) in
`engine/smartass/model/blocks/common.py`. Also confirm the code used for the bridge column. If it is `amount`,
`_num` already formats it as `heltall`.

- [ ] **Step 4: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_layout.py -q`
Expected: `5 passed`. If a title in `test_title_subtitle_and_sections` differs, because 1A made titles unique or
changed default titles, adapt the expected title to the model's `res.outs[i].title`, not the other way round.

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/render engine/tests/v2/test_v2_layout.py
git commit -m "feat(v2): pure sheet layout - formulas from the model's trees, TN style roles

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: Readback check (pure)

**Files:**
- Create: `engine/smartass/render/readback.py`
- Test: `engine/tests/v2/test_v2_readback.py`

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_readback.py
from smartass.render.readback import compare


class FakeCells:
    def __init__(self, d):
        self.d = d

    def get(self, k):
        return self.d[k]


def grid(rows, cols, values):
    """values: {(r, c): v} -> Value2-like tuple from A1 to (rows, cols)."""
    return tuple(tuple(values.get((r, c)) for c in range(1, cols + 1)) for r in range(1, rows + 1))


def test_equal_sheet_has_no_mismatch():
    addr = {"a": "C4", "b": "D4", "m": "C5"}
    cells = FakeCells({"a": -85.0, "b": None, "m": 0.1})
    sheet = grid(5, 4, {(4, 3): -85.0, (4, 4): 0.0, (5, 3): 0.1000000001})
    assert compare(addr, cells, sheet) == []


def test_mismatch_and_error_values_are_reported():
    addr = {"a": "C4", "b": "D4"}
    cells = FakeCells({"a": -85.0, "b": 2.0})
    sheet = grid(4, 4, {(4, 3): -84.0, (4, 4): -2146826281})        # #N/A comes back as an int
    out = compare(addr, cells, sheet)
    assert len(out) == 2 and out[0].startswith("C4: sheet -84")


def test_blank_and_empty_string_count_as_zero():
    addr = {"a": "C4"}
    assert compare(addr, FakeCells({"a": None}), grid(4, 3, {(4, 3): ""})) == []
```

- [ ] **Step 2: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_readback.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.render.readback'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/render/readback.py
"""'Sheet = model': every model cell's value on the sheet equals the model's value. The model's None (missing)
equals Excel's 0 or blank (SUMIFS over no rows is 0; an empty ratio is ""). An Excel error comes back from Value2
as a negative int code and is always a mismatch."""
from __future__ import annotations

TOL_ABS = 0.5          # NOK
TOL_REL = 1e-9
EXCEL_ERRORS = {-2146826281, -2146826246, -2146826259, -2146826288, -2146826265, -2146826273, -2146826252}


def _pos(a: str) -> tuple[int, int]:
    col = "".join(ch for ch in a if ch.isalpha())
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return int(a[len(col):]), n


def compare(addr: dict[str, str], cells, sheet: tuple) -> list[str]:
    """sheet: Value2 of the block from A1. Returns 'C4: sheet x, model y' per mismatch."""
    out = []
    for key, a in addr.items():
        r, c = _pos(a)
        try:
            s = sheet[r - 1][c - 1]
        except IndexError:
            s = None
        m = cells.get(key)
        if isinstance(s, int) and s in EXCEL_ERRORS:
            out.append(f"{a}: sheet shows an Excel error, model {m}")
            continue
        sv = 0.0 if s in (None, "") else s
        mv = 0.0 if m is None else m
        if not isinstance(sv, (int, float)):
            out.append(f"{a}: sheet {sv!r}, model {mv}")
            continue
        # amounts in NOK: half a krone; fractions (variance %, share): 1e-6, so a wrong % is never hidden
        tol = TOL_ABS if abs(mv) >= 1 or abs(sv) >= 1 else 1e-6
        if abs(sv - mv) > max(tol, TOL_REL * abs(mv)):
            out.append(f"{a}: sheet {sv:g}, model {mv:g}")
    return out
```

- [ ] **Step 4: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_readback.py -q`
Expected: `3 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/render/readback.py engine/tests/v2/test_v2_readback.py
git commit -m "feat(v2): readback check - sheet equals model

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: COM writer

**Files:**
- Create: `engine/smartass/render/excel.py`
- Test: `engine/tests/v2/test_v2_excel_writer.py`

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_excel_writer.py
from smartass.model.frame import Fact, Frame
from smartass.render.excel import data_grid, write_data
from smartass.render.layout import SheetPlan


class Rng:
    def __init__(self, ws, a, b):
        self.ws, self.a, self.b = ws, a, b

    def __setattr__(self, k, v):
        if k in ("ws", "a", "b"):
            object.__setattr__(self, k, v)
        else:
            self.ws.log.append((self.a, self.b, k, v))


class Lists:
    def __init__(self, ws):
        self.ws = ws

    def Add(self, src_type, rng, link, headers):
        self.ws.log.append(("table", rng.a, rng.b))
        ws = self.ws

        class T:
            def __setattr__(self, k, v):
                ws.log.append(("table." + k, v))
        return T()


class WS:
    def __init__(self):
        self.log = []
        self.ListObjects = Lists(self)

    def Cells(self, r, c):
        return (r, c)

    def Range(self, a, b=None):
        return Rng(self, a, b if b is not None else a)


def test_data_grid_header_and_rows():
    fr = Frame([Fact("FinMax", "fm", "ac.C30", "E1", "Sep 26", "Actual", "", 5.0, "Periodic")])
    g = data_grid(fr)
    assert g[0] == ("Cube", "Query", "Account", "Entity", "Period", "Scenario", "Version", "Balance", "Value")
    assert g[1][4] == "Sep 26" and g[1][-1] == 5.0


def test_write_data_sets_text_format_then_values_then_table():
    fr = Frame([Fact("FinMax", "fm", "ac.C30", "E1", "Sep 26", "Actual", "", 5.0, "Periodic")])
    ws = WS()
    write_data(ws, fr, 3)
    kinds = [x[2] if len(x) == 4 else x[0] for x in ws.log]
    assert kinds.index("NumberFormat") < kinds.index("Value2") < kinds.index("table")
    assert ("table.Name", "Data_Q3") in ws.log


def test_answer_cells_grid_is_one_rectangle():
    from smartass.render.excel import answer_grid
    p = SheetPlan(cells={(1, 2): "T", (4, 3): "=1+1"}, last_col="C", last_row=4)
    r0, c0, g = answer_grid(p)
    assert (r0, c0) == (1, 2) and g[0][0] == "T" and g[3][1] == "=1+1" and g[1][0] == ""
```

- [ ] **Step 2: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_excel_writer.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.render.excel'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/render/excel.py
"""The COM writer: Q<n> Data as an Excel table (label columns as text first, so 'Sep 26' never becomes a date),
Q<n> Answer as one rectangle of values and formulas (Formula2), then v1's TN styling. Runs inside a user-Excel job
(v1 host().call(..., user_excel=True)); every COM call goes through v1 jobs.com (busy retry)."""
from __future__ import annotations

from smartview_mcp import styling
from smartview_mcp.excel.jobs import com

from ..model.frame import COLUMNS, Frame
from .layout import SheetPlan

XL_SRC_RANGE = 1
XL_YES = 1


def data_grid(frame: Frame) -> list[tuple]:
    return [COLUMNS, *frame.rows()]


def write_data(ws, frame: Frame, q: int) -> None:
    g = data_grid(frame)
    n = len(g)
    labels = ws.Range(ws.Cells(1, 1), ws.Cells(n, len(COLUMNS) - 1))
    com(lambda: setattr(labels, "NumberFormat", "@"))
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(n, len(COLUMNS)))
    com(lambda: setattr(rng, "Value2", tuple(g)))
    t = com(lambda: ws.ListObjects.Add(XL_SRC_RANGE, rng, None, XL_YES))
    com(lambda: setattr(t, "Name", f"Data_Q{q}"))


def answer_grid(p: SheetPlan) -> tuple[int, int, list[list]]:
    rows = [r for r, _ in p.cells]
    cols = [c for _, c in p.cells]
    r0, c0 = min(rows), min(cols)
    g = [[""] * (max(cols) - c0 + 1) for _ in range(max(rows) - r0 + 1)]
    for (r, c), v in p.cells.items():
        g[r - r0][c - c0] = "" if v is None else v
    return r0, c0, g


def write_answer(ws, p: SheetPlan, warnings: list[str]) -> list[str]:
    """Write every cell in one Formula2 rectangle, then style. Returns the styling lint findings."""
    r0, c0, g = answer_grid(p)
    rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + len(g) - 1, c0 + len(g[0]) - 1))
    com(lambda: setattr(rng, "Formula2", tuple(tuple(x) for x in g)))
    roles = {"last_col": p.last_col, "label_cols": ["B"], "last_row": p.last_row, "col_widths": {"B": 46},
             "rows": p.rows}
    return styling.apply_report(ws, roles)


def read_block(ws, p: SheetPlan) -> tuple:
    rng = ws.Range(f"A1:{p.last_col}{p.last_row}")
    v = com(lambda: rng.Value2)
    return v if isinstance(v, tuple) else ((v,),)
```

`com` is imported from `smartview_mcp.excel.jobs`. If the import path differs (for example `jobs.com` re-exported
elsewhere), use the path `answer_write.py` imports it from.

- [ ] **Step 4: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_excel_writer.py -q`
Expected: `3 passed`. `com(fn)` calls `fn()` with a busy retry, so with the fake worksheet it just runs the lambda.

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/render/excel.py engine/tests/v2/test_v2_excel_writer.py
git commit -m "feat(v2): COM writer - data table as text labels, answer in one rectangle, TN style

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: Pipeline — prepare (fetch + compute) and reply

**Files:**
- Create: `engine/smartass/pipeline.py`
- Test: `engine/tests/v2/test_v2_pipeline.py`

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_pipeline.py
import pytest
from v2util import facts, resolver, spec_dict

from smartass.data.provider import FetchResult
from smartass.model.validate import GuardError
from smartass.pipeline import prepare, reply


class FakeProvider:
    def __init__(self, frame, warn=None):
        self.frame, self.warn, self.asked = frame, warn, []

    def closed_months(self):
        return {"FinMax": (2026, 9), "FinLite": (2026, 9), "FinPlan": (2026, 9)}

    def fetch(self, req):
        self.asked.append(req)
        got = [f for f in self.frame.facts if f.query == req.query and f.period in req.periods
               and f.balance == req.balance and f.account in req.accounts and f.entity in req.entities]
        return FetchResult(req, got, [self.warn] if self.warn else [], "http", 0.01)


def test_prepare_fetches_every_request_and_computes():
    prov = FakeProvider(facts())
    pre = prepare(spec_dict(), resolver(), prov)
    assert len(prov.asked) == len(pre.requests) >= 2
    assert pre.result.key_values["Key figures · EBITDA"]["Actual"] == 65.0
    assert pre.closed == {"FinMax": (2026, 9), "FinLite": (2026, 9), "FinPlan": (2026, 9)}
    assert pre.routes == {"http": len(prov.asked)}


def test_prepare_guards_before_fetching():
    prov = FakeProvider(facts())
    d = spec_dict()
    d["queries"][0]["entity"] = "un.SY.ZZ"
    with pytest.raises(GuardError):
        prepare(d, resolver(), prov)
    assert prov.asked == []


def test_reply_shape_and_fetch_warnings():
    pre = prepare(spec_dict(), resolver(), FakeProvider(facts(), warn="FinMax x: #NoAccess (left out)"))
    out = reply(pre, q=3, workbook="C:/x.xlsx", sheets=["Q3 Answer", "Q3 Data"], mismatches=[], lint=[])
    assert out["q"] == 3 and out["sheets"] == ["Q3 Answer", "Q3 Data"]
    assert "Key figures · EBITDA" in out["key_values"]
    assert any("#NoAccess" in w for w in out["warnings"])
    assert out["checks"] == {"sheet_equals_model": "OK"}
    assert out["saved_user_workbook"] is False


def test_lines_block_is_refused_in_1b():
    d = spec_dict(blocks=[{"type": "lines", "cell": {"All Accounts": "ac.C60"}}])
    with pytest.raises(GuardError, match="lines"):
        prepare(d, resolver(), FakeProvider(facts()))
```

- [ ] **Step 2: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_pipeline.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'smartass.pipeline'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/pipeline.py
"""fin_analyze in three steps. prepare(): validate, guard, fetch every grid request in parallel (HTTP, no Excel),
compute the model. excel_job(): runs in the user's Excel (v1 host job) - target, Q<n>, Data + Answer sheets,
styling, readback, spec saved in _svmeta, sheet order, save our own file. reply(): the result dict the agent
quotes from (key_values only - never numbers it computed itself)."""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field

from .model.analyze import Result, compute, plan_requests
from .model.frame import Frame
from .model.requests import GridRequest
from .model.spec import AnalysisSpec, LinesBlock
from .model.validate import GuardError

MAX_PARALLEL = 6


@dataclass
class Prepared:
    spec: AnalysisSpec
    raw: dict
    closed: dict
    requests: list[GridRequest]
    frame: Frame
    result: Result
    warnings: list[str] = field(default_factory=list)
    routes: dict[str, int] = field(default_factory=dict)
    fetch_seconds: float = 0.0
    aggregation: dict[str, str] = field(default_factory=dict)


def prepare(raw: dict, resolver, provider) -> Prepared:
    spec = AnalysisSpec.model_validate(raw)
    if any(isinstance(b, LinesBlock) for b in spec.blocks):
        raise GuardError("the 'lines' block needs fin_drill, which arrives in Plan 1C",
                         "leave the lines block out for now")
    closed = provider.closed_months()
    reqs = plan_requests(spec, resolver, closed)
    with ThreadPoolExecutor(max_workers=MAX_PARALLEL) as ex:
        got = list(ex.map(provider.fetch, reqs))
    frame = Frame([f for g in got for f in g.facts])
    warnings = [w for g in got for w in g.warnings]
    routes: dict[str, int] = {}
    for g in got:
        routes[g.route] = routes.get(g.route, 0) + 1
    result = compute(spec, resolver, closed, frame)
    aggregation = {m: resolver.aggregation(m) for q in spec.queries for m in q.metrics}
    return Prepared(spec, raw, closed, reqs, frame, result, warnings, routes, sum(g.seconds for g in got),
                    aggregation)


def reply(pre: Prepared, q: int, workbook: str, sheets: list[str], mismatches: list[str], lint: list[str]) -> dict:
    res = pre.result
    return {
        "q": q,
        "workbook": workbook,
        "sheets": sheets,
        "key_values": res.key_values,
        "material": res.material,
        "findings": [{"severity": f.severity, "text": f.text, "where": f.where} for f in res.findings],
        "checks": {"sheet_equals_model": "OK" if not mismatches else f"DIFF in {len(mismatches)} cells"},
        "mismatches": mismatches[:20],
        "warnings": pre.warnings + lint,
        "routes": pre.routes,
        "closed": {c: f"{m[1]:02d}/{m[0]}" for c, m in pre.closed.items()},
        "fetch_seconds": round(pre.fetch_seconds, 2),
        "saved_user_workbook": False,
    }
```

- [ ] **Step 4: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_pipeline.py -q`
Expected: `4 passed`. If `compute` has a different signature after the 1A fixes (for example `closed` given as a
dict), follow the code in `engine/smartass/model/analyze.py`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/pipeline.py engine/tests/v2/test_v2_pipeline.py
git commit -m "feat(v2): pipeline prepare (guard, parallel fetch, compute) and reply

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 8: Pipeline — the Excel job

**Files:**
- Modify: `engine/smartass/pipeline.py` (append `excel_job`)
- Test: `engine/tests/v2/test_v2_pipeline_job.py`

- [ ] **Step 1: Write the failing test**

The job is COM glue. This test pins the order of the steps and the rollback, with every v1 call replaced by a
recorder.

```python
# engine/tests/v2/test_v2_pipeline_job.py
import pytest
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass import pipeline


class T:
    path = "C:/a/v2/x.xlsx"
    created = True
    wb = object()
    app = object()


def wire(monkeypatch, log, fail_at=None):
    def rec(name, ret=None):
        def f(*a, **k):
            log.append(name)
            if name == fail_at:
                raise RuntimeError("boom")
            return ret
        return f
    monkeypatch.setattr(pipeline, "_target", rec("target", T()))
    monkeypatch.setattr(pipeline, "_next_q", rec("next_q", 4))
    monkeypatch.setattr(pipeline, "_add_sheet", rec("add_sheet", object()))
    monkeypatch.setattr(pipeline, "write_data", rec("write_data"))
    monkeypatch.setattr(pipeline, "write_answer", rec("write_answer", []))
    monkeypatch.setattr(pipeline, "_calculate", rec("calculate"))
    monkeypatch.setattr(pipeline, "read_block", rec("read_block", ()))
    monkeypatch.setattr(pipeline, "compare", rec("compare", []))
    monkeypatch.setattr(pipeline, "_save_spec", rec("save_spec"))
    monkeypatch.setattr(pipeline, "_arrange", rec("arrange"))
    monkeypatch.setattr(pipeline, "_save", rec("save"))
    monkeypatch.setattr(pipeline, "_rollback", rec("rollback"))


def test_job_order(monkeypatch):
    log = []
    wire(monkeypatch, log)
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    out = pipeline.excel_job(None, pre, target="session")
    assert log == ["target", "next_q", "add_sheet", "add_sheet", "write_data", "write_answer", "calculate",
                   "read_block", "compare", "save_spec", "arrange", "save"]
    assert out["q"] == 4 and out["sheets"] == ["Q4 Answer", "Q4 Data"]


def test_job_rolls_back_on_failure(monkeypatch):
    log = []
    wire(monkeypatch, log, fail_at="write_answer")
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    with pytest.raises(RuntimeError):
        pipeline.excel_job(None, pre, target="session")
    assert log[-1] == "rollback"
```

- [ ] **Step 2: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_pipeline_job.py -q`
Expected: FAIL with `AttributeError: ... has no attribute '_target'`

- [ ] **Step 3: Implement**

Append to `engine/smartass/pipeline.py`:

```python
# ---------------------------------------------------------------- the Excel job (runs on a v1 user-Excel job thread)
from .render.excel import read_block, write_answer, write_data  # noqa: E402
from .render.layout import layout  # noqa: E402
from .render.readback import compare  # noqa: E402

DATA_TAB = 0x808080      # grey tab for data sheets (BGR)


def _target(h, target: str, title: str):
    from smartview_mcp.excel import target as xltarget
    return xltarget.resolve_target(h, target, allow_archive=False, title=title)


def _next_q(t) -> int:
    from smartview_mcp.excel import sheets
    return sheets.next_question(t)


def _add_sheet(t, name: str, q: int, color=None):
    from smartview_mcp.excel import sheets
    return sheets.add_sheet(t, name, after_q=q, color=color)


def _calculate(t) -> None:
    from smartview_mcp.excel import sheets
    from smartview_mcp.excel.jobs import com
    com(lambda: t.app.Calculate())
    sheets.wait_calculated(t.app)


def _save_spec(t, q: int, raw: dict) -> None:
    from smartview_mcp.excel import sheets
    sheets.save_spec(t, q, {"v2": raw}, ids={}, owned=None)


def _arrange(t) -> None:
    from smartview_mcp import sheetorder
    sheetorder.arrange_sheets(t.wb, mode="full")


def _save(t) -> None:
    from smartview_mcp.excel.jobs import com
    if getattr(t, "created", False):           # only the engine's own analysis file, never the user's workbook
        com(lambda: t.wb.Save())


def _rollback(t) -> None:
    from smartview_mcp.excel import sheets
    sheets.rollback(t)


def excel_job(h, pre: Prepared, target: str = "session") -> dict:
    t = _target(h, target, pre.spec.question[:60])
    try:
        q = _next_q(t)
        ws_answer = _add_sheet(t, f"Q{q} Answer", q)
        ws_data = _add_sheet(t, f"Q{q} Data", q, DATA_TAB)
        write_data(ws_data, pre.frame, q)
        cubes = sorted({x.cube for x in pre.spec.queries})
        closed = ", ".join(f"{c} {m[1]:02d}/{m[0]}" for c, m in sorted(pre.closed.items()) if c in cubes)
        plan = layout(q, pre.spec, pre.result, stock=_stock(pre), subtitle=f"NOK · {' + '.join(cubes)} · closed {closed}")
        lint = write_answer(ws_answer, plan, pre.warnings)
        _calculate(t)
        mismatches = compare(plan.addr, pre.result.cells, read_block(ws_answer, plan))
        _save_spec(t, q, pre.raw)
        _arrange(t)
        _save(t)
    except Exception:
        _rollback(t)
        raise
    return reply(pre, q, str(t.path), [f"Q{q} Answer", f"Q{q} Data"], mismatches, list(lint or []))


def _stock(pre: Prepared) -> set[str]:
    """Stock metrics (FTE): shown with one decimal (tall1), never summed over months."""
    return {m for m, agg in pre.aggregation.items() if agg == "last"}
```

- [ ] **Step 4: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_pipeline_job.py tests/v2/test_v2_pipeline.py -q`
Expected: `6 passed`

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/pipeline.py engine/tests/v2/test_v2_pipeline.py engine/tests/v2/test_v2_pipeline_job.py
git commit -m "feat(v2): Excel job - data and answer sheets, readback, spec saved, rollback on failure

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 9: The `smartass2` server and its three tools

**Files:**
- Create: `engine/smartass/server.py`
- Create: `engine/child_server_v2.py`
- Create: `engine/run_server_v2.py`
- Test: `engine/tests/v2/test_v2_server.py`

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_server.py
import pytest
from v2util import resolver

from smartass import server


def test_three_tools_registered():
    assert server.TOOL_NAMES == ("fin_analyze", "fin_lookup", "fin_status")


def test_lookup_metrics_and_members(monkeypatch):
    monkeypatch.setattr(server, "_resolver", lambda: resolver())
    monkeypatch.setattr(server, "_search_members", lambda text, cube, limit: [
        {"connection": "FinMax", "dimension": "All Entities", "name": "un.SY.2BUS", "label": "2BUS Business Norway"}])
    out = server.lookup("opex", kind="any")
    assert out["metrics"][0]["name"] == "Opex"
    assert out["metrics"][0]["members"]["FinMax"] == ["ac.C50S", "ac.C60"]
    assert out["members"][0]["name"] == "un.SY.2BUS"


def test_lookup_kind_metric_only(monkeypatch):
    monkeypatch.setattr(server, "_resolver", lambda: resolver())
    monkeypatch.setattr(server, "_search_members", lambda *a: pytest.fail("members not asked"))
    assert server.lookup("ebitda", kind="metric")["members"] == []


def test_instructions_are_short():
    assert len(server.INSTRUCTIONS) < 2500
```

- [ ] **Step 2: Run it and see it fail**

Run: `uv run pytest tests/v2/test_v2_server.py -q`
Expected: FAIL with `ImportError: cannot import name 'server'`

- [ ] **Step 3: Implement**

```python
# engine/smartass/server.py
"""smartass2: the v2 MCP server (built beside v1 until v2 passes the test set). Three tools in Plan 1B:
fin_analyze, fin_lookup, fin_status. Error text (guard), Excel jobs (_call_user, progress notes) and the lifespan
(warm-up, closing the hidden Excel) are v1's via the legacy seam; answer files and session are v2's own."""
from __future__ import annotations

import asyncio
import contextlib
import datetime as dt

from mcp.server.mcpserver import Context, MCPServer

from smartview_mcp import atlas
from smartview_mcp import server as v1

from . import legacy
from .knowledge.atlas_resolver import AtlasResolver
from .knowledge.interim import catalog_metrics
from .model.spec import AnalysisSpec

TOOL_NAMES = ("fin_analyze", "fin_lookup", "fin_status")

INSTRUCTIONS = """smartass2 (v2, read-only towards Oracle EPM: FinMax/FinLite actuals, FinPlan plan/forecast).
Loop: 1) fin_lookup to find metric names and members (entities 'un.SY.<code>' in FinMax, '<code> <name>' in FinPlan).
2) Write an AnalysisSpec (version 2): plan (depth quick|standard|deep, purpose, hypotheses), queries (cube, metrics,
entity, periods like 'closed-ytd', compare ly | budget | forecast - budget/forecast only in FinPlan), blocks (table,
kpi_tiles, trend, matrix, bridge, ranking, drivers, commentary). 3) fin_analyze(spec): builds Q<n> Answer (live
formulas) + Q<n> Data in the session workbook; a refused spec says why and how to fix it. 4) Quote numbers only from
key_values; explain material lines from drivers evidence or label them as a hypothesis. Sheet text is English;
answer the user in their language. The user's own workbooks are never saved."""

mcp = MCPServer("smartass2", instructions=INSTRUCTIONS, lifespan=v1.lifespan)


def _register(annotations):
    def deco(fn):
        mcp.tool(annotations=annotations)(v1.guard(fn))
        return fn
    return deco


def _resolver(forecast_year: int | None = None) -> AtlasResolver:
    year = forecast_year if forecast_year is not None else dt.date.today().year % 100
    return AtlasResolver(atlas.connect(), catalog_metrics(), forecast_year=year)


def _search_members(text: str, cube: str | None, limit: int) -> list[dict]:
    with contextlib.closing(atlas.connect()) as db:
        return atlas.search(db, text, connection=cube, limit=limit)


def _metric_defs(r) -> list:
    m = r.metrics
    return list(m() if callable(m) else m)


def lookup(query: str, kind: str = "any", cube: str | None = None, limit: int = 15) -> dict:
    q = query.lower()
    metrics = []
    if kind in ("any", "metric"):
        for m in _metric_defs(_resolver()):
            if q in m.name.lower():
                metrics.append({"name": m.name, "members": {c: list(v) for c, v in m.members.items()},
                                "sign": m.sign, "balance": m.balance, "aggregation": m.aggregation})
    members = _search_members(query, cube, limit) if kind in ("any", "member") else []
    return {"metrics": metrics, "members": members}


@_register(v1.LOCAL)
def fin_lookup(query: str, kind: str = "any", cube: str | None = None) -> dict:
    """Find metric names (kind 'metric') and cube members (kind 'member') by text; 'any' returns both.
    cube: FinMax | FinLite | FinPlan to narrow members. Local atlas, instant."""
    return lookup(query, kind, cube)


@_register(v1.WRITE)
async def fin_analyze(spec: AnalysisSpec, target: str = "session", ctx: Context | None = None) -> dict:
    """Run one analysis: validate the spec, fetch the data over Smart View HTTP, compute, and build Q<n> Answer
    (live formulas on table Data_Q<n>) + Q<n> Data in the session workbook (target 'session', 'new' or 'active').
    Returns key_values (quote only these), material lines, findings, checks and warnings."""
    from smartview_mcp import svhttp

    from .data.provider import SvHttpProvider
    from .pipeline import excel_job, prepare
    legacy.configure_v2()
    provider = SvHttpProvider(svhttp.shared())
    closed = await asyncio.to_thread(provider.closed_months)
    resolver = _resolver(closed["FinPlan"][0] % 100)
    pre = await asyncio.to_thread(prepare, spec.model_dump(), resolver, provider)
    return await v1._call_user(lambda h: excel_job(h, pre, target), what="fin_analyze")


@_register(v1.LOCAL)
def fin_status() -> dict:
    """Health: atlas status, metric count, and where v2 writes its answer files."""
    from smartview_mcp import config as v1config
    legacy.configure_v2()
    with contextlib.closing(atlas.connect()) as db:
        st = atlas.status(db)
    return {"analyser": str(v1config.ANALYSER_DIR), "atlas": st, "metrics": len(catalog_metrics())}


def main():
    legacy.configure_v2()
    mcp.run()
```

```python
# engine/child_server_v2.py
"""The v2 MCP server (stdio), started by run_server_v2.py."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from smartass.server import main  # noqa: E402

if __name__ == "__main__":
    main()
```

```python
# engine/run_server_v2.py
"""Stable stdio entry for smartass2: v1's relay (run_server.py) pointed at child_server_v2.py, with its own log
and a fingerprint that also covers engine/smartass/. Register with:
claude mcp add smartass2 -- <engine>\\.venv\\Scripts\\python.exe <engine>\\run_server_v2.py"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import run_server as rs  # noqa: E402

rs.CHILD = rs.HERE / "child_server_v2.py"
rs.LOG = rs.LOG.with_name("proxy_v2.log")
_v1_fingerprint = rs.fingerprint


def fingerprint() -> tuple:
    extra = []
    for f in sorted((rs.HERE / "smartass").rglob("*.py")):
        try:
            st = f.stat()
            extra.append((str(f.relative_to(rs.HERE)), st.st_mtime_ns, st.st_size))
        except OSError:
            pass
    return _v1_fingerprint() + tuple(extra)


rs.fingerprint = fingerprint

if __name__ == "__main__":
    rs.Relay().run()
```

- [ ] **Step 4: Run it and see it pass**

Run: `uv run pytest tests/v2/test_v2_server.py -q`
Expected: `4 passed`. Then start the server once and confirm it imports and stops cleanly:
`timeout 5 uv run python child_server_v2.py < /dev/null`. It should exit without an import error; a timeout exit
is fine.

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/server.py engine/child_server_v2.py engine/run_server_v2.py engine/tests/v2/test_v2_server.py
git commit -m "feat(v2): smartass2 server with fin_analyze, fin_lookup, fin_status

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 10: Full suite, lint, changelog

- [ ] **Step 1:** Run `uv run pytest -q`. Expected: all green (the 1A total plus about 30 new tests).
- [ ] **Step 2:** Run `uv run ruff check .`. Expected: `All checks passed!`
- [ ] **Step 3:** Under `## Unreleased` / `### Added` in `CHANGELOG.md` add:

```markdown
- v2 first answer (`smartass2` server, beside v1): `fin_analyze`, `fin_lookup`, `fin_status`; Smart View HTTP
  provider, interim atlas resolver, Q<n> Answer as live formulas on table Data_Q<n>, "Sheet = model" readback;
  answer files in `<analyser>\v2`. Not yet: follow-ups, packs, drill, xl_* tools, ad hoc fallback (Plan 1C).
```

- [ ] **Step 4: Commit**

```bash
git add CHANGELOG.md
git commit -m "docs(v2): changelog for the first answer

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 11: Register `smartass2` (needs Tom's OK)

The controller asks Tom before running this, because it changes his Claude Code configuration:

```bash
claude mcp add smartass2 -s user -- "C:\dev\smartass\engine\.venv\Scripts\python.exe" "C:\dev\smartass\engine\run_server_v2.py"
```

Then Tom runs `/mcp` and checks that `smartass2` lists three tools.

### Task 12: First live answer (only on Tom's go)

Hidden Edge for the Smart View sign-in; the answer opens in Tom's Excel as a new v2 session workbook. Run
`fin_analyze` with the B02-style spec below and check:
- `checks.sheet_equals_model == "OK"`.
- `key_values` EBITDA YTD matches v1's `sv_answer` for the same question (v1 bench B02: YTD 3 496 062 985).

```json
{"version": 2, "question": "EBITDA YTD for Business Norway vs last year",
 "plan": {"depth": "standard", "purpose": "Business Norway EBITDA YTD vs last year", "hypotheses": []},
 "queries": [{"id": "fm", "cube": "FinMax", "metrics": ["Revenue", "Opex", "EBITDA"], "entity": "un.SY.2BUS",
              "periods": "closed-ytd", "compare": ["ly"]}],
 "blocks": [{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA", "Revenue"]},
            {"type": "table", "query": "fm", "title": "P&L YTD"},
            {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]},
            {"type": "drivers", "query": "fm", "metric": "Opex", "depth": 2}]}
```

Whatever fails live becomes a task in Plan 1C, with the observed error. Never patch it silently.

---

## Live result (2026-11-04, Tom's go)

- Registered as `smartass2` (user scope) with `SV_DRILL_KEEP_WARM=0` and `SV_DRILL_PROFILE_SLOTS=4`.
- Q1: Business Norway EBITDA YTD Oct 26 vs LY, built with Sheet = model OK (HTTP, 2 grids, 0.73 s).
- Q2: the B02 check (YTD Sep 26) gives EBITDA 3 496 062 985, variance 89 346 421. Both are identical to v1's
  bench B02.

## Carried to Plan 1C

1. Every server keeps its own warm drill browser, and there are only three profile slots. With three v1
   sessions open, v2 got "another sv_drill is running". The Smart View HTTP route should not need a browser slot
   per server: share one signed-in browser across servers, or raise the slots by default.
2. Drivers walk through single-child chains (2BUS → 4BUS_B → 50BUS_B with the same numbers). Skip pass-through
   nodes, so that depth is spent where the variance splits.
3. A bridge whose steps leave out P&L lines gets a large "Other / unexplained" (97.7 MNOK on Q1). The finding
   fires correctly. The agent skill should default to the full P&L chain for an EBITDA bridge (Revenue, COGS,
   Opex, Other items).
