# Stage 2+3: warehouse, calc engine and the first slice `fin_pnl`: implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task by task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** "P&L for <business area>" in one call: `fin_pnl(area)` answers month and YTD against last year, budget and
the latest forecast, with materiality flags, an EBITDA bridge by line and by sub-unit, KPIs and the year-end landing.
The numbers are decided in milliseconds from a local DuckDB warehouse. The delivered workbook still proves them:
live Smart View source sheets, every answer number a formula, and a "warehouse = source sheet" check per sheet.

**Architecture:**
- `warehouse.py` stores cube cells in DuckDB (`%LOCALAPPDATA%\smartass\warehouse.duckdb`). It holds no business
  logic. `MemoryWarehouse` has the same API in plain Python, for tests and for the live-pull fallback.
- `extract.py` builds the grids from the catalogue, reads them over `svhttp` (COM fallback) and stores and reconciles
  them. `refresh.py` decides when to fill: a new closed month, an atlas restatement, a new forecast name, plus the
  question's live top-up.
- `analytics\` is the pure calc engine (no I/O, no pandas). It works on a `Frame` and every number is a `Fact` with
  its inputs.
- `fin_pnl.py` computes the slice. `fin_spec.py` turns the result into an `sv_ws_build` request plus the expected
  value of every source cell. `tools_fin.py` registers `fin_pnl` and `sv_warehouse`.

**Tech stack:** Python 3.12, DuckDB 2.0.0 (new), `tomllib`, pytest + anyio, ruff. No pandas.

**Background (read first):**
- `docs\specs\2026-10-25-stage2-3-warehouse-calc-design.md`: the approved design for this plan.
- `docs\specs\2026-10-25-finance-expert-design.md`: the stages.
- `docs\plans\2026-10-25-catalogue-and-svhttp.md`: the previous plan (svhttp, catalogue).

**How the code was checked:**
- Every code block in Tasks 1–21 was applied to a scratch copy of `main` (16b8bde) while this plan was written. There
  the suite gave `569 passed`, and ruff was clean.
- The live scripts (Tasks 22–23) compile, but they have not been run.
- If a step's code does not apply cleanly, the code on `main` has moved: adapt the edit, not the tests.

**Owner decisions this plan implements exactly (Tom, 2026-10-25):**
- **Warehouse content:** all business areas plus the company total × the catalogue's P&L lines (incl. FTE) × 36
  months. FinMax Actual; FinPlan Actual/Working, Target/Final and the latest Forecast/Final.
- **Refresh:** automatic on a new closed month and on an atlas restatement. The open month and the latest forecast
  are read live, and every answer carries a freshness stamp.
- **Database:** DuckDB.
- **First slice:** `fin_pnl(area)` = the business area's P&L, month and YTD vs LY, budget and forecast, with:
  - materiality flags;
  - an EBITDA bridge by line and by sub-unit;
  - KPIs;
  - the year-end landing.

  It is delivered through the existing `sv_ws_build` pipeline with live Smart View source sheets plus
  warehouse-vs-sheet reconciliation checks.

---

## Working environment (read before Task 0)

- **Worktree:** implement in the git worktree `C:\dev\smartass-wt\stage2-3` on branch `feat/stage2-3`. Never in
  `C:\dev\smartass`: that checkout runs the user's live MCP server, which hot-reloads on every file change.
- **Python (`PY`):** every command below writes `PY` for the shared venv,
  `C:\dev\smartass\engine\.venv\Scripts\python.exe`, and runs from `C:\dev\smartass-wt\stage2-3\engine`. Define it
  once per shell:
  - PowerShell: `function PY { & C:\dev\smartass\engine\.venv\Scripts\python.exe @args }`
  - bash: `PY() { /c/dev/smartass/engine/.venv/Scripts/python.exe "$@"; }`
- **Imports:** the venv has an editable install that points at the MAIN checkout. Always run `PY -m pytest` from the
  worktree's `engine` dir: `-m` puts the current directory first on `sys.path`, so the worktree's `smartview_mcp`
  wins. Never call `pytest.exe` directly. Check it once:
  `PY -c "import smartview_mcp; print(smartview_mcp.__file__)"` must print a path under `C:\dev\smartass-wt`.
- **Tests:**
  - unit tests: `PY -m pytest -q` (491 pass at the start);
  - lint: `PY -m ruff check .`;
  - before every commit: `PY ..\scripts\check_staged.py`.
- **Data:** never commit real figures. Tests use invented numbers only, and live scripts print PASS/FAIL, counts and
  seconds, never amounts.
- **Read-only towards EPM, always.** svxml's allowlist is the guard; this plan only adds the read-only `EnumMembers`.
- **The LLM never computes a number.** Every engine number is a `Fact` with its inputs, every number in the workbook
  is a formula, and the factual key points are drafted from facts.
- **Live tasks:** Tasks 22–23 need EPM (SSO) and the user's Excel. Only the **controller** runs them, not a
  subagent. Every other task runs offline on fakes.

## Facts this plan depends on (probed live, 2026-10-25; read-only)

**Closed month over HTTP.** After `ConnectToProvider` → `OpenApplication(FinMax)` → `OpenCube(FinMax)`:
- The request is:
  ```xml
  <req_EnumMembers><sID>{sid}</sID><dim>TimePeriods</dim><memberFilter><filter name="Children"><arg id="0">CurrRepMonth</arg></filter></memberFilter><getAtts>0</getAtts><alsTbl>none</alsTbl><allGenerations>0</allGenerations><includeDescriptionInLabel>0</includeDescriptionInLabel></req_EnumMembers>
  ```
- The answer is `<res_EnumMembers><overflow>0</overflow><mbrs enc="1">[CurrRepMonth].[Y26M09]</mbrs>...`, where
  `Y26M09` means `Sep 26`.
- The provider dimension is **`TimePeriods`**. The atlas name `All Time Periods` gives errcode 4997 ("dimension
  invalid").
- The member is the **name** `CurrRepMonth`. The alias `Current Reporting Month` gives errcode 4996.

**Atlas dimensions** (local atlas, today):
- FinMax: `All Accounts, All Balance Types, All Counterparts, All Currencies, All Entities, All Functions, All GL
  Sources, All Legal Entities, All Products, All Projects, All Reserves, All Scenarios, All Systems, All Time Periods`.
- FinPlan: `Account, All Balance Types, All Products, Entity, Period, Scenario, Version, Years`.
- FinPlan `Period` members are `Jan`..`Dec`, and `Years` members are `FY24`..`FY31`.
- Both cubes have `Periodic` and `Accumulated` under `All Balance Types`.

**Entity tree:**
- `un.SY.2WHS` has ONE child (`un.SY.3WHS_W`), whose 10 children are the units. The sub-unit bridge therefore goes
  down through single-child levels.
- `un.SY.2TEC` has 14 children. `un.SY` has 22 (more than the catalogue's areas: eliminations and other units exist).

**DuckDB 2.0.0** (scratch benchmark, 6 500 rows):
- Upserts:
  - `executemany` took over two minutes;
  - chunked multi-row `VALUES` took 3.4 s;
  - one `INSERT OR REPLACE ... SELECT unnest($1::VARCHAR[]), ...` with one list per column took 0.35 s.
- A read-only and a read-write connection to one file cannot coexist in one process (`ConnectionException`).
- Another process's lock raises `IOException`.
- `list_contains(?, col)` works for list filters.

**Existing contracts reused** (signatures checked against the code on `main`, 16b8bde):
- `svhttp.SvHttp.retrieve(conn, pov, rows, cols, suppress_missing, suppress_zeros) -> (Layout, Parsed, secs)`;
  `svhttp.shared()`; `SvHttpError`.
- `grid.Axis(dimension, members)`, `grid.Layout(pov, rows, cols)` (`.pov`, `.nr`, `.nc`, `.row_tuples`,
  `.col_tuples`), and `grid.Parsed` (`.rows = [*row members, *values]`).
- `ops.query(h, conn, pov, rows, cols, suppress_missing, suppress_zeros) -> (Layout, Parsed, secs)`. It runs on the
  host thread via `h.on_host(fn, timeout)`.
- `catalog.load()`, `resolve_entities`, `Catalog.pnl[cube]` (`Line`: key, account, kind, label, better, num, concept,
  measure), `checks`, `materiality` (pct, abs_nok, always, rule, pct_floor_nok) and `kpis` (`Kpi`: numerator,
  denominator, unit, basis, annualise).
- From `templates`: `closed_month(h, connection)`, `latest_forecast(yy)`, `split_period`, `same_period_ly`,
  `months_to`, `MONTHS`, `_fmt`, `_pct`, `_clean`.
- `tools_ws._build_job(h, tm, question, pulls, spec, target, allow_archive, q, replace=False)`, `tools_ws.PullIn`,
  `userexcel.Pull` (sheet, layout, parsed, row_labels, col_labels) and `userexcel._make_pull`.
- `specs.expand/evaluate`; `answer.py` cell types (`pull`, `var`, `varpct`, `expr`, `sum`, `ref`, `value`, `text`);
  `formulas.sum_rows`.
- `atlas.connect/children/ancestors`, `atlas_history.take_snapshot/_diff/_rows`.
- `drill_browser.ProfileLock`, `doctor.Check`, `server.tool/_call_user/_err/WRITE`.

## Design ambiguities resolved in this plan

1. **FinPlan periods.** The warehouse key has no Years column, so FinPlan months are stored like FinMax's
   (`Sep 26` = Period `Sep` of `FY26`).
2. **"36 months"** is the 36 months ending with the closed month (`Oct 23..Sep 26`). That window covers the current
   and last year's YTD, the LTM and every opening December the averages need.
3. **FinPlan "FY-1, FY, FY+1":**
   - Actual/Working for FY-1 and FY (there are no actuals in FY+1);
   - Target/Final for all three years;
   - the latest forecast for every year where `templates.latest_forecast(yy)` names one.
4. **What is stored:**
   - FinPlan and FinMax flows are stored Periodic only; YTD is the sum of the months.
   - FinMax adds one Accumulated spot column at the closed month (for the months = YTD check).
   - FTE (stock) is stored Accumulated.
   - Accumulated source-sheet cells are reconciled against the sum of the months.
5. **Materiality:**
   - It is judged on the YTD variance, and only for NOK amounts (FTE is not an amount).
   - The flag column is text set by the engine ("material" / "comment"). A note says it does not change when the
     sheets are refreshed.
6. **No constants in formulas** (tn-formelstil blocks them). The annualising factor and the run-rate's 12 and n are
   labelled input cells (`Months in a year`, `Months in the period`). Every figure is still a formula.
7. **Sub-units:** the area's children, going down single-child levels; at most 40. The bridge shows the top 5 by
   |variance| plus "Other units" = total − top 5, so it closes exactly.
8. **Reconciliation:** checked right after the pulls, before the answer sheet. Each pull's grid values (`Pull.parsed`,
   what Smart View wrote on the sheet) are compared with the engine's facts within NOK 1. The result is one static
   OK/DIFF check per source sheet, added by a new `before_answer` hook in `_build_job`.
9. **Live reads:** the open month and the latest forecast are read live at every `fin_pnl` call (`top_up`) and
   stored with `closed_month = 'open'`. Closed-period queries pass `closed_only=True`.
10. **Bridges and landing:**
    - The line bridge is FinMax EBITDA YTD vs last year (the catalogue chain: rev, cogs, pers, oth, c73).
    - The landing uses FinPlan EBITDA (the grand line).
    - The run-rate falls back to YTD / n × 12 when last year's YTD share is not positive.
11. **`compare`** limits which comparisons are flagged and ranked. The workbook always shows all three, because
    the landing needs FinPlan anyway.
12. **Restatement:**
    - It is detected on FinMax `All Entities` snapshots.
    - Touched areas are found through the members' parents and ancestors.
    - The company total is re-read with them.
13. **Failed fills:**
    - A failed fill (any grid error) keeps the old state key, so the next call retries.
    - A failed identity check stores the data and marks the fill `warn`; the doctor shows it.
14. **No warehouse:**
    - When the warehouse is missing, locked or DuckDB will not load, `fin_pnl` reads that area live into a
      `MemoryWarehouse` and says so in the notes and the result.
    - On a fresh install the doctor's `warehouse` check is OK ("not filled yet"), so the fresh-install live scenario
      stays green.
15. **Closed month:** `templates.closed_month` itself becomes HTTP first (new `SvHttp.closed_month`), then COM, with
    the daily cache. There is no separate probe task; the live check is in Task 22.

---

## File structure

| File | Responsibility |
|---|---|
| Modify `engine\pyproject.toml`, `engine\uv.lock` | `duckdb==2.0.0` (via `uv add --no-sync`) |
| Modify `engine\smartview_mcp\config.py` | `WAREHOUSE_DB` (`paths.warehouse`), `WAREHOUSE_LOCK_WAIT_S` |
| Modify `scripts\check_staged.py` | `.duckdb` files never go in git |
| Modify `engine\smartview_mcp\svxml.py` | `EnumMembers` allowed; `enum_children()`, `members()` |
| Modify `engine\smartview_mcp\svhttp.py` | `SvHttp.closed_month()` |
| Modify `engine\smartview_mcp\templates.py` | `closed_month`: HTTP first, COM fallback; `division_pnl` marked for retirement |
| Modify `engine\smartview_mcp\formulas.py`, `answer.py`, `specs.py` | `{avg: [keys]}` cell (AVERAGE over lines) |
| Modify `engine\smartview_mcp\tools_ws.py` | `_build_job(..., before_answer=None)` hook; spec doc mentions `avg` |
| Modify `engine\smartview_mcp\atlas_history.py` | `latest_snapshot()`, `diff_snapshots()` |
| Create `engine\smartview_mcp\warehouse.py` | DuckDB storage: `Row`, `Fill`, `Warehouse`, `MemoryWarehouse`, `WarehouseUnavailable` |
| Create `engine\smartview_mcp\analytics\__init__.py`, `facts.py`, `periods.py`, `frame.py` | Facts, period arithmetic, the Frame |
| Create `engine\smartview_mcp\analytics\measures.py`, `kpis.py` | Flow / stock / n+1 average; catalogue KPIs |
| Create `engine\smartview_mcp\analytics\variance.py`, `materiality.py` | Variances with sign rules; the materiality rule |
| Create `engine\smartview_mcp\analytics\bridge.py`, `landing.py` | Line and sub-unit bridges; year-end landing |
| Create `engine\smartview_mcp\analytics\reconcile.py` | Identities, months = YTD, areas vs total |
| Create `engine\smartview_mcp\extract.py` | Grids from the catalogue, warehouse keys, `route()`, `fill()` |
| Create `engine\smartview_mcp\refresh.py` | `plan_refresh`, `touched_areas`, `refresh`, `sub_units`, `top_up` |
| Create `engine\smartview_mcp\fin_pnl.py` | `Scope`, `Result`, `compute()`, `summary()`, `freshness_text()` |
| Create `engine\smartview_mcp\fin_spec.py` | `build()` → `Request(pulls, spec, expected)`; `with_reconciliation()` |
| Create `engine\smartview_mcp\tools_fin.py` | Tools `fin_pnl`, `sv_warehouse`; `prepare()` |
| Modify `engine\smartview_mcp\server.py` | Import `tools_fin`; one INSTRUCTIONS bullet |
| Modify `engine\smartview_mcp\doctor.py` | `warehouse` check; duckdb in the bundle's versions |
| Tests (create) | `test_warehouse.py`, `test_analytics_*.py`, `analytics_util.py`, `test_extract.py`, `test_refresh.py`, `fin_util.py`, `test_fin_pnl.py`, `test_fin_spec.py`, `test_tools_fin.py`, `test_build_hook.py` |
| Tests (extend) | `svxml_util.py` (EchoTransport), `test_svxml.py`, `test_svhttp.py`, `test_templates.py`, `test_formulas.py`, `test_specs.py`, `test_atlas_history.py`, `test_doctor.py`, `test_config.py`, `test_check_staged.py`, `test_smoke_stdio.py` |
| Create `engine\tests\live\warehouse_live.py` | Live (controller): closed month, base fill + reconciliation, engine = direct retrieve, timing |
| Modify `engine\tests\live\scenarios.py`, `engine\tests\live\README.md` | Scenario `S11.fin_pnl.wholesale`; `DOCTOR_CHECKS` gains `warehouse` |
| Modify `CHANGELOG.md`, `skills\smartass\reference\operations.md` | Docs |

---

## Task 0: worktree and baseline (controller)

- [ ] **Step 1: Create the worktree** (from any shell; the main checkout's working tree is not touched)

```powershell
git -C C:\dev\smartass worktree add C:\dev\smartass-wt\stage2-3 -b feat/stage2-3
```

Expected: `Preparing worktree (new branch 'feat/stage2-3')`.

- [ ] **Step 2: Baseline**

Run from `C:\dev\smartass-wt\stage2-3\engine`:
- `PY -c "import smartview_mcp; print(smartview_mcp.__file__)"`. Expected: a path under `C:\dev\smartass-wt\stage2-3`.
- `PY -m pytest -q`. Expected: `491 passed`.
- `PY -m ruff check .`. Expected: `All checks passed!`

---

## Part A: foundations

### Task 1: DuckDB dependency, warehouse path, no `.duckdb` in git

**Files:**
- Modify: `engine\pyproject.toml`, `engine\uv.lock` (generated by uv)
- Modify: `engine\smartview_mcp\config.py` (after the `QUERY_LOG = ...` line, ~138)
- Modify: `scripts\check_staged.py:10`
- Test: `engine\tests\test_config.py`, `engine\tests\test_check_staged.py` (append)

- [ ] **Step 1: Write the failing tests**

Append to `engine\tests\test_config.py`:

```python
def test_the_warehouse_lives_in_the_home_folder_and_duckdb_is_installed():
    import duckdb
    assert config.WAREHOUSE_DB == config.HOME / "warehouse.duckdb" and config.WAREHOUSE_LOCK_WAIT_S > 0
    assert duckdb.__version__ == "2.0.0"
```

Append to `engine\tests\test_check_staged.py`:

```python
def test_blocks_the_duckdb_warehouse():
    assert cs.problems("x/warehouse.duckdb", b"") and cs.problems("x/warehouse.duckdb.wal", b"")
    assert not cs.problems("engine/smartview_mcp/warehouse.py", b"")
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_config.py tests/test_check_staged.py -q`
Expected: FAIL. `ModuleNotFoundError: No module named 'duckdb'` (or `no attribute 'WAREHOUSE_DB'`), and
`assert []` for the duckdb file name.

- [ ] **Step 3: Add the dependency** (it updates only `pyproject.toml` and `uv.lock`; `--no-sync` keeps uv from
  creating a second venv in the worktree), then install it into the shared venv:

```powershell
uv add --no-sync duckdb==2.0.0
uv pip install --python C:\dev\smartass\engine\.venv\Scripts\python.exe duckdb==2.0.0
```

Expected:
- `pyproject.toml`'s `dependencies` gains `"duckdb==2.0.0"`.
- `uv pip` reports `+ duckdb==2.0.0`.
- A later `uv sync` in the MAIN checkout (which does not list duckdb until the merge) would remove it again; if
  that happens, rerun the `uv pip install` line.

- [ ] **Step 4: Implement**

In `engine\smartview_mcp\config.py`, directly after the `QUERY_LOG = ...` line:

```python
# The local warehouse of cube cells (Stage 2, DuckDB). Configurable so a shared copy can come later.
WAREHOUSE_DB = _path(_pick("SV_WAREHOUSE", "paths.warehouse", HOME / "warehouse.duckdb"))
WAREHOUSE_LOCK_WAIT_S = float(os.environ.get("SV_WAREHOUSE_LOCK_WAIT", "10"))   # wait for another writer/reader
```

In `scripts\check_staged.py`, replace line 10:

```python
DB = re.compile(r"\.(db|sqlite3?|db-wal|db-shm|duckdb|duckdb\.wal)$", re.I)
```

- [ ] **Step 5: Run the tests again**

Run: `PY -m pytest tests/test_config.py tests/test_check_staged.py -q`. Expected: all passed.

- [ ] **Step 6: Commit**

```bash
git add engine/pyproject.toml engine/uv.lock engine/smartview_mcp/config.py scripts/check_staged.py engine/tests/test_config.py engine/tests/test_check_staged.py
PY ..\scripts\check_staged.py
git commit -m "warehouse: duckdb 2.0.0, paths.warehouse; .duckdb never in git"
```

### Task 2: `svxml.enum_children` and `members` (read-only EnumMembers)

**Files:**
- Modify: `engine\smartview_mcp\svxml.py:13` (`ALLOWED`), and add two functions before `grid_cells`
- Test: `engine\tests\test_svxml.py` (append)

- [ ] **Step 1: Write the failing tests** (append to `engine\tests\test_svxml.py`; `pytest` and `svxml` are already
  imported there)

```python
ENUM = ('<res_EnumMembers><overflow>0</overflow><mbrs enc="1">[CurrRepMonth].[Y26M09]</mbrs><atts enc="0">0</atts>'
        '<mbrDesc enc="1">Y26M09</mbrDesc></res_EnumMembers>')


def test_enum_children_is_read_only_and_names_the_provider_dimension():
    body = svxml.enum_children("S1", "TimePeriods", "CurrRepMonth")
    svxml.check_read_only(body)
    assert body.startswith("<req_EnumMembers><sID>S1</sID><dim>TimePeriods</dim>")
    assert '<filter name="Children"><arg id="0">CurrRepMonth</arg></filter>' in body


def test_members_parses_the_pipe_list():
    assert svxml.members(ENUM) == ["[CurrRepMonth].[Y26M09]"]
    assert svxml.members(ENUM.replace("[CurrRepMonth].[Y26M09]", "a&amp;b|c")) == ["a&b", "c"]
    assert svxml.members('<res_EnumMembers><mbrs enc="1"></mbrs></res_EnumMembers>') == []
    with pytest.raises(svxml.SvXmlError) as e:
        svxml.members('<exception errcode="4997" type="error"><desc>dimension invalid</desc></exception>')
    assert e.value.code == 4997
    with pytest.raises(svxml.SvXmlError, match="no member list"):
        svxml.members("<res_EnumMembers></res_EnumMembers>")
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_svxml.py -q`
Expected: FAIL, `AttributeError: module 'smartview_mcp.svxml' has no attribute 'enum_children'`

- [ ] **Step 3: Implement**

In `engine\smartview_mcp\svxml.py`, replace the `ALLOWED` line:

```python
ALLOWED = frozenset({"ConnectToProvider", "GetProvisionedDataSources", "OpenApplication", "OpenCube", "ExecuteGrid",
                     "EnumMembers"})
```

Insert directly before `def grid_cells(lay: Layout)`:

```python
def enum_children(sid: str, dim: str, member: str) -> str:
    """The children of `member` in provider dimension `dim` (read-only). Live 2026-10-25: the provider names
    dimensions without 'All ' and spaces ('TimePeriods'; the atlas name 'All Time Periods' gives errcode 4997) and
    wants the member NAME ('CurrRepMonth'; the alias 'Current Reporting Month' gives errcode 4996)."""
    return (f'<req_EnumMembers><sID>{_x(sid)}</sID><dim>{_x(dim)}</dim><memberFilter><filter name="Children">'
            f'<arg id="0">{_x(member)}</arg></filter></memberFilter><getAtts>0</getAtts><alsTbl>none</alsTbl>'
            '<allGenerations>0</allGenerations><includeDescriptionInLabel>0</includeDescriptionInLabel>'
            '</req_EnumMembers>')


def members(text: str) -> list[str]:
    """Member names from res_EnumMembers (<mbrs enc="1">a|b</mbrs>, pipe-separated); [] when there are none."""
    raise_for_exception(text)
    m = re.search(r"<mbrs[^>]*>(.*?)</mbrs>", text, re.S)
    if m is None:
        raise SvXmlError("no member list in the provider answer")
    return [html.unescape(x) for x in m.group(1).split("|")] if m.group(1) else []
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_svxml.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/svxml.py engine/tests/test_svxml.py
PY ..\scripts\check_staged.py
git commit -m "svxml: read-only EnumMembers (children of a member)"
```

### Task 3: `SvHttp.closed_month()`

**Files:**
- Modify: `engine\smartview_mcp\svhttp.py`: `import re`, a `_MONTHS` tuple, and two methods before `_execute`
- Test: `engine\tests\test_svhttp.py` (append; it has `FakeTransport`, `EXPIRED`, `_sv`, `db_path`)

- [ ] **Step 1: Write the failing tests** (append)

```python
ENUM_SEP = ('<res_EnumMembers><overflow>0</overflow><mbrs enc="1">[CurrRepMonth].[Y26M09]</mbrs>'
            '<atts enc="0">0</atts><mbrDesc enc="1">Y26M09</mbrDesc></res_EnumMembers>')


def test_closed_month_over_http(db_path):
    t = FakeTransport(ENUM_SEP)
    assert _sv(t, db_path).closed_month() == "Sep 26"
    assert t.tags() == ["GetProvisionedDataSources", "ConnectToProvider", "OpenApplication", "OpenCube", "EnumMembers"]
    assert "<dim>TimePeriods</dim>" in t.bodies[-1] and ">CurrRepMonth</arg>" in t.bodies[-1]


def test_closed_month_reconnects_once_but_not_for_a_bad_name(db_path):
    t = FakeTransport(EXPIRED, ENUM_SEP.replace("Y26M09", "Y26M12"))
    assert _sv(t, db_path).closed_month() == "Dec 26" and t.tags().count("ConnectToProvider") == 2
    bad = '<exception errcode="4996" type="error"><desc>member invalid</desc></exception>'
    t2 = FakeTransport(bad, ENUM_SEP)
    with pytest.raises(svhttp.SvHttpError, match="member invalid"):
        _sv(t2, db_path).closed_month()
    assert t2.tags().count("ConnectToProvider") == 1
    with pytest.raises(svhttp.SvHttpError, match="not understood"):
        _sv(FakeTransport(ENUM_SEP.replace("[CurrRepMonth].[Y26M09]", "x")), db_path).closed_month()
    with pytest.raises(svhttp.SvHttpError, match="FinPlan has no Current Reporting Month"):
        _sv(FakeTransport(), db_path).closed_month("FinPlan")
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_svhttp.py -q -k closed_month`
Expected: FAIL, `AttributeError: 'SvHttp' object has no attribute 'closed_month'`

- [ ] **Step 3: Implement**

In `engine\smartview_mcp\svhttp.py`:
- Add `import re` after `import logging`.
- Add this line right before `class SvHttpError(Exception):`:

```python
_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
```

Insert into `class SvHttp`, directly before `def _execute(self, conn: str, lay: Layout)`:

```python
    def closed_month(self, conn: str = "FinMax") -> str:
        """The last closed month ('Current Reporting Month') as 'Sep 26': the child of member CurrRepMonth in provider
        dimension TimePeriods, '[CurrRepMonth].[Y26M09]' (proven live 2026-10-25)."""
        if conn not in ("FinMax", "FinLite"):
            raise SvHttpError(f"{conn} has no Current Reporting Month")
        if not self._lock.acquire(timeout=config.SVHTTP_TIMEOUT_S):
            raise SvHttpError("the HTTP route is busy with another query")
        try:
            names = self._enum(conn, "TimePeriods", "CurrRepMonth")
        finally:
            self._lock.release()
        for n in names:
            m = re.search(r"Y(\d\d)M(\d\d)", n)
            if m and 1 <= int(m.group(2)) <= 12:
                return f"{_MONTHS[int(m.group(2)) - 1]} {m.group(1)}"
        raise SvHttpError(f"Current Reporting Month not understood: {names[:3]}")

    def _enum(self, conn: str, dim: str, member: str) -> list[str]:
        for attempt in (1, 2):
            try:
                base, _, sid = self._session(conn)
                ans = self.t.post(base, svxml.enum_children(sid, dim, member))
                self._note_sid(conn, ans)
                return svxml.members(ans)
            except svxml.SvXmlError as e:
                if e.code in (4996, 4997):                 # unknown member or dimension: a new session won't help
                    raise SvHttpError(f"Smart View provider: {e}") from None
                self._sids.pop(conn, None)                 # maybe an expired session: reconnect once
                self._bases = None
                if attempt == 2:
                    raise SvHttpError(f"Smart View provider: {e}") from None
                log.info("svhttp %s: %s - reconnecting once", conn, e)
        raise AssertionError("unreachable")
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_svhttp.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/svhttp.py engine/tests/test_svhttp.py
PY ..\scripts\check_staged.py
git commit -m "svhttp: closed month over HTTP (EnumMembers TimePeriods/CurrRepMonth)"
```

### Task 4: `templates.closed_month`: HTTP first, COM fallback, daily cache

**Files:**
- Modify: `engine\smartview_mcp\templates.py`:
  - add `import logging` and `log = logging.getLogger(__name__)`;
  - replace `closed_month` (lines ~41-66).
- Test: `engine\tests\test_templates.py` (append)

- [ ] **Step 1: Write the failing tests** (append; `templates` and `ops` are already imported there)

```python
# ------------------------------------------------------------------ closed month: HTTP first, COM fallback, daily cache
class _Http:
    def __init__(self, answer):
        self.answer, self.calls = answer, 0

    def closed_month(self, conn="FinMax"):
        self.calls += 1
        if isinstance(self.answer, Exception):
            raise self.answer
        return self.answer


class _Host:
    def __init__(self):
        self.calls = 0

    def on_host(self, fn, timeout=None):
        self.calls += 1
        return fn()


def test_closed_month_reads_http_first_and_caches_per_day(tmp_path, monkeypatch):
    from smartview_mcp import config, svhttp
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    http = _Http("Sep 26")
    monkeypatch.setattr(svhttp, "shared", lambda: http)
    h = _Host()
    assert templates.closed_month(h) == "Sep 26" and templates.closed_month(h) == "Sep 26"
    assert http.calls == 1 and h.calls == 0                       # second call from the cache, no Excel at all


def test_closed_month_falls_back_to_com(tmp_path, monkeypatch):
    from smartview_mcp import config, svhttp
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    monkeypatch.setattr(svhttp, "shared", lambda: _Http(svhttp.SvHttpError("signed out")))
    monkeypatch.setattr(ops, "find_member", lambda h, c, n: {"info": {"ChildMbrName": "[CurrRepMonth].[Y26M08]"}})
    h = _Host()
    assert templates.closed_month(h) == "Aug 26" and h.calls == 1
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_templates.py -q -k closed_month`
Expected:
- `test_closed_month_reads_http_first_and_caches_per_day` FAILS: today's code goes straight to the COM route, and
  `ops.find_member` fails on the fake host.
- `test_closed_month_falls_back_to_com` already passes (it pins the fallback behaviour).

- [ ] **Step 3: Implement**

In `engine\smartview_mcp\templates.py`, change the imports to start with:

```python
import datetime as dt
import logging
import re
```

Add this right after the imports, before `MONTHS = [...]`:

```python
log = logging.getLogger(__name__)
```

Replace the whole `def closed_month(...)` function with:

```python
def closed_month(h: ExcelHost, connection: str = "FinMax") -> str:
    """'Current Reporting Month' -> e.g. 'Sep 26'. Smart View's HTTP protocol first (svhttp, ~0.2 s, no Excel); the
    COM route (FinMax member ChildMbrName '[CurrRepMonth].[Y26M09]' in the hidden server Excel, ~10 s + a login on a
    fresh server) is the fallback. Cached per day on disk."""
    import json
    from . import config, svhttp
    cache = config.SVKIT / "closed_month.json"
    today = dt.date.today().isoformat()
    try:
        c = json.loads(cache.read_text(encoding="utf-8"))
        if c.get("date") == today and c.get(connection):
            return c[connection]
    except (OSError, ValueError):
        c = {}
    try:
        val = svhttp.shared().closed_month(connection)
    except Exception as e:   # noqa: BLE001 - any HTTP failure: the COM route
        log.info("closed month over HTTP failed (%s); asking Smart View in Excel", e)
        info = h.on_host(lambda: ops.find_member(h, connection, "Current Reporting Month")).get("info") or {}
        m = re.search(r"Y(\d\d)M(\d\d)", str(info.get("ChildMbrName", "")))
        if not m:
            raise HostError("Could not resolve 'Current Reporting Month' in FinMax; pass the period explicitly "
                            "(e.g. 'Sep 26').") from None
        val = f"{MONTHS[int(m.group(2)) - 1]} {m.group(1)}"
    try:
        c = c if c.get("date") == today else {"date": today}
        c[connection] = val
        cache.write_text(json.dumps(c), encoding="utf-8")
    except OSError:
        pass
    return val
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_templates.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/templates.py engine/tests/test_templates.py
PY ..\scripts\check_staged.py
git commit -m "templates: closed month over HTTP first, COM fallback, cached per day"
```

### Task 5: the `{avg: [keys]}` answer cell

**Files:**
- Modify: `engine\smartview_mcp\formulas.py` (`avg_rows` before `sum_cols`; `__all__`)
- Modify: `engine\smartview_mcp\answer.py` (the cell branch after `"sum"`, `cell_trace`, the docstring)
- Modify: `engine\smartview_mcp\specs.py` (`expand` patterns, `evaluate`, docstring)
- Modify: `engine\smartview_mcp\tools_ws.py` (`SPEC_DOC`)
- Test: `engine\tests\test_formulas.py`, `engine\tests\test_specs.py` (append)

- [ ] **Step 1: Write the failing tests**

Append to `engine\tests\test_formulas.py`:

```python
def test_avg_rows_is_an_average_over_ranges_and_passes_the_checker():
    assert fx.avg_rows("C", list(range(20, 30))) == "=AVERAGE(C20:C29)"
    assert fx.avg_rows("D", [5, 7, 8]) == "=AVERAGE(D5,D7:D8)"
    assert fx.check([("Q1 Svar", "C30", fx.avg_rows("C", list(range(20, 30))))])["ok"] is True
    import pytest
    with pytest.raises(ValueError):
        fx.avg_rows("C", [])
```

Append to `engine\tests\test_specs.py`:

```python
def test_avg_pattern_expands_and_evaluates():
    sess, ids = pulls()
    spec = {"columns": [{"header": "FTE"}], "sections": [{"title": "levels", "repeat": {"over": ["Dec 25", "Jan 26"],
            "line": {"key": "lv{i}", "label": "{v}", "cells": [{"pull": "m", "row": ["ac.EBITDA"], "col": ["{v}"]}]}},
            "lines": [{"key": "favg", "label": "Average", "cells": [{"avg": "lv*"}]}]}]}
    s = specs.expand(spec, ids, sess)
    assert s["sections"][0]["lines"][-1]["cells"][0] == {"avg": ["lv0", "lv1"]}
    assert specs.evaluate(s, sess)["favg"] == [(100.0 + 110.0) / 2]
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_formulas.py tests/test_specs.py -q`
Expected: FAIL, `AttributeError: module 'smartview_mcp.formulas' has no attribute 'avg_rows'`, and the avg pattern
stays a string.

- [ ] **Step 3: Implement**

`engine\smartview_mcp\formulas.py`: insert before `def sum_cols(`:

```python
def avg_rows(col: str, rows: list[int]) -> str:
    """Average of one column over the given rows: '=AVERAGE(C13:C21)' / '=AVERAGE(C10,C20:C25)' (e.g. the FTE
    average: the opening level and each month end)."""
    if not rows:
        raise ValueError("an average needs at least one line")
    parts = [f"{col}{a}" if a == b else f"{col}{a}:{col}{b}" for a, b in _runs(sorted(rows))]
    return f"=AVERAGE({','.join(parts)})"
```

…and change `__all__` to start with `["sum_rows", "avg_rows", "sum_cols", ...` (keep the rest).

`engine\smartview_mcp\answer.py`, in `_write`, directly after the `"sum"` branch:

```python
                elif "sum" in c:
                    f = fx.sum_rows(col, [rows_by_key[k] for k in c["sum"]])
                elif "avg" in c:                                   # {"avg": [keys]}: AVERAGE over those lines
                    f = fx.avg_rows(col, [rows_by_key[k] for k in c["avg"]])
```

…in `cell_trace`, directly after the `"sum"` branch:

```python
                elif "sum" in c and isinstance(c["sum"], list):
                    ent["deps"] = [(rows_by_key[k], col) for k in c["sum"] if k in rows_by_key]
                elif "avg" in c and isinstance(c["avg"], list):
                    ent["deps"] = [(rows_by_key[k], col) for k in c["avg"] if k in rows_by_key]
```

…and in the module docstring change `{"sum": ["rev", "cogs"]}, {"value": 123}` to
`{"sum": ["rev", "cogs"]}, {"avg": ["f0", "f1"]}, {"value": 123}`.

`engine\smartview_mcp\specs.py`, in `expand.fix_cell`, replace the two `"sum"` pattern lines:

```python
        for agg in ("sum", "avg"):
            if isinstance(c.get(agg), str):
                c = {**c, agg: [k for k in keys if fnmatch.fnmatchcase(k, c[agg])]}
```

…in `evaluate`, directly after the `"sum"` branch:

```python
                    elif "avg" in c:
                        xs = [vals[k][j] for k in c["avg"] if k in vals]
                        row[j] = None if not xs or any(x is None for x in xs) else sum(xs) / len(xs)
```

…and in the module docstring change the `{"sum": "m*"}` line to
`  * {"sum": "m*"}         - every line key matching the pattern, in spec order ({"avg": "f*"} likewise).`

`engine\smartview_mcp\tools_ws.py`, in `SPEC_DOC`: change `{expr: 'rev+cogs'} | {sum: [keys] | 'm*'} |` to
`{expr: 'rev+cogs'} | {sum: [keys] | 'm*'} | {avg: [keys] | 'f*'} |`.

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_formulas.py tests/test_specs.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/formulas.py engine/smartview_mcp/answer.py engine/smartview_mcp/specs.py engine/smartview_mcp/tools_ws.py engine/tests/test_formulas.py engine/tests/test_specs.py
PY ..\scripts\check_staged.py
git commit -m "answer: {avg: [keys]} cell (AVERAGE over lines, e.g. the n+1 FTE average)"
```

### Task 6: `_build_job(..., before_answer=None)` hook

**Files:**
- Modify: `engine\smartview_mcp\tools_ws.py` (`_build_job` signature, docstring, one call before `_answer_job`)
- Test: `engine\tests\test_build_hook.py` (new)

- [ ] **Step 1: Write the failing test**

```python
"""tools_ws._build_job: the before_answer hook sees the pulls by id and can change the spec (fin_pnl's checks)."""
from types import SimpleNamespace as NS

from smartview_mcp import tools_ws, userexcel


def test_before_answer_gets_the_pulls_by_id_and_its_spec_is_built(monkeypatch):
    t = NS(path="C:/x/a.xlsx", created=False)
    monkeypatch.setattr(userexcel, "resolve_target", lambda h, target, allow, title=None: t)
    monkeypatch.setattr(userexcel, "next_question", lambda t: 3)
    monkeypatch.setattr(userexcel, "pull", lambda h, t, q, conn, pov, rows, cols, title=None: NS(
        sheet=title, parsed=NS(pov_problems=[], unknown_members=[])))
    monkeypatch.setattr(userexcel, "log_source", lambda *a, **k: None)
    monkeypatch.setattr(tools_ws, "_tidy", lambda *a, **k: None)
    built = {}

    def answer_job(t, q, spec, ids, sess, tm, warns):
        built.update(spec=spec, ids=ids)
        return {"answer_sheet": f"Q{q} Svar", "warnings": []}
    monkeypatch.setattr(tools_ws, "_answer_job", answer_job)
    seen = {}

    def hook(spec, pulls):
        seen.update(pulls)
        return {**spec, "checks": [{"label": "warehouse = sheet", "static": "OK"}]}
    pull = tools_ws.PullIn(id="m", connection="FinMax", rows=[{"dimension": "All Accounts", "members": ["ac.C30"]}],
                           columns=[{"dimension": "All Time Periods", "members": ["Sep 26"]}], title="Q3 FinMax")
    tm = NS(summary=lambda: {})
    tools_ws._build_job(None, tm, "q", [pull], {"columns": []}, "session", False, None, before_answer=hook)
    assert set(seen) == {"m"} and seen["m"].sheet == "Q3 FinMax"
    assert built["spec"]["checks"] == [{"label": "warehouse = sheet", "static": "OK"}]
    assert built["ids"] == {"m": "Q3 FinMax"}
```

- [ ] **Step 2: Run it and check it fails**

Run: `PY -m pytest tests/test_build_hook.py -q`
Expected: FAIL, `TypeError: _build_job() got an unexpected keyword argument 'before_answer'`

- [ ] **Step 3: Implement**

In `engine\smartview_mcp\tools_ws.py`, replace the `_build_job` signature and docstring:

```python
def _build_job(h, tm, question: str, pulls: list, spec: dict, target: str, allow_archive: bool, q: int | None,
               replace: bool = False, before_answer=None) -> dict:
    """All pulls as live Smart View sheets, then the answer (shared by sv_ws_build, sv_concept and fin_pnl).
    replace=True (a whole question re-run, e.g. sv_concept with q): the question's old data sheets are parked first
    and deleted only after the new build succeeded (restored if it fails).
    before_answer(spec, {pull id: Pull}) -> spec runs after the pulls and before the answer sheet (fin_pnl adds its
    warehouse = source sheet checks there)."""
```

…and replace the line `        out = _answer_job(t, qn, spec, ids, sess, tm, warns)` with:

```python
        if before_answer is not None:
            spec = before_answer(spec, {pi.id: sess[ids[pi.id]] for pi in pulls})
        out = _answer_job(t, qn, spec, ids, sess, tm, warns)
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_build_hook.py -q`. Expected: `1 passed`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/tools_ws.py engine/tests/test_build_hook.py
PY ..\scripts\check_staged.py
git commit -m "tools_ws: before_answer hook in _build_job"
```

### Task 7: `atlas_history.latest_snapshot` and `diff_snapshots`

**Files:**
- Modify: `engine\smartview_mcp\atlas_history.py` (append two functions)
- Test: `engine\tests\test_atlas_history.py` (append; uses its `db`, `two_snaps`, `C`, `DIM`)

- [ ] **Step 1: Write the failing test** (append)

```python
def test_latest_snapshot_and_diff_between_two_ids(db, monkeypatch):
    assert h.latest_snapshot(db, C, DIM) is None
    s1, s2 = two_snaps(db, monkeypatch)
    assert h.latest_snapshot(db, C, DIM) == s2
    d = h.diff_snapshots(db, s1, s2)
    assert {m["name"] for m in d["moved"]} >= {"un.SY.4REV_W"}
    assert [m["name"] for m in d["added"]] == ["un.SY.2SSM"] and h.diff_snapshots(db, s2, s2)["moved"] == []
```

- [ ] **Step 2: Run it and check it fails**

Run: `PY -m pytest tests/test_atlas_history.py -q`
Expected: FAIL, `AttributeError: module 'smartview_mcp.atlas_history' has no attribute 'latest_snapshot'`

- [ ] **Step 3: Implement** (append to `engine\smartview_mcp\atlas_history.py`)

```python
def latest_snapshot(db, connection: str, dimension: str) -> int | None:
    """The id of the newest snapshot of a dimension (a new id only when the hierarchy changed), or None."""
    ensure_schema(db)
    r = db.execute("SELECT id FROM snapshot WHERE connection=? AND dimension=? ORDER BY taken_at DESC, id DESC LIMIT 1",
                   (connection, dimension)).fetchone()
    return r[0] if r else None


def diff_snapshots(db, old_id: int, new_id: int) -> dict:
    """added / removed / moved / relabelled members between two snapshots (as in changes())."""
    ensure_schema(db)
    return _diff(_rows(db, old_id), _rows(db, new_id))
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_atlas_history.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/atlas_history.py engine/tests/test_atlas_history.py
PY ..\scripts\check_staged.py
git commit -m "atlas_history: latest snapshot id and a diff between two snapshots"
```

---

## Part B: the warehouse

### Task 8: `warehouse.py`

**Files:**
- Create: `engine\smartview_mcp\warehouse.py`
- Test: `engine\tests\test_warehouse.py` (new; every test runs on a temporary DuckDB file AND on `MemoryWarehouse`)

- [ ] **Step 1: Write the failing tests**

```python
"""warehouse.py on a temporary DuckDB file, and MemoryWarehouse with the same semantics. Every figure is invented."""
import pytest

from smartview_mcp import config, warehouse as wh
from smartview_mcp.warehouse import OPEN, Fill, MemoryWarehouse, Row, Warehouse, WarehouseUnavailable


def k(entity="un.SY.2WHS", account="ac.C30", period="Sep 26", balance="Periodic", cube="FinMax", scenario="Actual",
      version=""):
    return (cube, scenario, version, entity, account, period, balance)


def r(key, value, closed="Sep 26", t=1000.0):
    return Row(key, value, "http", t, closed)


@pytest.fixture(params=["duckdb", "memory"])
def store(request, tmp_path):
    return Warehouse(tmp_path / "w.duckdb") if request.param == "duckdb" else MemoryWarehouse()


def test_put_get_and_upsert(store):
    assert store.put([r(k(), 100.0), r(k(account="ac.C40S"), -40.0), r(k(period="Aug 26"), None)]) == 3
    got = store.get("FinMax", ["un.SY.2WHS"], ["ac.C30", "ac.C40S"], ["Sep 26", "Aug 26"], "Actual", "", "Periodic")
    assert got == {k(): 100.0, k(account="ac.C40S"): -40.0, k(period="Aug 26"): None}
    store.put([r(k(), 120.0)])                                     # same member tuple: replaced
    assert store.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic") == {k(): 120.0}
    assert store.get("FinMax", ["un.SY.2BUS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic") == {}
    assert store.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Accumulated") == {}


def test_open_rows_are_kept_out_of_closed_history(store):
    store.put([r(k(period="Oct 26"), 5.0, closed=OPEN), r(k(), 1.0)])
    ask = ("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26", "Oct 26"], "Actual", "", "Periodic")
    assert set(store.get(*ask)) == {k(), k(period="Oct 26")}
    assert set(store.get(*ask, closed_only=True)) == {k()}


def test_state_fills_and_freshness(store):
    assert store.state() == {}
    store.set_state(closed_month="Sep 26", forecast="Forecast Q326")
    store.set_state(closed_month="Oct 26")
    assert store.state() == {"closed_month": "Oct 26", "forecast": "Forecast Q326"}
    store.put([r(k(), 1.0, t=10.0), r(k(account="ac.C40S"), 2.0, t=20.0, closed=OPEN)])
    f = store.freshness("FinMax", ["un.SY.2WHS"])
    assert (f["rows"], f["oldest"], f["newest"], f["open_rows"], f["closed_months"]) == (2, 10.0, 20.0, 1, ["Sep 26"])
    store.record_fill(Fill("base", "FinMax Actual", 5.0, 1.5, 2, "ok", {"problems": []}))
    store.record_fill(Fill("top_up", "forecast", 6.0, 0.2, 1, "warn", {"problems": ["x"]}))
    last = store.last_fills(1)[0]
    assert (last["job"], last["result"], last["detail"]) == ("top_up", "warn", {"problems": ["x"]})
    st = store.status()
    assert st["rows"] == 2 and st["state"]["closed_month"] == "Oct 26" and len(st["fills"]) == 2


def test_a_missing_file_reads_as_unavailable_and_status_says_so(tmp_path):
    w = Warehouse(tmp_path / "none.duckdb")
    with pytest.raises(WarehouseUnavailable, match="no warehouse yet"):
        w.get("FinMax", [], [], [], "Actual", "", "Periodic")
    assert w.state() == {} and w.status()["exists"] is False


def test_another_writer_makes_it_unavailable(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "WAREHOUSE_LOCK_WAIT_S", 0.2)
    w = Warehouse(tmp_path / "w.duckdb")
    w.put([r(k(), 1.0)])
    other = wh.Warehouse(tmp_path / "w.duckdb")
    monkeypatch.setattr(other._lock, "acquire", lambda: False)    # another process holds the lock file
    with pytest.raises(WarehouseUnavailable, match="another smartass process"):
        other.put([r(k(), 2.0)])
    assert w.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic") == {k(): 1.0}
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_warehouse.py -q`
Expected: FAIL, `ImportError: cannot import name 'warehouse' from 'smartview_mcp'`

- [ ] **Step 3: Implement**

```python
"""The local warehouse (Stage 2): cube cells in DuckDB at config.WAREHOUSE_DB, normally
%LOCALAPPDATA%\\smartass\\warehouse.duckdb (outside OneDrive, never in git). Storage only - no business logic.

- fact: one row per cube cell (cube, scenario, version, entity, account, period, balance); an upsert replaces it.
  FinMax rows have version ''. FinPlan periods are stored like FinMax's ('Sep 26' = Period Sep of Years FY26).
  closed_month is the closed month when the row was read, or 'open' for a row read live at question time (the open
  month, the latest forecast): such rows are never taken for closed history.
- fill: one row per extract job (what, when, cells, seconds, result, detail) - freshness and the doctor read it.
- state: small key/value facts of the refresh logic (last closed month filled, atlas snapshot, forecast name).

One writer per machine: writes hold an exclusive lock file (drill_browser.ProfileLock, freed when the process dies);
every call opens and closes its own DuckDB connection, so other processes can read in between. A missing or locked
warehouse raises WarehouseUnavailable; callers then answer from live pulls and say so. MemoryWarehouse has the
same API in plain Python (tests, and the live-pull fallback)."""
from __future__ import annotations

import json
import threading
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

from . import config

WhKey = tuple[str, str, str, str, str, str, str]      # cube, scenario, version, entity, account, period, balance
COLUMNS = ("cube", "scenario", "version", "entity", "account", "period", "balance", "value", "source", "fetched_at",
           "closed_month", "atlas_build")
TYPES = ("VARCHAR",) * 7 + ("DOUBLE", "VARCHAR", "DOUBLE", "VARCHAR", "VARCHAR")
SCHEMA = """
CREATE TABLE IF NOT EXISTS fact (
  cube VARCHAR, scenario VARCHAR, version VARCHAR, entity VARCHAR, account VARCHAR, period VARCHAR, balance VARCHAR,
  value DOUBLE, source VARCHAR, fetched_at DOUBLE, closed_month VARCHAR, atlas_build VARCHAR,
  PRIMARY KEY (cube, scenario, version, entity, account, period, balance));
CREATE SEQUENCE IF NOT EXISTS fill_id;
CREATE TABLE IF NOT EXISTS fill (
  id BIGINT DEFAULT nextval('fill_id') PRIMARY KEY, job VARCHAR, what VARCHAR, started_at DOUBLE, seconds DOUBLE,
  cells BIGINT, result VARCHAR, detail VARCHAR);
CREATE TABLE IF NOT EXISTS state (key VARCHAR PRIMARY KEY, value VARCHAR);
"""
OPEN = "open"                                          # closed_month of a row read live at question time


class WarehouseUnavailable(Exception):
    """The warehouse file is missing or another process holds it; answer from live pulls instead."""


@dataclass(frozen=True)
class Row:
    key: WhKey
    value: float | None
    source: str                 # 'http' | 'com'
    fetched_at: float
    closed_month: str           # the closed month at read time, or OPEN
    atlas_build: str = ""       # the atlas snapshot the members were checked against

    def record(self) -> tuple:
        return (*self.key, self.value, self.source, self.fetched_at, self.closed_month, self.atlas_build)


@dataclass(frozen=True)
class Fill:
    job: str                    # base | new_month | restatement | forecast | top_up
    what: str
    started_at: float
    seconds: float
    cells: int
    result: str                 # ok | warn | error
    detail: dict

    def as_dict(self) -> dict:
        return asdict(self)


def _match(key: WhKey, cube, entities, accounts, periods, scenario, version, balance) -> bool:
    c, s, v, e, a, p, b = key
    return (c == cube and s == scenario and v == version and b == balance and e in entities and a in accounts
            and p in periods)


class MemoryWarehouse:
    """The warehouse API over dicts: for tests and for answering from live pulls when the file is unavailable."""

    def __init__(self):
        self.rows: dict[WhKey, Row] = {}
        self.fills: list[Fill] = []
        self._state: dict[str, str] = {}

    def put(self, rows: Iterable[Row]) -> int:
        n = 0
        for r in rows:
            self.rows[r.key] = r
            n += 1
        return n

    def get(self, cube: str, entities, accounts, periods, scenario: str, version: str, balance: str,
            closed_only: bool = False) -> dict[WhKey, float | None]:
        ents, accs, pers = set(entities), set(accounts), set(periods)
        return {k: r.value for k, r in self.rows.items()
                if _match(k, cube, ents, accs, pers, scenario, version, balance)
                and not (closed_only and r.closed_month == OPEN)}

    def freshness(self, cube: str | None = None, entities=None) -> dict:
        rows = [r for k, r in self.rows.items() if (cube is None or k[0] == cube)
                and (entities is None or k[3] in set(entities))]
        return _freshness(len(rows), [r.fetched_at for r in rows], sum(r.closed_month == OPEN for r in rows),
                          {r.closed_month for r in rows if r.closed_month != OPEN})

    def record_fill(self, fill: Fill) -> None:
        self.fills.append(fill)

    def last_fills(self, n: int = 10) -> list[dict]:
        return [f.as_dict() for f in reversed(self.fills[-n:])]

    def state(self) -> dict[str, str]:
        return dict(self._state)

    def set_state(self, **values: str) -> None:
        self._state.update({k: str(v) for k, v in values.items()})

    def status(self) -> dict:
        return {"path": None, "exists": True, "rows": len(self.rows), "fills": self.last_fills(5),
                "state": self.state(), **self.freshness()}


def _freshness(n: int, stamps: list[float], open_rows: int, closed: set[str]) -> dict:
    return {"rows": n, "oldest": min(stamps) if stamps else None, "newest": max(stamps) if stamps else None,
            "open_rows": open_rows, "closed_months": sorted(closed)}


_mu = threading.RLock()                                # one DuckDB connection at a time in this process


class Warehouse:
    """The DuckDB file. Each method opens, works and closes; writes hold the machine-wide lock file."""

    def __init__(self, path: Path | None = None):
        from .drill_browser import ProfileLock
        self.path = Path(path or config.WAREHOUSE_DB)
        self._lock = ProfileLock(self.path)            # <path>.lock

    def _connect(self, write: bool):
        import duckdb
        if not write and not self.path.exists():
            raise WarehouseUnavailable(f"no warehouse yet ({self.path})")
        deadline = time.monotonic() + config.WAREHOUSE_LOCK_WAIT_S
        while True:
            try:
                if write:
                    self.path.parent.mkdir(parents=True, exist_ok=True)
                    db = duckdb.connect(str(self.path))
                    db.execute(SCHEMA)
                    return db
                return duckdb.connect(str(self.path), read_only=True)
            except (duckdb.IOException, duckdb.ConnectionException) as e:
                if time.monotonic() > deadline:
                    raise WarehouseUnavailable(f"warehouse {self.path} is in use: {str(e).splitlines()[0]}") from None
                time.sleep(0.1)

    def _write(self, fn):
        with _mu:
            deadline = time.monotonic() + config.WAREHOUSE_LOCK_WAIT_S
            while not self._lock.acquire():
                if time.monotonic() > deadline:
                    raise WarehouseUnavailable(f"another smartass process is writing {self.path}")
                time.sleep(0.1)
            try:
                db = self._connect(True)
                try:
                    return fn(db)
                finally:
                    db.close()
            finally:
                self._lock.release()

    def _read(self, fn):
        with _mu:
            db = self._connect(False)
            try:
                return fn(db)
            finally:
                db.close()

    def put(self, rows: Iterable[Row]) -> int:
        recs = [r.record() for r in rows]
        if not recs:
            return 0
        cols = [list(c) for c in zip(*recs)]
        sel = ", ".join(f"unnest(${i + 1}::{t}[])" for i, t in enumerate(TYPES))   # one statement: ~0.3 s/6500 rows

        def go(db):
            db.execute(f"INSERT OR REPLACE INTO fact SELECT {sel}", cols)
            return len(recs)
        return self._write(go)

    def get(self, cube: str, entities, accounts, periods, scenario: str, version: str, balance: str,
            closed_only: bool = False) -> dict[WhKey, float | None]:
        sql = ("SELECT cube, scenario, version, entity, account, period, balance, value FROM fact WHERE cube=? AND "
               "scenario=? AND version=? AND balance=? AND list_contains(?, entity) AND list_contains(?, account) "
               "AND list_contains(?, period)" + (" AND closed_month <> 'open'" if closed_only else ""))
        args = [cube, scenario, version, balance, list(entities), list(accounts), list(periods)]
        try:
            return self._read(lambda db: {tuple(r[:7]): r[7] for r in db.execute(sql, args).fetchall()})
        except WarehouseUnavailable:
            raise
        except Exception as e:                       # a damaged or foreign file: say so, never a silent gap
            raise WarehouseUnavailable(f"cannot read {self.path}: {e}") from None

    def freshness(self, cube: str | None = None, entities=None) -> dict:
        where, args = ["1=1"], []
        if cube is not None:
            where.append("cube=?")
            args.append(cube)
        if entities is not None:
            where.append("list_contains(?, entity)")
            args.append(list(entities))
        sql = (f"SELECT count(*), min(fetched_at), max(fetched_at), count(*) FILTER (WHERE closed_month='open'), "
               f"list(DISTINCT closed_month) FILTER (WHERE closed_month<>'open') FROM fact WHERE {' AND '.join(where)}")
        n, lo, hi, op, closed = self._read(lambda db: db.execute(sql, args).fetchone())
        return {"rows": n, "oldest": lo, "newest": hi, "open_rows": op or 0, "closed_months": sorted(closed or [])}

    def record_fill(self, fill: Fill) -> None:
        self._write(lambda db: db.execute(
            "INSERT INTO fill (job, what, started_at, seconds, cells, result, detail) VALUES (?,?,?,?,?,?,?)",
            [fill.job, fill.what, fill.started_at, fill.seconds, fill.cells, fill.result,
             json.dumps(fill.detail, ensure_ascii=False)]))

    def last_fills(self, n: int = 10) -> list[dict]:
        rows = self._read(lambda db: db.execute(
            "SELECT job, what, started_at, seconds, cells, result, detail FROM fill ORDER BY id DESC LIMIT ?",
            [n]).fetchall())
        return [{"job": j, "what": w, "started_at": s, "seconds": sec, "cells": c, "result": r,
                 "detail": json.loads(d or "{}")} for j, w, s, sec, c, r, d in rows]

    def state(self) -> dict[str, str]:
        try:
            return dict(self._read(lambda db: db.execute("SELECT key, value FROM state").fetchall()))
        except WarehouseUnavailable:
            if self.path.exists():
                raise
            return {}                                   # no file yet: nothing filled

    def set_state(self, **values: str) -> None:
        self._write(lambda db: db.execute(
            "INSERT OR REPLACE INTO state SELECT unnest($1::VARCHAR[]), unnest($2::VARCHAR[])",
            [list(values), [str(v) for v in values.values()]]))

    def status(self) -> dict:
        if not self.path.exists():
            return {"path": str(self.path), "exists": False, "rows": 0, "fills": [], "state": {}}
        return {"path": str(self.path), "exists": True, "fills": self.last_fills(5), "state": self.state(),
                **self.freshness()}
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_warehouse.py -q`. Expected: `8 passed` (three tests run on both stores).

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/warehouse.py engine/tests/test_warehouse.py
PY ..\scripts\check_staged.py
git commit -m "warehouse: DuckDB fact/fill/state store with a lock file; MemoryWarehouse"
```

---

## Part C: the calc engine (`analytics\`, pure)

### Task 9: facts, periods and the Frame

**Files:**
- Create: `engine\smartview_mcp\analytics\__init__.py`, `facts.py`, `periods.py`, `frame.py`
- Test: `engine\tests\test_analytics_base.py` (new)

- [ ] **Step 1: Write the failing tests**

```python
"""analytics: facts, periods and frames (synthetic figures only)."""
import pytest

from smartview_mcp import catalog
from smartview_mcp.analytics import periods
from smartview_mcp.analytics.facts import Fact, combine, missing, total
from smartview_mcp.analytics.frame import from_rows


def test_combine_and_total_carry_inputs_and_reasons():
    a = Fact("a", 10.0, "nok", "A", (("k1",),))
    b = Fact("b", 4.0, "nok", "B", (("k2",), ("k1",)))
    d = combine("d", "nok", "a - b", lambda x, y: x - y, a, b)
    assert (d.value, d.inputs) == (6.0, (("k1",), ("k2",)))
    gone = missing("m", "nok", "M", "not read")
    assert combine("x", "nok", "a/m", lambda x, y: x / y, a, gone).reason == "m: not read"
    zero = Fact("z", 0.0, "nok", "Z")
    q = combine("q", "pct", "a/z", lambda x, y: x / y, a, zero)
    assert q.value is None and q.reason == "division by zero"
    assert total("t", [a, b], "nok", "sum").value == 14.0 and total("t", [], "nok", "none").value == 0.0
    assert total("t", [a, gone], "nok", "sum").value is None
    assert d.as_dict()["inputs"] == [["k1"], ["k2"]] and "reason" not in d.as_dict()


def test_periods():
    assert periods.shift("Jan 26", -1) == "Dec 25" and periods.shift("Nov 26", 3) == "Feb 27"
    assert periods.ytd("Mar 26") == ["Jan 26", "Feb 26", "Mar 26"]
    assert periods.ly("Sep 26") == "Sep 25" and periods.opening("Sep 26") == "Dec 25"
    assert periods.ltm("Sep 26")[0] == "Oct 25" and periods.ltm("Sep 26")[-1] == "Sep 26"
    assert periods.remaining("Sep 26") == ["Oct 26", "Nov 26", "Dec 26"] and periods.remaining("Dec 26") == []
    w = periods.window("Sep 26")
    assert len(w) == 36 and w[0] == "Oct 23" and w[-1] == "Sep 26"
    assert periods.between("Sep 26", "Nov 26") == ["Oct 26", "Nov 26"] and periods.between("Sep 26", "Sep 26") == []
    assert periods.after("Oct 26", "Sep 26") and not periods.after("Sep 26", "Sep 26")
    assert periods.finplan("Sep 26") == ("Sep", "FY26") and periods.from_finplan("Sep", "FY26") == "Sep 26"
    assert periods.year(25)[0] == "Jan 25" and len(periods.year(25)) == 12


def test_frame_takes_flows_periodic_and_stocks_accumulated():
    lines = catalog.load().pnl["FinMax"]
    rows = {("FinMax", "Actual", "", "E", "ac.C30", "Sep 26", "Periodic"): 100.0,
            ("FinMax", "Actual", "", "E", "ac.C30", "Sep 26", "Accumulated"): 900.0,     # not a flow's month
            ("FinMax", "Actual", "", "E", "ac.STA20Y", "Sep 26", "Accumulated"): 12.5,
            ("FinMax", "Actual", "", "E", "ac.STA20Y", "Sep 26", "Periodic"): 0.5,        # a stock's movement
            ("FinMax", "Actual", "", "E", "ac.C40S", "Sep 26", "Periodic"): None,
            ("FinMax", "Budget", "", "E", "ac.C30", "Sep 26", "Periodic"): 1.0,           # unmapped scenario
            ("FinMax", "Actual", "", "E", "ac.XYZ", "Sep 26", "Periodic"): 1.0}           # not a catalogue line
    f = from_rows("FinMax", lines, rows, {("Actual", ""): "actual"})
    assert f.get("E", "rev", "Sep 26", "actual") == (100.0, ("FinMax", "Actual", "", "E", "ac.C30", "Sep 26",
                                                             "Periodic"))
    assert f.get("E", "fte", "Sep 26", "actual")[0] == 12.5
    assert f.get("E", "cogs", "Sep 26", "actual")[0] == 0.0          # read, but blank: counts as 0
    assert f.get("E", "rev", "Aug 26", "actual") is None             # never read
    assert len(f.values) == 3
    with pytest.raises(KeyError, match="no line 'nope'"):
        f.line("nope")
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_analytics_base.py -q`
Expected: FAIL, `ModuleNotFoundError: No module named 'smartview_mcp.analytics'`

- [ ] **Step 3: Implement**

`engine\smartview_mcp\analytics\__init__.py`:

```python
"""Deterministic calc engine (Stage 3): pure functions over small frames of warehouse values - no I/O, no Excel,
no pandas. Every number it produces is a facts.Fact with its definition and the warehouse cells it came from."""
```

`engine\smartview_mcp\analytics\facts.py`:

```python
"""Facts: every number the calc engine produces, with what it means and which warehouse cells it came from, so any
figure in an answer can be traced. A number that cannot be computed is a Fact with value None and a reason - never a
silent gap, never a division by zero."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Iterable

UNITS = ("nok", "fte", "pct", "nok_per_fte", "count")


@dataclass(frozen=True)
class Fact:
    id: str
    value: float | None
    unit: str                         # one of UNITS; pct is a ratio (0.1 = 10 %)
    definition: str
    inputs: tuple[tuple, ...] = ()    # warehouse keys (cube, scenario, version, entity, account, period, balance)
    reason: str | None = None         # why value is None

    def as_dict(self, max_inputs: int = 40) -> dict:
        out = {"id": self.id, "value": self.value, "unit": self.unit, "definition": self.definition,
               "inputs": [list(k) for k in self.inputs[:max_inputs]]}
        if len(self.inputs) > max_inputs:
            out["inputs_more"] = len(self.inputs) - max_inputs
        if self.reason:
            out["reason"] = self.reason
        return out


def missing(id: str, unit: str, definition: str, reason: str, inputs: Iterable[tuple] = ()) -> Fact:
    return Fact(id, None, unit, definition, tuple(inputs), reason)


def _inputs(facts: Iterable[Fact]) -> tuple[tuple, ...]:
    return tuple(dict.fromkeys(k for f in facts for k in f.inputs))


def total(id: str, facts: list[Fact], unit: str, definition: str) -> Fact:
    """The sum of facts (0 for none); missing if any part is."""
    gone = next((f for f in facts if f.value is None), None)
    if gone is not None:
        return missing(id, unit, definition, f"{gone.id}: {gone.reason}", _inputs(facts))
    return Fact(id, float(sum(f.value for f in facts)), unit, definition, _inputs(facts))


def combine(id: str, unit: str, definition: str, fn: Callable[..., float], *facts: Fact) -> Fact:
    """fn over the facts' values; missing when a part is missing or fn divides by zero."""
    gone = next((f for f in facts if f.value is None), None)
    if gone is not None:
        return missing(id, unit, definition, f"{gone.id}: {gone.reason}", _inputs(facts))
    try:
        return Fact(id, float(fn(*(f.value for f in facts))), unit, definition, _inputs(facts))
    except ZeroDivisionError:
        return missing(id, unit, definition, "division by zero", _inputs(facts))
```

`engine\smartview_mcp\analytics\periods.py`:

```python
"""Period arithmetic on FinMax month names ('Sep 26'), reusing templates' helpers. FinPlan months are stored the same
way in the warehouse ('Sep 26' = Period Sep of Years FY26)."""
from __future__ import annotations

from ..templates import MONTHS, months_to, same_period_ly, split_period


def label(month: int, yy: int) -> str:
    return f"{MONTHS[month - 1]} {yy:02d}"


def shift(p: str, n: int) -> str:
    """The month n months after p (before it for negative n)."""
    mo, yy = split_period(p)
    i = yy * 12 + mo - 1 + n
    return label(i % 12 + 1, i // 12)


def index(p: str) -> int:
    mo, yy = split_period(p)
    return yy * 12 + mo


def after(a: str, b: str) -> bool:
    return index(a) > index(b)


def ytd(p: str) -> list[str]:
    return months_to(p)


def ly(p: str) -> str:
    return same_period_ly(p)


def ltm(p: str) -> list[str]:
    return [shift(p, -k) for k in range(11, -1, -1)]


def year(yy: int) -> list[str]:
    return [label(m, yy) for m in range(1, 13)]


def remaining(p: str) -> list[str]:
    """The months after p in its year (none after December)."""
    mo, yy = split_period(p)
    return [label(m, yy) for m in range(mo + 1, 13)]


def opening(p: str) -> str:
    """The previous December: the opening level of a stock (FTE) for p's year."""
    return label(12, split_period(p)[1] - 1)


def window(closed: str, n: int = 36) -> list[str]:
    """The n months ending with the closed month, oldest first."""
    return [shift(closed, -k) for k in range(n - 1, -1, -1)]


def between(a: str, b: str) -> list[str]:
    """The months after a up to and including b (empty if b is not after a)."""
    return [shift(a, k) for k in range(1, index(b) - index(a) + 1)]


def finplan(p: str) -> tuple[str, str]:
    """'Sep 26' -> ('Sep', 'FY26'): FinPlan's Period and Years members."""
    mo, yy = split_period(p)
    return MONTHS[mo - 1], f"FY{yy:02d}"


def from_finplan(period: str, years: str) -> str:
    """('Sep', 'FY26') -> 'Sep 26'."""
    return f"{period} {years[2:]}"
```

`engine\smartview_mcp\analytics\frame.py`:

```python
"""A Frame: one cube's warehouse values keyed by (entity, line key, period, scenario key), with the catalogue lines
attached - the calc engine's only input. Scenario keys are the catalogue's ([scenarios]: actual, actual_finplan,
budget, forecast). Flow lines come from Periodic cells, stock lines (FTE) from Accumulated cells (the month-end
level)."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Mapping

from ..catalog import Line

Key = tuple[str, str, str, str]                                # entity, line key, period, scenario key
WhKey = tuple[str, str, str, str, str, str, str]               # warehouse key (see warehouse.py)


@dataclass(frozen=True)
class Frame:
    cube: str
    lines: tuple[Line, ...]
    values: Mapping[Key, float | None]
    sources: Mapping[Key, WhKey] = field(default_factory=dict)

    def line(self, key: str) -> Line:
        for x in self.lines:
            if x.key == key:
                return x
        raise KeyError(f"{self.cube} has no line '{key}'")

    def get(self, entity: str, line: str, period: str, scenario: str) -> tuple[float, WhKey | tuple] | None:
        """(value, warehouse key) - a cell read as blank (#Missing) counts as 0 - or None when it was never read."""
        k = (entity, line, period, scenario)
        if k not in self.values:
            return None
        v = self.values[k]
        return (0.0 if v is None else float(v)), self.sources.get(k, (self.cube, scenario, *k[:3]))


def balance_of(line: Line) -> str:
    return "Accumulated" if line.measure == "stock" else "Periodic"


def from_rows(cube: str, lines, rows: Mapping[WhKey, float | None], scenarios: Mapping[tuple[str, str], str]) -> Frame:
    """Frame from warehouse rows. scenarios maps (scenario member, version) to a scenario key, e.g.
    {('Actual', ''): 'actual'} for FinMax or {('Target', 'Final'): 'budget'} for FinPlan; other rows are ignored."""
    by_account = {x.account: x for x in lines}
    values, sources = {}, {}
    for key, v in rows.items():
        c, scen, ver, ent, account, period, bal = key
        line = by_account.get(account)
        s = scenarios.get((scen, ver))
        if c != cube or line is None or s is None or bal != balance_of(line):
            continue
        values[(ent, line.key, period, s)] = v
        sources[(ent, line.key, period, s)] = key
    return Frame(cube, tuple(lines), values, sources)
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_analytics_base.py -q`. Expected: `3 passed`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/analytics engine/tests/test_analytics_base.py
PY ..\scripts\check_staged.py
git commit -m "analytics: facts, periods and the Frame"
```

### Task 10: measures and KPIs (the n+1 average, annualising)

**Files:**
- Create: `engine\smartview_mcp\analytics\measures.py`, `engine\smartview_mcp\analytics\kpis.py`
- Create: `engine\tests\analytics_util.py` (test helper, imported like `svxml_util`)
- Test: `engine\tests\test_analytics_measures.py` (new)

- [ ] **Step 1: Write the helper and the failing tests**

`engine\tests\analytics_util.py`:

```python
"""Synthetic frames for the analytics tests (imported like svxml_util; every figure is invented)."""
from smartview_mcp import catalog
from smartview_mcp.analytics import periods
from smartview_mcp.analytics.frame import Frame

E = "un.SY.2WHS"


def frame(cube="FinMax", values=None):
    return Frame(cube, catalog.load().pnl[cube], dict(values or {}))


def months(values: dict, line: str, scenario: str, start: str, amounts: list, entity: str = E) -> dict:
    """values with one entry per month from `start` (inclusive)."""
    for i, v in enumerate(amounts):
        values[(entity, line, periods.shift(start, i), scenario)] = v
    return values
```

`engine\tests\test_analytics_measures.py`:

```python
"""analytics.measures and analytics.kpis: flows sum, stocks take levels, the n+1 average, annualising."""
from analytics_util import E, frame, months

from smartview_mcp import catalog
from smartview_mcp.analytics.kpis import kpi
from smartview_mcp.analytics.measures import amount, average_level, flow, level


def _fte_frame():
    v = months({}, "fte", "actual", "Dec 25", [10.0, 12.0, 14.0, 16.0])      # Dec 25 opening, Jan..Mar 26
    months(v, "opex", "actual", "Jan 26", [-300.0, -300.0, -300.0])
    months(v, "rev", "actual", "Jan 26", [1000.0, 0.0, 1000.0])
    months(v, "ebitda", "actual", "Jan 26", [100.0, -50.0, 150.0])
    return frame(values=v)


def test_flow_sums_months_and_names_its_inputs():
    f = flow(_fte_frame(), E, "opex", "actual", ["Jan 26", "Feb 26", "Mar 26"])
    assert f.value == -900.0 and f.unit == "nok" and len(f.inputs) == 3
    assert flow(_fte_frame(), E, "opex", "actual", []).value == 0.0
    gone = flow(_fte_frame(), E, "opex", "actual", ["Mar 26", "Apr 26"])
    assert gone.value is None and "Apr 26" in gone.reason


def test_stock_level_and_the_n_plus_1_average():
    fr = _fte_frame()
    assert level(fr, E, "fte", "actual", "Mar 26").value == 16.0
    avg = average_level(fr, E, "fte", "actual", ["Jan 26", "Feb 26", "Mar 26"])
    assert avg.value == (10 + 12 + 14 + 16) / 4 and avg.unit == "fte" and len(avg.inputs) == 4
    assert "4 levels (Dec 25 opening + 3 month ends)" in avg.definition
    assert amount(fr, E, "fte", "actual", ["Jan 26", "Feb 26", "Mar 26"]).value == 16.0   # stock: period end
    assert average_level(fr, E, "fte", "actual", ["Jan 25"]).value is None               # Dec 24 never read


def test_kpis_margin_and_annualised_cost_per_fte():
    cat = catalog.load()
    k = {x.key: x for x in cat.kpis}
    fr = _fte_frame()
    q1 = ["Jan 26", "Feb 26", "Mar 26"]
    m = kpi(fr, k["ebitda_margin"], E, "actual", q1)
    assert m.value == 200.0 / 2000.0 and m.unit == "pct"
    c = kpi(fr, k["opex_per_fte"], E, "actual", q1)
    assert c.value == -900.0 / 13.0 * 12 / 3 and c.unit == "nok_per_fte" and "x 12/3" in c.definition
    zero = kpi(fr, k["ebitda_margin"], E, "actual", ["Feb 26"])
    assert zero.value is None and zero.reason == "division by zero"
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_analytics_measures.py -q`
Expected: FAIL, `ModuleNotFoundError: No module named 'smartview_mcp.analytics.kpis'`

- [ ] **Step 3: Implement**

`engine\smartview_mcp\analytics\measures.py`:

```python
"""Measures over a Frame: a flow sums its months; a stock (FTE) takes the period-end level; the average of a stock is
the opening level (the previous December) plus each month end - n+1 points (owner decision 2026-10-25)."""
from __future__ import annotations

from . import periods
from .facts import Fact, missing
from .frame import Frame


def _unit(frame: Frame, line: str) -> str:
    return "fte" if frame.line(line).measure == "stock" else "nok"


def _span(months: list[str]) -> str:
    if not months:
        return "none"
    return months[0] if len(months) == 1 else f"{months[0]}..{months[-1]}"


def _values(frame: Frame, entity: str, line: str, scenario: str, months: list[str], id: str, definition: str):
    """((values, warehouse keys), None), or (None, a missing Fact naming the first cell never read)."""
    vals, keys = [], []
    for m in months:
        got = frame.get(entity, line, m, scenario)
        if got is None:
            return None, missing(id, _unit(frame, line), definition,
                                 f"no warehouse value for {entity} {line} {m} ({scenario})", keys)
        vals.append(got[0])
        keys.append(got[1])
    return (vals, tuple(keys)), None


def flow(frame: Frame, entity: str, line: str, scenario: str, months: list[str], id: str | None = None) -> Fact:
    """The sum of the months (0 for no months, e.g. the remaining months after December)."""
    id = id or f"{scenario}.{entity}.{line}.{_span(months)}"
    definition = f"{frame.cube} {line} {scenario} {entity}, sum of {_span(months)}"
    got, gone = _values(frame, entity, line, scenario, months, id, definition)
    if gone:
        return gone
    vals, keys = got
    return Fact(id, float(sum(vals)), _unit(frame, line), definition, keys)


def level(frame: Frame, entity: str, line: str, scenario: str, period: str, id: str | None = None) -> Fact:
    id = id or f"{scenario}.{entity}.{line}.{period}"
    definition = f"{frame.cube} {line} {scenario} {entity}, level at the end of {period}"
    got, gone = _values(frame, entity, line, scenario, [period], id, definition)
    if gone:
        return gone
    vals, keys = got
    return Fact(id, vals[0], _unit(frame, line), definition, keys)


def average_level(frame: Frame, entity: str, line: str, scenario: str, months: list[str],
                  id: str | None = None) -> Fact:
    """The average of the opening level (the December before months[0]) and each month end: n+1 points."""
    points = [periods.opening(months[0])] + list(months)
    id = id or f"{scenario}.{entity}.{line}.avg.{_span(months)}"
    definition = (f"{frame.cube} {line} {scenario} {entity}, average of {len(points)} levels "
                  f"({points[0]} opening + {len(months)} month ends)")
    got, gone = _values(frame, entity, line, scenario, points, id, definition)
    if gone:
        return gone
    vals, keys = got
    return Fact(id, sum(vals) / len(vals), _unit(frame, line), definition, keys)


def amount(frame: Frame, entity: str, line: str, scenario: str, months: list[str], id: str | None = None) -> Fact:
    """A line over a period by its measure: flow = the sum of the months, stock = the level at the last month."""
    if frame.line(line).measure == "stock":
        return level(frame, entity, line, scenario, months[-1], id)
    return flow(frame, entity, line, scenario, months, id)
```

`engine\smartview_mcp\analytics\kpis.py`:

```python
"""Catalogue KPIs: numerator / denominator over one cube's lines. basis average_stock divides by the average level
(opening December + each month end, n+1 points); annualise scales a part-year numerator to twelve months."""
from __future__ import annotations

from ..catalog import Kpi
from .facts import Fact, combine
from .frame import Frame
from .measures import amount, average_level


def kpi(frame: Frame, k: Kpi, entity: str, scenario: str, months: list[str], id: str | None = None) -> Fact:
    id = id or f"kpi.{k.key}.{scenario}.{months[-1]}"
    num = amount(frame, entity, k.numerator, scenario, months)
    if k.basis == "average_stock":
        den = average_level(frame, entity, k.denominator, scenario, months)
    else:
        den = amount(frame, entity, k.denominator, scenario, months)
    factor = 12 / len(months) if k.annualise else 1.0
    how = f"{k.numerator} / {k.denominator}" + (f" x 12/{len(months)}" if k.annualise else "")
    unit = "pct" if k.unit == "pct" else ("nok_per_fte" if k.basis == "average_stock" else "nok")
    return combine(id, unit, f"{k.label}: {how}; {den.definition}", lambda a, b: a / b * factor, num, den)
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_analytics_measures.py -q`. Expected: `3 passed`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/analytics/measures.py engine/smartview_mcp/analytics/kpis.py engine/tests/analytics_util.py engine/tests/test_analytics_measures.py
PY ..\scripts\check_staged.py
git commit -m "analytics: flows, stock levels, the n+1 average and catalogue KPIs"
```

### Task 11: variance and materiality

**Files:**
- Create: `engine\smartview_mcp\analytics\variance.py`, `engine\smartview_mcp\analytics\materiality.py`
- Test: `engine\tests\test_analytics_variance.py` (new)

- [ ] **Step 1: Write the failing tests**

```python
"""analytics.variance and analytics.materiality: sign conventions, zero comparisons, the 10 % / NOK 500k floor /
NOK 5m rule and the always-lines."""
import dataclasses

import pytest

from smartview_mcp import catalog
from smartview_mcp.analytics.facts import Fact, missing
from smartview_mcp.analytics.materiality import assess, flags
from smartview_mcp.analytics.variance import variance

LINES = {x.key: x for x in catalog.load().pnl["FinMax"]}
M = catalog.load().materiality


def v(line, cur, cmp, unit="nok"):
    return variance(LINES[line], Fact("c", cur, unit, "cur"), Fact("p", cmp, unit, "cmp"), line)


def test_sign_convention_costs_negative_higher_is_better():
    assert v("rev", 120.0, 100.0).favourable is True
    cost = v("oth", -120.0, -100.0)                       # costs up by 20 (more negative)
    assert cost.abs.value == -20.0 and cost.favourable is False and cost.pct.value == pytest.approx(-0.2)
    assert v("fte", 12.0, 10.0, "fte").favourable is None  # FTE has no better direction
    lower = variance(dataclasses.replace(LINES["oth"], better="lower"), Fact("c", 5.0, "nok", ""),
                     Fact("p", 7.0, "nok", ""), "x")
    assert lower.favourable is True


def test_zero_comparison_has_no_pct_and_missing_parts_are_explicit():
    z = v("oth", -1000.0, 0.0)
    assert z.abs.value == -1000.0 and z.pct.value is None and "zero" in z.pct.reason
    gone = variance(LINES["rev"], Fact("c", 1.0, "nok", ""), missing("p", "nok", "", "not read"), "rev")
    assert gone.abs.value is None and gone.pct.value is None and gone.favourable is None


@pytest.mark.parametrize("cur, cmp, material", [
    (-1_200_000.0, -1_000_000.0, False),    # 20 % but only NOK 200k: under the 500k floor
    (-6_000_000.0, -5_000_000.0, True),     # 20 % and NOK 1m: the % rule
    (-106_000_000.0, -100_000_000.0, True),  # 6 % but NOK 6m: the NOK 5m rule
    (-104_000_000.0, -100_000_000.0, False),  # 4 % and NOK 4m: neither
    (-5_000_000.0, 0.0, True),              # zero comparison: only the NOK rule, and 5m is enough
    (-400_000.0, 0.0, False),
])
def test_materiality_rule_either_with_the_floor(cur, cmp, material):
    assert assess(v("oth", cur, cmp), M).material is material


def test_rule_both_needs_both_and_always_lines_are_commented():
    both = dataclasses.replace(M, rule="both")
    assert assess(v("oth", -6_000_000.0, -5_000_000.0), both).material is False   # 1m < 5m
    assert assess(v("oth", -60_000_000.0, -50_000_000.0), both).material is True
    small = assess(v("rev", 101.0, 100.0), M)
    assert small.material is False and small.always is True and small.comment is True
    assert assess(v("fte", 30.0, 10.0, "fte"), M).reason == "not an amount: no NOK rule"


def test_flags_rank_by_size():
    fs = flags([v("rev", 101.0, 100.0), v("oth", -9e6, -1e6), v("cogs", -3e6, -1e6)], M)
    assert [f.line for f in fs] == ["oth", "cogs", "rev"] and fs[0].size == 8e6
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_analytics_variance.py -q`
Expected: FAIL, `ModuleNotFoundError: No module named 'smartview_mcp.analytics.materiality'`

- [ ] **Step 3: Implement**

`engine\smartview_mcp\analytics\variance.py`:

```python
"""Variance = current - comparison, and % of |comparison|. FinMax signs: costs are negative, so for a line with
better = 'higher' (every P&L line) a positive variance is better. A zero comparison has no %."""
from __future__ import annotations

from dataclasses import dataclass

from ..catalog import Line
from .facts import Fact, combine, missing


@dataclass(frozen=True)
class Variance:
    line: str
    current: Fact
    comparison: Fact
    abs: Fact
    pct: Fact
    favourable: bool | None           # None: unknown, zero, or the line has no better direction (FTE)


def variance(line: Line, current: Fact, comparison: Fact, id: str) -> Variance:
    a = combine(f"{id}.var", current.unit, f"{current.id} - {comparison.id}", lambda x, y: x - y,
                current, comparison)
    if a.value is None:
        pct = missing(f"{id}.pct", "pct", "variance / |comparison|", a.reason or "no variance", a.inputs)
    elif comparison.value == 0:
        pct = missing(f"{id}.pct", "pct", "variance / |comparison|", "the comparison is zero: no %", a.inputs)
    else:
        pct = combine(f"{id}.pct", "pct", f"{a.id} / |{comparison.id}|", lambda v, c: v / abs(c), a, comparison)
    fav = None
    if a.value is not None and a.value != 0 and line.better != "none":
        fav = (a.value > 0) if line.better == "higher" else (a.value < 0)
    return Variance(line.key, current, comparison, a, pct, fav)
```

`engine\smartview_mcp\analytics\materiality.py`:

```python
"""Materiality, the catalogue rule (owner decisions 2026-10-25): a variance is material at pct of |comparison| when
it is also at least pct_floor_nok, and/or (rule either | both) at abs_nok. A zero comparison has no %, so only the
NOK amount counts. Lines in `always` are always commented. Only NOK amounts are judged (FTE is not an amount)."""
from __future__ import annotations

from dataclasses import dataclass

from ..catalog import Materiality
from .variance import Variance


@dataclass(frozen=True)
class Flag:
    line: str
    material: bool
    always: bool
    reason: str
    size: float                       # |variance| (0 when unknown): the ranking key

    @property
    def comment(self) -> bool:
        return self.material or self.always


def _nok(v: float) -> str:
    return f"NOK {v:,.0f}".replace(",", " ")


def assess(v: Variance, m: Materiality) -> Flag:
    always = v.line in m.always
    a, c = v.abs.value, v.comparison.value
    if a is None:
        return Flag(v.line, False, always, f"variance unknown ({v.abs.reason})", 0.0)
    if v.abs.unit != "nok":
        return Flag(v.line, False, always, "not an amount: no NOK rule", abs(a))
    size = abs(a)
    base = abs(c) if c is not None else 0.0
    pct_hit = base > 0 and size >= m.pct * base and size >= m.pct_floor_nok
    abs_hit = size >= m.abs_nok
    material = (pct_hit or abs_hit) if m.rule == "either" else (pct_hit and abs_hit)
    why = []
    if pct_hit:
        why.append(f"{size / base:.0%} of the comparison (counted from {_nok(m.pct_floor_nok)})")
    if abs_hit:
        why.append(f"at least {_nok(m.abs_nok)}")
    if not why:
        why.append("below the thresholds" + ("; the comparison is zero: no %" if base == 0 else ""))
    return Flag(v.line, material, always, "; ".join(why), size)


def flags(variances: list[Variance], m: Materiality) -> list[Flag]:
    """One flag per variance, largest |variance| first (a stable sort keeps catalogue order among equals)."""
    return sorted((assess(v, m) for v in variances), key=lambda f: -f.size)
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_analytics_variance.py -q`. Expected: `10 passed`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/analytics/variance.py engine/smartview_mcp/analytics/materiality.py engine/tests/test_analytics_variance.py
PY ..\scripts\check_staged.py
git commit -m "analytics: variances with the sign convention; the materiality rule with the NOK floor"
```

### Task 12: bridges and the year-end landing

**Files:**
- Create: `engine\smartview_mcp\analytics\bridge.py`, `engine\smartview_mcp\analytics\landing.py`
- Test: `engine\tests\test_analytics_bridge.py` (new)

- [ ] **Step 1: Write the failing tests**

```python
"""analytics.bridge and analytics.landing: bridges close exactly; the landing and its run-rate alternative."""
import pytest

from smartview_mcp import catalog
from smartview_mcp.analytics.bridge import by_line, by_unit, components
from smartview_mcp.analytics.facts import Fact, missing
from smartview_mcp.analytics.landing import landing
from smartview_mcp.analytics.variance import variance

CAT = catalog.load()


def f(id, value):
    return Fact(id, value, "nok", id)


def test_components_follow_the_catalogue_chain():
    fm = components(CAT.checks["FinMax"], CAT.pnl["FinMax"], "ebitda")
    assert fm == [(1, "rev"), (1, "cogs"), (1, "pers"), (1, "oth"), (1, "c73")]
    fp = components(CAT.checks["FinPlan"], CAT.pnl["FinPlan"], "ebitda")
    assert [k for _, k in fp] == ["rev", "cogs", "opex", "oth"]
    with pytest.raises(ValueError, match="no identity check defines 'rev'"):
        components(CAT.checks["FinMax"], CAT.pnl["FinMax"], "rev")


def test_line_bridge_closes_when_the_identities_hold():
    lines = {x.key: x for x in CAT.pnl["FinMax"]}
    cur = {"rev": 1000.0, "cogs": -400.0, "pers": -200.0, "oth": -100.0, "c73": -10.0}
    ly = {"rev": 900.0, "cogs": -350.0, "pers": -210.0, "oth": -80.0, "c73": 0.0}
    vs = {k: variance(lines[k], f(f"c.{k}", cur[k]), f(f"l.{k}", ly[k]), k) for k in cur}
    b = by_line(f("ly", sum(ly.values())), f("cur", sum(cur.values())), vs,
                components(CAT.checks["FinMax"], CAT.pnl["FinMax"], "ebitda"), {"rev": "Revenue"})
    assert [s.key for s in b.steps] == ["rev", "cogs", "pers", "oth", "c73"] and b.steps[0].label == "Revenue"
    assert b.residual() == 0.0 and b.closes()
    broken = by_line(f("ly", 0.0), f("cur", 999.0), vs, [(1, "rev")], {})
    assert not broken.closes()


def test_unit_bridge_top_n_plus_other_sums_exactly():
    units = [(f"u{i}", f"Unit {i}", f(f"c{i}", c), f(f"l{i}", p))
             for i, (c, p) in enumerate([(10.0, 0.0), (-50.0, 0.0), (5.0, 4.0), (30.0, 0.0), (1.0, 1.0)])]
    units.append(("ux", "Unknown", missing("cx", "nok", "", "not read"), f("lx", 1.0)))
    b = by_unit(f("start", 100.0), f("end", 103.0), units, top_n=2)
    assert [s.key for s in b.steps] == ["u1", "u3", "other"]
    assert b.steps[-1].fact.value == 3.0 - (-50.0 + 30.0) and b.residual() == 0.0
    assert by_unit(missing("s", "nok", "", "x"), f("end", 1.0), units).steps[-1].fact.value is None


def test_landing_forecast_and_seasonal_run_rate():
    ld = landing(f("ytd", 900.0), f("rem", 250.0), f("tgt", 1200.0), f("ly_ytd", 750.0), f("ly_fy", 1000.0), 9)
    assert ld.forecast.value == 1150.0 and ld.gap_forecast.value == -50.0
    assert ld.run_rate.value == pytest.approx(900.0 / 0.75) and ld.seasonal and "75.0%" in ld.method
    assert ld.gap_run_rate.value == pytest.approx(1200.0 - 1200.0)


def test_landing_falls_back_to_a_plain_run_rate():
    ld = landing(f("ytd", 900.0), f("rem", 250.0), f("tgt", 1200.0), f("ly_ytd", 750.0), f("ly_fy", 0.0), 9)
    assert ld.run_rate.value == 1200.0 and not ld.seasonal and "YTD / 9 x 12" in ld.method
    neg = landing(f("ytd", 900.0), f("rem", 0.0), f("tgt", 1.0), f("ly_ytd", 750.0), f("ly_fy", -10.0), 9)
    assert not neg.seasonal                                       # a negative share is no seasonality profile
    gone = landing(missing("ytd", "nok", "", "not read"), f("rem", 1.0), f("tgt", 1.0), f("a", 1.0), f("b", 2.0), 9)
    assert gone.forecast.value is None and gone.run_rate.value is None and gone.gap_forecast.value is None
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_analytics_bridge.py -q`
Expected: FAIL, `ModuleNotFoundError: No module named 'smartview_mcp.analytics.bridge'`

- [ ] **Step 3: Implement**

`engine\smartview_mcp\analytics\bridge.py`:

```python
"""Bridges: an EBITDA walk by P&L line (the leaf lines of the catalogue's identity chain) and by sub-unit (the top N
children by |variance| plus 'Other units' = the total variance less the top N, so that walk closes exactly)."""
from __future__ import annotations

import re
from dataclasses import dataclass

from .facts import Fact, combine, missing, total
from .variance import Variance


@dataclass(frozen=True)
class Step:
    key: str
    label: str
    fact: Fact


@dataclass(frozen=True)
class Bridge:
    start: Fact                       # the comparison total
    steps: tuple[Step, ...]
    end: Fact                         # the current total

    def residual(self) -> float | None:
        vals = [self.start.value, self.end.value, *(s.fact.value for s in self.steps)]
        if any(v is None for v in vals):
            return None
        return self.end.value - self.start.value - sum(s.fact.value for s in self.steps)

    def closes(self, tol: float = 0.5) -> bool:
        r = self.residual()
        return r is not None and abs(r) <= tol


def _terms(expr: str) -> list[tuple[int, str]]:
    return [(-1 if s == "-" else 1, k) for s, k in re.findall(r"([+-]?)\s*([a-z_]\w*)", expr)]


def components(checks, lines, total_key: str) -> list[tuple[int, str]]:
    """The leaf lines (with sign) that add up to total_key, from identity checks of the form 'a+b-c' (c = a + b),
    in catalogue line order. FinMax ebitda -> rev, cogs, pers, oth, c73."""
    defs = {}
    for c in checks:
        terms = _terms(c.expr)
        neg = [k for s, k in terms if s < 0]
        if len(neg) == 1:
            defs[neg[0]] = [(s, k) for s, k in terms if s > 0]
    if total_key not in defs:
        raise ValueError(f"no identity check defines '{total_key}'")

    def expand(key: str, sign: int, depth: int) -> list[tuple[int, str]]:
        if key not in defs or depth > len(defs):
            return [(sign, key)]
        return [x for s, k in defs[key] for x in expand(k, sign * s, depth + 1)]
    order = {x.key: i for i, x in enumerate(lines)}
    return sorted(expand(total_key, 1, 0), key=lambda t: order.get(t[1], len(order)))


def by_line(start: Fact, end: Fact, variances: dict[str, Variance], leaves: list[tuple[int, str]],
            labels: dict[str, str]) -> Bridge:
    """start -> end through the variance of each leaf line."""
    steps = []
    for sign, key in leaves:
        v = variances.get(key)
        f = v.abs if v is not None else missing(f"bridge.{key}", "nok", f"variance of {key}", "no variance")
        if sign < 0:
            f = combine(f.id + ".neg", f.unit, f"-({f.definition})", lambda x: -x, f)
        steps.append(Step(key, labels.get(key, key), f))
    return Bridge(start, tuple(steps), end)


def by_unit(start: Fact, end: Fact, units: list[tuple[str, str, Fact, Fact]], top_n: int = 5,
            id: str = "units") -> Bridge:
    """units: (member, label, current, comparison). The top_n units by |variance| (known variances only), then
    'Other units' = (end - start) - sum(top)."""
    vs = [(name, label, combine(f"{id}.{name}.var", "nok", f"{cur.id} - {cmp.id}", lambda a, b: a - b, cur, cmp))
          for name, label, cur, cmp in units]
    known = sorted((x for x in vs if x[2].value is not None), key=lambda x: -abs(x[2].value))
    top = known[:top_n]
    change = combine(f"{id}.total.var", "nok", f"{end.id} - {start.id}", lambda a, b: a - b, end, start)
    shown = total(f"{id}.top", [f for _, _, f in top], "nok", f"sum of the top {len(top)} units")
    other = combine(f"{id}.other", "nok", "total variance less the units shown", lambda a, b: a - b, change, shown)
    return Bridge(start, tuple(Step(n, lab, f) for n, lab, f in top) + (Step("other", "Other units", other),), end)
```

`engine\smartview_mcp\analytics\landing.py`:

```python
"""Year-end landing: YTD actual + the remaining months of the latest forecast; and a run-rate alternative, the YTD
scaled by last year's seasonality (last year's YTD share of its full year), or YTD / n x 12 when last year gives no
usable share. Both are shown with their gap to target."""
from __future__ import annotations

from dataclasses import dataclass

from .facts import Fact, combine


@dataclass(frozen=True)
class Landing:
    forecast: Fact                    # YTD actual + remaining forecast
    run_rate: Fact
    target: Fact
    gap_forecast: Fact                # forecast landing - target
    gap_run_rate: Fact
    seasonal: bool                    # True: the run rate uses last year's YTD share
    method: str


def landing(ytd_actual: Fact, remaining_forecast: Fact, target_fy: Fact, ly_ytd: Fact, ly_fy: Fact,
            months_done: int, id: str = "land") -> Landing:
    fc = combine(f"{id}.forecast", "nok", "YTD actual + remaining months of the latest forecast",
                 lambda a, b: a + b, ytd_actual, remaining_forecast)
    share = ly_ytd.value / ly_fy.value if ly_ytd.value is not None and ly_fy.value else None
    if share is not None and share > 0:
        rr = combine(f"{id}.run_rate", "nok", "YTD actual / last year's YTD share of its full year",
                     lambda a, y, f: a / (y / f), ytd_actual, ly_ytd, ly_fy)
        seasonal, method = True, f"seasonality: last year's YTD was {share:.1%} of its full year"
    else:
        rr = combine(f"{id}.run_rate", "nok", f"YTD actual / {months_done} x 12",
                     lambda a: a / months_done * 12, ytd_actual)
        seasonal = False
        method = f"plain run rate (YTD / {months_done} x 12): last year gives no usable seasonality share"
    gap_fc = combine(f"{id}.gap_forecast", "nok", "forecast landing - target", lambda a, b: a - b, fc, target_fy)
    gap_rr = combine(f"{id}.gap_run_rate", "nok", "run-rate landing - target", lambda a, b: a - b, rr, target_fy)
    return Landing(fc, rr, target_fy, gap_fc, gap_rr, seasonal, method)
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_analytics_bridge.py -q`. Expected: `5 passed`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/analytics/bridge.py engine/smartview_mcp/analytics/landing.py engine/tests/test_analytics_bridge.py
PY ..\scripts\check_staged.py
git commit -m "analytics: EBITDA bridges by line and sub-unit (exact closure); year-end landing"
```

### Task 13: reconciliation (pure)

**Files:**
- Create: `engine\smartview_mcp\analytics\reconcile.py`
- Test: `engine\tests\test_analytics_reconcile.py` (new)

- [ ] **Step 1: Write the failing tests**

```python
"""analytics.reconcile: identities per entity and month, months = YTD, areas vs the company total."""
from analytics_util import E, frame

from smartview_mcp import catalog
from smartview_mcp.analytics import reconcile

CHECKS = catalog.load().checks["FinMax"]
GOOD = {"rev": 100.0, "cogs": -40.0, "gm": 60.0, "pers": -20.0, "oth": -10.0, "opex": -30.0, "boi": 30.0,
        "c73": -5.0, "ebitda": 25.0}


def test_identities_pass_and_name_the_broken_cell():
    v = {(E, k, "Sep 26", "actual"): x for k, x in GOOD.items()}
    v.update({(E, k, "Aug 26", "actual"): x for k, x in {**GOOD, "gm": 70.0}.items()})
    n, problems = reconcile.identities(frame(values=v), CHECKS)
    assert n == 8 and len(problems) == 2                 # Aug 26: C30 + C40S = C48 and C48 + C70 = EBITDA_BOI
    assert problems[0].startswith(f"FinMax {E} Aug 26 actual: C30 + C40S = C48 is off by")
    partial = {(E, "rev", "Jul 26", "actual"): 1.0}
    assert reconcile.identities(frame(values=partial), CHECKS) == (0, [])


def test_months_vs_ytd():
    key = lambda p, b, a="ac.C30": ("FinMax", "Actual", "", E, a, p, b)  # noqa: E731
    periodic = {key("Jan 26", "Periodic"): 10.0, key("Feb 26", "Periodic"): None, key("Mar 26", "Periodic"): 5.0,
                key("Jan 26", "Periodic", "ac.C60"): 1.0}
    acc = {key("Mar 26", "Accumulated"): 15.0, key("Mar 26", "Accumulated", "ac.C60"): 9.0}
    n, problems = reconcile.months_vs_ytd(periodic, acc, ["Jan 26", "Feb 26", "Mar 26"])
    assert n == 1 and problems == []                      # ac.C60 lacks Feb and Mar: not checked
    acc[key("Mar 26", "Accumulated")] = 16.0
    assert "differs from the sum of Jan 26..Mar 26 by 1" in reconcile.months_vs_ytd(
        periodic, acc, ["Jan 26", "Feb 26", "Mar 26"])[1][0]


def test_areas_vs_total_is_information():
    v = {("A", "ebitda", "Sep 26", "actual"): 10.0, ("B", "ebitda", "Sep 26", "actual"): 5.0,
         ("un.SY", "ebitda", "Sep 26", "actual"): 16.0}
    info = reconcile.areas_vs_total(frame(values=v), ["A", "B"], "un.SY", "ebitda", "Sep 26", "actual")
    assert "less the 2 areas is 1 " in info
    assert reconcile.areas_vs_total(frame(values=v), ["A", "C"], "un.SY", "ebitda", "Sep 26", "actual") is None
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_analytics_reconcile.py -q`
Expected: FAIL, `ImportError: cannot import name 'reconcile' from 'smartview_mcp.analytics'`

- [ ] **Step 3: Implement**

```python
"""Reconciliation of warehouse values (pure): the catalogue's identity checks per entity and month, the sum of the
Periodic months against Accumulated, and the areas against the company total (information only: eliminations and
units outside the areas exist)."""
from __future__ import annotations

import re

from .frame import Frame

TOL = 1.0                                              # NOK


def _eval(expr: str, vals: dict[str, float]) -> float | None:
    out = 0.0
    for sign, key in re.findall(r"([+-]?)\s*([a-z_]\w*)", expr):
        if key not in vals:
            return None
        out += -vals[key] if sign == "-" else vals[key]
    return out


def _n(v: float) -> str:
    return f"{v:,.0f}".replace(",", " ")


def identities(frame: Frame, checks, tol: float = TOL) -> tuple[int, list[str]]:
    """(checks evaluated, problems) for every (entity, period, scenario) of the frame that has the check's lines."""
    groups: dict[tuple[str, str, str], dict[str, float]] = {}
    for (ent, line, period, scen), v in frame.values.items():
        groups.setdefault((ent, period, scen), {})[line] = 0.0 if v is None else v
    n, problems = 0, []
    for (ent, period, scen), vals in sorted(groups.items()):
        for c in checks:
            d = _eval(c.expr, vals)
            if d is None:
                continue
            n += 1
            if abs(d) >= tol:
                problems.append(f"{frame.cube} {ent} {period} {scen}: {c.label} is off by {_n(d)}")
    return n, problems


def months_vs_ytd(periodic: dict, accumulated: dict, months: list[str], tol: float = TOL) -> tuple[int, list[str]]:
    """Accumulated at months[-1] = the sum of the Periodic months, per (entity, account). Both are warehouse rows
    {key: value}; a pair without every month is not checked."""
    sums: dict[tuple[str, str], list[float]] = {}
    for k, v in periodic.items():
        if k[5] in months:
            sums.setdefault((k[3], k[4]), []).append(0.0 if v is None else v)
    n, problems = 0, []
    for k, v in sorted(accumulated.items()):
        parts = sums.get((k[3], k[4]))
        if k[5] != months[-1] or parts is None or len(parts) != len(months):
            continue
        n += 1
        d = (0.0 if v is None else v) - sum(parts)
        if abs(d) >= tol:
            problems.append(f"{k[0]} {k[3]} {k[4]}: Accumulated {months[-1]} differs from the sum of "
                            f"{months[0]}..{months[-1]} by {_n(d)}")
    return n, problems


def areas_vs_total(frame: Frame, areas: list[str], company: str, line: str, period: str,
                   scenario: str) -> str | None:
    """'The areas against the company total' as an information line, or None when a value is missing."""
    vals = [frame.get(a, line, period, scenario) for a in [*areas, company]]
    if any(v is None for v in vals):
        return None
    diff = vals[-1][0] - sum(v[0] for v in vals[:-1])
    return (f"{frame.cube} {line} {period} {scenario}: the company total less the {len(areas)} areas is {_n(diff)} "
            "(units outside the areas and eliminations)")
```

- [ ] **Step 4: Run the tests again, then the whole engine so far**

Run: `PY -m pytest tests/test_analytics_reconcile.py -q`. Expected: `3 passed`.
Run: `PY -m pytest -q` and `PY -m ruff check .`. Expected: all passed, `All checks passed!`

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/analytics/reconcile.py engine/tests/test_analytics_reconcile.py
PY ..\scripts\check_staged.py
git commit -m "analytics: identity checks, months = YTD, areas vs company total"
```

---

## Part D: extract and refresh

### Task 14: test helper, a fake provider that answers any grid (`EchoTransport`)

**Files:**
- Modify: `engine\tests\svxml_util.py` (append)
- Test: `engine\tests\test_svxml.py` (append one test of the helper)

- [ ] **Step 1: Write the helper's test** (append to `engine\tests\test_svxml.py`)

```python
def test_echo_transport_answers_a_retrieve_with_the_value_function(tmp_path):
    from svxml_util import EchoTransport, atlas_with_dimensions, pnl_value

    from smartview_mcp import svhttp
    from smartview_mcp.grid import Axis
    sv = svhttp.SvHttp(EchoTransport(pnl_value), home="https://plan1.example",
                       db_path=atlas_with_dimensions(tmp_path / "a.db"))
    lay, p, _ = sv.retrieve("FinMax", {"All Balance Types": "Accumulated"},
                            [Axis("All Entities", ["un.SY.2WHS"]), Axis("All Accounts", ["ac.C30", "ac.STA20Y"])],
                            [Axis("All Time Periods", ["Mar 26"])], False, False)
    assert p.rows == [["un.SY.2WHS", "ac.C30", 3000.0], ["un.SY.2WHS", "ac.STA20Y", 10.0]]
    assert sv.closed_month() == "Sep 26"
```

- [ ] **Step 2: Run it and check it fails**

Run: `PY -m pytest tests/test_svxml.py -q -k echo`
Expected: FAIL, `ImportError: cannot import name 'EchoTransport' from 'svxml_util'`

- [ ] **Step 3: Implement** (append to `engine\tests\svxml_util.py`)

```python
SOURCES = ('<res_GetProvisionedDataSources><Product id="HP"><Server name="FinMax" '
           'context="https://plan1.example/HyperionPlanning/SmartView" appName="FinMax"/></Product>'
           '<Product id="HP"><Server name="FinPlan_for_DataLoad" '
           'context="https://strat.example/HyperionPlanning/SmartView" appName="FinPlan"/></Product>'
           '</res_GetProvisionedDataSources>')


def _request_matrix(body):
    import html
    import re
    h, w = (int(x) for x in re.search(r'<slice rows="(\d+)" cols="(\d+)">', body).groups())
    vals = [html.unescape(x) for x in re.search(r"<vals[^>]*>(.*?)</vals>", body, re.S).group(1).split("|")]
    types = re.search(r"<types[^>]*>(.*?)</types>", body, re.S).group(1).split("|")
    return h, w, vals, types


class EchoTransport:
    """A provider that answers every retrieve by echoing the requested grid and filling each data cell with
    value(row_members, col_members, pov_members) - a float or None (blank). Labels come back as the member names.
    Figures are whatever the test's value function invents."""

    def __init__(self, value, closed="[CurrRepMonth].[Y26M09]"):
        self.value, self.closed, self.grids = value, closed, []

    def post(self, base, body):
        import re

        from smartview_mcp import svxml
        svxml.check_read_only(body)
        tag = re.match(r"\s*<req_(\w+)>", body).group(1)
        if tag == "GetProvisionedDataSources":
            return SOURCES
        if tag in ("ConnectToProvider", "OpenApplication"):
            return f"<res_{tag}><sID>s1</sID></res_{tag}>"
        if tag == "OpenCube":
            return "<res_OpenCube></res_OpenCube>"
        if tag == "EnumMembers":
            return f'<res_EnumMembers><mbrs enc="1">{self.closed}</mbrs></res_EnumMembers>'
        h, w, vals, types = _request_matrix(body)
        cells = [[None if vals[r * w + c] == "" else vals[r * w + c] for c in range(w)] for r in range(h)]
        data = [(r, c) for r in range(h) for c in range(w) if types[r * w + c] == "2"]
        r0, c0 = min(r for r, _ in data), min(c for _, c in data)          # 1 + column dims, row dims
        pov = tuple(v for v in cells[0][c0:] if v is not None)
        for r, c in data:
            cells[r][c] = self.value(tuple(cells[r][:c0]), tuple(cells[d][c] for d in range(1, r0)), pov)
        self.grids.append((base, len(data)))
        return answer(cells)


FINMAX_DIMS = ["All Accounts", "All Balance Types", "All Currencies", "All Entities", "All Scenarios",
               "All Time Periods"]
FINPLAN_DIMS = ["Account", "All Balance Types", "Entity", "Period", "Scenario", "Version", "Years"]


def atlas_with_dimensions(path):
    """A temporary atlas that knows the two cubes' dimensions (all svhttp needs to lay out a grid)."""
    from smartview_mcp import atlas
    db = atlas.connect(path)
    for conn, dims in (("FinMax", FINMAX_DIMS), ("FinPlan", FINPLAN_DIMS)):
        for d in dims:
            atlas.store(db, conn, d, [{"name": d, "label": d, "parent": None, "depth": 0, "order": 0}], 0.1)
    db.close()
    return path


# A consistent synthetic P&L for EchoTransport (invented figures; the catalogue identities hold).
FM = {"ac.C30": 1000.0, "ac.C40S": -400.0, "ac.C48": 600.0, "ac.C50S": -200.0, "ac.C60": -100.0, "ac.C70": -300.0,
      "ac.EBITDA_BOI": 300.0, "ac.C73": -10.0, "ac.EBITDA": 290.0, "ac.C75": -50.0, "ac.STA20Y": 10.0}
FP = {"Revenue": 1000.0, "COGS": -400.0, "Gross_Profit": 600.0, "OPEX": -300.0, "EBITDA before Other Items": 300.0,
      "Total Other Items": -10.0, "EBITDA": 290.0}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def pnl_value(rows, cols, pov):
    """A consistent synthetic P&L: the same every month; an Accumulated flow = month number x the month (the sum of
    the months) in both cubes; FTE is a level. The account may be on rows, columns or in the POV."""
    cells = (*rows, *cols, *pov)
    account = next(x for x in cells if x in FM or x in FP)
    v = FM.get(account, FP.get(account))
    month = next((x[:3] for x in cells if x[:3] in MONTHS and (len(x) == 3 or x[3] == " ")), None)
    if "Accumulated" in cells and account != "ac.STA20Y" and month:
        return v * (MONTHS.index(month) + 1)
    return v
```

- [ ] **Step 4: Run it again**

Run: `PY -m pytest tests/test_svxml.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/tests/svxml_util.py engine/tests/test_svxml.py
PY ..\scripts\check_staged.py
git commit -m "tests: EchoTransport, a fake provider that answers any grid with invented figures"
```

### Task 15: `extract.py`, grids, warehouse keys, the route and a fill

**Files:**
- Create: `engine\smartview_mcp\extract.py`
- Test: `engine\tests\test_extract.py` (new)

- [ ] **Step 1: Write the failing tests**

```python
"""extract.py: grids from the catalogue, warehouse keys, a fill over the svhttp fake transport, the COM fallback.
Every figure is invented by the value function below."""
import pytest
from svxml_util import EchoTransport, atlas_with_dimensions, pnl_value as value

from smartview_mcp import catalog, extract, ops, svhttp
from smartview_mcp.grid import Axis
from smartview_mcp.host import InputError
from smartview_mcp.warehouse import MemoryWarehouse

CAT = catalog.load()


@pytest.fixture
def http(tmp_path):
    return svhttp.SvHttp(EchoTransport(value), home="https://plan1.example",
                         db_path=atlas_with_dimensions(tmp_path / "a.db"))


def test_areas_skip_abolished_and_put_the_company_first():
    fm = [e.key for e in extract.areas(CAT, "FinMax")]
    assert fm[0] == "company" and "consumer" not in fm and "ig" not in fm and "wholesale" in fm
    assert [e.finplan for e in extract.areas(CAT, "FinPlan")][0] == "BANORWAY"


def test_base_grids_cover_36_months_levels_spot_and_finplan():
    grids = extract.base_grids(CAT, "Sep 26", lambda yy: "Forecast Q326" if yy == 26 else None)
    fm = [g for g in grids if g.conn == "FinMax"]
    assert len(fm) == 3 + 1 + 1                                     # 36 months in 12s, FTE levels, YTD spot
    assert fm[0].cols[0].members[0] == "Oct 23" and fm[2].cols[0].members[-1] == "Sep 26"
    levels = fm[3].cols[0].members
    assert levels[0] == "Dec 22" and fm[3].pov["All Balance Types"] == "Accumulated"
    assert fm[4].cols[0].members == ["Sep 26"] and "ac.STA20Y" not in fm[4].rows[1].members
    fp = [(g.pov["Scenario"], g.pov["Version"], g.pov["Years"]) for g in grids if g.conn == "FinPlan"]
    assert fp == [("Actual", "Working", "FY25"), ("Actual", "Working", "FY26"), ("Target", "Final", "FY25"),
                  ("Target", "Final", "FY26"), ("Target", "Final", "FY27"), ("Forecast Q326", "Final", "FY26")]
    only = extract.base_grids(CAT, "Sep 26", lambda yy: None, keys={"wholesale"})
    assert {tuple(g.rows[0].members) for g in only} == {("un.SY.2WHS",), ("2WHS Wholesale",)}


def test_wh_key_for_both_cubes():
    assert extract.wh_key("FinMax", {"All Scenarios": "Actual", "All Entities": "un.SY.2WHS", "All Accounts": "ac.C30",
                                     "All Time Periods": "Sep 26", "All Balance Types": "Periodic",
                                     "All Currencies": "NOK"}) == (
        "FinMax", "Actual", "", "un.SY.2WHS", "ac.C30", "Sep 26", "Periodic")
    assert extract.wh_key("FinPlan", {"Scenario": "Target", "Version": "Final", "Entity": "2WHS Wholesale",
                                      "Account": "EBITDA", "Period": "Sep", "Years": "FY26",
                                      "All Balance Types": "Periodic"}) == (
        "FinPlan", "Target", "Final", "2WHS Wholesale", "EBITDA", "Sep 26", "Periodic")


def test_a_fill_over_http_stores_every_cell_and_reconciles(http):
    store = MemoryWarehouse()
    grids = extract.base_grids(CAT, "Sep 26", lambda yy: "Forecast Q326" if yy == 26 else None)
    res = extract.fill(store, extract.route(http=http), CAT, grids, "base", "Sep 26", now=lambda: 100.0)
    assert res.ok and res.fill.result == "ok", res.fill.detail
    assert res.fill.cells == sum(g.cells for g in grids) and res.fill.detail["routes"] == ["http"]
    assert res.fill.detail["checks"] > 1000 and res.fill.detail["problems"] == []
    assert "less the 8 areas" in res.fill.detail["info"][0]
    got = store.get("FinPlan", ["2WHS Wholesale"], ["EBITDA"], ["Sep 26"], "Target", "Final", "Periodic")
    assert got == {("FinPlan", "Target", "Final", "2WHS Wholesale", "EBITDA", "Sep 26", "Periodic"): 290.0}
    assert store.last_fills(1)[0]["job"] == "base"


def test_a_broken_identity_is_a_warning_and_a_failed_grid_an_error(http):
    def bad(rows, cols, pov):
        return 999.0 if rows[1] == "ac.C48" and cols == ("Mar 26",) else value(rows, cols, pov)
    http.t.value = bad
    store = MemoryWarehouse()
    grids = extract.finmax_grids(CAT, ["Jan 26", "Feb 26", "Mar 26"], ["un.SY.2WHS"])
    res = extract.fill(store, extract.route(http=http), CAT, grids, "new_month", "Mar 26")
    assert res.fill.result == "warn" and "un.SY.2WHS Mar 26 actual: C30 + C40S = C48" in res.fill.detail["problems"][0]

    def flaky(conn, pov, rows, cols):
        if pov["All Balance Types"] == "Accumulated":
            raise RuntimeError("provider down")
        return extract.route(http=http)(conn, pov, rows, cols)
    res = extract.fill(store, flaky, CAT, grids, "new_month", "Mar 26")
    assert res.fill.result == "error" and "provider down" in res.fill.detail["errors"][0] and res.fill.cells == 30


def test_route_falls_back_to_com_but_not_for_input_errors(monkeypatch):
    class Down:
        def retrieve(self, *a):
            raise svhttp.SvHttpError("signed out")

    class Host:
        def on_host(self, fn, timeout=None):
            return fn()
    monkeypatch.setattr(ops, "query", lambda h, c, pov, rows, cols, sm, sz: ("lay", "parsed", 1.0))
    r = extract.route(http=Down(), host=lambda: Host())
    assert r("FinMax", {}, [Axis("All Accounts", ["ac.C30"])], [Axis("All Time Periods", ["Sep 26"])])[3] == "com"
    with pytest.raises(svhttp.SvHttpError):
        extract.route(http=Down())("FinMax", {}, [], [])

    class Wrong:
        def retrieve(self, *a):
            raise InputError("dimension used twice")
    with pytest.raises(InputError):
        extract.route(http=Wrong(), host=lambda: Host())("FinMax", {}, [], [])
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_extract.py -q`
Expected: FAIL, `ImportError: cannot import name 'extract' from 'smartview_mcp'`

- [ ] **Step 3: Implement**

```python
"""Extract jobs (Stage 2): the grids that fill the warehouse, built from the catalogue, read over Smart View's HTTP
protocol with the COM route as fallback (like tools_query._retrieve), stored as warehouse rows and reconciled.

Base fill: every business area (not abolished) plus the company total x the catalogue's lines x 36 months.
- FinMax Actual: flow lines Periodic (YTD is the sum of the months); stock lines (FTE) Accumulated = the month-end
  level, plus the opening December; one Accumulated spot column at the closed month for the months = YTD check.
- FinPlan: Actual/Working (FY-1, FY), Target/Final and the latest Forecast/Final (FY-1, FY, FY+1, where the forecast
  member exists), Periodic months. FinPlan periods are stored as 'Sep 26'."""
from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from typing import Callable

from . import config, ops
from .analytics import periods, reconcile
from .analytics.frame import from_rows
from .catalog import Catalog, Entity
from .grid import Axis, Layout, Parsed
from .host import InputError
from .warehouse import Fill, Row, WhKey

log = logging.getLogger(__name__)

Retrieve = Callable[[str, dict, list[Axis], list[Axis]], tuple[Layout, Parsed, float, str]]
FINMAX_POV = {"All Scenarios": "Actual", "All Currencies": "NOK"}
CHUNK = 12                                        # months per FinMax grid: about 1 100 cells, well under a second
FINMAX_SCENARIOS = {("Actual", ""): "actual"}


@dataclass(frozen=True)
class Grid:
    conn: str
    pov: dict
    rows: list[Axis]
    cols: list[Axis]
    what: str

    @property
    def cells(self) -> int:
        n = 1
        for a in self.rows + self.cols:
            n *= len(a.members)
        return n


def finplan_scenarios(forecast: str | None) -> dict[tuple[str, str], str]:
    """(FinPlan scenario member, version) -> catalogue scenario key, for analytics.frame.from_rows."""
    out = {("Actual", "Working"): "actual_finplan", ("Target", "Final"): "budget"}
    if forecast:
        out[(forecast, "Final")] = "forecast"
    return out


def areas(cat: Catalog, cube: str) -> list[Entity]:
    """The entities a base fill covers in `cube`: every area that is not abolished, the company total first."""
    attr = "finmax" if cube == "FinMax" else "finplan"
    ents = [e for e in cat.entities if not e.abolished and getattr(e, attr)]
    return sorted(ents, key=lambda e: e.key != "company")


def _norm(d: str) -> str:
    return d.removeprefix("All ").replace(" ", "").casefold()


def wh_key(conn: str, members: dict[str, str]) -> WhKey:
    """The warehouse key of one cell, from its full member set {dimension: member} (any dimension spelling)."""
    m = {_norm(d): v for d, v in members.items()}
    if conn == "FinPlan":
        return (conn, m["scenario"], m["version"], m["entity"], m["account"],
                periods.from_finplan(m["period"], m["years"]), m["balancetypes"])
    return (conn, m["scenarios"], "", m["entities"], m["accounts"], m["timeperiods"], m["balancetypes"])


def to_rows(lay: Layout, parsed: Parsed, conn: str, route: str, fetched_at: float, closed_label: str,
            atlas_build: str = "") -> tuple[list[Row], list[str]]:
    """Warehouse rows for every data cell of a retrieved grid (parsed with suppression off), and the problems:
    cells that are not data (#NoAccess, #Invalid) are not stored."""
    pov = dict(lay.pov)
    rows, problems = [], []
    for r in parsed.rows:
        combo = r[:lay.nr]
        for j, ctuple in enumerate(lay.col_tuples):
            v = r[lay.nr + j]
            members = {**pov, **{a.dimension: x for a, x in zip(lay.rows, combo)},
                       **{a.dimension: x for a, x in zip(lay.cols, ctuple)}}
            key = wh_key(conn, members)
            if isinstance(v, str):
                problems.append(f"{' / '.join(combo)} / {' / '.join(ctuple)}: {v}")
                continue
            rows.append(Row(key, None if v is None else float(v), route, fetched_at, closed_label, atlas_build))
    return rows, problems


def _chunks(xs: list, n: int) -> list[list]:
    return [xs[i:i + n] for i in range(0, len(xs), n)]


def finmax_grids(cat: Catalog, months: list[str], entities: list[str], spot: str | None = None) -> list[Grid]:
    lines = cat.pnl["FinMax"]
    flows = [x.account for x in lines if x.measure == "flow"]
    stocks = [x.account for x in lines if x.measure == "stock"]
    out = []
    for ch in _chunks(months, CHUNK):
        out.append(Grid("FinMax", {**FINMAX_POV, "All Balance Types": "Periodic"},
                        [Axis("All Entities", entities), Axis("All Accounts", flows)],
                        [Axis("All Time Periods", ch)], f"FinMax Actual {ch[0]}..{ch[-1]} Periodic"))
    if stocks:
        levels = sorted(set(months) | {periods.opening(months[0])}, key=periods.index)
        out.append(Grid("FinMax", {**FINMAX_POV, "All Balance Types": "Accumulated"},
                        [Axis("All Entities", entities), Axis("All Accounts", stocks)],
                        [Axis("All Time Periods", levels)], f"FinMax Actual levels {levels[0]}..{levels[-1]}"))
    if spot:
        out.append(Grid("FinMax", {**FINMAX_POV, "All Balance Types": "Accumulated"},
                        [Axis("All Entities", entities), Axis("All Accounts", flows)],
                        [Axis("All Time Periods", [spot])], f"FinMax Actual YTD {spot} (months = YTD check)"))
    return out


def finplan_grids(cat: Catalog, plans: list[tuple[str, str, int]], entities: list[str]) -> list[Grid]:
    """plans: (scenario member, version, two-digit year); one grid of entities x lines x the 12 months each."""
    accts = [x.account for x in cat.pnl["FinPlan"]]
    return [Grid("FinPlan", {"Scenario": s, "Version": v, "Years": f"FY{yy:02d}", "All Balance Types": "Periodic"},
                 [Axis("Entity", entities), Axis("Account", accts)], [Axis("Period", list(periods.MONTHS))],
                 f"FinPlan {s}/{v} FY{yy:02d}") for s, v, yy in plans]


def base_plans(closed: str, forecast_of: Callable[[int], str | None]) -> list[tuple[str, str, int]]:
    yy = periods.split_period(closed)[1]
    plans = [("Actual", "Working", y) for y in (yy - 1, yy)] + [("Target", "Final", y) for y in (yy - 1, yy, yy + 1)]
    return plans + [(fc, "Final", y) for y in (yy - 1, yy, yy + 1) if (fc := forecast_of(y))]


def base_grids(cat: Catalog, closed: str, forecast_of: Callable[[int], str | None],
               keys: set[str] | None = None) -> list[Grid]:
    """The base fill (all areas, or the area keys in `keys`)."""
    fm = [e.finmax for e in areas(cat, "FinMax") if keys is None or e.key in keys]
    fp = [e.finplan for e in areas(cat, "FinPlan") if keys is None or e.key in keys]
    out = finmax_grids(cat, periods.window(closed), fm, spot=closed) if fm else []
    return out + (finplan_grids(cat, base_plans(closed, forecast_of), fp) if fp else [])


def route(http=None, host: Callable[[], object] | None = None) -> Retrieve:
    """Smart View's HTTP protocol first (svhttp.shared() unless `http` is given); on any failure that is not an input
    error, the COM route through the hidden Excel when `host` (a function returning the ExcelHost) is given."""
    def retrieve(conn, pov, rows, cols):
        from . import svhttp
        try:
            lay, p, secs = (http or svhttp.shared()).retrieve(conn, pov, rows, cols, False, False)
            return lay, p, secs, "http"
        except InputError:
            raise
        except Exception as e:   # noqa: BLE001 - the HTTP route failed: Smart View in Excel, if we may
            if host is None:
                raise
            log.warning("warehouse extract: HTTP route failed (%s); using Smart View in Excel", e)
        h = host()
        lay, p, secs = h.on_host(lambda: ops.query(h, conn, pov, rows, cols, False, False),
                                 timeout=config.CALL_TIMEOUT_S)
        return lay, p, secs, "com"
    return retrieve


@dataclass
class FillResult:
    fill: Fill
    rows: list[Row] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.fill.result != "error"


def check_batch(cat: Catalog, rows: list[Row], closed: str) -> tuple[int, list[str], list[str]]:
    """(checks run, problems, information) over the rows of one fill: the identity checks per entity and month in
    both cubes, the Periodic months against the Accumulated spot column, and the areas against the company total."""
    data = {r.key: r.value for r in rows}
    n, problems, info = 0, [], []
    fm = from_rows("FinMax", cat.pnl["FinMax"], data, FINMAX_SCENARIOS)
    forecasts = {k[1] for k in data if k[0] == "FinPlan" and k[1].startswith("Forecast")}
    fp = from_rows("FinPlan", cat.pnl["FinPlan"], data,
                   {("Actual", "Working"): "actual_finplan", ("Target", "Final"): "budget",
                    **{(f, "Final"): f for f in forecasts}})
    for frame in (fm, fp):
        k, p = reconcile.identities(frame, cat.checks[frame.cube])
        n, problems = n + k, problems + p
    periodic = {k: v for k, v in data.items() if k[0] == "FinMax" and k[6] == "Periodic"}
    spot = {k: v for k, v in data.items() if k[0] == "FinMax" and k[6] == "Accumulated" and k[5] == closed
            and k[4] in {x.account for x in cat.pnl["FinMax"] if x.measure == "flow"}}
    k, p = reconcile.months_vs_ytd(periodic, spot, periods.ytd(closed))
    n, problems = n + k, problems + p
    company = next((e for e in cat.entities if e.key == "company"), None)
    others = [e.finmax for e in areas(cat, "FinMax") if e.key != "company"]
    if company and company.finmax:
        line = next((x.key for x in cat.pnl["FinMax"] if x.kind == "grand"), None)
        text = reconcile.areas_vs_total(fm, others, company.finmax, line, closed, "actual") if line else None
        if text:
            info.append(text)
    return n, problems, info


def fill(store, retrieve: Retrieve, cat: Catalog, grids: list[Grid], job: str, closed: str,
         closed_label: str | None = None, atlas_build: str = "", now: Callable[[], float] = time.time) -> FillResult:
    """Read every grid, store the rows, reconcile them and record the job in `fill`. A grid that fails is named in
    the result (the others are still stored); a failed job leaves the refresh state as it was, so it is retried."""
    t0, started = time.monotonic(), now()
    rows, errors, routes = [], [], set()
    for g in grids:
        try:
            lay, parsed, _secs, how = retrieve(g.conn, g.pov, g.rows, g.cols)
        except Exception as e:   # noqa: BLE001 - one failed grid must not lose the others
            errors.append(f"{g.what}: {type(e).__name__}: {str(e).splitlines()[0][:200] if str(e) else ''}")
            continue
        routes.add(how)
        got, bad = to_rows(lay, parsed, g.conn, how, now(), closed_label or closed, atlas_build)
        rows += got
        errors += [f"{g.what}: {b}" for b in bad[:5]] + [f"{g.what}: {w}" for w in parsed.pov_problems]
        errors += [f"{g.what}: unknown members {', '.join(parsed.unknown_members)}"] if parsed.unknown_members else []
    store.put(rows)
    n, problems, info = check_batch(cat, rows, closed) if rows else (0, [], [])
    result = "error" if errors else ("warn" if problems else "ok")
    detail = {"grids": len(grids), "routes": sorted(routes), "errors": errors[:20], "checks": n,
              "problems": problems[:20], "info": info}
    f = Fill(job, f"{len(grids)} grids", started, round(time.monotonic() - t0, 2), len(rows), result, detail)
    store.record_fill(f)
    return FillResult(f, rows)
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_extract.py -q`. Expected: `6 passed`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/extract.py engine/tests/test_extract.py
PY ..\scripts\check_staged.py
git commit -m "extract: base-fill grids from the catalogue, HTTP-first route, fill with reconciliation"
```

### Task 16: `refresh.py`, triggers, restatement areas, refresh and the live top-up

**Files:**
- Create: `engine\smartview_mcp\refresh.py`
- Test: `engine\tests\test_refresh.py` (new)

- [ ] **Step 1: Write the failing tests**

```python
"""refresh.py: which jobs run when (pure), restatement areas from atlas snapshots, a refresh and a top-up over the
svhttp fake transport. Every figure is invented."""
import sqlite3

import pytest
from svxml_util import EchoTransport, atlas_with_dimensions, pnl_value as value

from smartview_mcp import atlas, atlas_history, catalog, extract, refresh, svhttp
from smartview_mcp.warehouse import OPEN, MemoryWarehouse

CAT = catalog.load()
FC = {26: "Forecast Q326"}.get


def kinds(jobs):
    return [j.kind for j in jobs]


def test_first_call_is_the_base_fill():
    jobs = refresh.plan_refresh(CAT, {}, "Sep 26", 7, "Forecast Q326", set(), FC)
    assert kinds(jobs) == ["base"] and jobs[0].state == {"closed_month": "Sep 26", "entities_snapshot": "7",
                                                          "forecast": "Forecast Q326"}


def test_nothing_changed_means_no_job():
    state = {"closed_month": "Sep 26", "entities_snapshot": "7", "forecast": "Forecast Q326"}
    assert refresh.plan_refresh(CAT, state, "Sep 26", 7, "Forecast Q326", set(), FC) == []


def test_a_new_closed_month_rereads_the_two_before():
    state = {"closed_month": "Sep 26", "entities_snapshot": "7", "forecast": "Forecast Q326"}
    jobs = refresh.plan_refresh(CAT, state, "Oct 26", 7, "Forecast Q326", set(), FC)
    assert kinds(jobs) == ["new_month"] and jobs[0].state == {"closed_month": "Oct 26"}
    fm = [g for g in jobs[0].grids if g.conn == "FinMax"]
    assert fm[0].cols[0].members == ["Aug 26", "Sep 26", "Oct 26"] and fm[-1].cols[0].members == ["Oct 26"]
    two = refresh.plan_refresh(CAT, {**state, "closed_month": "Aug 26"}, "Oct 26", 7, "Forecast Q326", set(), FC)
    assert two[0].grids[0].cols[0].members == ["Aug 26", "Sep 26", "Oct 26"]


def test_restatement_and_forecast_change():
    state = {"closed_month": "Sep 26", "entities_snapshot": "7", "forecast": "Forecast Q226"}
    jobs = refresh.plan_refresh(CAT, state, "Sep 26", 8, "Forecast Q326", {"wholesale"}, FC)
    assert kinds(jobs) == ["restatement", "forecast"]
    ents = {m for g in jobs[0].grids for m in g.rows[0].members}
    assert ents == {"un.SY", "un.SY.2WHS", "BANORWAY", "2WHS Wholesale"}
    assert [g.pov["Scenario"] for g in jobs[1].grids] == ["Forecast Q326"]
    quiet = refresh.plan_refresh(CAT, {**state, "forecast": "Forecast Q326"}, "Sep 26", 8, "Forecast Q326", set(), FC)
    assert kinds(quiet) == ["restatement"] and quiet[0].grids == () and quiet[0].state == {"entities_snapshot": "8"}


@pytest.fixture
def snap_db(tmp_path):
    db = sqlite3.connect(tmp_path / "atlas.db")
    db.executescript(atlas.SCHEMA)

    def load(rows):
        db.execute("DELETE FROM member")
        db.executemany("INSERT INTO member VALUES ('FinMax','All Entities',?,?,?,0,0)", rows)
        db.commit()
        return atlas_history.take_snapshot(db, "FinMax", "All Entities")
    base = [("un.SY", "SY", ""), ("un.SY.2WHS", "WHS", "un.SY"), ("un.SY.3WHS_W", "W", "un.SY.2WHS"),
            ("un.SY.2BUS", "BUS", "un.SY"), ("un.SY.K1", "K1", "un.SY.3WHS_W")]
    s1 = load(base)
    s2 = load(base[:-1] + [("un.SY.K1", "K1", "un.SY.2BUS")])          # K1 moved from Wholesale to Business
    return db, s1, s2


def test_touched_areas_follow_moves_up_the_atlas(snap_db):
    db, s1, s2 = snap_db
    assert refresh.touched_areas(CAT, db, str(s1), s2) == {"wholesale", "business"}
    assert refresh.touched_areas(CAT, db, None, s2) == set() and refresh.touched_areas(CAT, db, str(s2), s2) == set()


def test_sub_units_skip_single_child_levels():
    tree = {"un.SY.2WHS": ["un.SY.3WHS_W"], "un.SY.3WHS_W": ["a", "b", "c"], "un.SY.2BUS": ["x", "y"]}
    kids = lambda m: tree.get(m, [])  # noqa: E731
    assert refresh.sub_units(kids, "un.SY.2WHS") == ["a", "b", "c"]
    assert refresh.sub_units(kids, "un.SY.2BUS", limit=1) == ["x"] and refresh.sub_units(kids, "leaf") == []


@pytest.fixture
def retrieve(tmp_path):
    http = svhttp.SvHttp(EchoTransport(value), home="https://plan1.example",
                         db_path=atlas_with_dimensions(tmp_path / "a.db"))
    return extract.route(http=http)


def test_refresh_runs_the_base_fill_once_then_nothing(retrieve, snap_db):
    db, _, _ = snap_db
    store = MemoryWarehouse()
    first = refresh.refresh(store, retrieve, CAT, "Sep 26", db, FC)
    assert [r.fill.job for r in first] == ["base"] and first[0].ok
    assert store.state()["closed_month"] == "Sep 26" and store.state()["forecast"] == "Forecast Q326"
    assert refresh.refresh(store, retrieve, CAT, "Sep 26", db, FC) == []


def test_a_failed_job_keeps_the_old_state(snap_db):
    db, _, _ = snap_db
    store = MemoryWarehouse()

    def down(*a):
        raise RuntimeError("no network")
    res = refresh.refresh(store, down, CAT, "Sep 26", db, FC)
    assert res[0].fill.result == "error" and store.state() == {}


def test_top_up_reads_forecast_and_open_months_live_and_units_as_history(retrieve):
    store = MemoryWarehouse()
    whs = next(e for e in CAT.entities if e.key == "wholesale")
    res = refresh.top_up(store, retrieve, CAT, whs, "Oct 26", "Sep 26", "Forecast Q326", ["u1", "u2"])
    assert [r.ok for r in res] == [True, True]
    fc = ("FinPlan", "Forecast Q326", "Final", "2WHS Wholesale", "EBITDA", "Dec 26", "Periodic")
    assert store.rows[fc].closed_month == OPEN
    assert store.rows[("FinMax", "Actual", "", "un.SY.2WHS", "ac.C30", "Oct 26", "Periodic")].closed_month == OPEN
    unit = store.rows[("FinMax", "Actual", "", "u1", "ac.EBITDA", "Oct 25", "Accumulated")]
    assert unit.closed_month == "Sep 26" and unit.value == 290.0 * 10
    assert refresh.top_up(store, retrieve, CAT, whs, "Sep 26", "Sep 26", None, []) == []
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_refresh.py -q`
Expected: FAIL, `ImportError: cannot import name 'refresh' from 'smartview_mcp'`

- [ ] **Step 3: Implement**

```python
"""Warehouse refresh triggers (Stage 2), checked at the start of every fin_* call. The check itself makes no EPM call
when nothing changed: the closed month is cached per day (templates.closed_month), the atlas snapshot and the
latest forecast name are local.
- No base fill yet: the base fill.
- A new closed month: fill the new month(s) and re-read the two before (late postings), with the YTD spot column.
- The atlas shows changed members under an area (atlas_history, i.e. a reorg or restatement): re-read that area's
  36 months (and the company total).
- The latest forecast name changed: re-read FinPlan's forecast.
A job that fails leaves its state key as it was, so the next call tries again; the answer uses the stale data and
says so. top_up() reads what a question needs live: the latest forecast, an open month and the sub-units."""
from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Callable

from . import atlas, atlas_history
from .analytics import periods
from .catalog import Catalog, Entity
from .extract import FillResult, Grid, areas, base_grids, base_plans, fill, finmax_grids, finplan_grids
from .grid import Axis
from .warehouse import OPEN

CLOSED, SNAPSHOT, FORECAST = "closed_month", "entities_snapshot", "forecast"


@dataclass(frozen=True)
class Job:
    kind: str                         # base | new_month | restatement | forecast
    grids: tuple[Grid, ...]
    why: str
    state: dict                       # state keys to set when the job succeeds


def plan_refresh(cat: Catalog, state: dict[str, str], closed: str, snapshot: int | None, forecast: str | None,
                 touched: set[str], forecast_of: Callable[[int], str | None]) -> list[Job]:
    """The jobs to run (pure). state: the warehouse's state keys; snapshot: the newest FinMax entity snapshot id;
    forecast: the latest forecast name for the closed month's year; touched: area keys with changed members."""
    snap = "" if snapshot is None else str(snapshot)
    if not state.get(CLOSED):
        return [Job("base", tuple(base_grids(cat, closed, forecast_of)), "no base fill yet",
                    {CLOSED: closed, SNAPSHOT: snap, FORECAST: forecast or ""})]
    jobs = []
    last = state[CLOSED]
    if last != closed:
        start = last if periods.after(closed, last) else closed
        months = sorted(set(periods.between(start, closed)) | {periods.shift(closed, -2), periods.shift(closed, -1),
                                                               closed}, key=periods.index)
        fm = [e.finmax for e in areas(cat, "FinMax")]
        fp = [e.finplan for e in areas(cat, "FinPlan")]
        years = sorted({periods.split_period(m)[1] for m in months})
        jobs.append(Job("new_month", tuple(finmax_grids(cat, months, fm, spot=closed)
                                           + finplan_grids(cat, [("Actual", "Working", y) for y in years], fp)),
                        f"closed month {last} -> {closed}", {CLOSED: closed}))
    if snap != state.get(SNAPSHOT, "") and touched:
        jobs.append(Job("restatement", tuple(base_grids(cat, closed, forecast_of, keys=touched | {"company"})),
                        f"the atlas changed under {', '.join(sorted(touched))}", {SNAPSHOT: snap}))
    elif snap != state.get(SNAPSHOT, ""):
        jobs.append(Job("restatement", (), "the atlas changed outside the areas", {SNAPSHOT: snap}))
    if (forecast or "") != state.get(FORECAST, ""):
        yy = periods.split_period(closed)[1]
        fp = [e.finplan for e in areas(cat, "FinPlan")]
        plans = [p for p in base_plans(closed, forecast_of) if p[0] == forecast and p[2] == yy]
        jobs.append(Job("forecast", tuple(finplan_grids(cat, plans, fp)),
                        f"latest forecast {state.get(FORECAST) or 'none'} -> {forecast or 'none'}",
                        {FORECAST: forecast or ""}))
    return jobs


def touched_areas(cat: Catalog, db, old: str | None, new: int | None) -> set[str]:
    """Area keys whose members changed between two FinMax entity snapshots (added, removed, moved or relabelled
    members, and the parents they moved from and to, traced up the atlas). Empty when either snapshot is unknown."""
    if not old or new is None or str(new) == old:
        return set()
    d = atlas_history.diff_snapshots(db, int(old), new)
    names: set[str] = set()
    for kind in ("added", "removed"):
        for m in d.get(kind, []):
            names |= {m["name"], *m.get("parents", [])}
    for m in d.get("moved", []):
        names |= {m["name"], *m.get("from", []), *m.get("to", [])}
    names |= {m["name"] for m in d.get("relabelled", [])}
    by_member = {e.finmax: e.key for e in areas(cat, "FinMax") if e.key != "company"}
    out = set()
    for n in names:
        for x in [n, *atlas.ancestors(db, "FinMax", n)]:
            if x in by_member:
                out.add(by_member[x])
    return out


def refresh(store, retrieve, cat: Catalog, closed: str, db, forecast_of: Callable[[int], str | None],
            now: Callable[[], float] = time.time) -> list[FillResult]:
    """Run what plan_refresh asks for; each successful job updates its state keys."""
    state = store.state()
    snapshot = atlas_history.latest_snapshot(db, "FinMax", "All Entities")
    forecast = forecast_of(periods.split_period(closed)[1])
    touched = touched_areas(cat, db, state.get(SNAPSHOT), snapshot)
    out = []
    for job in plan_refresh(cat, state, closed, snapshot, forecast, touched, forecast_of):
        if job.grids:
            res = fill(store, retrieve, cat, list(job.grids), job.kind, closed,
                       atlas_build="" if snapshot is None else str(snapshot), now=now)
            out.append(res)
            if not res.ok:
                continue
        store.set_state(**job.state)
    return out


def sub_units(children: Callable[[str], list[str]], entity: str, limit: int = 40) -> list[str]:
    """The sub-units of an area for its bridge: its children, going down through single-child levels (2WHS has
    one child, 3WHS_W, whose children are the units)."""
    kids = children(entity)
    seen = {entity}
    while len(kids) == 1 and kids[0] not in seen:
        seen.add(kids[0])
        kids = children(kids[0])
    return kids[:limit]


def top_up(store, retrieve, cat: Catalog, entity: Entity, period: str, closed: str, forecast: str | None,
           units: list[str], ebitda: str = "ac.EBITDA", now: Callable[[], float] = time.time) -> list[FillResult]:
    """Read live what one question needs and store it: the area's latest forecast and, when `period` is after the
    closed month, its open months (both with closed_month 'open'); and the sub-units' YTD EBITDA now and last year
    (closed history, stored as it comes)."""
    yy = periods.split_period(period)[1]
    live: list[Grid] = []
    if forecast and entity.finplan:
        live += finplan_grids(cat, [(forecast, "Final", yy)], [entity.finplan])
    if periods.after(period, closed):
        if entity.finmax:
            live += finmax_grids(cat, periods.between(closed, period), [entity.finmax])
        if entity.finplan:
            live += finplan_grids(cat, [("Actual", "Working", yy)], [entity.finplan])
    out = []
    if live:
        out.append(fill(store, retrieve, cat, live, "top_up", closed, closed_label=OPEN, now=now))
    if units:
        g = Grid("FinMax", {"All Scenarios": "Actual", "All Currencies": "NOK", "All Accounts": ebitda,
                            "All Balance Types": "Accumulated"}, [Axis("All Entities", units)],
                 [Axis("All Time Periods", [period, periods.ly(period)])], "FinMax sub-unit EBITDA YTD")
        out.append(fill(store, retrieve, cat, [g], "top_up", closed, now=now))
    return out
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_refresh.py -q`. Expected: `9 passed`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/refresh.py engine/tests/test_refresh.py
PY ..\scripts\check_staged.py
git commit -m "refresh: new month, restatement and forecast triggers; live top-up of forecast, open month, units"
```

---

## Part E: the first slice `fin_pnl`

### Task 17: `fin_pnl.compute`

**Files:**
- Create: `engine\smartview_mcp\fin_pnl.py`
- Create: `engine\tests\fin_util.py` (test helper: a synthetic Wholesale warehouse)
- Test: `engine\tests\test_fin_pnl.py` (new)

- [ ] **Step 1: Write the helper and the failing tests**

`engine\tests\fin_util.py`:

```python
"""A synthetic Wholesale warehouse for the fin_pnl tests: a base fill and a top-up over EchoTransport (svhttp's fake
provider). Every figure is invented: the pnl_value P&L, x1.1 in 2026, Target x0.9, the forecast x1.05."""
from svxml_util import EchoTransport, atlas_with_dimensions, pnl_value

from smartview_mcp import catalog, extract, refresh, svhttp
from smartview_mcp.fin_pnl import Scope
from smartview_mcp.warehouse import MemoryWarehouse

CAT = catalog.load()
WHS = next(e for e in CAT.entities if e.key == "wholesale")
FC = "Forecast Q326"
UNITS = (("un.SY.U1", "Unit one"), ("un.SY.U2", "Unit two"))


def value(rows, cols, pov):
    v = pnl_value(rows, cols, pov)
    year = 1.1 if any(c.endswith(" 26") for c in cols) or "FY26" in pov or "FY26" in cols else 1.0
    plan = 0.9 if "Target" in pov or "Target" in cols else (1.05 if FC in pov or FC in cols else 1.0)
    return v * year * plan


def http(tmp_path):
    return svhttp.SvHttp(EchoTransport(value), home="https://plan1.example",
                         db_path=atlas_with_dimensions(tmp_path / "a.db"))


def filled(tmp_path, period="Sep 26", closed="Sep 26"):
    """(store, retrieve, scope) after a base fill and the question's top-up."""
    store, retrieve = MemoryWarehouse(), extract.route(http=http(tmp_path))
    extract.fill(store, retrieve, CAT, extract.base_grids(CAT, closed, {26: FC}.get), "base", closed)
    refresh.top_up(store, retrieve, CAT, WHS, period, closed, FC, [u for u, _ in UNITS])
    return store, retrieve, Scope(WHS, period, closed, FC, UNITS)
```

`engine\tests\test_fin_pnl.py`:

```python
"""fin_pnl.compute on a synthetic warehouse: variances, flags, bridges, KPIs and landing (invented figures)."""
import dataclasses

import pytest
from fin_util import CAT, filled

from smartview_mcp import fin_pnl

YR = 1.1


def test_finmax_month_and_ytd_vs_last_year(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)
    e = res.fm["ebitda"]
    assert e["m"].current.value == pytest.approx(290 * YR) and e["m"].comparison.value == 290
    assert e["ytd"].current.value == pytest.approx(290 * YR * 9) and e["ytd"].abs.value == pytest.approx(261)
    assert e["ytd"].favourable is True and len(e["ytd"].current.inputs) == 9
    fte = res.fm["fte"]
    assert fte["m"].current.value == pytest.approx(11.0) and fte["ytd"].current.value == pytest.approx(
        (10 + 11 * 9) / 10)                                         # Dec 25 = 10 (2025), Jan..Sep 26 = 11


def test_finplan_budget_forecast_and_landing(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)
    b = res.budget["ebitda"]["ytd"]
    assert b.current.value == pytest.approx(290 * YR * 9) and b.comparison.value == pytest.approx(290 * YR * 9 * 0.9)
    assert res.forecast["ebitda"]["ytd"].comparison.value == pytest.approx(290 * YR * 9 * 1.05)
    ld = res.landing
    assert ld.forecast.value == pytest.approx(290 * YR * 9 + 290 * YR * 1.05 * 3)
    assert ld.seasonal and ld.run_rate.value == pytest.approx(290 * YR * 9 / 0.75)
    assert ld.target.value == pytest.approx(290 * YR * 0.9 * 12)


def test_flags_bridges_and_kpis(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)
    assert not any(f.material for f in res.flags["ly"].values())          # small invented figures
    assert {k for c, k, f in res.ranking() if c == "ly"} == {"rev", "pers", "ebitda"}   # the always-lines
    assert res.line_bridge.closes() and [s.key for s in res.line_bridge.steps] == ["rev", "cogs", "pers", "oth", "c73"]
    ub = res.unit_bridge
    assert ub.closes() and [s.key for s in ub.steps] == ["un.SY.U1", "un.SY.U2", "other"]
    assert ub.steps[-1].fact.value == pytest.approx(261 - 2 * 261)
    m = res.kpis["ebitda_margin"]["ytd.cur"]
    assert m.value == pytest.approx(290 / 1000)
    c = res.kpis["opex_per_fte"]["ytd.cur"]
    assert c.value == pytest.approx(-300 * YR * 9 / ((10 + 11 * 9) / 10) * 12 / 9)
    assert "m.cur" not in res.kpis["opex_per_fte"]                        # per-FTE KPIs only for the YTD


def test_compare_limits_the_flags_and_a_missing_month_is_explicit(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, dataclasses.replace(scope, compare=("ly",)), store)
    assert set(res.flags) == {"ly"}
    later = fin_pnl.compute(CAT, dataclasses.replace(scope, period="Nov 26"), store)
    assert later.fm["ebitda"]["ytd"].current.value is None and "Oct 26" in later.fm["ebitda"]["ytd"].current.reason


def test_an_open_month_is_read_from_the_top_up(tmp_path):
    store, _, scope = filled(tmp_path, period="Oct 26")
    res = fin_pnl.compute(CAT, scope, store)
    assert res.fm["ebitda"]["m"].current.value == pytest.approx(290 * YR)
    assert res.freshness["FinMax"]["open_rows"] > 0
    assert "open month read live" in fin_pnl.freshness_text(res, "warehouse")
    s = fin_pnl.summary(res, "warehouse", [])
    assert s["area"] == "Wholesale" and s["landing"]["method"].startswith("seasonality")
    assert all(not f["id"].endswith(".pct") for f in s["facts"])
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_fin_pnl.py -q`
Expected: FAIL, `ModuleNotFoundError: No module named 'smartview_mcp.fin_pnl'`

- [ ] **Step 3: Implement**

Fact ids are a contract with `fin_spec.py` (Task 18) and the live checks (Tasks 22–23); keep them exactly:
- FinMax: `fm.<line>.m.cur|m.ly|ytd.cur|ytd.ly`, with `.var` and `.pct` on the variances; FTE levels are
  `fm.<line>.lvl.<month>`.
- FinPlan: `fp.<line>.actual.m|ytd|ly_ytd|ly_fy`, `fp.<line>.budget.m|ytd|fy`, `fp.<line>.forecast.m|ytd|fy`, and
  `fp.<line>.forecast.remaining`.
- Sub-units: `unit.<member>.cur|ly`.
- KPIs: `kpi.<key>.ytd.cur|ytd.ly|m.cur|m.ly`.
- Landing: `land.forecast|run_rate|gap_forecast|gap_run_rate`.
- Bridge steps: `units.*`.

```python
"""fin_pnl, the first slice (Stage 2+3): a business area's P&L for a month and YTD against last year (FinMax) and
budget and the latest forecast (FinPlan: actual = version Working, plan = Final), with materiality flags, an EBITDA
bridge by line and by sub-unit, KPIs and the year-end landing - decided by the calc engine from the warehouse, in
milliseconds and without Excel. fin_spec.py turns the Result into the sv_ws_build request (live source sheets +
formulas); the engine's numbers then reconcile the sheets. templates.division_pnl is this slice's predecessor."""
from __future__ import annotations

from dataclasses import dataclass, field

from .analytics import periods
from .analytics.bridge import Bridge, by_line, by_unit, components
from .analytics.facts import Fact, missing
from .analytics.frame import Frame, from_rows
from .analytics.kpis import kpi
from .analytics.landing import Landing, landing
from .analytics.materiality import Flag, assess
from .analytics.measures import amount, average_level, flow, level
from .analytics.variance import Variance, variance
from .catalog import Catalog, Entity
from .extract import FINMAX_SCENARIOS, finplan_scenarios

COMPARISONS = ("ly", "budget", "forecast")


@dataclass(frozen=True)
class Scope:
    entity: Entity
    period: str                                     # 'Sep 26'
    closed: str                                     # the last closed month
    forecast: str | None                            # the latest 'Forecast Q<q><yy>' for the period's year
    units: tuple[tuple[str, str], ...] = ()         # (member, label) sub-units for the bridge
    compare: tuple[str, ...] = COMPARISONS


@dataclass
class Result:
    scope: Scope
    ly: str
    ytd: list[str]
    ly_ytd: list[str]
    facts: dict[str, Fact] = field(default_factory=dict)
    fm: dict[str, dict[str, Variance]] = field(default_factory=dict)       # line -> {"m", "ytd"} vs last year
    budget: dict[str, dict[str, Variance]] = field(default_factory=dict)   # FinPlan line -> {"m", "ytd"}
    forecast: dict[str, dict[str, Variance]] = field(default_factory=dict)
    flags: dict[str, dict[str, Flag]] = field(default_factory=dict)        # comparison -> line -> Flag (YTD)
    line_bridge: Bridge | None = None
    unit_bridge: Bridge | None = None
    kpis: dict[str, dict[str, Fact]] = field(default_factory=dict)         # kpi -> {"m.cur", "ytd.ly", ...}
    landing: Landing | None = None
    freshness: dict = field(default_factory=dict)

    def ranking(self) -> list[tuple[str, str, Flag]]:
        """(comparison, line, flag) for every line to comment (material or always), largest |variance| first."""
        out = [(c, k, f) for c, fs in self.flags.items() for k, f in fs.items() if f.comment]
        return sorted(out, key=lambda x: -x[2].size)


def grand(lines) -> str:
    return next(x.key for x in lines if x.kind == "grand")


def _keep(res: Result, *facts: Fact) -> None:
    for f in facts:
        res.facts[f.id] = f


def _cell(rows: dict, key: tuple, id: str, definition: str) -> Fact:
    if key not in rows:
        return missing(id, "nok", definition, f"no warehouse value for {' / '.join(key[3:6])}")
    v = rows[key]
    return Fact(id, 0.0 if v is None else float(v), "nok", definition, (key,))


def _finmax(cat: Catalog, s: Scope, store, res: Result) -> Frame:
    e, p, ly = s.entity.finmax, s.period, res.ly
    lines = cat.pnl["FinMax"]
    closed_only = not periods.after(p, s.closed)
    months = res.ytd + res.ly_ytd
    flows = [x.account for x in lines if x.measure == "flow"]
    stocks = [x.account for x in lines if x.measure == "stock"]
    rows = store.get("FinMax", [e], flows, months, "Actual", "", "Periodic", closed_only=closed_only)
    if stocks:
        rows |= store.get("FinMax", [e], stocks, months + [periods.opening(p), periods.opening(ly)], "Actual", "",
                          "Accumulated", closed_only=closed_only)
    fm = from_rows("FinMax", lines, rows, FINMAX_SCENARIOS)
    for x in lines:
        b = f"fm.{x.key}"
        m_cur = amount(fm, e, x.key, "actual", [p], f"{b}.m.cur")
        m_ly = amount(fm, e, x.key, "actual", [ly], f"{b}.m.ly")
        if x.measure == "stock":
            y_cur = average_level(fm, e, x.key, "actual", res.ytd, f"{b}.ytd.cur")
            y_ly = average_level(fm, e, x.key, "actual", res.ly_ytd, f"{b}.ytd.ly")
            for m in [periods.opening(p), *res.ytd, periods.opening(ly), *res.ly_ytd]:
                _keep(res, level(fm, e, x.key, "actual", m, f"{b}.lvl.{m}"))
        else:
            y_cur = flow(fm, e, x.key, "actual", res.ytd, f"{b}.ytd.cur")
            y_ly = flow(fm, e, x.key, "actual", res.ly_ytd, f"{b}.ytd.ly")
        res.fm[x.key] = {"m": variance(x, m_cur, m_ly, f"{b}.m"), "ytd": variance(x, y_cur, y_ly, f"{b}.ytd")}
        for v in res.fm[x.key].values():
            _keep(res, v.current, v.comparison, v.abs, v.pct)
    res.flags["ly"] = {k: assess(v["ytd"], cat.materiality) for k, v in res.fm.items()}
    return fm


def _finplan(cat: Catalog, s: Scope, store, res: Result) -> Frame | None:
    ef = s.entity.finplan
    if not ef:
        return None
    lines = cat.pnl["FinPlan"]
    accts = [x.account for x in lines]
    yy = periods.split_period(s.period)[1]
    year, last = periods.year(yy), periods.year(yy - 1)
    closed_only = not periods.after(s.period, s.closed)
    rows = store.get("FinPlan", [ef], accts, year + last, "Actual", "Working", "Periodic", closed_only=closed_only)
    rows |= store.get("FinPlan", [ef], accts, year, "Target", "Final", "Periodic")
    if s.forecast:
        rows |= store.get("FinPlan", [ef], accts, year, s.forecast, "Final", "Periodic")
    fp = from_rows("FinPlan", lines, rows, finplan_scenarios(s.forecast))
    for x in lines:
        b = f"fp.{x.key}"
        a = {"m": flow(fp, ef, x.key, "actual_finplan", [s.period], f"{b}.actual.m"),
             "ytd": flow(fp, ef, x.key, "actual_finplan", res.ytd, f"{b}.actual.ytd"),
             "ly_ytd": flow(fp, ef, x.key, "actual_finplan", periods.ytd(res.ly), f"{b}.actual.ly_ytd"),
             "ly_fy": flow(fp, ef, x.key, "actual_finplan", last, f"{b}.actual.ly_fy")}
        _keep(res, *a.values())
        plans = (("budget", "budget", res.budget), ("forecast", "forecast", res.forecast))
        for name, scen, target in plans:
            if scen == "forecast" and not s.forecast:
                continue
            pm = {"m": flow(fp, ef, x.key, scen, [s.period], f"{b}.{name}.m"),
                  "ytd": flow(fp, ef, x.key, scen, res.ytd, f"{b}.{name}.ytd"),
                  "fy": flow(fp, ef, x.key, scen, year, f"{b}.{name}.fy")}
            _keep(res, *pm.values())
            target[x.key] = {"m": variance(x, a["m"], pm["m"], f"{b}.{name}.m"),
                             "ytd": variance(x, a["ytd"], pm["ytd"], f"{b}.{name}.ytd")}
            for v in target[x.key].values():
                _keep(res, v.abs, v.pct)
    res.flags["budget"] = {k: assess(v["ytd"], cat.materiality) for k, v in res.budget.items()}
    if res.forecast:
        res.flags["forecast"] = {k: assess(v["ytd"], cat.materiality) for k, v in res.forecast.items()}
    g = grand(lines)
    rem = (flow(fp, ef, g, "forecast", periods.remaining(s.period), f"fp.{g}.forecast.remaining") if s.forecast
           else missing(f"fp.{g}.forecast.remaining", "nok", "remaining forecast months", "no latest forecast"))
    res.landing = landing(res.facts[f"fp.{g}.actual.ytd"], rem, res.facts.get(f"fp.{g}.budget.fy") or missing(
        f"fp.{g}.budget.fy", "nok", "target full year", "no target"), res.facts[f"fp.{g}.actual.ly_ytd"],
        res.facts[f"fp.{g}.actual.ly_fy"], len(res.ytd))
    _keep(res, rem, res.landing.forecast, res.landing.run_rate, res.landing.gap_forecast, res.landing.gap_run_rate)
    return fp


def compute(cat: Catalog, s: Scope, store) -> Result:
    """Every number of the answer, from the warehouse (store: warehouse.Warehouse or MemoryWarehouse)."""
    ly = periods.ly(s.period)
    res = Result(s, ly, periods.ytd(s.period), periods.ytd(ly))
    fm = _finmax(cat, s, store, res)
    _finplan(cat, s, store, res)
    for c in COMPARISONS:
        if c not in s.compare:
            res.flags.pop(c, None)
    lines = cat.pnl["FinMax"]
    g = grand(lines)
    total_ly, total_cur = res.facts[f"fm.{g}.ytd.ly"], res.facts[f"fm.{g}.ytd.cur"]
    leaves = components(cat.checks["FinMax"], lines, g)
    res.line_bridge = by_line(total_ly, total_cur, {k: v["ytd"] for k, v in res.fm.items()}, leaves,
                              {x.key: x.label for x in lines})
    if s.units:
        acct = next(x.account for x in lines if x.key == g)
        keys = [("FinMax", "Actual", "", u, acct, per, "Accumulated") for u, _ in s.units for per in (s.period, ly)]
        rows = store.get("FinMax", [u for u, _ in s.units], [acct], [s.period, ly], "Actual", "", "Accumulated")
        units = [(u, label, _cell(rows, keys[2 * i], f"unit.{u}.cur", f"{u} {g} YTD {s.period}"),
                  _cell(rows, keys[2 * i + 1], f"unit.{u}.ly", f"{u} {g} YTD {ly}"))
                 for i, (u, label) in enumerate(s.units)]
        for _, _, cur, cmp in units:
            _keep(res, cur, cmp)
        res.unit_bridge = by_unit(total_ly, total_cur, units)
    for b in (res.line_bridge, res.unit_bridge):
        if b is not None:
            _keep(res, *(st.fact for st in b.steps))
    e = s.entity.finmax
    for k in cat.kpis:
        if k.connection != "FinMax":
            continue
        out = {"ytd.cur": kpi(fm, k, e, "actual", res.ytd, f"kpi.{k.key}.ytd.cur"),
               "ytd.ly": kpi(fm, k, e, "actual", res.ly_ytd, f"kpi.{k.key}.ytd.ly")}
        if k.basis == "period":
            out["m.cur"] = kpi(fm, k, e, "actual", [s.period], f"kpi.{k.key}.m.cur")
            out["m.ly"] = kpi(fm, k, e, "actual", [ly], f"kpi.{k.key}.m.ly")
        res.kpis[k.key] = out
        _keep(res, *out.values())
    res.freshness = {"FinMax": store.freshness("FinMax", [e])}
    if s.entity.finplan:
        res.freshness["FinPlan"] = store.freshness("FinPlan", [s.entity.finplan])
    return res


def freshness_text(res: Result, mode: str) -> str:
    """'Data: warehouse, closed history to Sep 26 read 2026-10-25 08:12; the latest forecast and any open month read
    live at 10:31.'"""
    import datetime as dt

    def ts(x):
        return dt.datetime.fromtimestamp(x).strftime("%Y-%m-%d %H:%M") if x else "never"
    stamps = [f for f in res.freshness.values() if f.get("rows")]
    oldest = min((f["oldest"] for f in stamps), default=None)
    newest = max((f["newest"] for f in stamps), default=None)
    return (f"Data: {mode}; closed history to {res.scope.closed}, oldest cell read {ts(oldest)}; the latest forecast "
            f"and any open month read live (newest {ts(newest)}).")


def summary(res: Result, mode: str, fills: list) -> dict:
    """What the agent writes the answer from: facts (ids, values, definitions), flags, landing, freshness."""
    keep = ("fm.", "fp.", "kpi.", "land.", "units.")
    facts = [f.as_dict(max_inputs=3) for k, f in res.facts.items()
             if k.startswith(keep) and ".lvl." not in k and not k.endswith(".pct")]
    flags = [{"comparison": c, "line": k, "material": f.material, "always": f.always, "reason": f.reason,
              "variance": f.size} for c, k, f in res.ranking()]
    out = {"area": res.scope.entity.label, "period": res.scope.period, "closed_month": res.scope.closed,
           "forecast": res.scope.forecast, "freshness": freshness_text(res, mode), "flags": flags, "facts": facts,
           "fills": [{"job": r.fill.job, "result": r.fill.result, "cells": r.fill.cells, "seconds": r.fill.seconds,
                      "errors": r.fill.detail.get("errors", [])[:3], "problems": r.fill.detail.get("problems", [])[:3]}
                     for r in fills]}
    if res.landing is not None:
        out["landing"] = {"forecast": res.landing.forecast.value, "run_rate": res.landing.run_rate.value,
                          "target": res.landing.target.value, "gap_forecast": res.landing.gap_forecast.value,
                          "gap_run_rate": res.landing.gap_run_rate.value, "method": res.landing.method}
    return out
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_fin_pnl.py -q`. Expected: `5 passed`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/fin_pnl.py engine/tests/fin_util.py engine/tests/test_fin_pnl.py
PY ..\scripts\check_staged.py
git commit -m "fin_pnl: the business-area P&L computed from the warehouse (variances, flags, bridges, KPIs, landing)"
```

### Task 18: `fin_spec.py`, the workbook request and warehouse = sheet

The source sheets are simulated in the tests: each pull is retrieved from the same fake provider. That proves three
things without Excel:
- every pull cell the spec names exists;
- the spec, evaluated in Python (`specs.evaluate`, the same rules as the sheet formulas), gives the engine's numbers;
- the reconciliation finds a changed sheet value.

**Files:**
- Create: `engine\smartview_mcp\fin_spec.py`
- Test: `engine\tests\test_fin_spec.py` (new)

- [ ] **Step 1: Write the failing tests**

```python
"""fin_spec: the sv_ws_build request for fin_pnl. The source sheets are simulated by retrieving each pull from the
same fake provider; the spec evaluated in Python must give the engine's numbers, and warehouse = sheet must hold.
Every figure is invented."""
import pytest
from fin_util import CAT, filled, http

from smartview_mcp import fin_pnl, fin_spec, specs
from smartview_mcp.grid import Axis
from smartview_mcp.userexcel import _make_pull


def sheets(tmp_path, req):
    """{pull id: Pull} as the source sheets would hold them (same provider, suppression off)."""
    sv, out = http(tmp_path / "sheets"), {}
    (tmp_path / "sheets").mkdir(exist_ok=True)
    for i, pl in enumerate(req.pulls):
        rows = [Axis(a["dimension"], a["members"]) for a in pl["rows"]]
        cols = [Axis(a["dimension"], a["members"]) for a in pl["columns"]]
        lay, parsed, _ = sv.retrieve(pl["connection"], pl["pov"], rows, cols, False, False)
        out[pl["id"]] = _make_pull(f"Q1 {pl['connection']} {pl['id']}", pl["connection"], lay, parsed, 0.0)
    return out


@pytest.fixture
def built(tmp_path):
    store, _, scope = filled(tmp_path)
    res = fin_pnl.compute(CAT, scope, store)
    req = fin_spec.build(res, CAT, now="2026-10-25 12:00")
    return res, req, sheets(tmp_path, req)


def test_the_request_has_the_five_source_sheets_and_formula_cells_only(built):
    res, req, pulls = built
    assert [p["id"] for p in req.pulls] == ["pnl", "fte", "fpa", "fpp", "units"]
    assert req.spec["auto_points"] is False and req.spec["chart"]["type"] == "waterfall"
    titles = [s["title"] for s in req.spec["sections"]]
    assert titles[0].startswith("P&L vs last year (FinMax)") and any(t.startswith("Year-end landing") for t in titles)
    for sec in req.spec["sections"]:
        for ln in sec["lines"]:
            for c in ln["cells"]:
                if isinstance(c, dict) and "value" in c:            # only the labelled calendar inputs
                    assert ln["key"] in ("months_year", "months_ytd")
    labels = [c["label"] for c in req.spec["checks"]]
    assert "C30 + C40S = C48" in labels and "Revenue + COGS = Gross profit" in labels
    assert "Line bridge closes" in labels and "Sub-unit bridge closes" in labels and labels[-1] == "Period is closed"


def test_the_spec_evaluates_to_the_engine_numbers(built):
    res, req, pulls = built
    ids = {pid: p.sheet for pid, p in pulls.items()}
    sess = {p.sheet: p for p in pulls.values()}
    s = specs.expand(req.spec, ids, sess)
    v = specs.evaluate(s, sess)
    assert v["ebitda"][4] == pytest.approx(res.fm["ebitda"]["ytd"].current.value)
    assert v["ebitda"][6] == pytest.approx(res.fm["ebitda"]["ytd"].abs.value)
    assert v["fte"][4] == pytest.approx(res.fm["fte"]["ytd"].current.value)          # the n+1 average
    assert v["k_opex_per_fte"][4] == pytest.approx(res.kpis["opex_per_fte"]["ytd.cur"].value)
    assert v["k_ebitda_margin"][0] == pytest.approx(res.kpis["ebitda_margin"]["m.cur"].value)
    assert v["b_ebitda"][6] == pytest.approx(res.budget["ebitda"]["ytd"].abs.value)
    assert v["l_fc"][4] == pytest.approx(res.landing.forecast.value)
    assert v["l_rr"][4] == pytest.approx(res.landing.run_rate.value)
    assert v["l_gap"][4] == pytest.approx(res.landing.gap_forecast.value)
    assert v["lb_sum"][4] == pytest.approx(v["lb_end"][4]) and v["ub_sum"][4] == pytest.approx(v["ub_end"][4])
    assert v["u_other"][6] == pytest.approx(res.unit_bridge.steps[-1].fact.value)
    by_line = {ln["key"]: ln for sec in s["sections"] for ln in sec["lines"]}
    assert by_line["rev"]["cells"][8] == {"text": "comment"}                          # an always-line


def test_warehouse_equals_every_source_sheet_and_a_difference_is_named(built):
    res, req, pulls = built
    spec = fin_spec.with_reconciliation(req.spec, req.expected, pulls)
    recon = [c for c in spec["checks"] if c["label"].startswith("Warehouse = source sheet")]
    assert len(recon) == 5 and all(c["static"] == "OK" for c in recon), recon
    assert sum(len(v) for v in req.expected.values()) > 60
    p = pulls["pnl"]
    row = next(r for r in p.parsed.rows if r[0] == "ac.C30")
    row[p.layout.nr + 2] += 5.0                                                       # Accumulated, Sep 26
    again = fin_spec.with_reconciliation(req.spec, req.expected, pulls)
    bad = next(c for c in again["checks"] if c["label"] == f"Warehouse = source sheet '{p.sheet}'")
    assert bad["static"] == "DIFF" and "ac.C30 Accumulated/Sep 26: sheet" in bad["detail"]


def test_key_points_are_factual(built):
    res, req, _ = built
    kp = req.spec["key_points"]
    assert kp[0].startswith("EBITDA YTD Sep 26: 2 871 vs 2 610 YTD Sep 25 (variance 261,")
    assert any(k.startswith("Year-end landing FY26:") for k in kp)
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_fin_spec.py -q`
Expected: FAIL, `ImportError: cannot import name 'fin_spec' from 'smartview_mcp'`

- [ ] **Step 3: Implement**

```python
"""The workbook for fin_pnl: the sv_ws_build request (live Smart View source sheets + an answer spec whose every
number is a formula over them) and the expected value of every source cell the answer uses, from the calc engine's
facts - so the build can prove warehouse = sheet (with_reconciliation, run as _build_job's before_answer hook).

Answer columns (one set for every section; the section title says what the comparison is):
  0 month | 1 comparison month | 2 variance | 3 variance % | 4 YTD | 5 comparison YTD | 6 variance | 7 variance % |
  8 material (YTD). Bridges and the landing use column 4."""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass

from .analytics import periods
from .catalog import Catalog
from .fin_pnl import Result, freshness_text, grand
from .templates import _fmt, _pct

TOL = 1.0                                                  # NOK: warehouse = sheet
COLUMNS = ["{p}", "Comparison", "Variance", "Variance %", "YTD {p}", "Comparison YTD", "Variance YTD",
           "Variance % YTD", "Material (YTD)"]
KINDS = ["heltall", "heltall", "heltall", "prosent1", "heltall", "heltall", "heltall", "prosent1", "heltall"]


@dataclass
class Request:
    question: str
    pulls: list[dict]                                        # PullIn fields: id, connection, rows, columns, pov
    spec: dict
    expected: dict[str, list[tuple[list[str], list[str], float | None, str]]]   # pull id -> (row, col, value, fact)


def _ax(dim: str, members: list[str]) -> dict:
    return {"dimension": dim, "members": list(dict.fromkeys(members))}


def _pull(pull: str, row: list[str], col: list[str]) -> dict:
    return {"pull": pull, "row": row, "col": col}


def _flag_text(res: Result, comparison: str, key: str):
    f = res.flags.get(comparison, {}).get(key)
    if f is None or not f.comment:
        return None
    return {"text": "material" if f.material else "comment"}


def _value(res: Result, fid: str):
    f = res.facts.get(fid)
    return None if f is None else f.value


def _prefixed(expr: str, prefix: str) -> str:
    return re.sub(r"[a-z_]\w*", lambda m: prefix + m.group(0), expr)


def build(res: Result, cat: Catalog, now: str | None = None) -> Request:
    s = res.scope
    p, ly = s.period, res.ly
    mon, fy = periods.finplan(p)
    fy_ly = periods.finplan(ly)[1]
    fm_lines, fp_lines = cat.pnl["FinMax"], cat.pnl["FinPlan"]
    g, gp = grand(fm_lines), grand(fp_lines)
    stocks = [x for x in fm_lines if x.measure == "stock"]
    e = s.entity.finmax
    base = {"All Entities": e, "All Scenarios": "Actual", "All Currencies": "NOK"}
    pulls = [{"id": "pnl", "connection": "FinMax", "pov": base,
              "rows": [_ax("All Accounts", [x.account for x in fm_lines])],
              "columns": [_ax("All Balance Types", ["Periodic", "Accumulated"]), _ax("All Time Periods", [p, ly])]}]
    expected: dict[str, list] = {"pnl": []}
    cur_pts = [periods.opening(p), *res.ytd]
    ly_pts = [periods.opening(ly), *res.ly_ytd]
    if stocks:
        pulls.append({"id": "fte", "connection": "FinMax", "pov": {**base, "All Balance Types": "Accumulated"},
                      "rows": [_ax("All Accounts", [x.account for x in stocks])],
                      "columns": [_ax("All Time Periods", cur_pts + ly_pts)]})
        expected["fte"] = []
    sections, checks = [], []

    # ---- A. FinMax P&L vs last year
    lines = []
    for x in fm_lines:
        a = [x.account]
        if x.measure == "stock":
            cells = [_pull("pnl", a, ["Accumulated", p]), _pull("pnl", a, ["Accumulated", ly]),
                     {"var": [0, 1]}, {"varpct": [0, 1]}, {"ref": {"key": f"{x.key}_avg", "col": 4}},
                     {"ref": {"key": f"{x.key}_avg", "col": 5}}, {"var": [4, 5]}, {"varpct": [4, 5]}, None]
            expected["pnl"] += [(a, ["Accumulated", p], _value(res, f"fm.{x.key}.m.cur"), f"fm.{x.key}.m.cur"),
                                (a, ["Accumulated", ly], _value(res, f"fm.{x.key}.m.ly"), f"fm.{x.key}.m.ly")]
            label = f"{x.label} (month end; YTD = average)"
        else:
            cells = [_pull("pnl", a, ["Periodic", p]), _pull("pnl", a, ["Periodic", ly]), {"var": [0, 1]},
                     {"varpct": [0, 1]}, _pull("pnl", a, ["Accumulated", p]),
                     _pull("pnl", a, ["Accumulated", ly]), {"var": [4, 5]}, {"varpct": [4, 5]},
                     _flag_text(res, "ly", x.key)]
            for bal, per, part in (("Periodic", p, "m.cur"), ("Periodic", ly, "m.ly"), ("Accumulated", p, "ytd.cur"),
                                   ("Accumulated", ly, "ytd.ly")):
                fid = f"fm.{x.key}.{part}"
                expected["pnl"].append((a, [bal, per], _value(res, fid), fid))
            label = x.label
        ln = {"key": x.key, "label": label, "kind": x.kind, "cells": cells}
        if x.num:
            ln["num"] = x.num
        lines.append(ln)
    sections.append({"title": f"P&L vs last year (FinMax): {p} vs {ly}, YTD vs YTD {ly}", "lines": lines})
    checks += [{"label": c.label, "expr": c.expr, "columns": [0, 1, 4, 5]} for c in cat.checks["FinMax"]]

    # ---- FTE levels for the n+1 average
    for x in stocks:
        lv = []
        for i, (c, l) in enumerate(zip(cur_pts, ly_pts)):
            lv.append({"key": f"{x.key}_l{i}", "label": f"{x.label} {c} / {l}", "kind": "line", "aux": True,
                       "num": x.num or "tall1",
                       "cells": [None] * 4 + [_pull("fte", [x.account], [c]), _pull("fte", [x.account], [l])]})
            expected["fte"] += [([x.account], [c], _value(res, f"fm.{x.key}.lvl.{c}"), f"fm.{x.key}.lvl.{c}"),
                                ([x.account], [l], _value(res, f"fm.{x.key}.lvl.{l}"), f"fm.{x.key}.lvl.{l}")]
        keys = [ln["key"] for ln in lv]
        lv.append({"key": f"{x.key}_avg", "label": f"Average {x.label} ({len(keys)} points: opening + month ends)",
                   "kind": "sum", "aux": True, "num": x.num or "tall1",
                   "cells": [None] * 4 + [{"avg": keys}, {"avg": keys}, {"var": [4, 5]}, {"varpct": [4, 5]}]})
        sections.append({"title": f"{x.label} levels for the average (opening December + each month end)",
                         "lines": lv})

    # ---- inputs and KPIs
    need_inputs = any(k.annualise for k in cat.kpis if k.connection == "FinMax")
    kl = []
    for k in cat.kpis:
        if k.connection != "FinMax" or k.key not in res.kpis:
            continue
        f = res.kpis[k.key]
        den = f"{k.denominator}_avg" if k.basis == "average_stock" else k.denominator
        expr = f"{k.numerator}/{den}" + ("*months_year/months_ytd" if k.annualise else "")

        def cell(part):
            fact = f.get(part)
            if fact is None:
                return None
            return {"expr": expr} if fact.value is not None else {"text": "n/a"}
        cells = [cell("m.cur"), cell("m.ly"), None, None, cell("ytd.cur"), cell("ytd.ly"), None, None, None]
        if cells[0] and cells[1] and "expr" in cells[0] and "expr" in cells[1]:
            cells[2] = {"var": [0, 1]}
        if cells[4] and cells[5] and "expr" in cells[4] and "expr" in cells[5]:
            cells[6] = {"var": [4, 5]}
        kl.append({"key": f"k_{k.key}", "label": k.label + (" (annualised)" if k.annualise else ""), "kind": "line",
                   "num": "prosent1" if k.unit == "pct" else "heltall", "cells": cells})
    if res.landing is not None and not res.landing.seasonal:
        need_inputs = True
    if need_inputs:
        sections.append({"title": "Inputs", "lines": [
            {"key": "months_year", "label": "Months in a year", "kind": "line", "aux": True,
             "cells": [None] * 4 + [{"value": 12}, {"value": 12}]},
            {"key": "months_ytd", "label": f"Months in the period (YTD {p})", "kind": "line", "aux": True,
             "cells": [None] * 4 + [{"value": len(res.ytd)}, {"value": len(res.ly_ytd)}]}]})
    if kl:
        sections.append({"title": "KPIs (FinMax; variance in points or NOK)", "lines": kl})

    # ---- B/C. FinPlan vs budget and the latest forecast
    ef = s.entity.finplan
    if ef and res.budget:
        years = [fy, fy_ly]
        months = [mon] if mon == "Dec" else [mon, "Dec"]
        pulls.append({"id": "fpa", "connection": "FinPlan",
                      "pov": {"Entity": ef, "Scenario": "Actual", "Version": "Working"},
                      "rows": [_ax("Account", [x.account for x in fp_lines])],
                      "columns": [_ax("Years", years), _ax("All Balance Types", ["Periodic", "Accumulated"]),
                                  _ax("Period", months)]})
        scen = ["Target"] + ([s.forecast] if s.forecast and res.forecast else [])
        pulls.append({"id": "fpp", "connection": "FinPlan", "pov": {"Entity": ef, "Years": fy, "Version": "Final"},
                      "rows": [_ax("Account", [x.account for x in fp_lines])],
                      "columns": [_ax("Scenario", scen), _ax("All Balance Types", ["Periodic", "Accumulated"]),
                                  _ax("Period", months)]})
        expected["fpa"], expected["fpp"] = [], []
        bl, fl = [], []
        for x in fp_lines:
            a = [x.account]
            bl.append({"key": f"b_{x.key}", "label": x.label, "kind": x.kind, "cells": [
                _pull("fpa", a, [fy, "Periodic", mon]), _pull("fpp", a, ["Target", "Periodic", mon]),
                {"var": [0, 1]}, {"varpct": [0, 1]}, _pull("fpa", a, [fy, "Accumulated", mon]),
                _pull("fpp", a, ["Target", "Accumulated", mon]), {"var": [4, 5]}, {"varpct": [4, 5]},
                _flag_text(res, "budget", x.key)]})
            b = f"fp.{x.key}"
            expected["fpa"] += [(a, [fy, "Periodic", mon], _value(res, f"{b}.actual.m"), f"{b}.actual.m"),
                                (a, [fy, "Accumulated", mon], _value(res, f"{b}.actual.ytd"), f"{b}.actual.ytd")]
            expected["fpp"] += [(a, ["Target", "Periodic", mon], _value(res, f"{b}.budget.m"), f"{b}.budget.m"),
                                (a, ["Target", "Accumulated", mon], _value(res, f"{b}.budget.ytd"), f"{b}.budget.ytd")]
            if len(scen) > 1:
                fc = s.forecast
                fl.append({"key": f"f_{x.key}", "label": x.label, "kind": x.kind, "cells": [
                    {"ref": {"key": f"b_{x.key}", "col": 0}}, _pull("fpp", a, [fc, "Periodic", mon]),
                    {"var": [0, 1]}, {"varpct": [0, 1]}, {"ref": {"key": f"b_{x.key}", "col": 4}},
                    _pull("fpp", a, [fc, "Accumulated", mon]), {"var": [4, 5]}, {"varpct": [4, 5]},
                    _flag_text(res, "forecast", x.key)]})
                expected["fpp"] += [
                    (a, [fc, "Periodic", mon], _value(res, f"{b}.forecast.m"), f"{b}.forecast.m"),
                    (a, [fc, "Accumulated", mon], _value(res, f"{b}.forecast.ytd"), f"{b}.forecast.ytd")]
        bl.append({"key": "recon", "label": "EBITDA: FinPlan actual less FinMax (definitions differ)", "kind": "line",
                   "aux": True, "cells": [{"expr": f"b_{gp}-{g}"}, None, None, None, {"expr": f"b_{gp}-{g}"}]})
        sections.append({"title": f"P&L vs budget (FinPlan; actual = Working, Target = Final): {mon} {fy}",
                         "lines": bl})
        checks += [{"label": c.label, "expr": _prefixed(c.expr, "b_"), "columns": [0, 1, 4, 5]}
                   for c in cat.checks["FinPlan"]]
        if fl:
            sections.append({"title": f"P&L vs the latest forecast ({s.forecast}, Final)", "lines": fl})

    # ---- D. EBITDA bridge by line (FinMax YTD)
    chart = None
    if res.line_bridge is not None:
        steps = [st.key for st in res.line_bridge.steps]
        if any(sign < 0 for sign, _ in _leaves(cat)):
            raise ValueError("a negative leaf in the EBITDA chain is not supported by the line bridge")
        bl = [{"key": "lb_start", "label": f"EBITDA YTD {ly}", "kind": "line", "cells": [None] * 4 + [
            {"ref": {"key": g, "col": 5}}]}]
        bl += [{"key": f"lb_{k}", "label": next(x.label for x in fm_lines if x.key == k), "kind": "line",
                "cells": [None] * 4 + [{"ref": {"key": k, "col": 6}}]} for k in steps]
        bl.append({"key": "lb_end", "label": f"EBITDA YTD {p}", "kind": "sum",
                   "cells": [None] * 4 + [{"ref": {"key": g, "col": 4}}]})
        bl.append({"key": "lb_sum", "label": "Start + steps (check)", "kind": "line", "aux": True,
                   "cells": [None] * 4 + [{"sum": ["lb_start"] + [f"lb_{k}" for k in steps]}]})
        sections.append({"title": f"EBITDA bridge by line, YTD {ly} → YTD {p} (FinMax)", "lines": bl})
        checks.append({"label": "Line bridge closes", "expr": "lb_sum-lb_end", "columns": [4]})
        chart = {"type": "waterfall", "lines": ["lb_start"] + [f"lb_{k}" for k in steps] + ["lb_end"], "column": 4,
                 "title": f"EBITDA bridge YTD {ly} → YTD {p}"}

    # ---- E. EBITDA bridge by sub-unit
    if res.unit_bridge is not None and s.units:
        acct = next(x.account for x in fm_lines if x.key == g)
        pulls.append({"id": "units", "connection": "FinMax",
                      "pov": {"All Scenarios": "Actual", "All Currencies": "NOK", "All Accounts": acct,
                              "All Balance Types": "Accumulated"},
                      "rows": [_ax("All Entities", [u for u, _ in s.units])],
                      "columns": [_ax("All Time Periods", [p, ly])]})
        expected["units"] = [r for u, _ in s.units for r in (
            ([u], [p], _value(res, f"unit.{u}.cur"), f"unit.{u}.cur"),
            ([u], [ly], _value(res, f"unit.{u}.ly"), f"unit.{u}.ly"))]
        top = [st for st in res.unit_bridge.steps if st.key != "other"]
        ul = [{"key": f"u{i}", "label": st.label, "kind": "line", "cells": [None] * 4 + [
            _pull("units", [st.key], [p]), _pull("units", [st.key], [ly]), {"var": [4, 5]},
            {"varpct": [4, 5]}]} for i, st in enumerate(top, 1)]
        rest = "-".join([g] + [f"u{i}" for i in range(1, len(top) + 1)])
        ul.append({"key": "u_other", "label": "Other units", "kind": "line",
                   "cells": [None] * 4 + [{"expr": rest}, {"expr": rest}, {"var": [4, 5]}, {"varpct": [4, 5]}]})
        sections.append({"title": f"EBITDA by sub-unit, YTD {p} vs YTD {ly} (largest variances)", "lines": ul})
        ub = [{"key": "ub_start", "label": f"EBITDA YTD {ly}", "kind": "line",
               "cells": [None] * 4 + [{"ref": {"key": g, "col": 5}}]}]
        ub += [{"key": f"ub{i}", "label": st.label, "kind": "line",
                "cells": [None] * 4 + [{"ref": {"key": f"u{i}", "col": 6}}]} for i, st in enumerate(top, 1)]
        ub.append({"key": "ub_other", "label": "Other units", "kind": "line",
                   "cells": [None] * 4 + [{"ref": {"key": "u_other", "col": 6}}]})
        ub.append({"key": "ub_end", "label": f"EBITDA YTD {p}", "kind": "sum",
                   "cells": [None] * 4 + [{"ref": {"key": g, "col": 4}}]})
        ub.append({"key": "ub_sum", "label": "Start + steps (check)", "kind": "line", "aux": True,
                   "cells": [None] * 4 + [{"sum": ["ub_start"] + [f"ub{i}" for i in range(1, len(top) + 1)]
                                          + ["ub_other"]}]})
        sections.append({"title": f"EBITDA bridge by sub-unit, YTD {ly} → YTD {p}", "lines": ub})
        checks.append({"label": "Sub-unit bridge closes", "expr": "ub_sum-ub_end", "columns": [4]})

    # ---- F. Year-end landing (FinPlan EBITDA)
    if res.landing is not None and ef and res.budget:
        acct = next(x.account for x in fp_lines if x.key == gp)
        ld = res.landing
        lines = [{"key": "l_ytd", "label": f"Actual YTD {mon} {fy}", "kind": "line",
                  "cells": [None] * 4 + [{"ref": {"key": f"b_{gp}", "col": 4}}]}]
        if res.forecast:
            fc = s.forecast
            lines += [
                {"key": "l_fc_fy", "label": f"{fc}: full year", "kind": "line", "aux": True,
                 "cells": [None] * 4 + [_pull("fpp", [acct], [fc, "Accumulated", "Dec"])]},
                {"key": "l_fc_ytd", "label": f"{fc}: YTD {mon}", "kind": "line", "aux": True,
                 "cells": [None] * 4 + [{"ref": {"key": f"f_{gp}", "col": 5}}]},
                {"key": "l_rem", "label": "Remaining months of the forecast", "kind": "line",
                 "cells": [None] * 4 + [{"expr": "l_fc_fy-l_fc_ytd"}]},
                {"key": "l_fc", "label": "Landing: actual YTD + remaining forecast", "kind": "sum",
                 "cells": [None] * 4 + [{"expr": "l_ytd+l_rem"}]}]
            expected["fpp"].append(([acct], [fc, "Accumulated", "Dec"], _value(res, f"fp.{gp}.forecast.fy"),
                                    f"fp.{gp}.forecast.fy"))
        lines.append({"key": "l_tgt", "label": "Target (budget): full year", "kind": "line",
                      "cells": [None] * 4 + [_pull("fpp", [acct], ["Target", "Accumulated", "Dec"])]})
        expected["fpp"].append(([acct], ["Target", "Accumulated", "Dec"], _value(res, f"fp.{gp}.budget.fy"),
                                f"fp.{gp}.budget.fy"))
        if res.forecast:
            lines.append({"key": "l_gap", "label": "Landing less target", "kind": "line",
                          "cells": [None] * 4 + [{"expr": "l_fc-l_tgt"}]})
        if ld.seasonal:
            lines += [
                {"key": "l_ly_ytd", "label": f"Last year: actual YTD {mon} {fy_ly}", "kind": "line", "aux": True,
                 "cells": [None] * 4 + [_pull("fpa", [acct], [fy_ly, "Accumulated", mon])]},
                {"key": "l_ly_fy", "label": f"Last year: actual full year {fy_ly}", "kind": "line", "aux": True,
                 "cells": [None] * 4 + [_pull("fpa", [acct], [fy_ly, "Accumulated", "Dec"])]},
                {"key": "l_share", "label": "Last year's YTD share of its full year", "kind": "line", "aux": True,
                 "num": "prosent1", "cells": [None] * 4 + [{"expr": "l_ly_ytd/l_ly_fy"}]},
                {"key": "l_rr", "label": "Run-rate landing (YTD / last year's share)", "kind": "sum",
                 "cells": [None] * 4 + [{"expr": "l_ytd/l_share"}]}]
            expected["fpa"] += [
                ([acct], [fy_ly, "Accumulated", mon], _value(res, f"fp.{gp}.actual.ly_ytd"), f"fp.{gp}.actual.ly_ytd"),
                ([acct], [fy_ly, "Accumulated", "Dec"], _value(res, f"fp.{gp}.actual.ly_fy"), f"fp.{gp}.actual.ly_fy")]
        else:
            lines.append({"key": "l_rr", "label": f"Run-rate landing (YTD / {len(res.ytd)} x 12)", "kind": "sum",
                          "cells": [None] * 4 + [{"expr": "l_ytd/months_ytd*months_year"}]})
        lines.append({"key": "l_rr_gap", "label": "Run-rate landing less target", "kind": "line",
                      "cells": [None] * 4 + [{"expr": "l_rr-l_tgt"}]})
        sections.append({"title": f"Year-end landing {fy} (FinPlan EBITDA; values in the YTD column)",
                         "lines": lines})

    closed_ok = not periods.after(p, s.closed)
    checks.append({"label": "Period is closed", "static": "OK" if closed_ok else "DIFF",
                   "detail": f"{p} vs Current Reporting Month {s.closed}"})
    m = cat.materiality
    notes = [freshness_text(res, "warehouse"),
             "Signs as in FinMax: costs are negative; variance = current − comparison (positive = better).",
             f"Material (YTD): {m.pct:.0%} of |comparison| from NOK {m.pct_floor_nok:,.0f}, "
             f"{'or' if m.rule == 'either' else 'and'} NOK {m.abs_nok:,.0f}; 'comment' = always commented "
             f"({', '.join(m.always)}). The flags are set by smartass from the warehouse; they do not change when the "
             "sheets are refreshed.".replace(",", " "),
             "Actual vs budget and forecast is read entirely in FinPlan (its definitions differ from FinMax; the "
             "EBITDA difference line shows by how much).",
             "History is restated to today's organisation (e.g. the 1 June 2026 reorg)."]
    if not closed_ok:
        notes.insert(0, f"{p} is after the Current Reporting Month ({s.closed}): the period is still open.")
    title = f"{s.entity.label} — P&L {p} and YTD vs last year, budget and forecast"
    spec = {"title": title,
            "subtitle": f"NOK · FinMax and FinPlan · {s.entity.finmax} / {s.entity.finplan or '-'}",
            "auto_points": False, "key_points": key_points(res, cat),
            "columns": [{"header": h.format(p=p), "kind": k} for h, k in zip(COLUMNS, KINDS)],
            "sections": sections, "checks": checks, "notes": notes,
            "source": f"Source: Oracle EPM Cloud (FinMax, FinPlan) via Smart View, retrieved "
                      f"{now or dt.datetime.now().strftime('%Y-%m-%d %H:%M')}.",
            "chart": chart}
    return Request(f"{s.entity.label} P&L {p}", pulls, spec, expected)


def _leaves(cat: Catalog):
    from .analytics.bridge import components
    return components(cat.checks["FinMax"], cat.pnl["FinMax"], grand(cat.pnl["FinMax"]))


def key_points(res: Result, cat: Catalog) -> list[str]:
    """Factual points drafted from the facts (interpretation is the agent's, Stage 5)."""
    g, gp = grand(cat.pnl["FinMax"]), grand(cat.pnl["FinPlan"])
    labels = {**{x.key: x.label for x in cat.pnl["FinPlan"]}, **{x.key: x.label for x in cat.pnl["FinMax"]}}
    out = []
    v = res.fm.get(g, {}).get("ytd")
    if v is not None:
        out.append(f"EBITDA YTD {res.scope.period}: {_fmt(v.current.value)} vs {_fmt(v.comparison.value)} YTD "
                   f"{res.ly} (variance {_fmt(v.abs.value)}, {_pct(v.current.value, v.comparison.value)}).")
    b = res.budget.get(gp, {}).get("ytd")
    if b is not None:
        out.append(f"Against budget (FinPlan): EBITDA YTD {_fmt(b.current.value)} vs target {_fmt(b.comparison.value)} "
                   f"(variance {_fmt(b.abs.value)}, {_pct(b.current.value, b.comparison.value)}).")
    names = {"ly": "last year", "budget": "budget", "forecast": "forecast"}
    mat = [(c, k, f) for c, k, f in res.ranking() if f.material][:3]
    if mat:
        out.append("Material YTD variances: " + "; ".join(
            f"{labels.get(k, k)} vs {names[c]} {_fmt(_var(res, c, k))}" for c, k, f in mat) + ".")
    ld = res.landing
    if ld is not None and ld.run_rate.value is not None:
        text = f"Year-end landing {periods.finplan(res.scope.period)[1]}: "
        if ld.forecast.value is not None:
            text += f"{_fmt(ld.forecast.value)} with the latest forecast, "
        text += f"{_fmt(ld.run_rate.value)} at run rate; target {_fmt(ld.target.value)}."
        out.append(text)
    ub = res.unit_bridge
    if ub is not None and len(ub.steps) > 1 and ub.steps[0].fact.value is not None:
        out.append(f"Largest sub-unit EBITDA variance vs YTD {res.ly}: {ub.steps[0].label} "
                   f"{_fmt(ub.steps[0].fact.value)}.")
    return out


def _var(res: Result, comparison: str, key: str) -> float | None:
    src = {"ly": res.fm, "budget": res.budget, "forecast": res.forecast}[comparison]
    v = src.get(key, {}).get("ytd")
    return None if v is None else v.abs.value


def _sheet_value(p, row: list[str], col: list[str]):
    """The pulled cell as Smart View put it on the source sheet: a number, 0.0 for blank, a label such as
    '#NoAccess', or KeyError when the sheet has no such cell."""
    j = p.layout.col_tuples.index(tuple(col))
    for r in p.parsed.rows:
        if tuple(r[:p.layout.nr]) == tuple(row):
            v = r[p.layout.nr + j]
            return 0.0 if v is None else v
    raise KeyError(tuple(row))


def reconcile_pull(p, cells: list[tuple[list[str], list[str], float | None, str]], tol: float = TOL) -> dict:
    """The static check 'warehouse = source sheet' for one pull."""
    worst, bad = 0.0, []
    for row, col, want, fid in cells:
        try:
            got = _sheet_value(p, row, col)
        except (KeyError, ValueError):
            bad.append(f"{'/'.join(row)} {'/'.join(col)} not on the sheet")
            continue
        if want is None:
            bad.append(f"{'/'.join(row)} {'/'.join(col)}: no warehouse value ({fid})")
            continue
        if not isinstance(got, (int, float)):
            bad.append(f"{'/'.join(row)} {'/'.join(col)}: the sheet shows {got}")
            continue
        d = abs(float(got) - want)
        worst = max(worst, d)
        if d > tol:
            bad.append(f"{'/'.join(row)} {'/'.join(col)}: sheet {_fmt(got)}, warehouse {_fmt(want)}")
    detail = f"{len(cells)} cells, largest difference NOK {worst:,.0f}".replace(",", " ")
    if bad:
        detail += f"; {len(bad)} off, first: {bad[0]}"
    return {"label": f"Warehouse = source sheet '{p.sheet}'", "static": "DIFF" if bad else "OK", "detail": detail}


def with_reconciliation(spec: dict, expected: dict, pulls_by_id: dict) -> dict:
    """_build_job's before_answer hook: one 'warehouse = sheet' check per source sheet, after the others."""
    extra = [reconcile_pull(pulls_by_id[pid], cells) for pid, cells in expected.items() if pid in pulls_by_id]
    return {**spec, "checks": list(spec.get("checks") or []) + extra}
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_fin_spec.py -q`. Expected: `4 passed`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/fin_spec.py engine/tests/test_fin_spec.py
PY ..\scripts\check_staged.py
git commit -m "fin_spec: sv_ws_build request for fin_pnl, formulas only, warehouse = source sheet checks"
```

### Task 19: tools `fin_pnl` and `sv_warehouse`

**Files:**
- Create: `engine\smartview_mcp\tools_fin.py`
- Modify: `engine\smartview_mcp\server.py` (import the module after `tools_ws`)
- Test: `engine\tests\test_tools_fin.py` (new); `engine\tests\test_smoke_stdio.py:7-8`

- [ ] **Step 1: Write the failing tests**

`engine\tests\test_tools_fin.py`:

```python
"""tools_fin: area resolution, prepare() with the warehouse and with the live-pull fallback, and the fin_pnl tool
handing its request to _build_job with the reconciliation hook. Every figure is invented."""
import sqlite3

import pytest
from fin_util import CAT, FC, WHS, http
from mcp.server.mcpserver.exceptions import ToolError

from smartview_mcp import atlas, extract, fin_pnl, templates, tools_fin, tools_ws
from smartview_mcp.warehouse import MemoryWarehouse, WarehouseUnavailable


@pytest.fixture
def env(tmp_path, monkeypatch):
    db = sqlite3.connect(tmp_path / "atlas.db")
    db.executescript(atlas.SCHEMA)
    db.executemany("INSERT INTO member VALUES ('FinMax','All Entities',?,?,?,1,0)",
                   [("un.SY.2WHS", "2WHS Wholesale (un.SY)", "un.SY"),
                    ("un.SY.3WHS_W", "3WHS_W W (un.SY)", "un.SY.2WHS"),
                    ("un.SY.U1", "U1 Unit one (un.SY)", "un.SY.3WHS_W"),
                    ("un.SY.U2", "U2 Unit two (un.SY)", "un.SY.3WHS_W")])
    db.commit()
    monkeypatch.setattr(templates, "closed_month", lambda h, c="FinMax": "Sep 26")
    monkeypatch.setattr(templates, "latest_forecast", lambda yy: FC if yy == 26 else None)
    return extract.route(http=http(tmp_path)), db


@pytest.mark.parametrize("text, msg", [("Narnia", "names no business area"),
                                       ("Wholesale vs Business", "names several areas"),
                                       ("Consumer", "abolished on 2026-06-01")])
def test_resolve_area_explains(text, msg):
    with pytest.raises(ToolError, match=msg):
        tools_fin.resolve_area(CAT, text)
    assert tools_fin.resolve_area(CAT, "grossist").key == "wholesale"


def test_prepare_fills_the_warehouse_then_computes(env):
    retrieve, db = env
    res, mode, fills = tools_fin.prepare(None, CAT, WHS, None, fin_pnl.COMPARISONS, store=MemoryWarehouse(),
                                         retrieve=retrieve, db=db)
    assert mode == "warehouse" and [f.fill.job for f in fills] == ["base", "top_up", "top_up"]
    assert res.scope.units == (("un.SY.U1", "U1 Unit one"), ("un.SY.U2", "U2 Unit two"))
    assert res.unit_bridge.closes() and res.scope.forecast == FC


def test_prepare_answers_from_live_pulls_when_the_warehouse_is_unavailable(env):
    class Locked(MemoryWarehouse):
        def state(self):
            raise WarehouseUnavailable("another smartass process is writing")
    retrieve, db = env
    res, mode, fills = tools_fin.prepare(None, CAT, WHS, "Sep 26", ("ly",), store=Locked(), retrieve=retrieve, db=db)
    assert mode.startswith("live pulls (the warehouse was unavailable: another smartass")
    assert fills[0].fill.job == "live" and res.fm["ebitda"]["ytd"].current.value is not None


@pytest.mark.anyio
async def test_the_tool_builds_through_build_job_with_the_hook(env, monkeypatch, tmp_path):
    seen = {}

    def build_job(h, tm, question, pulls, spec, target, allow, q, replace=False, before_answer=None):
        seen.update(question=question, ids=[p.id for p in pulls], hook=before_answer, spec=spec)
        return {"workbook": "C:/x/a.xlsx", "question": "Q1", "checks": {}, "timings": {}}

    async def call_user(fn, what):
        return fn(None)
    monkeypatch.setattr(tools_ws, "_build_job", build_job)
    monkeypatch.setattr(tools_fin, "_call_user", call_user)
    monkeypatch.setattr(tools_fin, "Warehouse", MemoryWarehouse)
    retrieve, db = env
    real = atlas.connect                       # svhttp still opens its own test atlas by path
    monkeypatch.setattr(extract, "route", lambda http=None, host=None: retrieve)
    monkeypatch.setattr(atlas, "connect", lambda path=None: real(path) if path else real(tmp_path / "atlas.db"))
    out = await tools_fin.fin_pnl("Wholesale")
    assert seen["ids"] == ["pnl", "fte", "fpa", "fpp", "units"] and seen["hook"] is not None
    assert out["area"] == "Wholesale" and out["workbook"] == "C:/x/a.xlsx" and "freshness" in out
    assert seen["spec"]["key_points"][0].startswith("EBITDA YTD Sep 26:")


def test_a_failed_job_is_a_visible_warning():
    from smartview_mcp.extract import FillResult
    from smartview_mcp.warehouse import Fill
    ok = FillResult(Fill("top_up", "x", 0.0, 0.1, 1, "ok", {}))
    bad = FillResult(Fill("new_month", "x", 0.0, 0.1, 0, "error", {"errors": ["FinMax: provider down"]}))
    assert tools_fin.stale_note([ok]) is None
    assert "(new_month: FinMax: provider down)" in tools_fin.stale_note([ok, bad])
```

In `engine\tests\test_smoke_stdio.py`, replace the two assertions on the tool names:

```python
    assert len(names) == 32, sorted(names)
    assert not any(n.startswith("sv_adhoc") or n == "sv_show" for n in names)
    assert {"sv_status", "sv_query", "sv_ws_build", "sv_drill", "xl_read", "xl_revert", "fin_pnl",
            "sv_warehouse"} <= names
```

- [ ] **Step 2: Run them and check they fail**

Run: `PY -m pytest tests/test_tools_fin.py tests/test_smoke_stdio.py -q`
Expected: FAIL, `ImportError: cannot import name 'tools_fin'` and `assert 30 == 32`.

- [ ] **Step 3: Implement**

`engine\smartview_mcp\tools_fin.py`:

```python
"""Finance tools on the warehouse and the calc engine: fin_pnl (the first slice) and sv_warehouse. Registered on the
shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

import contextlib
from typing import Literal

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError

from . import atlas, catalog, extract, fin_spec, qlog, refresh, templates
from . import fin_pnl as pnl                      # the tool below is called fin_pnl
from .host import HostError
from .server import WRITE, _call_user, _err, tool
from .timing import Timer, stage
from .warehouse import MemoryWarehouse, Warehouse, WarehouseUnavailable

Comparison = Literal["ly", "budget", "forecast"]


def resolve_area(cat: catalog.Catalog, area: str) -> catalog.Entity:
    """The one business area `area` names; a readable ToolError for none, several, or an abolished one."""
    found = catalog.resolve_entities(cat, area)
    known = ", ".join(e.label for e in cat.entities if not e.abolished)
    if not found:
        raise ToolError(f"'{area}' names no business area. Next step: call fin_pnl with one of: {known}.")
    if len(found) > 1:
        raise ToolError(f"'{area}' names several areas ({', '.join(e.label for e in found)}). Next step: ask the "
                        "user which one, then call fin_pnl once per area.")
    e = found[0]
    if e.abolished:
        succ = ", ".join(x.label for x in cat.entities if x.key in e.successors)
        raise ToolError(f"{e.label} was abolished on {e.abolished}: {e.note} Next step: call fin_pnl for one of "
                        f"its successors: {succ}.")
    return e


def stale_note(fills) -> str | None:
    """A visible warning when a refresh or top-up job failed: the answer then uses the earlier warehouse data."""
    failed = [r.fill for r in fills if r.fill.result == "error"]
    if not failed:
        return None
    first = (failed[0].detail.get("errors") or ["?"])[0]
    return (f"Warning: {len(failed)} warehouse job(s) failed this time ({failed[0].job}: {first[:140]}); figures "
            "that were not refreshed come from the earlier read shown in the data line.")


def _units(db, entity: catalog.Entity) -> tuple[tuple[str, str], ...]:
    """(member, label) of the area's sub-units from the atlas, for the sub-unit bridge."""
    labels = {}

    def kids(m):
        rows = atlas.children(db, "FinMax", m)
        labels.update({r["name"]: templates._clean(r["label"]) or r["name"] for r in rows})
        return [r["name"] for r in rows]
    names = refresh.sub_units(kids, entity.finmax) if entity.finmax else []
    return tuple((n, labels.get(n, n)) for n in names)


def prepare(h, cat: catalog.Catalog, entity: catalog.Entity, period: str | None, compare: tuple[str, ...],
            store=None, retrieve=None, db=None):
    """Refresh check + top-up + calc engine: (Result, store mode, fills). The warehouse when it can be opened,
    otherwise live pulls of this area into memory (slower, and the answer says so). db: the atlas (default: the
    user's)."""
    own = db is None
    db = atlas.connect() if own else db
    try:
        return _prepare(h, cat, entity, period, compare, store, retrieve, db)
    finally:
        if own:
            db.close()


def _prepare(h, cat, entity, period, compare, store, retrieve, db):
    with stage("closed month"):
        closed = templates.closed_month(h)
    p = period or closed
    retrieve = retrieve or extract.route(host=lambda: h)
    forecast = templates.latest_forecast(templates.split_period(p)[1])
    units = _units(db, entity)
    scope = pnl.Scope(entity, p, closed, forecast, units, compare)

    def run(st, fills):
        with stage("top up (live)"):
            fills += refresh.top_up(st, retrieve, cat, entity, p, closed, forecast, [u for u, _ in units])
        with stage("calc engine"):
            return pnl.compute(cat, scope, st), fills
    try:
        st = store if store is not None else Warehouse()
        with stage("warehouse refresh"):
            fills = refresh.refresh(st, retrieve, cat, closed, db, templates.latest_forecast)
        res, fills = run(st, fills)
        return res, "warehouse", fills
    except (WarehouseUnavailable, ImportError) as e:
        mem = MemoryWarehouse()
        with stage("live pulls (no warehouse)"):
            fills = [extract.fill(mem, retrieve, cat, extract.base_grids(cat, closed, templates.latest_forecast,
                                                                         keys={entity.key}), "live", closed)]
        res, fills = run(mem, fills)
        return res, f"live pulls (the warehouse was unavailable: {str(e)[:160]})", fills


@tool(annotations=WRITE)
async def fin_pnl(area: str, period: str | None = None, compare: list[Comparison] | None = None,
                  target: str = "session", allow_archive: bool = False, q: int | None = None,
                  ctx: Context | None = None) -> dict:
    """ONE call for 'P&L for <business area>': the month and YTD against last year (FinMax), budget and the latest
    forecast (FinPlan), materiality flags (catalogue rule), an EBITDA bridge by line and by sub-unit, KPIs (margins,
    cost per FTE) and the year-end landing. The numbers come from the local warehouse (refreshed automatically when
    a month closes or the atlas shows a restatement; the open month and the latest forecast are read live) and are
    delivered as an analysis in the session's workbook: live Smart View source sheets, every answer cell a formula,
    and checks that the warehouse equals each source sheet. area: e.g. 'Wholesale', 'Business', '2TEC';
    period: 'Sep 26' (default: the last closed month); compare: any of 'ly', 'budget', 'forecast' (default all).
    Write the reply from the returned key_points, flags, facts and freshness - never compute a number yourself."""
    cat = catalog.load()
    entity = resolve_area(cat, area)
    cmp = tuple(compare or pnl.COMPARISONS)

    def job(h):
        from .tools_ws import PullIn, _build_job
        tm = Timer.start()
        res, mode, fills = prepare(h, cat, entity, period, cmp)
        req = fin_spec.build(res, cat)
        if mode != "warehouse":
            req.spec["notes"] = [n.replace("Data: warehouse", f"Data: {mode}") for n in req.spec["notes"]]
        stale = stale_note(fills)
        if stale:
            req.spec["notes"].insert(0, stale)
        pulls = [PullIn.model_validate(x) for x in req.pulls]
        out = _build_job(h, tm, req.question, pulls, req.spec, target, allow_archive, q, replace=bool(q),
                         before_answer=lambda spec, by_id: fin_spec.with_reconciliation(spec, req.expected, by_id))
        if stale:
            out["warnings"] = [stale] + list(out.get("warnings") or [])
        return {**out, **pnl.summary(res, mode, fills)}
    try:
        out = await _call_user(job, what="fin_pnl")
    except (HostError, ToolError) as e:
        qlog.write({"tool": "fin_pnl", "area": area, "period": period, "error": str(e)})
        raise _err(e)
    qlog.write({"tool": "fin_pnl", "area": area, "period": out["period"], "workbook": out["workbook"],
                "q": out["question"], "checks": out["checks"], "timings": out["timings"]})
    return out


@tool(annotations=WRITE)
async def sv_warehouse(action: Literal["status", "refresh"] = "status") -> dict:
    """The local warehouse behind fin_pnl. status: file, rows, closed month, last fills and their reconciliation
    (cheap, no EPM). refresh: run the refresh check now (a base fill the first time, about a minute; afterwards
    only what changed), read-only against EPM."""
    if action == "status":
        try:
            return Warehouse().status()
        except (WarehouseUnavailable, ImportError) as e:
            raise ToolError(f"warehouse unavailable: {e}. Next step: retry in a minute (another smartass process "
                            "may be writing it); if it repeats, call sv_status.") from None
    cat = catalog.load()

    def job(h):
        Timer.start()
        with stage("closed month"):
            closed = templates.closed_month(h)
        with stage("warehouse refresh"), contextlib.closing(atlas.connect()) as db:
            fills = refresh.refresh(Warehouse(), extract.route(host=lambda: h), cat, closed, db,
                                    templates.latest_forecast)
        return {"closed_month": closed, "jobs": [{"job": r.fill.job, "result": r.fill.result, "cells": r.fill.cells,
                                                  "seconds": r.fill.seconds, **r.fill.detail} for r in fills]
                or "nothing to do: the warehouse is up to date"}
    try:
        return await _call_user(job, what="sv_warehouse refresh")
    except (HostError, ToolError, WarehouseUnavailable) as e:
        raise _err(e)
```

In `engine\smartview_mcp\server.py`, after `from . import tools_ws  # noqa: E402,F401`:

```python
from . import tools_fin  # noqa: E402,F401
```

- [ ] **Step 4: Run the tests again, then everything**

Run: `PY -m pytest tests/test_tools_fin.py tests/test_smoke_stdio.py -q`. Expected: all passed.
Run: `PY -m pytest -q` and `PY -m ruff check .`. Expected: all passed, `All checks passed!`

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/tools_fin.py engine/smartview_mcp/server.py engine/tests/test_tools_fin.py engine/tests/test_smoke_stdio.py
PY ..\scripts\check_staged.py
git commit -m "tools: fin_pnl (warehouse + calc engine + sv_ws_build) and sv_warehouse"
```

---

## Part F: ops, docs and live verification

### Task 20: doctor `warehouse` check

**Files:**
- Modify: `engine\smartview_mcp\doctor.py`:
  - `_warehouse()` before `_output()`;
  - `CHECKS`;
  - `_versions()` adds duckdb.
- Modify: `engine\tests\live\scenarios.py:738` (`DOCTOR_CHECKS` gains `"warehouse"`)
- Test: `engine\tests\test_doctor.py` (append)

- [ ] **Step 1: Write the failing test** (append)

```python
def test_warehouse_check(monkeypatch, tmp_path):
    from smartview_mcp import warehouse as wh
    monkeypatch.setattr(config, "WAREHOUSE_DB", tmp_path / "w.duckdb")
    c = doctor._warehouse()
    assert c.status == "OK" and "not filled yet" in c.detail             # a fresh install
    w = wh.Warehouse(tmp_path / "w.duckdb")
    w.put([wh.Row(("FinMax", "Actual", "", "E", "ac.C30", "Sep 26", "Periodic"), 1.0, "http", 1.0, "Sep 26")])
    assert doctor._warehouse().status == "WARN"                          # cells but no fill recorded
    import time
    w.set_state(closed_month="Sep 26")
    w.record_fill(wh.Fill("base", "x", time.time(), 1.0, 1, "ok", {}))
    c = doctor._warehouse()
    assert c.status == "OK" and "closed month Sep 26" in c.detail and "(base: ok)" in c.detail
    w.record_fill(wh.Fill("new_month", "x", time.time(), 1.0, 0, "error", {"errors": ["FinMax: provider down"]}))
    c = doctor._warehouse()
    assert c.status == "WARN" and "provider down" in c.detail and "sv_warehouse" in c.fix
```

- [ ] **Step 2: Run it and check it fails**

Run: `PY -m pytest tests/test_doctor.py -q`
Expected: FAIL, `AttributeError: module 'smartview_mcp.doctor' has no attribute '_warehouse'`

- [ ] **Step 3: Implement**

In `engine\smartview_mcp\doctor.py`, insert before `def _output() -> Check:`:

```python
def _warehouse(max_age_days: float = 35) -> Check:
    """The local warehouse behind fin_pnl: present, last fill, its result and reconciliation. A fresh install has
    none yet - that is fine: it fills on the first fin_pnl question."""
    from .warehouse import Warehouse, WarehouseUnavailable
    try:
        st = Warehouse().status()
    except ImportError as e:
        return Check("warehouse", "FAIL", f"DuckDB is not installed ({e})", "Re-run installer\\install.ps1.")
    except WarehouseUnavailable as e:
        return Check("warehouse", "WARN", str(e), "Another smartass process is using it; run the doctor again.")
    if not st["exists"]:
        return Check("warehouse", "OK", f"not filled yet ({st['path']}); it fills on the first fin_pnl question")
    last = st["fills"][0] if st["fills"] else None
    closed = st["state"].get("closed_month", "none")
    if last is None:
        return Check("warehouse", "WARN", f"{st['rows']} cells, no fill recorded",
                     "Call sv_warehouse(action='refresh').")
    when = time.strftime("%Y-%m-%d %H:%M", time.localtime(last["started_at"]))
    detail = f"{st['rows']} cells, closed month {closed}, last fill {when} ({last['job']}: {last['result']})"
    if last["result"] == "error":
        err = (last["detail"].get("errors") or ["?"])[0]
        return Check("warehouse", "WARN", f"{detail}; {err}",
                     "Call sv_warehouse(action='refresh'); answers use the older data and say so. " + BUNDLE_FIX)
    if last["result"] == "warn":
        prob = (last["detail"].get("problems") or ["?"])[0]
        return Check("warehouse", "WARN", f"{detail}; reconciliation: {prob}",
                     "Check the cube for that cell (sv_query); the warehouse holds what EPM returned.")
    if (time.time() - last["started_at"]) / 86400 > max_age_days:
        return Check("warehouse", "WARN", f"{detail}; older than {max_age_days:.0f} days",
                     "Call sv_warehouse(action='refresh').")
    return Check("warehouse", "OK", detail)
```

Replace the `CHECKS` line:

```python
CHECKS = (_config, _excel, _smartview, _epm, _drill_browser, _atlas, _catalog, _warehouse, _output, _http, _update)
```

In `_versions()`, add `"duckdb"` to the package tuple:
`for name in ("mcp", "pydantic", "pywin32", "psutil", "openpyxl", "playwright", "duckdb"):`

In `engine\tests\live\scenarios.py:738`:

```python
DOCTOR_CHECKS = {"config", "excel", "smartview", "epm", "drill_browser", "atlas", "catalog", "warehouse", "output",
                 "http_api", "update"}
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest tests/test_doctor.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/doctor.py engine/tests/test_doctor.py engine/tests/live/scenarios.py
PY ..\scripts\check_staged.py
git commit -m "doctor: warehouse check (present, last fill, reconciliation, age)"
```

### Task 21: instructions, `division_pnl` marked for retirement, docs

**Files:**
- Modify: `engine\smartview_mcp\server.py` (`INSTRUCTIONS`: one bullet before "- Analyses: ONE call")
- Modify: `engine\smartview_mcp\templates.py` (`division_pnl` docstring)
- Modify: `engine\smartview_mcp\tools_ws.py` (`sv_analyse` docstring, the `division_pnl` line)
- Modify: `CHANGELOG.md` (Unreleased), `skills\smartass\reference\operations.md`
- Test: `engine\tests\test_templates.py` (append)

- [ ] **Step 1: Write the failing test** (append to `engine\tests\test_templates.py`)

```python
def test_fin_pnl_is_announced_and_division_pnl_marked_for_retirement():
    from smartview_mcp import server, tools_ws
    assert "fin_pnl(area)" in server.INSTRUCTIONS
    assert "fin_pnl" in templates.division_pnl.__doc__ and "retired" in templates.division_pnl.__doc__
    assert "to be retired: prefer fin_pnl" in tools_ws.sv_analyse.__doc__
    assert "division_pnl" in templates.TEMPLATES                      # still works until the retirement
```

- [ ] **Step 2: Run it and check it fails**

Run: `PY -m pytest tests/test_templates.py -q -k retirement`
Expected: FAIL, `assert 'fin_pnl(area)' in ...`

- [ ] **Step 3: Implement**

In `engine\smartview_mcp\server.py`, insert this bullet into `INSTRUCTIONS` directly before
`- Analyses: ONE call - sv_ws_build(question, pulls, spec)`:

```text
- P&L for a business area (month and YTD vs last year, budget, latest forecast; materiality; EBITDA bridge by line
  and sub-unit; KPIs; year-end landing): ONE call - fin_pnl(area). It reads the local warehouse (refreshed by itself)
  and builds the workbook with live source sheets; write the reply from its key_points, flags, facts, freshness.
```

In `engine\smartview_mcp\templates.py`, add a docstring as the first statement of `division_pnl`:

```python
    """To be retired: fin_pnl (fin_pnl.py, Stage 2+3) answers the same question from the warehouse with budget,
    forecast, materiality, bridges, KPIs and landing. Kept until fin_pnl has covered it in the live matrix."""
```

In `engine\smartview_mcp\tools_ws.py`, the `sv_analyse` docstring, replace the `division_pnl:` line with:

```text
      division_pnl:   (to be retired: prefer fin_pnl) entity (e.g. 'un.SY.2TEC'), period ('Sep 26'; default last
                      closed), basis 'ytd'|'month',
```

(The next line, `compare 'ly'|None, drivers bool`, stays.)

In `CHANGELOG.md`, append under `## Unreleased`:

```markdown
- `fin_pnl(area)`: a business area's P&L in one call - month and YTD against last year (FinMax), budget and the
  latest forecast (FinPlan), materiality flags, an EBITDA bridge by line and by sub-unit, KPIs and the year-end
  landing. The numbers come from a local DuckDB warehouse (`%LOCALAPPDATA%\smartass\warehouse.duckdb`,
  `paths.warehouse`), refreshed by itself when a month closes or the atlas shows a restatement; the open month and
  the latest forecast are read live and every answer has a freshness stamp. The workbook keeps live Smart View
  source sheets, formulas only, and a "warehouse = source sheet" check per sheet.
- `sv_warehouse(action='status'|'refresh')`; `smartass doctor` has a `warehouse` check.
- The closed month is read over Smart View's HTTP protocol first (EnumMembers), with Smart View in Excel as the
  fallback.
- `sv_analyse` template `division_pnl` is to be retired in favour of `fin_pnl`.
- New dependency: `duckdb==2.0.0`.
```

In `skills\smartass\reference\operations.md`, add after the `sv_query reads over Smart View's HTTP protocol` bullet:

```markdown
- fin_pnl(area) answers "P&L for <business area>" from the local warehouse (DuckDB, `%LOCALAPPDATA%\smartass\
  warehouse.duckdb`): the first call fills it (about a minute), later calls refresh only what changed (a new closed
  month, an atlas restatement, a new forecast) and read the open month and the latest forecast live. If the warehouse
  is locked or missing it answers from live pulls and says so. `sv_warehouse(action='status')` shows the last fills
  and their reconciliation.
```

- [ ] **Step 4: Run the tests again**

Run: `PY -m pytest -q` and `PY -m ruff check .`. Expected: all passed (569 passed), `All checks passed!`

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/server.py engine/smartview_mcp/templates.py engine/smartview_mcp/tools_ws.py engine/tests/test_templates.py CHANGELOG.md skills/smartass/reference/operations.md
PY ..\scripts\check_staged.py
git commit -m "docs: fin_pnl and the warehouse; division_pnl marked for retirement"
```

### Task 22 (LIVE, controller only): closed month, base fill on the real cubes, engine = direct retrieve, timing

This task needs SSO to EPM (the drill's hidden Edge). If the user's live MCP server holds the drill browser profile,
close its browser first (it idles out after `SV_DRILL_IDLE_S`), or expect the COM fallback route in the output.

**Files:**
- Create: `engine\tests\live\warehouse_live.py`

- [ ] **Step 1: Write the live script**

```python
"""Live (controller only): the closed month over HTTP, a base fill of a scratch warehouse on the real cubes with its
reconciliation, fin_pnl's engine numbers for Wholesale against direct retrieves, and timings. Read-only towards EPM.
Prints no figures (only PASS/FAIL, counts and seconds).
Run from the worktree's engine dir: PY tests\\live\\warehouse_live.py"""
from __future__ import annotations

import contextlib
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from smartview_mcp import atlas, catalog, config, extract, fin_pnl, refresh, svhttp, templates  # noqa: E402
from smartview_mcp.analytics import periods  # noqa: E402
from smartview_mcp.grid import Axis  # noqa: E402
from smartview_mcp.warehouse import Warehouse  # noqa: E402


def same(a, b) -> bool:
    return a is not None and b is not None and abs(a - b) <= 1.0


def where(problem: str) -> str:
    """A reconciliation problem without its amount."""
    return problem.split(" is off by")[0].split(" differs from")[0]


def main() -> int:
    bad = 0
    http = svhttp.shared()
    t0 = time.monotonic()
    closed = http.closed_month()
    print(f"PASS  closed month over HTTP: {closed} ({time.monotonic() - t0:.1f} s)")

    path = config.HOME / "live" / "warehouse_check.duckdb"
    path.parent.mkdir(parents=True, exist_ok=True)
    for f in path.parent.glob("warehouse_check.duckdb*"):
        f.unlink()
    cat, store, retrieve = catalog.load(), Warehouse(path), extract.route()
    with contextlib.closing(atlas.connect()) as db:
        fills = refresh.refresh(store, retrieve, cat, closed, db, templates.latest_forecast)
        base = fills[0].fill
        ok = base.result == "ok"
        bad += not ok
        print(f"{'PASS' if ok else 'FAIL'}  base fill: {base.detail['grids']} grids, {base.cells} cells, "
              f"{base.seconds:.1f} s, routes {base.detail['routes']}, {base.detail['checks']} checks, "
              f"{len(base.detail['problems'])} problems, {len(base.detail['errors'])} errors")
        for p in base.detail["problems"][:5]:
            print("        problem at", where(p))
        for e in base.detail["errors"][:5]:
            print("        error:", e[:160])
        t0 = time.monotonic()
        again = refresh.refresh(store, retrieve, cat, closed, db, templates.latest_forecast)
        ok = again == []
        bad += not ok
        print(f"{'PASS' if ok else 'FAIL'}  warm refresh check: {len(again)} jobs ({time.monotonic() - t0:.2f} s)")

        whs = catalog.resolve_entity(cat, "Wholesale")
        kids = refresh.sub_units(lambda m: [r["name"] for r in atlas.children(db, "FinMax", m)], whs.finmax)
    fc = templates.latest_forecast(periods.split_period(closed)[1])
    t0 = time.monotonic()
    tops = refresh.top_up(store, retrieve, cat, whs, closed, closed, fc, kids)
    top_s = time.monotonic() - t0
    ok = all(r.ok for r in tops)
    bad += not ok
    print(f"{'PASS' if ok else 'FAIL'}  top-up: {sum(r.fill.cells for r in tops)} cells, {top_s:.1f} s")
    t0 = time.monotonic()
    res = fin_pnl.compute(cat, fin_pnl.Scope(whs, closed, closed, fc, tuple((k, k) for k in kids)), store)
    calc_ms = (time.monotonic() - t0) * 1000
    mon, fy = periods.finplan(closed)
    fm = {"All Entities": whs.finmax, "All Scenarios": "Actual", "All Currencies": "NOK"}
    fp = {"Entity": whs.finplan, "Years": fy, "Period": mon, "All Balance Types": "Accumulated"}
    cases = {
        "EBITDA YTD (FinMax)": ("fm.ebitda.ytd.cur", "FinMax", {**fm, "All Balance Types": "Accumulated"},
                                "All Accounts", "ac.EBITDA", "All Time Periods", closed),
        "EBITDA YTD last year": ("fm.ebitda.ytd.ly", "FinMax", {**fm, "All Balance Types": "Accumulated"},
                                 "All Accounts", "ac.EBITDA", "All Time Periods", res.ly),
        "Revenue, the month": ("fm.rev.m.cur", "FinMax", {**fm, "All Balance Types": "Periodic"},
                               "All Accounts", "ac.C30", "All Time Periods", closed),
        "FTE at month end": ("fm.fte.m.cur", "FinMax", {**fm, "All Balance Types": "Accumulated"},
                             "All Accounts", "ac.STA20Y", "All Time Periods", closed),
        "FinPlan actual EBITDA YTD": ("fp.ebitda.actual.ytd", "FinPlan", {**fp, "Version": "Working"},
                                      "Account", "EBITDA", "Scenario", "Actual"),
        "Target EBITDA YTD": ("fp.ebitda.budget.ytd", "FinPlan", {**fp, "Version": "Final"},
                              "Account", "EBITDA", "Scenario", "Target"),
    }
    if fc:
        cases["Forecast EBITDA YTD"] = ("fp.ebitda.forecast.ytd", "FinPlan", {**fp, "Version": "Final"},
                                        "Account", "EBITDA", "Scenario", fc)
    for name, (fid, conn, pov, rdim, rmem, cdim, cmem) in cases.items():
        _, p, _ = http.retrieve(conn, pov, [Axis(rdim, [rmem])], [Axis(cdim, [cmem])], False, False)
        got = p.rows[0][1] if p.rows else None
        got = 0.0 if got is None else got
        ok = isinstance(got, float) and same(res.facts[fid].value, got)
        bad += not ok
        print(f"{'PASS' if ok else 'FAIL'}  {name}: the engine equals a direct retrieve")
    ok = res.line_bridge.closes() and (res.unit_bridge is None or res.unit_bridge.closes())
    bad += not ok
    print(f"{'PASS' if ok else 'FAIL'}  bridges close ({len(kids)} sub-units)")
    print(f"timing: base fill {base.seconds:.1f} s, top-up {top_s:.1f} s, calc engine {calc_ms:.0f} ms")
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 2: Run it**

Run: `PY tests\live\warehouse_live.py`

Expected: only `PASS` lines, and:
- `closed month over HTTP: <the month>`;
- `base fill: … routes ['http'] …, 0 problems, 0 errors`;
- `warm refresh check: 0 jobs`;
- the engine equals a direct retrieve for the 6–7 cases;
- `bridges close`;
- a timing line; expected about 10–60 s for the base fill and under 100 ms for the calc engine.

If it fails:
- **A reconciliation problem:** check that cell with `sv_query` (or `smartass doctor`). The script prints only where
  the problem is, never an amount. An identity that really does not hold in EPM goes in the catalogue notes, not in
  the code.
- **The engine ≠ direct for a FinPlan YTD case:** FinPlan's Accumulated is then not the sum of its Periodic months
  for that line. Add a unit test that reproduces it with invented figures, then store and use FinPlan Accumulated
  for that case.
- **COM instead of HTTP:** it works, but slowly; record it.

- [ ] **Step 3: Record and commit**

Record the PASS/FAIL lines, cell counts and timings (**no figures**) in a new section "Stage 2+3 live check (Task
22)" at the end of `docs\specs\2026-10-25-stage2-3-warehouse-calc-design.md`, then:

```bash
git add engine/tests/live/warehouse_live.py docs/specs/2026-10-25-stage2-3-warehouse-calc-design.md
PY ..\scripts\check_staged.py
git commit -m "live: warehouse base fill reconciles and equals direct retrieves; timings"
```

### Task 23 (LIVE, controller only): matrix scenario `S11.fin_pnl.wholesale`

It runs the real tool through a fresh MCP server from the worktree (`tests\live\matrix.py` starts
`run_server.py` from the worktree's `engine`). It needs Excel and Smart View, like the S2 scenarios.

**Files:**
- Modify: `engine\tests\live\scenarios.py` (append the scenario after `s10_http`)
- Modify: `engine\tests\live\README.md` (one table row)

- [ ] **Step 1: Write the scenario** (append to `engine\tests\live\scenarios.py`)

```python
# ================================================================== S11 fin_pnl (Stage 2+3)
def _fin_query(conn: str, pov: dict, row: tuple[str, str], col: tuple[str, str]) -> dict:
    return _q([{"dimension": row[0], "members": [row[1]]}], [{"dimension": col[0], "members": [col[1]]}], pov, conn)


@scenario("S11.fin_pnl.wholesale", ["fin"], budget=240)
def s11_fin_pnl(ctx):
    out, secs = ctx.client.must("fin_pnl", {"area": "Wholesale", "target": "new"})
    _answer_ok(out, "fin_pnl")
    assert out["formula_style"].get("ok"), f"formula_style {out['formula_style']}"
    recon = {k: v for k, v in out["checks"].items() if k.startswith("Warehouse = source sheet")}
    assert len(recon) == 5 and _all_ok(recon), f"warehouse = sheet: {recon}"
    facts = {f["id"]: f["value"] for f in out["facts"]}
    p = out["period"]
    ly = f"{p[:3]} {int(p[4:]) - 1:02d}"
    fm = {"All Entities": "un.SY.2WHS", "All Scenarios": "Actual", "All Currencies": "NOK"}
    fp = {"Entity": "2WHS Wholesale", "Years": f"FY{p[4:]}", "Period": p[:3], "All Balance Types": "Accumulated"}
    cases = [("fm.ebitda.ytd.cur", "FinMax", {**fm, "All Balance Types": "Accumulated"},
              ("All Accounts", "ac.EBITDA"), ("All Time Periods", p)),
             ("fm.ebitda.ytd.ly", "FinMax", {**fm, "All Balance Types": "Accumulated"},
              ("All Accounts", "ac.EBITDA"), ("All Time Periods", ly)),
             ("fm.rev.m.cur", "FinMax", {**fm, "All Balance Types": "Periodic"},
              ("All Accounts", "ac.C30"), ("All Time Periods", p)),
             ("fp.ebitda.budget.ytd", "FinPlan", {**fp, "Version": "Final"}, ("Account", "EBITDA"),
              ("Scenario", "Target"))]
    bad = []
    for fid, conn, pov, row, col in cases:
        b, _ = ctx.client.must("sv_query", _fin_query(conn, pov, row, col))
        v = helpers.number(b["table"])
        if facts.get(fid) is None or not _close(v, facts[fid], 1.0):
            bad.append(fid)
    assert not bad, f"fin_pnl facts differ from sv_query: {bad}"
    again, warm = ctx.client.must("fin_pnl", {"area": "Wholesale", "target": out["workbook"]})
    _answer_ok(again, "fin_pnl warm")
    assert all(f["job"] == "top_up" for f in again["fills"]), f"warm call refilled: {again['fills']}"
    return {"notes": f"cold {secs:.0f}s, warm {warm:.0f}s; {len(out['checks'])} checks OK; "
                     f"{len(cases)} values = sv_query"}
```

In `engine\tests\live\README.md`, add a row to the Scenarios table after S10:

```markdown
| S11 | `fin` | `fin_pnl` Wholesale: workbook checks incl. warehouse = sheet, facts = `sv_query`, warm call refills nothing |
```

- [ ] **Step 2: Run it**

Run: `PY tests\live\matrix.py --only S11 --keep`

Expected:
- `S11.fin_pnl.wholesale PASS`, with notes like `cold …s, warm …s; N checks OK; 4 values = sv_query`;
- the warm call is clearly faster than the cold one; the build time is dominated by the five Smart View sheets in
  Excel.

If it fails:
- **A formula_style block:** fix the spec in `fin_spec.py` and add a unit test (Task 18's tests) for that formula
  first.
- **DIFF on a warehouse = sheet check:** the detail names the first cell; compare it with Task 22.

- [ ] **Step 3: Commit and push**

```bash
git add engine/tests/live/scenarios.py engine/tests/live/README.md
PY ..\scripts\check_staged.py
git commit -m "live: matrix scenario S11 fin_pnl Wholesale"
git push -u origin feat/stage2-3
```

Then hand over with superpowers:finishing-a-development-branch (merge or PR). Restart the MCP server after the
merge (`/mcp` reconnect) so `fin_pnl` is listed.

---

## Done when

- `PY -m pytest -q` is green (569 passed: the 491 existing + 78 new), `PY -m ruff check .` is clean, and
  `PY ..\scripts\check_staged.py` passes.
- `tests\live\warehouse_live.py` prints only PASS lines (Task 22), and the results are recorded in the design doc.
- `S11.fin_pnl.wholesale` passes in the live matrix (Task 23): every check is OK, including the five
  "warehouse = source sheet" checks, and the facts equal `sv_query`.
- `smartass doctor` shows `warehouse OK`.
- `templates.division_pnl` still works and is marked for retirement; delete it in a later plan, once `fin_pnl` has
  covered it in the matrix.

## Not in this plan (next plans)

- The shared warehouse; the policy for cube data on local disk (confirm with IT before colleagues get it).
- The balance sheet, capex, price/volume/mix bridges, and cost-centre anomalies (z-score).
- Merging the number formatters into `numfmt.py`, and rebuilding the other `sv_analyse` templates on the engine.
- Stage 4: `fin_plan` / `fin_run` / `fin_memo` and the memo validator. Stage 5: playbooks. Stage 6: the add-in.
