# Phase 6–7 "Retire + harden" Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task by task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** finish the rework. When this plan is done there is one pipeline (`sv_gate` → `sv_answer`, with follow-ups
and recipes) plus a few utilities. The old overlapping tools, the DuckDB warehouse and the Norwegian sheet-name
compatibility are deleted. The engine's Excel login is hardened (§3.2 of the design). The skill and the docs
describe the pipeline. The live bench route `core` and the gate stay green. A PR `rework/pipeline` → `main` is
opened.

**Architecture:**
- The answer spec v1 (`core/contract.py`) stays the only contract.
- Tools go from **35 to 25**: 10 are deleted and none are added.
- The pnl recipe reads its facts **live**, through an in-memory fact store. That is the path a sub-unit already
  takes today (`tools_fin._prepare.live`). Its reconciliation check becomes "Live read = source sheet", which
  `fin_spec.check_source` already emits for live modes.
- Helpers that tools_core and tools_drill borrow from `tools_ws` move to `core/reply.py`.
- `userexcel.py` becomes the package `smartview_mcp/excel/`. `livexl.py` becomes the package
  `smartview_mcp/workbench/`. `answer.py` is split in two.
- Login hardening adds three things:
  - hung-window detection (`IsHungAppWindow`);
  - a `ProviderEmpty` error that restarts **only the engine's hidden Excel** once;
  - progress notes to the user (`timing.say`), plus a fast message when the user's Excel is the one that is stuck.
- The local HTTP API gains `POST /gate` and `POST /answer`.

**Tech Stack:** Python 3.12, pydantic v2, pywin32 COM (Excel), Playwright (hidden Edge, svhttp), pytest + anyio,
ruff, uv. The live bench needs EPM and Excel.

**Run everything from `C:\dev\smartass\engine`** (PowerShell 5.1: no `&&`; use `;` and `if ($?) {...}`):
- unit tests: `uv run pytest -q -p no:cacheprovider <files>`;
- full suite: `uv run pytest -q -p no:cacheprovider` (about 75 s; **871 passing** before this plan, on HEAD
  `b0c142b`);
- lint: `uv run ruff check smartview_mcp tests`.

Commit after each task (from `C:\dev\smartass`):
`git add -A; git commit -m "<message>" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`.
Never edit files with `Get-Content | Set-Content` (it mangles UTF-8); use the edit/create tools, and `git mv` for
moves.

**Line count** (Python under `engine/smartview_mcp`, newlines counted by `git show`; run from `C:\dev\smartass`):

```powershell
@'
import subprocess, sys
def count(ref):
    out = subprocess.run(["git", "ls-tree", "-r", "--name-only", ref, "engine/smartview_mcp"],
                         capture_output=True, text=True, check=True).stdout.split()
    files = [f for f in out if f.endswith(".py")]
    n = sum(subprocess.run(["git", "show", f"{ref}:{f}"], capture_output=True, check=True).stdout.count(b"\n")
            for f in files)
    return n, len(files)
for ref in sys.argv[1:]:
    n, k = count(ref)
    print(f"{ref}: {n} lines, {k} files")
'@ | python - origin/main HEAD
```

Baseline, measured 2026-11-02:

| Ref | Lines | Files |
|---|---|---|
| `origin/main` (`22a2856`) | **16 718** | 63 |
| `HEAD` (`b0c142b`, rework/pipeline) | **18 884** | 86 |
| Owner target: ≥ 40 % fewer than main | **≤ 10 031** | – |
| Projected after this plan (D14) | about 16 400–16 700 | about 80 |

The 40 % target is **not reachable** with the tools the owner keeps; see D14. This plan's hard gate is
**≤ 16 718 lines (at or below main)**, and Task 16 records the actual numbers next to the owner's target.

---

## Decisions taken in this plan (owner may override)

| # | Decision | Why |
|---|---|---|
| D1 | **Tool verdicts** (35 → 25). **Keep**: `sv_gate`, `sv_remember`, `sv_answer`; `sv_status`, `sv_restart`; `sv_atlas_build`, `sv_atlas_status`, `sv_atlas_changes`; `sv_dimensions`, `sv_find_member`, `sv_members`, `sv_member_info`; `sv_drill`; and the 12 `xl_*` tools (`xl_workbooks`, `xl_overview`, `xl_read`, `xl_write`, `xl_rows`, `xl_add_sheet`, `xl_format`, `xl_format_sheet`, `xl_trace`, `xl_check`, `xl_changes`, `xl_revert`). **Delete**: `sv_ws_pull`, `sv_ws_build`, `sv_ws_answer`, `sv_analyse`, `sv_concept`, `fin_pnl`, `sv_warehouse`, `sv_scope`, `sv_query`, `sv_refresh_workbook`. | Kept: the pipeline (gate, defaults, answer), engine health and restart, the atlas behind find_member and the gate, member lookups for writing specs, drill (B12), and the workbench for edits the pipeline does not make. Deleted: each of them overlaps `sv_answer` (a pull, a free-form spec, a template, a recipe), the gate (`sv_scope`), or a deleted subsystem (`sv_warehouse`, `sv_refresh_workbook`; see D3 and D5). |
| D2 | **`sv_query` is deleted.** A quick chat lookup is a one-pull, one-line `sv_answer` spec (the gate makes it scoped). The bench oracle moves to an in-process svhttp fetch that runs **after** the MCP server has exited (D4). | Every caller is covered: the bench oracle (D4), the live matrix (pruned in Task 1), and chat lookups (`sv_answer`). Keeping it would also keep `tools_query._retrieve` and its route tests. `ops.query` stays: it is the COM fallback of `extract.route`, which the pnl recipe uses. |
| D3 | **`sv_refresh_workbook` is deleted** (with `ops.refresh_workbook`, `userexcel.refresh_pull`/`_refresh_pull` if nothing else calls them, and its tests). | New-pipeline source sheets are value grids written from HTTP; Smart View Refresh does not touch them, and a "refresh" is a follow-up that re-fetches (design §3.2). The tool only served users' own Smart View workbooks, which is outside the pipeline. The skill's description loses "refreshing a workbook's Smart View sheets". **Owner to confirm.** |
| D4 | **Bench oracle = in-process `svhttp.shared().fetch` after the server is gone.** `run_case` records the oracle requests and the numbers it found. After `client.close()`, `bench.oracle_pass` fetches each cell over HTTP in the bench process: a fresh hidden Edge on the same profile, whose lock is free once the server exits. Then it scores. The `baseline` route, `--route` and every old-tool function are removed. | Running in the same process while the server holds the drill profile would fail on `ProfileLock`. Running after it keeps the oracle independent: it cannot warm the route, and it uses a different process and browser. |
| D5 | **Warehouse deleted; the pnl recipe is live-only.** <br>• `fin_pnl.py` stays as the P&L engine (`compute`, `Scope`, `Result`, text helpers that `fin_spec` uses). <br>• Its tool-only functions are deleted (`summary`, `table`, `_cells`, `tool_warnings`, `landing_warnings`) if nothing else uses them. <br>• `tools_fin.py` is replaced by `pnl_facts.py`: `resolve_area`, `is_unit`, `resolve_scope`, `_entity_names`, `_finplan_of`, `stale_note`, `_units`, `_aliases`, `_forecast`, `check_period`, `prepare` (live only) and `top_up` (from `refresh.py`). <br>• The in-memory store (`Row`, `Fill`, `OPEN`, `WhKey`, `MemoryWarehouse` → renamed `FactStore`) moves to `factstore.py`. <br>• Deleted: `warehouse.py`, `refresh.py`, `doctor._warehouse`, `config.WAREHOUSE_DB`/`WAREHOUSE_LOCK_WAIT_S`, `duckdb` in `pyproject.toml`/`uv.lock`. <br>• B01/B15 keep the check "Live read = source sheet" (`layout.expect.source = "Live read"`). | The simplest option that keeps a reconciliation on B01/B15: the expectation cells come from the same `Result` facts, now read live. The live path already exists and is tested (sub-units, "no warehouse" fallback). Cost: B01/B15 now pull the area live on every call. Measure it in Task 17; B13 (a sub-unit) already takes this path. |
| D6 | **English-only sheet names:** `Q<n> Answer`, `Q<n> Lines` (Excel table `Lines_Q<n>[_alias]`), `Sources` (header A1 `Fetched`; a user's own unrelated `Sources` pushes ours to `Sources SV`, the same protection as today). New-file title `Analysis`. Every Norwegian status, note or warning string in code is translated. Old names (`Svar`, `Linjer`, `Kilder`, `Kilder SV`, `Hentet`) are no longer recognised, so `sheetorder`, `next_question` and `park_question_data` treat them as the user's sheets. | Owner decision 3. Kept as they are (not sheet names): the Norwegian **synonyms** the gate understands (`core/period.py`, `core/scope.py`, `concepts.py`, `core/semantic.py`), the TN cell-style names (`Hovedsum`, `Nøkkel` in `livexl.TN_EMPH` and styling) and the OneDrive folder name in config. |
| D7 | **`auditlib.py` moves to `engine/tests/live/auditlib.py`.** It is used only by the live matrix and by `test_matrix_logic.py`. `hsgetvalue.py` is deleted: form B was never wired, and `spike_r1_hsgetvalue.py` goes with it. | They are not engine code. The live matrix keeps the reliability scenarios (SSO expiry, dialogs, killed Excel), which exercise the login hardening. |
| D8 | **The live matrix is pruned, not deleted.** Kept: S0, S5 (drill), S6 (xl, with `sv_ws_pull` replaced by `sv_answer`), S8 (inject, with `sv_query` replaced by `helpers.tiny_answer`), S9 and the fresh-install scenarios. Removed: the S1 query scenarios, S2, S3, S4, S7, and the `fin_pnl` scenarios. `warehouse_live.py`, `svhttp_vs_com.py` and `spike_r1/r2/r3*.py` are deleted. | The bench covers answers; the matrix covers reliability. One harness per purpose. |
| D9 | **Module size ≤ ~400 lines for structurally edited modules.** <br>• `userexcel.py` (1 452) → package `excel/`: `jobs.py`, `login.py`, `instances.py`, `target.py`, `sheets.py`, `adhoc.py`. <br>• `livexl.py` (1 044) → package `workbench/`: `books.py`, `changes.py`, `fmt.py`. <br>• `answer.py` (605) → `answer.py` + `answer_write.py`. <br>• `host.py` (678) → `host.py` + `hostproc.py`; host keeps re-imported names so tests that monkeypatch `host.kill_ours`/`host.windows_of` keep working. <br>• `ops.py`: if it is above 420 lines after the deletions, the member-rejection bisection moves to `ops_reject.py`. <br>• Modules touched only by deletions or docstring edits are exempt: `styling.py`, `fin_spec.py`, `fin_pnl.py`, `drill_browser.py`, `catalog.py`, `vendor/*`. | Splitting is churn; it is done where the code is already being restructured or hardened. No compatibility re-export modules: callers and test monkeypatch targets are updated. |
| D10 | **Login hardening, hidden Excel:** a Login window that Windows reports as hung (`IsHungAppWindow`) is stuck at once (detected within about 5–7 s by the existing scan). `LOGIN_GRACE_S` drops from 45 to 30 s (SSO once took 18 s). An rc=752/57704 that survives the one in-place retry raises `ProviderEmpty`. `ExcelHost.call` then sends a progress note, restarts the hidden Excel once and retries; after a user-Excel job it only restarts and asks the caller to ask again. | Design §3.2: detect within seconds, tell the user, restart only our Excel. A transient 752 (B12 in phase 5) still clears on the in-place retry, which now tells the user what is happening. |
| D11 | **Login hardening, user's Excel (never restarted):** <br>• `excel.login.wait_login_cleared` sends a note after 5 s of an open Login ("waiting for the automatic SSO click"). <br>• A Login that is hung for 10 s raises a readable `HostError` at once instead of waiting 90 s. <br>• The ad hoc pull (`excel.adhoc`) says so on the first "Empty response" connect, and retries a transient retrieve once (today it fails at once). | The user learns within seconds what blocks the call, and nothing of theirs is closed. |
| D12 | **HTTP API mirrors the pipeline:** `POST /gate {question}` and `POST /answer {spec\|followup, q, target, allow_archive}` behind the same secret. They run through `tools_core.gate()` / `tools_core.answer()` (plain async functions that `sv_gate`/`sv_answer` now wrap), one request at a time. A second concurrent request gets 409, a body over 1 MB gets 413, bad JSON gets 400, and an `InputError` gets 400 with its message. | The future Excel add-in calls the same pipeline as the agent. Handlers are injectable (like `health`), so the tests need no Excel. |
| D13 | **Docs:** `SKILL.md` is rewritten around gate → answer → follow-ups/recipes → findings/commentary, with drill as a utility. `reference/operations.md` lists errors of the surviving tools only. README, DESIGN, CHANGELOG, troubleshooting and the server `INSTRUCTIONS` are updated. Specs, plans and audits dated before 2026-11-02 move to `docs/archive/{specs,plans,audit}/` (with `git mv`), and links are fixed. | Phase 7 of the design. |
| D14 | **Line target: the owner must choose.** Measured: main 16 718, HEAD 18 884; −40 % of main = 10 031. This plan deletes about 2 300–2 500 lines and adds about 300 (login, HTTP), so it lands about 16 400–16 700. That is at or below main, and about −12 % against HEAD. To reach 10 031 another ~6 400 lines would have to go, which means dropping kept capabilities. The largest are: drill (`drill_browser`, `drill_xl`, `drill_lines`, `drill_store`, `drill_split`, `tools_drill`: about 1 750); the xl workbench (`livexl` + `tools_xl`: about 1 180); the P&L engine (`fin_pnl`, `fin_spec`, `analytics/`, `extract`: about 2 300, replaceable by a formula-only pnl recipe without materiality flags, bridge or landing); `vendor/` (about 1 030). | The plan does not cut capabilities the owner listed as kept. Options for the owner: (a) accept "≤ main" as the target; (b) also replace the P&L engine with a formula-only recipe (about −2 000; loses B15 detail); (c) exclude `vendor/` from the count. |

---

## File structure

| File | Action | Responsibility |
|---|---|---|
| `engine/tests/live/bench.py` | modify | core route only; deferred svhttp oracle (`oracle_pass`) |
| `engine/tests/live/helpers.py` | modify | `oracle_value`, `tiny_answer`; `number()` removed; imports local `auditlib` |
| `engine/tests/live/auditlib.py` | move (from `smartview_mcp/auditlib.py`) | matrix verdicts, engine-Excel detection |
| `engine/tests/live/scenarios.py`, `matrix.py`, `README.md` | modify | pruned scenarios (D8) |
| `engine/tests/live/warehouse_live.py`, `svhttp_vs_com.py`, `spike_r1_hsgetvalue.py`, `spike_r2_svhttp.py`, `spike_r3_livewrite.py` | delete | old tools / finished spikes |
| `engine/smartview_mcp/core/reply.py` | create | `finish`, `result`, `tidy`, `answer_markdown` (from `tools_ws`) |
| `engine/smartview_mcp/tools_ws.py` | delete | `sv_ws_pull/build/answer`, `sv_analyse`, `sv_concept` |
| `engine/smartview_mcp/tools_query.py` | modify | keep `sv_dimensions`, `sv_find_member`, `sv_members`, `sv_member_info` |
| `engine/smartview_mcp/ops.py` | modify | dead code out; `ProviderEmpty` after a failed transient retry |
| `engine/smartview_mcp/pnl_facts.py` | create | live P&L facts for recipe pnl (from `tools_fin` + `refresh.top_up`) |
| `engine/smartview_mcp/factstore.py` | create | in-memory fact store (`FactStore`, `Row`, `Fill`, `OPEN`, `WhKey`) |
| `engine/smartview_mcp/tools_fin.py`, `warehouse.py`, `refresh.py`, `hsgetvalue.py` | delete | retired |
| `engine/smartview_mcp/extract.py`, `fin_pnl.py`, `doctor.py`, `config.py`, `core/recipes/pnl.py`, `tools_core.py` | modify | no warehouse |
| `engine/pyproject.toml`, `engine/uv.lock` | modify | drop `duckdb` |
| `engine/smartview_mcp/sheetorder.py`, `answer.py`, `drill_xl.py`, `tools_drill.py` | modify | English sheet names (D6) |
| `engine/smartview_mcp/excel/{__init__,jobs,login,instances,target,sheets,adhoc}.py` | create (split of `userexcel.py`) | user-Excel jobs, login, instances, target, sheets/meta, ad hoc pull |
| `engine/smartview_mcp/workbench/{__init__,books,changes,fmt}.py` | create (split of `livexl.py`) | xl_* back end |
| `engine/smartview_mcp/answer_write.py` | create (from `answer.py`) | cell writing, named ranges, charts |
| `engine/smartview_mcp/hostproc.py` | create (from `host.py`) | process registry, own-Excel detection, `ComThread` |
| `engine/smartview_mcp/dialogs.py`, `host.py`, `timing.py`, `config.py` | modify | hung detection, `ProviderEmpty`, `say()` |
| `engine/smartview_mcp/httpapi.py` | modify | `POST /gate`, `POST /answer` |
| `engine/smartview_mcp/server.py` | modify | tool imports, `INSTRUCTIONS` |
| `engine/tests/test_*.py` | modify/delete | tests follow their code (per task) |
| `engine/tests/test_smoke_stdio.py` | modify | tool count 35 → 30 → 27 → 25 |
| `skills/smartass/SKILL.md`, `skills/smartass/reference/{operations,recipes,drill,cubes}.md` | modify | pipeline skill |
| `README.md`, `docs/DESIGN.md`, `CHANGELOG.md`, `docs/troubleshooting.md` | modify | phase 7 docs |
| `docs/archive/{specs,plans,audit}/` | create (moves) | pre-2026-11-02 docs |
| `docs/audit/2026-11-02-bench-core-final.{md,json}` | create (by the bench) | final live run |

## Task 1: Bench oracle over in-process svhttp; old routes and the old live harness out

**Files:**
- Modify: `engine/tests/live/bench.py`, `engine/tests/live/helpers.py`, `engine/tests/live/scenarios.py`,
  `engine/tests/live/matrix.py`, `engine/tests/live/README.md`, `engine/tests/test_matrix_logic.py`
- Move: `engine/smartview_mcp/auditlib.py` → `engine/tests/live/auditlib.py`
- Delete: `engine/tests/live/warehouse_live.py`, `svhttp_vs_com.py`, `spike_r1_hsgetvalue.py`,
  `spike_r2_svhttp.py`, `spike_r3_livewrite.py`

- [ ] **Step 1: Write the failing test** for the pure oracle helper in `engine/tests/test_matrix_logic.py`. It
  already puts `tests/live` on `sys.path`. Append:

```python
def test_oracle_cell_reads_the_one_data_cell_of_a_provider_grid():
    import helpers
    from smartview_mcp.grid import Axis, Layout

    lay = Layout({"Entity": "un.SY"}, [Axis("Account", ["ac.EBITDA"])], [Axis("Period", ["Sep 26"])])
    m = lay.matrix()
    values = [list(r) for r in m]
    values[1 + lay.nc][lay.nr] = 1234.5
    assert helpers.oracle_cell(lay, values) == 1234.5
    values[1 + lay.nc][lay.nr] = "#Missing"
    assert helpers.oracle_cell(lay, values) is None
```

  Also change the auditlib import at the top of that file from
  `from smartview_mcp.auditlib import error_class, ...` to a local import after the `sys.path` insert:
  `from auditlib import error_class, is_engine_excel, merge_before, poll_kill, stale_suspended, verdict`. Move the
  `sys.path.insert(0, os.path.join(os.path.dirname(__file__), "live"))` line to module level, above that import.

- [ ] **Step 2: Run it to verify it fails**

Run: `uv run pytest -q -p no:cacheprovider tests\test_matrix_logic.py`
Expected: an ImportError on `from auditlib import ...` (the module is not moved yet). After Step 3's move, the
new test fails with `AttributeError: module 'helpers' has no attribute 'oracle_cell'`.

- [ ] **Step 3: Move auditlib; add the oracle helpers.** Run from `C:\dev\smartass`:
  `git mv engine\smartview_mcp\auditlib.py engine\tests\live\auditlib.py`.
  In `tests/live/helpers.py`:
  - replace `from smartview_mcp.auditlib import is_engine_excel, poll_kill, stale_suspended` with
    `from auditlib import is_engine_excel, poll_kill, stale_suspended`;
  - delete `number()` (it parsed `sv_query` markdown);
  - add:

```python
def oracle_cell(lay, values: list[list]) -> float | None:
    """The single data cell of a one-row, one-column provider grid (row 1 = POV, then the column headers, then the
    data rows; the data starts after the row-label columns). None for #Missing/#NoAccess or a blank."""
    v = values[1 + lay.nc][lay.nr]
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def oracle_value(connection: str, pov: dict, row: tuple[str, str], col: tuple[str, str]) -> float | None:
    """One cube cell over the Smart View HTTP provider, in THIS process (bench.oracle_pass runs it after the MCP
    server exited, so its hidden Edge is free to start on the drill profile)."""
    from smartview_mcp import svhttp
    from smartview_mcp.grid import Axis
    f = svhttp.shared().fetch(connection, pov, [Axis(row[0], [row[1]])], [Axis(col[0], [col[1]])])
    return oracle_cell(f.layout, f.values)


TINY_SPEC = {"version": 1, "question": "LIVE CHECK: EBITDA un.SY Sep 26",
             "gate": {"outcome": "clear", "checklist": {
                 "period": "Sep 26", "scenario": "Actual", "entity": "un.SY", "measure": "EBITDA",
                 "currency_sign": "NOK, EPM signs", "granularity": "one cell"}},
             "pulls": [{"id": "p", "connection": "FinMax",
                        "rows": [{"dimension": "All Accounts", "members": ["ac.EBITDA"]}],
                        "columns": [{"dimension": "All Time Periods", "members": ["Sep 26"]}],
                        "pov": {"All Entities": "un.SY", "All Scenarios": "Actual", "All Currencies": "NOK",
                                "All Balance Types": "Accumulated"}}],
             "layout": {"title": "EBITDA", "columns": [{"header": "Sep 26", "kind": "number"}],
                        "sections": [{"title": "EBITDA", "lines": [
                            {"key": "ebitda", "label": "EBITDA", "cells": ["=p!ac.EBITDA|Sep 26"]}]}]}}


def tiny_answer(client, target: str = "new", timeout: float = 600):
    """The smallest pipeline call (one pull, one line) - what the matrix used sv_query for."""
    return client.call("sv_answer", {"spec": TINY_SPEC, "target": target}, timeout=timeout)
```

  Before writing `TINY_SPEC`, check the pull-cell syntax against the B02 spec that bench builds
  (`scenarios._bus_spec()`), and copy that cell form exactly. The spec must pass `pipeline.parse_spec`.
  Verify with:

```powershell
uv run python -c "import sys; sys.path.insert(0, r'tests\live'); import helpers; from smartview_mcp.core import pipeline; pipeline.parse_spec(helpers.TINY_SPEC); print('ok')"
```

  Expected: `ok`.

- [ ] **Step 4: Rewrite the bench to route core only, with a deferred oracle.** In `tests/live/bench.py`:
  - Docstring: one route (the pipeline). The oracle is "an independent single-cell HTTP fetch in the bench process,
    run after the MCP server exited". Usage: `uv run python tests\live\bench.py [--only B01,B02] [--keep]
    [--out NAME]`, with `--out` defaulting to `bench-core`.
  - `cell(conn, pov, row, col)` returns `{"connection": conn, "pov": pov, "row": row, "col": col}`.
  - `Case`: drop `baseline` and `core`; add `run: Callable | None = None` (fn(ctx) → Run dict). `case(...)`
    stores the decorated function as `run`. Delete `core_route`.
  - Delete the baseline functions and helpers: `b01`, `b02` (the `sv_ws_build` one), `b03`, `b04` (`sv_ws_answer`),
    `_analyse`, `b09` (`sv_concept`), `b13` and `b15` (`fin_pnl`), and every `CASES[...].core = ...` line.
    - Each core function becomes the `case` itself: `case("B01", ..., oracle=...)(_core("B01", "pnl", {...}))`, the
      same pattern B14 uses today.
    - B12 keeps `b12` (`sv_drill`) as its `run`.
  - Keep `S._bus_pulls`/`S._bus_spec`, `S.CELL_12881`, `S.FP_POV`, `S.values`, `S._all_ok`,
    `S._no_formula_errors` and `S.Ctx`. Where bench uses the oracle query helper `S._q`, use `cell(...)` instead.
  - `run_case(c, ctx)` (no `route`):
    - agent cases always run `sv_gate`;
    - `c.run is None` → `UNSUPPORTED`;
    - instead of the `sv_query` loop, store `r["_oracle_q"] = c.oracle(ctx, run) if c.oracle else {}` and
      `r["_nums"] = nums`;
    - set a provisional `r["result"] = verdict(r)`. The oracle is not scored yet, and `all({}.values())` is True, so
      `needs` chains such as B03 → B02 still work.
  - Add:

```python
def oracle_pass(results: list[dict]) -> None:
    """Score the oracle cells after the MCP server has exited (its hidden Edge freed the drill profile): one
    in-process HTTP fetch per cell, so the oracle never warms the route it checks. Re-scores every result."""
    from smartview_mcp import drill_browser
    try:
        for r in results:
            nums = r.pop("_nums", None) or []
            for label, q in (r.pop("_oracle_q", None) or {}).items():
                try:
                    v = helpers.oracle_value(**q)
                except Exception as e:  # noqa: BLE001 - one bad oracle cell must not stop the scoring
                    r["oracle"][label], r["match"][label] = False, f"oracle failed: {str(e)[:120]}"
                    continue
                how = found(v, nums) if v is not None else None
                r["oracle"][label] = how is not None
                r["match"][label] = how or ("no oracle value" if v is None else "missing")
            if r["result"] in ("PASS", "FAIL"):
                r["result"] = verdict(r)
    finally:
        drill_browser.shutdown_shared()
```

  - `main`:
    - drop `--route`;
    - after the `finally` that closes the client, call `oracle_pass(results)`, then write the report;
    - print the per-case line after `oracle_pass`, not inside the loop. Inside the loop, print
      `f"ran {c.id} ({r['result']} before oracle)"`.
    - `write_report(name, started, results)` keeps the `route` key as the constant `"core"`, so the JSON shape
      matches the phase 5 reports.

- [ ] **Step 5: Prune the live matrix** (D8). In `tests/live/scenarios.py`:
  - delete the scenarios that call `sv_query` for lookups (the S1 query/input ones), S2 (`sv_ws_build`,
    `sv_ws_answer`), S3 (`sv_analyse`), S4 (`sv_scope`, `sv_concept`), S7 (`sv_refresh_workbook`), and the S1x
    scenarios that call `fin_pnl`;
  - in S6, replace `sv_ws_pull` with `helpers.tiny_answer(ctx.client)`, and use its `workbook`/`answer_sheet` for
    the xl_* steps;
  - in the S8 inject scenarios, replace every `client.call("sv_query", ...)` with `helpers.tiny_answer(client)`,
    keeping the same expectations: success after recovery, or the named error class;
  - in S5, replace any `sv_ws_answer` call with `sv_answer` (followup), or drop that step if it only checked the old
    answer tool;
  - delete helpers left unused (`_ebitda_q`, `WHS_POV` if unused, ...);
  - in `matrix.py` and `README.md`, update the tag table (S1 lookup keeps `sv_find_member` only).
  
  Delete the files: from `C:\dev\smartass`, run
  `git rm engine\tests\live\warehouse_live.py engine\tests\live\svhttp_vs_com.py engine\tests\live\spike_r1_hsgetvalue.py engine\tests\live\spike_r2_svhttp.py engine\tests\live\spike_r3_livewrite.py`.

- [ ] **Step 6: Verify nothing references the removed pieces**

Run (from `engine`):
`Select-String tests\live\*.py -Pattern '"sv_query"|"sv_ws_|"sv_analyse"|"sv_concept"|"fin_pnl"|"sv_scope"|"sv_refresh_workbook"|"sv_warehouse"|smartview_mcp.auditlib|baseline'`
Expected: no output.
Run: `uv run python -c "import sys; sys.path.insert(0, r'tests\live'); import bench, matrix; print(len(bench.CASES))"`
Expected: the same case count as before (B01–B15 plus T01–T10), with no import errors.

- [ ] **Step 7: Run the tests and lint**

Run: `uv run pytest -q -p no:cacheprovider tests\test_matrix_logic.py tests\test_live_helpers.py; uv run ruff check smartview_mcp tests`
Expected: all pass; `All checks passed!`. Then run the full suite. Expected: 872 passed (871 + 1 new), 0 failed.

- [ ] **Step 8: Commit**

`git add -A; git commit -m "test(bench): svhttp oracle after the server exits; core route only; prune live matrix" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 2: Move the reply helpers out of tools_ws

**Files:**
- Create: `engine/smartview_mcp/core/reply.py`
- Modify: `engine/smartview_mcp/tools_core.py`, `engine/smartview_mcp/tools_drill.py`, `engine/smartview_mcp/tools_ws.py`
  (temporarily: it imports the moved helpers back)
- Test: `engine/tests/test_core_skeleton.py`

- [ ] **Step 1: List what tools_core and tools_drill use from tools_ws**

Run: `Select-String smartview_mcp\tools_core.py,smartview_mcp\tools_drill.py -Pattern 'tools_ws\.\w+|ws\.\w+' -AllMatches | % { $_.Matches.Value } | Sort-Object -Unique`
Expected: `tools_ws._finish`, `tools_ws._result`, and in tools_drill the `ws.*` names it uses (e.g. `ws._finish`,
`ws._result`, `ws._tidy`, `ws._answer_markdown`, `ws.PullIn`). Every name listed moves in Step 3, together with the
private helpers those functions call (follow them in `tools_ws.py`).

- [ ] **Step 2: Write the failing test** in `engine/tests/test_core_skeleton.py`:

```python
def test_reply_helpers_live_in_core_reply():
    from smartview_mcp.core import reply
    assert callable(reply.finish) and callable(reply.result)
```

Run: `uv run pytest -q -p no:cacheprovider tests\test_core_skeleton.py -k reply`
Expected: FAIL with `ImportError: cannot import name 'reply'`.

- [ ] **Step 3: Create `core/reply.py`.** Move the functions found in Step 1 verbatim, renaming them public:
  `_finish` → `finish`, `_result` → `result`, `_tidy` → `tidy`, `_answer_markdown` → `answer_markdown`. Bring the
  imports they need. The module docstring is: `"""The tool reply of an answer: save + activate the workbook, and
  the result dict (table, checks, warnings, sheet, workbook) that sv_answer and sv_drill return."""`.
  In `tools_ws.py`, replace the moved bodies with `from .core.reply import finish as _finish, result as _result, ...`
  so it keeps working until Task 3 deletes it. In `tools_core.py`, use `from .core import reply`, then
  `reply.finish(...)` and `reply.result(...)`, and drop `tools_ws` from the imports. In `tools_drill.py`, replace
  `from . import tools_ws as ws` with `from .core import reply` and call the helpers through `reply`. `PullIn`, if
  tools_drill uses it, moves to `core/reply.py` too.

- [ ] **Step 4: Run the tests**

Run: `uv run pytest -q -p no:cacheprovider; uv run ruff check smartview_mcp tests`
Expected: 873 passed, 0 failed; `All checks passed!`. Tests that monkeypatch `tools_ws._finish` for sv_answer or
sv_drill paths now patch `reply.finish`: grep with `Select-String tests\*.py -Pattern 'tools_ws\._finish|tools_ws\._result'`
and update every hit that tests tools_core or tools_drill.

- [ ] **Step 5: Commit**

`git add -A; git commit -m "refactor: answer reply helpers move to core/reply.py" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 3: Delete the tools_ws group (sv_ws_pull, sv_ws_build, sv_ws_answer, sv_analyse, sv_concept)

**Files:**
- Delete: `engine/smartview_mcp/tools_ws.py`
- Modify: `engine/smartview_mcp/server.py` (import list, `INSTRUCTIONS` mentions; the full rewrite is Task 14),
  `engine/smartview_mcp/tools_core.py` (the `_followup_spec` message), `engine/tests/test_smoke_stdio.py`
- Delete or modify tests: `test_build_hook.py`, `test_logic.py`, `test_templates.py`, `test_followup.py`,
  `test_recipes.py`, `test_core_skeleton.py` (only the tests that call the deleted tools)

- [ ] **Step 1: Update the smoke test first (it fails until the tools are gone)**

In `tests/test_smoke_stdio.py`, change `assert len(names) == 35` to `assert len(names) == 30`, and add
`assert not {"sv_ws_pull", "sv_ws_build", "sv_ws_answer", "sv_analyse", "sv_concept"} & set(names)`.
Run: `uv run pytest -q -p no:cacheprovider tests\test_smoke_stdio.py`
Expected: FAIL (`35 == 30`).

- [ ] **Step 2: Delete the module and its registration.** From `C:\dev\smartass`:
  `git rm engine\smartview_mcp\tools_ws.py`. In `server.py`, delete `from . import tools_ws  # noqa: E402,F401`. In
  `tools_core._followup_spec`, end the "not drawn by sv_answer" message at "ask it as a new question with spec=..."
  (remove "or use sv_ws_answer ...").

- [ ] **Step 3: Find what still imports it**

Run: `Select-String smartview_mcp\*.py,smartview_mcp\core\*.py,smartview_mcp\core\recipes\*.py,tests\*.py -Pattern 'tools_ws|sv_ws_|sv_analyse|sv_concept|_build_job|PullIn'`
Expected hits:
- `tools_fin.py` (`from .tools_ws import PullIn, _build_job`, used by `fin_pnl`; Task 5 deletes it). For now,
  change the import to a local import inside the `fin_pnl` tool function. If `_build_job` is only defined in
  tools_ws, delete the `fin_pnl` and `sv_warehouse` tools now instead: move Task 5 Step 1 (smoke count 25) forward
  and do the whole of Task 5 right after this task.
- Tests that call the deleted tools: delete those test functions, and delete whole files whose every test targets
  them (`test_build_hook.py` if it only drives `_build_job`). Tests of kept modules that merely used a tools_ws
  helper now import `core.reply`.
- `server.py`'s `INSTRUCTIONS` text: replace each deleted tool name with `sv_answer` (Task 14 rewrites the text).

Re-run the grep. Expected: no hits outside `tools_fin.py` (and its tests, which Task 5 deletes).

- [ ] **Step 4: Delete the now-dead helpers that only these tools used.** Check each name with
  `Select-String smartview_mcp\*.py,smartview_mcp\core\*.py,smartview_mcp\core\recipes\*.py -Pattern '\b<name>\b'`.
  Delete the name if the only hit is its definition. Candidates:
  - `templates.check_pulls`, `templates.pnl_lines`, `templates.pnl_checks`;
  - `specs.merge` (keep it if `core/contract.patched` uses it; D2 of widen says it does);
  - `userexcel.park_question_data`, `userexcel.load_spec`, `userexcel.excel_turn`, `userexcel._q`;
  - `grid.to_markdown` (keep it while `tools_query.sv_query` uses it; Task 4 deletes it);
  - `concepts.*` functions not reached from `core/recipes/concept.py`.
  
  Delete their unit tests too.

- [ ] **Step 5: Run the tests**

Run: `uv run pytest -q -p no:cacheprovider; uv run ruff check smartview_mcp tests`
Expected: all pass, 0 failed (the count drops by the deleted tests). Ruff reports no unused imports.

- [ ] **Step 6: Commit**

`git add -A; git commit -m "refactor!: delete sv_ws_pull, sv_ws_build, sv_ws_answer, sv_analyse, sv_concept" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 4: Delete sv_query, sv_scope, sv_refresh_workbook

**Files:**
- Modify: `engine/smartview_mcp/tools_query.py`, `engine/smartview_mcp/ops.py`, `engine/smartview_mcp/userexcel.py`,
  `engine/smartview_mcp/grid.py`, `engine/smartview_mcp/server.py`, `engine/tests/test_smoke_stdio.py`
- Delete or modify tests: `test_retrieve_route.py`, `test_errors.py`, `test_refresh.py` (the refresh_workbook
  part), `test_logic.py`

- [ ] **Step 1: Smoke count first**

In `tests/test_smoke_stdio.py`: `assert len(names) == 27`, and add
`assert not {"sv_query", "sv_scope", "sv_refresh_workbook"} & set(names)`.
Run: `uv run pytest -q -p no:cacheprovider tests\test_smoke_stdio.py`. Expected: FAIL (`30 == 27`).

- [ ] **Step 2: Delete the three tools from `tools_query.py`**: `sv_query`, its `_retrieve` helper, `sv_scope`,
  `sv_refresh_workbook`, and their imports. Keep `sv_dimensions`, `sv_find_member`, `sv_members`, `sv_member_info`.
  Update the module docstring to say "member lookups for writing answer specs".

- [ ] **Step 3: Delete the code only they reached.** Use the same grep rule as Task 3 Step 4 for each name.
  Candidates:
  - `ops.refresh_workbook`, `ops.save_table`, `ops.live_confirm`, `ops.canonical_pov`, `ops.check_members`,
    `ops._missing_live`, `ops._atlas_top`, `ops._explain_status`;
  - `userexcel.refresh_pull`, `userexcel._refresh_pull`;
  - `grid.to_markdown`;
  - `core/scope.to_contract` (`sv_scope`'s output, if only it used it), `core/semantic.problems`,
    `catalog.resolve_entity`.
  
  `ops.query`, `ops._query_once`, `ops.is_transient`, `ops.is_rc1000` and `ops._rejected_members` **stay**: they are
  `extract.route`'s COM fallback.

- [ ] **Step 4: Tests.** Delete `test_retrieve_route.py` tests that call `tools_query._retrieve`/`sv_query`, and
  keep any that test `extract.route`. Delete the `sv_query` and `sv_refresh_workbook` cases in `test_errors.py`,
  and the `refresh_workbook` tests in `test_refresh.py`. Run:
  `Select-String tests\*.py -Pattern 'sv_query|sv_scope|refresh_workbook|_retrieve\b|to_markdown'`
  Expected: only `what="sv_query"` strings in `test_host_dialogs.py` (a label passed to `call`; change them to
  `what="sv_answer"`).

- [ ] **Step 5: Run the tests**

Run: `uv run pytest -q -p no:cacheprovider; uv run ruff check smartview_mcp tests`
Expected: all pass, 0 failed; `All checks passed!`.

- [ ] **Step 6: Commit**

`git add -A; git commit -m "refactor!: delete sv_query, sv_scope, sv_refresh_workbook" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 5: Warehouse out; the pnl recipe reads live (fin_pnl, sv_warehouse deleted)

**Files:**
- Create: `engine/smartview_mcp/factstore.py`, `engine/smartview_mcp/pnl_facts.py`, `engine/tests/test_factstore.py`,
  `engine/tests/test_pnl_facts.py`
- Modify: `engine/smartview_mcp/extract.py`, `engine/smartview_mcp/fin_pnl.py`, `engine/smartview_mcp/doctor.py`,
  `engine/smartview_mcp/config.py`, `engine/smartview_mcp/core/recipes/pnl.py`, `engine/smartview_mcp/tools_core.py`,
  `engine/smartview_mcp/server.py`, `engine/pyproject.toml`, `engine/uv.lock`, `engine/tests/test_smoke_stdio.py`,
  `engine/tests/fin_util.py`
- Delete: `engine/smartview_mcp/tools_fin.py`, `warehouse.py`, `refresh.py`; tests `test_warehouse.py`,
  `test_refresh.py`, `test_tools_fin.py` (after their kept cases move, Step 4)

- [ ] **Step 1: Smoke count first.** In `tests/test_smoke_stdio.py`: `assert len(names) == 25`, and add
  `assert not {"fin_pnl", "sv_warehouse"} & set(names)`. Run it. Expected: FAIL (`27 == 25`).

- [ ] **Step 2: Write the failing tests for the moved code.** `tests/test_factstore.py`:

```python
from smartview_mcp.factstore import OPEN, FactStore, Row


def test_factstore_upsert_and_read_back():
    s = FactStore()
    r = Row(cube="FinMax", scenario="Actual", version="Final", entity="un.SY", account="ac.EBITDA",
            period="Sep 26", value=10.0)
    s.upsert([r])
    assert s.read(cube="FinMax", entity="un.SY", periods=["Sep 26"])[0].value == 10.0
    assert OPEN
```

  Before writing it, open `warehouse.py` and copy the real `Row` fields and the `MemoryWarehouse` method names; the
  test above names them as in `tests/fin_util.py`. Use the names the code actually has.
  `tests/test_pnl_facts.py`: move the cases of `test_tools_fin.py` that test `resolve_area`, `is_unit`,
  `resolve_scope`, `check_period`, `stale_note` and the **live** branch of `prepare` (those with
  `monkeypatch.setattr(..., "WAREHOUSE_DB", ...)` set to a missing file, or a sub-unit). Rewrite their imports to
  `from smartview_mcp import pnl_facts`, and assert the reworded message of `check_period` ("outside the live read
  window" instead of "the warehouse holds"). Add:

```python
def test_prepare_never_touches_a_warehouse(monkeypatch, fake_retrieve):
    from smartview_mcp import pnl_facts
    prep = pnl_facts.prepare("FinMax", "un.SY", "Sep 26", retrieve=fake_retrieve)
    assert prep.mode.startswith("live")
```

  `fake_retrieve` is the fixture the old live-branch test used; move it into `tests/fin_util.py` if it was local.
  Run: `uv run pytest -q -p no:cacheprovider tests\test_factstore.py tests\test_pnl_facts.py`.
  Expected: FAIL with `ModuleNotFoundError: smartview_mcp.factstore`.

- [ ] **Step 3: Create the modules.**
  - `factstore.py`: move `Row`, `Fill`, `OPEN`, `WhKey` and `MemoryWarehouse` (renamed `FactStore`) verbatim from
    `warehouse.py`. The docstring is: `"""In-memory facts for one pnl recipe call: rows read live by extract.fill,
    read back by fin_pnl.compute."""`.
  - `pnl_facts.py`: move from `tools_fin.py` the functions `resolve_area`, `is_unit`, `_entity_names`,
    `_finplan_of`, `resolve_scope`, `stale_note`, `_units`, `_aliases`, `_forecast`, `check_period`, and
    `prepare`/`_prepare` with **only** the live path:
    `store = FactStore(); extract.fill(store, retrieve, cat, extract.base_grids(cat, closed, templates.latest_forecast, entities=[entity]), "live", closed)`,
    then `run()` with `top_up` + `_forecast` + `fin_pnl.compute`. Move `refresh.top_up` (and its private helpers)
    into `pnl_facts.py`. Delete the warehouse open/lock/`WarehouseUnavailable` branches. `mode` is `"live"`, or
    `"live (sub-unit)"` where it is today, so `fin_spec.check_source(mode)` still yields `"Live read"`.
  - `core/recipes/pnl.py` and `tools_core._pnl_facts`: import `pnl_facts` instead of `tools_fin`.
  - `extract.py`: delete the `WarehouseUnavailable` handling and any import of `warehouse`; type `fill`'s store
    parameter as `FactStore`.
  - `fin_pnl.py`: delete `summary`, `table`, `_cells`, `tool_warnings`, `landing_warnings` where the grep rule
    (Task 3 Step 4) finds no other user. `fin_spec` needs `NAMES`, `PROVISIONAL`, `Result`, `finplan_warnings`,
    `flag_text`, `freshness_text`, `grand`, `CUBE`, `compute`, `Scope`, so those stay.
  - `doctor.py`: delete `_warehouse` and its entry in the checks list. `config.py`: delete `WAREHOUSE_DB` and
    `WAREHOUSE_LOCK_WAIT_S`.
  - `server.py`: delete `from . import tools_fin`. `git rm` the files `tools_fin.py`, `warehouse.py`, `refresh.py`.
  - `pyproject.toml`: delete the `duckdb==2.0.0` line; run `uv lock`. Expected: `Removed duckdb v2.0.0`.

- [ ] **Step 4: Tests follow the code.**
  - `git rm tests\test_warehouse.py tests\test_refresh.py tests\test_tools_fin.py`, once their kept cases are in
    `test_pnl_facts.py`/`test_factstore.py`.
  - In `test_fin_pnl.py`, `test_fin_spec.py`, `test_check_staged.py` and `tests/fin_util.py`, replace
    `MemoryWarehouse` with `FactStore` and `warehouse.` with `factstore.`.
  - In `test_fin_spec.py`/`test_recipes.py`, a test that expects `layout.expect[...]["source"] == "Warehouse"` now
    expects `"Live read"`.
  - In `test_config.py` and `test_doctor.py`, delete the warehouse cases.
  - In `test_extract.py`, delete the `WarehouseUnavailable` cases.
  
  Run: `Select-String smartview_mcp\*.py,smartview_mcp\core\*.py,smartview_mcp\core\recipes\*.py,tests\*.py,tests\live\*.py -Pattern 'warehouse|Warehouse|duckdb|tools_fin|fin_pnl\.summary|WAREHOUSE'`
  Expected: no hits, except the word "warehouse" in comments that explain the change. Rewrite those to describe the
  live read, so the expected result is **no hits at all**.

- [ ] **Step 5: Run the tests**

Run: `uv run pytest -q -p no:cacheprovider; uv run ruff check smartview_mcp tests`
Expected: all pass, 0 failed; `All checks passed!`.

- [ ] **Step 6: Commit**

`git add -A; git commit -m "refactor!: delete the DuckDB warehouse, fin_pnl and sv_warehouse tools; pnl recipe reads live" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 6: English-only sheet names and strings

**Files:**
- Modify: `engine/smartview_mcp/sheetorder.py`, `engine/smartview_mcp/answer.py`, `engine/smartview_mcp/drill_xl.py`,
  `engine/smartview_mcp/userexcel.py`, `engine/smartview_mcp/tools_drill.py`, `engine/smartview_mcp/tools_core.py`,
  `engine/smartview_mcp/styling.py`
- Test: `engine/tests/test_sheetorder.py`, `test_drill.py`, `test_followup.py`, `test_formulas.py`,
  `test_live_helpers.py`, `test_templates.py`

- [ ] **Step 1: Write the failing tests** in `tests/test_sheetorder.py`:

```python
import pytest
from smartview_mcp import sheetorder


@pytest.mark.parametrize("name", ["Q3 Answer", "Q3 Lines", "Sources", "Sources SV", "Q3 FinMax"])
def test_our_sheets_are_english(name):
    assert sheetorder.is_ours(name)


@pytest.mark.parametrize("name", ["Q3 Svar", "Q3 Linjer", "Kilder", "Kilder SV"])
def test_norwegian_names_are_the_users(name):
    assert not sheetorder.is_ours(name)
```

  Use the module's real predicate name (open `sheetorder.py`; the `_OURS` regex is matched by a public helper; use
  that one). Run: `uv run pytest -q -p no:cacheprovider tests\test_sheetorder.py`. Expected: the second test FAILS
  (Norwegian names are still recognised).

- [ ] **Step 2: Change the code.**
  - `sheetorder.py`: `_OURS` becomes `(Answer|FinMax|FinLite|FinPlan|Lines)`; `SOURCES = ("Sources", "Sources SV")`;
    the A1 marker is `"Fetched"`. Delete every `Svar|Linjer|Kilder|Hentet` alternative and the docstring lines that
    explain the compatibility.
  - `answer.py`: the default suffix is `"Answer"` (lines ~47, 76, 250, 323); translate the docstring and the status
    strings (~270, 326).
  - `drill_xl.py`: `Lines_Q{q}` and `Q{q} Lines`; translate the snapshot text (~140) and the strings at ~151, ~232.
  - `userexcel.py`:
    - `"Analyse"` → `"Analysis"` (~710, 923);
    - `Kilder` → `Sources` (~762, 770, 943);
    - `_Q` regex and ~967: English names only;
    - ~1014–1024 and the status strings ~1193, 1255 translated;
    - `_kilder_name` → `_sources_name`;
    - the Sources headers become `"Fetched", "Question no", "Question", "Cube", "POV", ...` (copy the full English
      list, matching the columns).
  - `tools_drill.py` ~276–277, 384–387 and `tools_core.py` ~201 (`'Kilder'` → `'Sources'`).
  - `styling.py`: delete `_lint`/`_lint_text` if only `vendor/tn_rapportstil`'s lint used them (Task 7 deletes that
    lint). Otherwise translate their messages.

- [ ] **Step 3: Find what is left**

Run: `Select-String smartview_mcp -Recurse -Include *.py -Pattern 'Svar|Linjer|Kilder|Hentet|Spørsmål|Analyse\b|Kube|feilet|mangler'`
Expected: hits only in the synonym tables (`core/period.py`, `core/scope.py`, `concepts.py`, `core/semantic.py`),
the TN style names, and `config.py`'s folder path. Translate any other hit.

- [ ] **Step 4: Tests.** In the test files listed above, replace the Norwegian sheet names and strings with the
  English ones (`"Q1 Svar"` → `"Q1 Answer"`, `"Kilder"` → `"Sources"`, `"Linjer_Q1"` → `"Lines_Q1"`, ...). Delete
  the tests that only proved old names are still read (search for `compat`/`legacy`/`old name` in their names and
  docstrings).
  Run: `uv run pytest -q -p no:cacheprovider; uv run ruff check smartview_mcp tests`
  Expected: all pass, 0 failed.

- [ ] **Step 5: Commit**

`git add -A; git commit -m "refactor!: English-only sheet names (Answer, Lines, Sources) and strings" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 7: Delete the remaining dead code

**Files:**
- Delete: `engine/smartview_mcp/hsgetvalue.py`, `engine/tests/test_hsgetvalue.py`
- Modify (deletions only): `vendor/tn_rapportstil.py`, `styling.py`, `fin_spec.py`, `tools_drill.py`, `server.py`,
  `analytics/periods.py`, `core/contract.py`, `ops.py`, `userexcel.py`, `templates.py`, `catalog.py`, and their tests

- [ ] **Step 1: List the candidates and prove each is dead.** For each name, run
  `Select-String smartview_mcp,tests -Recurse -Include *.py -Pattern '\b<name>\b'`. The name is dead when it has
  only its definition plus tests of it. Candidates:
  - `hsgetvalue.py` (whole module);
  - the lint section of `vendor/tn_rapportstil.py` (about 250 lines);
  - `fin_spec.with_reconciliation`;
  - `tools_drill._drill_rows`, `tools_drill._has_drill_cells`;
  - `server._safe_name`;
  - `analytics.periods.ltm`;
  - `contract.json_schema`;
  - anything Tasks 3–5 listed but kept "for now".
  
  Write the verdict for each name into the commit message body ("deleted: ...").

- [ ] **Step 2: Delete them**, and the tests that only cover them. Run `uv run ruff check smartview_mcp tests`; it
  flags the imports this leaves unused. Fix those too.

- [ ] **Step 3: Run the tests**

Run: `uv run pytest -q -p no:cacheprovider; uv run ruff check smartview_mcp tests`
Expected: all pass, 0 failed; `All checks passed!`.

- [ ] **Step 4: Commit**

`git add -A; git commit -m "refactor: delete code only the retired tools used" -m "deleted: <the names from Step 1>" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`
(write the actual names; Step 1 has them).

---

## Task 8: Split userexcel.py into the package `excel/`

**Files:**
- Create: `engine/smartview_mcp/excel/__init__.py` (docstring only), `jobs.py`, `login.py`, `instances.py`,
  `target.py`, `sheets.py`, `adhoc.py`
- Delete: `engine/smartview_mcp/userexcel.py`
- Modify: every importer of `userexcel` (engine and tests)

Map (line numbers are from HEAD `b0c142b`; after Tasks 3–7 locate by function name):

| New module | Contents |
|---|---|
| `excel/jobs.py` | `com`, `JobToken`, `bind_job`, `Turn`, `hold_turns`/`release_turns`, `fast_excel`, `rollback` (lines 40–245) |
| `excel/login.py` | `login_open`, `SmartViewBoxes`, `NOT_SIGNED_IN`, `recover_ie_login`, `wait_login_cleared` (245–352) |
| `excel/instances.py` | `XlInstance` … `bridge` (355–598) |
| `excel/target.py` | `Target`, `resolve_target`, new-file code (598–954) |
| `excel/sheets.py` | `next_question`, `unique_sheet_name`, `add_sheet`, `Pull`, meta/spec/pulls, `fte_rows`, `_make_pull`, `lookup_formula`, `log_source`, `wait_calculated` (954–1163, 1366–1452) |
| `excel/adhoc.py` | `pull`, `_pull`, `_write_and_retrieve`, `_drop_rejected`, sheet mode (1164–1365) |

- [ ] **Step 1: Map the importers**

Run: `Select-String smartview_mcp,tests -Recurse -Include *.py -Pattern 'userexcel' | % { "$($_.Path):$($_.LineNumber): $($_.Line.Trim())" }`
Expected: a list of import lines and `userexcel.<name>` uses. Save it to
`C:\Users\t992115\.copilot\session-state\09673b0c-ab80-492c-a857-e1daf56b6ab1\files\userexcel-uses.txt`, and use it
as the checklist for Step 3.

- [ ] **Step 2: Move the code.** Create each module with its functions moved verbatim, plus the imports they use.
  Cross-module references are absolute: `from . import jobs`, then `jobs.com(...)`, so test monkeypatches hit one
  object. Each module stays ≤ 400 lines. If `target.py` or `sheets.py` goes above that, move the new-file code to
  `excel/newfile.py`. `git rm engine\smartview_mcp\userexcel.py`.

- [ ] **Step 3: Update the importers** from the Step 1 list. Production code: `from .excel import jobs, login, ...`.
  Tests: `monkeypatch.setattr(userexcel, "wait_login_cleared", ...)` becomes `monkeypatch.setattr(login, ...)` **on
  the module that calls it**. If `adhoc._pull` calls `login.wait_login_cleared(...)`, patching `login` reaches it,
  because the call goes through the module attribute.

- [ ] **Step 4: Verify**

Run: `Select-String smartview_mcp,tests -Recurse -Include *.py -Pattern 'userexcel'`. Expected: no output.
Run the line counts:
`Get-ChildItem smartview_mcp\excel\*.py | % { "{0,5} {1}" -f (Get-Content $_.FullName -Encoding UTF8).Count, $_.Name }`
Expected: each ≤ 400.
Run: `uv run pytest -q -p no:cacheprovider; uv run ruff check smartview_mcp tests`. Expected: all pass, 0 failed.

- [ ] **Step 5: Commit**

`git add -A; git commit -m "refactor: split userexcel.py into the excel package" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 9: Split livexl.py, answer.py, host.py (and ops.py if needed)

**Files:**
- Create: `engine/smartview_mcp/workbench/{__init__,books,changes,fmt}.py`, `engine/smartview_mcp/answer_write.py`,
  `engine/smartview_mcp/hostproc.py`, possibly `engine/smartview_mcp/ops_reject.py`
- Delete: `engine/smartview_mcp/livexl.py`
- Modify: `tools_xl.py`, `answer.py`, `host.py`, `ops.py`, and their importers and tests

- [ ] **Step 1: livexl → workbench/.**
  - `books.py`: `Book`, `find_book`, `workbooks`, `read`, `overview` (lines 44–333).
  - `changes.py`: snapshots, `write`, `rows_cols`, `add_sheet`, `changes`, `revert` (334–720, 925–1044).
  - `fmt.py`: `format_range`, `format_sheet`, `trace`, `check`, `TN_EMPH` (721–924).
  - Shared helpers used by two or more of them go to `books.py`, which the others import.
  - If `changes.py` exceeds 400 lines, move `changes`/`revert` into `workbench/history.py`.
  - `git rm engine\smartview_mcp\livexl.py`; update `tools_xl.py` and the tests (`test_livexl*.py` →
    patch targets in `workbench.*`).

- [ ] **Step 2: answer → answer + answer_write.** Move `_Cells`, `_write`, `_named` and the chart code (from
  line ~286) into `answer_write.py`; `answer.py` keeps layout and rendering and imports `answer_write`. Both
  ≤ 400 lines.

- [ ] **Step 3: host → host + hostproc.** Move the process registry, own-Excel detection and `ComThread`
  (lines ~70–206) to `hostproc.py`. In `host.py` add `from .hostproc import ComThread, kill_ours, windows_of, ...`
  (the names the rest of `host.py` uses), and make `host.py`'s internal calls go through those module-level names, so
  `monkeypatch.setattr(host, "kill_ours", ...)` keeps working. Check with
  `Select-String tests\*.py -Pattern 'setattr\(host, "(\w+)"' | % { $_.Matches.Groups[1].Value } | Sort-Object -Unique`:
  every name listed must still be a module attribute of `host` that host's own code calls.

- [ ] **Step 4: ops.** Count its lines. If it is above 420, move `_rejected_members` and the bisection helpers to
  `ops_reject.py`, and import them in `ops.py`.

- [ ] **Step 5: Verify**

Run: `Get-ChildItem smartview_mcp\workbench\*.py, smartview_mcp\answer*.py, smartview_mcp\host*.py, smartview_mcp\ops*.py | % { "{0,5} {1}" -f (Get-Content $_.FullName -Encoding UTF8).Count, $_.Name }`
Expected: each ≤ 400 (ops ≤ 420).
Run: `Select-String smartview_mcp,tests -Recurse -Include *.py -Pattern '\blivexl\b'`. Expected: no output.
Run: `uv run pytest -q -p no:cacheprovider; uv run ruff check smartview_mcp tests`. Expected: all pass, 0 failed.

- [ ] **Step 6: Commit**

`git add -A; git commit -m "refactor: split livexl (workbench package), answer, host into modules of 400 lines or less" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 10: Login hardening – hung Login windows count as stuck; shorter grace

**Files:**
- Modify: `engine/smartview_mcp/dialogs.py`, `engine/smartview_mcp/host.py` (`signin_message`),
  `engine/smartview_mcp/config.py`
- Test: `engine/tests/test_dialogs.py`, `engine/tests/test_host_dialogs.py`

- [ ] **Step 1: Write the failing tests.** In `tests/test_dialogs.py`:

```python
from smartview_mcp.dialogs import Dialog, blocking


def test_a_hung_webview2_login_is_stuck_at_once():
    d = Dialog(21, "#32770", "Login", "", "login", False, hung=True)
    assert blocking([d], {}, now=100.0, grace=30.0) is d


def test_a_responsive_young_webview2_login_still_gets_the_grace():
    d = Dialog(22, "#32770", "Login", "", "login", False)
    assert blocking([d], {}, now=100.0, grace=30.0) is None
```

  In `tests/test_host_dialogs.py`:

```python
@pytest.mark.anyio
async def test_hung_login_restarts_excel_and_names_the_cause(monkeypatch):
    hung = Dialog(13, "#32770", "Login", "", "login", False, hung=True)
    xl = FakeExcel(monkeypatch, [[hung], [hung]])
    with pytest.raises(SignInBlocked) as e:
        await xl.h.call(xl.hang, timeout=5, what="sv_answer")
    assert "stopped responding" in str(e.value)
    assert xl.killed == [1001, 1002]
```

  Run: `uv run pytest -q -p no:cacheprovider tests\test_dialogs.py tests\test_host_dialogs.py`
  Expected: FAIL with `TypeError: Dialog.__init__() got an unexpected keyword argument 'hung'`.

- [ ] **Step 2: Implement.**
  - `dialogs.Dialog`: add the field `hung: bool = False  # Windows reports the window as not responding`.
  - Add a thin helper:

```python
def is_hung(hwnd: int) -> bool:
    """Windows' own 'Not responding' test (no input processed for 5 s)."""
    import ctypes
    try:
        return bool(ctypes.windll.user32.IsHungAppWindow(hwnd))
    except OSError:
        return False
```

  - `scan`: build each dialog with `hung=kind == "login" and is_hung(hwnd)`.
  - `blocking`: a login is stuck when `d.ie or d.hung or now - first_seen[d.hwnd] >= grace`. Update the docstring.
  - `host.signin_message`: when `d.hung`, the cause is `"Smart View's sign-in window stopped responding (Windows
    marks it 'Not responding')"`; the rest of the message (restart, next step, `smartass doctor --bundle`) is
    unchanged.
  - `config.py`: `LOGIN_GRACE_S = float(os.environ.get("SV_LOGIN_GRACE", "30"))`. Update the comment above it: "SSO
    once needed 18 s; a hung window is stuck at once".

- [ ] **Step 3: Run the tests**

Run: `uv run pytest -q -p no:cacheprovider tests\test_dialogs.py tests\test_host_dialogs.py tests\test_config.py`
Expected: all pass. If a `test_config.py` case asserts the default 45, change it to 30.

- [ ] **Step 4: Commit**

`git add -A; git commit -m "feat(login): a hung Smart View Login window is stuck at once; grace 45 -> 30 s" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 11: Login hardening – progress notes and ProviderEmpty restarts only the hidden Excel

**Files:**
- Modify: `engine/smartview_mcp/timing.py`, `engine/smartview_mcp/host.py`, `engine/smartview_mcp/ops.py`
- Test: `engine/tests/test_timing.py` (create it if absent), `engine/tests/test_host_dialogs.py`,
  `engine/tests/test_ops_transient.py` (create)

- [ ] **Step 1: Write the failing tests.** `tests/test_timing.py`:

```python
from smartview_mcp import timing


def test_say_reaches_the_client_of_the_running_job():
    notes = []
    timing.Timer.note_notify(notes.append)
    timing.Timer.start()
    timing.say("EPM gave an empty response; retrying in 3 s")
    assert notes == ["EPM gave an empty response; retrying in 3 s"]


def test_say_without_a_job_or_with_a_broken_callback_is_silent():
    timing._local.timer = None
    timing.say("nobody listens")
    timing.Timer.note_notify(lambda t: 1 / 0)
    timing.Timer.start()
    timing.say("still fine")
```

  `tests/test_ops_transient.py`:

```python
import pytest

from smartview_mcp import config, ops
from smartview_mcp.host import HostError, ProviderEmpty


class _H:
    pid = 4242

    def check_signin(self):
        pass


def test_a_752_that_survives_the_retry_raises_provider_empty(monkeypatch):
    monkeypatch.setattr(config, "TRANSIENT_RETRY_WAIT_S", 0)
    monkeypatch.setattr(ops, "ensure_grid_sheet", lambda h, conn, fresh=False: None)

    def fail(*a, **k):
        raise HostError("Retrieve failed rc=752: Empty response from provider")
    monkeypatch.setattr(ops, "_query_once", fail)
    with pytest.raises(ProviderEmpty) as e:
        ops.query(_H(), "FinMax", {}, [], [], True, True)
    assert e.value.pid == 4242


def test_a_752_that_clears_on_the_retry_returns(monkeypatch):
    monkeypatch.setattr(config, "TRANSIENT_RETRY_WAIT_S", 0)
    monkeypatch.setattr(ops, "ensure_grid_sheet", lambda h, conn, fresh=False: None)
    calls = []

    def once(*a, **k):
        calls.append(1)
        if len(calls) == 1:
            raise HostError("rc=752 Empty response from provider")
        return "grid"
    monkeypatch.setattr(ops, "_query_once", once)
    assert ops.query(_H(), "FinMax", {}, [], [], True, True) == "grid"
```

  In `tests/test_host_dialogs.py`:

```python
@pytest.mark.anyio
async def test_provider_empty_restarts_the_hidden_excel_once_and_says_so(monkeypatch):
    xl = FakeExcel(monkeypatch, [[], []])
    notes = []

    def fn(h):
        if h.pid == 1001:
            raise host.ProviderEmpty("rc=752 Empty response from provider", pid=1001, notify=notes.append)
        return "ok"

    assert await xl.h.call(fn, timeout=5, what="sv_answer") == "ok"
    assert xl.killed == [1001]
    assert any("yours is untouched" in n for n in notes)


@pytest.mark.anyio
async def test_provider_empty_twice_gives_the_next_step(monkeypatch):
    xl = FakeExcel(monkeypatch, [[], []])

    def fn(h):
        raise host.ProviderEmpty("rc=752 Empty response from provider", pid=h.pid, notify=None)

    with pytest.raises(HostError) as e:
        await xl.h.call(fn, timeout=5, what="sv_answer")
    assert "Empty response from provider" in str(e.value) and "smartass doctor --bundle" in str(e.value)
    assert xl.killed == [1001, 1002]


@pytest.mark.anyio
async def test_provider_empty_in_a_user_excel_job_restarts_only_ours_and_does_not_retry(monkeypatch):
    xl = FakeExcel(monkeypatch, [[]])
    xl.h._ensure_started()                    # the hidden Excel exists (pid 1001)
    runs = []

    def fn(h):
        runs.append(1)
        raise host.ProviderEmpty("rc=752 Empty response from provider", pid=1001, notify=None)

    with pytest.raises(HostError) as e:
        await xl.h.call(fn, timeout=5, what="sv_answer", user_excel=True)
    assert len(runs) == 1 and xl.killed == [1001]
    assert "ask again" in str(e.value)
```

  Check the user-Excel branch of `_call_once` before writing the third test. If a user-Excel job does not start the
  hidden Excel, the fake's `_ensure` call sets it up; otherwise adapt the setup so that pid 1001 exists, and keep the
  assertions.
  Run: `uv run pytest -q -p no:cacheprovider tests\test_timing.py tests\test_ops_transient.py tests\test_host_dialogs.py`
  Expected: FAIL (`timing.say` and `host.ProviderEmpty` do not exist).

- [ ] **Step 2: Implement.**
  - `timing.py`:

```python
def say(text: str) -> None:
    """A progress note to the user of the running job (if any), and the log. Never raises."""
    log.info("note: %s", text)
    t = Timer.current()
    if t is not None and t.notify is not None:
        try:
            t.notify(text)
        except Exception:   # noqa: BLE001 - a note must never break the job
            log.debug("progress note failed", exc_info=True)
```

  - `host.py`, after `SignInBlocked`:

```python
class ProviderEmpty(HostError):
    """EPM kept answering 'Empty response from provider' (rc=752/57704) after the in-place retry: the Smart View
    session of the engine's hidden Excel (pid) is broken. notify = the job's progress callback, kept for call()."""

    def __init__(self, msg: str, pid: int | None, notify=None):
        super().__init__(msg)
        self.pid, self.notify = pid, notify


PROVIDER_EMPTY = ("EPM kept answering 'Empty response from provider' {where}. This usually clears within a minute: "
                  "wait, then ask again. If it persists, sign in to Smart View once in your own Excel (Smart View "
                  "ribbon -> Panel -> FinMax) and run `smartass doctor --bundle`.")
```

  - `ExcelHost.call`: add a branch next to the `SignInBlocked` one:

```python
            except ProviderEmpty as e:
                if e.notify is not None:
                    try:
                        e.notify("EPM gave no response; restarting the engine's hidden Excel (yours is untouched)")
                    except Exception:   # noqa: BLE001
                        pass
                await asyncio.to_thread(self._hard_reset, f"{what}: Empty response from provider", e.pid)
                if user_excel:
                    raise HostError("EPM answered 'Empty response from provider'; the engine restarted its own "
                                    "hidden Excel (yours is untouched) - ask again.") from None
                if attempt == 2:
                    raise HostError(PROVIDER_EMPTY.format(where="after a restart of the engine's Excel")) from None
                log.warning("Empty response from provider during %s; restarted the hidden Excel, retrying once", what)
```

    Update the docstring of `call`.
  - `ops.query`: on a transient error, call
    `timing.say(f"EPM gave an empty response on {conn}; retrying in {config.TRANSIENT_RETRY_WAIT_S:.0f} s")`
    before the sleep. Wrap the retry:

```python
            try:
                return _query_once(h, conn, pov, rows, cols, suppress_missing, suppress_zeros)
            except HostError as e2:
                if is_transient(e2):
                    t = timing.Timer.current()
                    raise ProviderEmpty(str(e2), getattr(h, "pid", None), t.notify if t else None) from e2
                raise
```

  - `ops._connect_with_retry`: on the last attempt, when the error is transient, raise `ProviderEmpty` the same way.
    On the first transient attempt, call `timing.say("connecting to <conn>: EPM gave an empty response; retrying")`.
    Import `timing` and `ProviderEmpty` in `ops.py`.

- [ ] **Step 3: Run the tests**

Run: `uv run pytest -q -p no:cacheprovider tests\test_timing.py tests\test_ops_transient.py tests\test_host_dialogs.py tests\test_errors.py`
Expected: all pass. Then run the full suite. Expected: all pass, 0 failed.

- [ ] **Step 4: Commit**

`git add -A; git commit -m "feat(login): ProviderEmpty restarts only the engine's hidden Excel once; progress notes" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 12: Login hardening – the user's Excel: early notice, hung Login, retrieve retry

**Files:**
- Modify: `engine/smartview_mcp/excel/login.py`, `engine/smartview_mcp/excel/adhoc.py`
- Test: `engine/tests/test_excel_login.py` (create)

- [ ] **Step 1: Write the failing tests.** `tests/test_excel_login.py`:

```python
import pytest

from smartview_mcp import timing
from smartview_mcp.excel import adhoc, login
from smartview_mcp.host import HostError


class Clock:
    def __init__(self):
        self.t = 0.0

    def now(self):
        return self.t

    def sleep(self, s):
        self.t += s


def _fake(monkeypatch, open_for: float, hung: bool = False):
    c = Clock()
    monkeypatch.setattr(login, "login_open", lambda pid: [7] if c.t < open_for else [])
    monkeypatch.setattr(login.dialogs, "scan", lambda pid: [])
    monkeypatch.setattr(login.dialogs, "is_hung", lambda hwnd: hung)
    return c


def test_a_slow_sso_click_tells_the_user_once(monkeypatch):
    c = _fake(monkeypatch, open_for=12)
    notes = []
    timing.Timer.note_notify(notes.append)
    timing.Timer.start()
    login.wait_login_cleared(1, clock=c.now, sleep=c.sleep)
    assert len(notes) == 1 and "Login" in notes[0]


def test_a_hung_login_in_the_users_excel_fails_fast_and_closes_nothing(monkeypatch):
    c = _fake(monkeypatch, open_for=999, hung=True)
    closed = []
    monkeypatch.setattr(login.dialogs, "close", closed.append)
    with pytest.raises(HostError) as e:
        login.wait_login_cleared(1, clock=c.now, sleep=c.sleep)
    assert "not responding" in str(e.value) and "nothing was closed" in str(e.value)
    assert c.t <= login.LOGIN_HUNG_S + 1 and closed == []


def test_a_transient_retrieve_is_retried_once_and_said():
    said, calls = [], []

    def do():
        calls.append(1)
        if len(calls) == 1:
            raise HostError("Retrieve on 'Q1 FinMax' failed: rc=752 Empty response from provider")
        return "ok"
    assert adhoc.retrieve_with_retry(do, say=said.append, wait=lambda s: None) == "ok"
    assert len(calls) == 2 and len(said) == 1


def test_a_non_transient_retrieve_error_is_not_retried():
    calls = []

    def do():
        calls.append(1)
        raise HostError("Retrieve on 'Q1 FinMax' failed: rc=1000 member not found")
    with pytest.raises(HostError):
        adhoc.retrieve_with_retry(do, say=lambda t: None, wait=lambda s: None)
    assert len(calls) == 1
```

  Run: `uv run pytest -q -p no:cacheprovider tests\test_excel_login.py`
  Expected: FAIL (`wait_login_cleared` has no `clock` parameter; `retrieve_with_retry` does not exist).

- [ ] **Step 2: Implement in `excel/login.py`.**
  - Import `dialogs` and `timing` at module level, so the tests can patch `login.dialogs`.
  - `login_open(pid) -> list[int]` returns the hwnds. The list is truthy when one is open, so existing callers keep
    working.
  - Add the constants `LOGIN_NOTICE_S = 5` and `LOGIN_HUNG_S = 10`, and
    `LOGIN_HUNG = ("The Smart View Login window in your Excel is not responding. Nothing was closed. Close that Login
    window (or the Excel it belongs to), then ask again.")`.
  - `wait_login_cleared(pid, appear_grace=0.0, limit=90.0, clock=time.monotonic, sleep=time.sleep)`: use
    `clock()`/`sleep()` in place of `time.monotonic()`/`time.sleep()`.
    - When a login has been open for `LOGIN_NOTICE_S`, send one
      `timing.say("Smart View's Login is open in your Excel; waiting for the automatic SSO click")`.
    - Track `hung_since`. When any hwnd from `login_open` is `dialogs.is_hung`, keep the first time; reset it when
      none is hung. If `clock() - hung_since >= LOGIN_HUNG_S`, raise `HostError(LOGIN_HUNG)`.
    - The IE recovery path is unchanged.

- [ ] **Step 3: Implement in `excel/adhoc.py`.**

```python
def retrieve_with_retry(do, say, wait=time.sleep):
    """Run a retrieve in the user's Excel; a transient 'Empty response' (rc=752/57704) is retried once after
    TRANSIENT_RETRY_WAIT_S, and the user is told. The user's Excel is never restarted."""
    try:
        return do()
    except HostError as e:
        if not ops.is_transient(e):
            raise
        say("EPM gave an empty response in your Excel; retrying once")
        wait(config.TRANSIENT_RETRY_WAIT_S)
        return do()
```

  - In `_pull`, wrap the retrieve call: `retrieve_with_retry(lambda: <the existing retrieve>, timing.say)`.
  - In the connect loop, on the first "Empty response" connect, call
    `timing.say("connecting in your Excel: EPM gave an empty response; trying again")` before the 4 s sleep. The
    `NOT_SIGNED_IN` error after the second one stays.
  - Check that `ops` imports nothing from `excel` (no import cycle): `ops` → `host`; `excel.adhoc` → `ops`.

- [ ] **Step 4: Run the tests**

Run: `uv run pytest -q -p no:cacheprovider tests\test_excel_login.py; uv run pytest -q -p no:cacheprovider; uv run ruff check smartview_mcp tests`
Expected: all pass, 0 failed; `All checks passed!`.

- [ ] **Step 5: Commit**

`git add -A; git commit -m "feat(login): user's Excel - early Login notice, fast hung-Login error, one retry on Empty response" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 13: HTTP API – POST /gate and POST /answer

**Files:**
- Modify: `engine/smartview_mcp/httpapi.py`, `engine/smartview_mcp/tools_core.py`
- Test: `engine/tests/test_httpapi.py`

- [ ] **Step 1: Write the failing tests.** In `tests/test_httpapi.py`, add the fixture and the tests:

```python
import time


def _post(url, body, secret=SECRET, raw=None):
    data = raw if raw is not None else json.dumps(body).encode()
    req = urllib.request.Request(url, data=data, method="POST",
                                 headers={"Content-Type": "application/json",
                                          **({"X-Smartass-Secret": secret} if secret else {})})
    return urllib.request.build_opener(urllib.request.ProxyHandler({})).open(req, timeout=5)


@pytest.fixture
def pipe():
    seen = []

    async def gate(question):
        seen.append(("gate", question))
        return {"outcome": "clear", "question": question}

    async def answer(**kw):
        seen.append(("answer", kw))
        if kw.get("spec") == "slow":
            time.sleep(0.5)
        return {"sheet": "Q1 Answer"}

    s = httpapi.make_server(0, SECRET, health=lambda: {}, gate=gate, answer=answer)
    threading.Thread(target=s.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{s.server_address[1]}", seen
    s.shutdown()
    s.server_close()


def test_gate_and_answer_run_the_pipeline_functions(pipe):
    url, seen = pipe
    assert json.loads(_post(url + "/gate", {"question": "EBITDA Sep 26?"}).read())["outcome"] == "clear"
    assert json.loads(_post(url + "/answer", {"spec": {"version": 1}, "target": "new"}).read()) == {"sheet": "Q1 Answer"}
    assert seen == [("gate", "EBITDA Sep 26?"), ("answer", {"spec": {"version": 1}, "target": "new"})]


def test_post_needs_the_secret_valid_json_and_known_fields(pipe):
    url, _ = pipe
    for kwargs, code in (({"body": {"question": "x"}, "secret": None}, 401),
                         ({"body": None, "raw": b"{not json"}, 400),
                         ({"body": {"question": "x", "evil": 1}}, 400),
                         ({"body": None, "raw": b"{" + b" " * (httpapi.MAX_BODY + 1) + b"}"}, 413)):
        with pytest.raises(urllib.error.HTTPError) as e:
            _post(url + "/gate", **kwargs)
        assert e.value.code == code


def test_one_request_at_a_time(pipe):
    url, _ = pipe
    t = threading.Thread(target=lambda: _post(url + "/answer", {"spec": "slow"}))
    t.start()
    time.sleep(0.1)
    with pytest.raises(urllib.error.HTTPError) as e:
        _post(url + "/gate", {"question": "x"})
    assert e.value.code == 409
    t.join()
```

  Run: `uv run pytest -q -p no:cacheprovider tests\test_httpapi.py`
  Expected: FAIL (`make_server() got an unexpected keyword argument 'gate'`).

- [ ] **Step 2: Plain pipeline functions in `tools_core.py`.**
  - Rename the bodies: `async def gate(question: str) -> dict` (the current `sv_gate` body) and
    `async def answer(spec=None, followup=None, q=None, target="active", allow_archive=False) -> dict` (the current
    `sv_answer` body; copy the exact parameter list and defaults from `sv_answer`).
  - `sv_gate` and `sv_answer` stay registered tools with the same signatures and docstrings, and only
    `return await gate(question)` / `return await answer(...)`.
  - Errors are unchanged: the tool decorator's guard still turns `InputError`/`HostError` into tool errors for MCP.

- [ ] **Step 3: Implement in `httpapi.py`.**
  - `MAX_BODY = 1_000_000`.
  - `_FIELDS = {"/gate": {"question"}, "/answer": {"spec", "followup", "q", "target", "allow_archive"}}`.
  - `make_server(port, secret, health=None, gate=None, answer=None)`. The defaults are lazy imports of
    `tools_core.gate` / `tools_core.answer`.
  - One `threading.Lock` per server: `busy.acquire(blocking=False)` or 409 `{"error": "busy: one request at a
    time"}`.
  - `do_POST`:
    - check the secret (401) and the path (404);
    - check `Content-Length` (413 when above `MAX_BODY`, before reading);
    - `json.loads` (400 `{"error": "body is not JSON"}`); the body must be a dict with keys within `_FIELDS` (400,
      naming the unknown keys);
    - run `asyncio.run(fn(**body))` and send 200 with the result;
    - `InputError` → 400 `{"error": str(e)}`; `HostError` → 503 `{"error": str(e)}`; anything else → 500
      `{"error": "internal error; see server.log"}`, logged with `log.exception`.
  - Factor the secret check into `_authorised(self)`, used by both `do_GET` and `do_POST`.
  - Update the module docstring: "GET /health, POST /gate, POST /answer (the same pipeline calls as the MCP tools
    sv_gate/sv_answer) ... one request at a time".

- [ ] **Step 4: Run the tests**

Run: `uv run pytest -q -p no:cacheprovider tests\test_httpapi.py tests\test_core_skeleton.py; uv run pytest -q -p no:cacheprovider; uv run ruff check smartview_mcp tests`
Expected: all pass, 0 failed.

- [ ] **Step 5: Commit**

`git add -A; git commit -m "feat(http): POST /gate and /answer mirror the pipeline" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 14: Server instructions and the skill around the pipeline

**Files:**
- Modify: `engine/smartview_mcp/server.py` (`INSTRUCTIONS`), `skills/smartass/SKILL.md`,
  `skills/smartass/reference/operations.md`, `skills/smartass/reference/recipes.md`,
  `skills/smartass/reference/drill.md`, `skills/smartass/reference/cubes.md`
- Test: `engine/tests/test_skill_docs.py` (create)

- [ ] **Step 1: Write the failing test** `tests/test_skill_docs.py`. It keeps the docs honest about the tool set:

```python
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
GONE = ("sv_query", "sv_ws_pull", "sv_ws_build", "sv_ws_answer", "sv_analyse", "sv_concept", "fin_pnl",
        "sv_scope", "sv_warehouse", "sv_refresh_workbook")


def _docs():
    yield ROOT / "skills" / "smartass" / "SKILL.md"
    yield from (ROOT / "skills" / "smartass" / "reference").glob("*.md")
    yield ROOT / "README.md"
    yield ROOT / "docs" / "DESIGN.md"


def test_docs_name_no_retired_tool():
    hits = [f"{p.name}: {t}" for p in _docs() for t in GONE
            if re.search(rf"\b{t}\b", p.read_text(encoding="utf-8"))]
    assert hits == []


def test_server_instructions_lead_with_gate_then_answer():
    from smartview_mcp import server
    text = server.INSTRUCTIONS
    assert text.index("sv_gate") < text.index("sv_answer")
    assert not [t for t in GONE if re.search(rf"\b{t}\b", text)]
```

  Run: `uv run pytest -q -p no:cacheprovider tests\test_skill_docs.py`
  Expected: FAIL, listing the retired names still in the docs. README/DESIGN hits are fixed in Task 15; until then,
  run with `-k "not retired"` for this task's commit, or fix the README/DESIGN mentions here too.

- [ ] **Step 2: Rewrite `server.INSTRUCTIONS`.** The order is:
  1. `sv_gate` with the user's question; ask the questions it returns, or `sv_remember` a default.
  2. `sv_answer` with a spec (pulls + layout, or a recipe: `pnl`, `actual_vs_plan`, `monthly_trend`, `cost_base`,
     `concept`). A follow-up is `sv_answer(followup=<patch>, q=<n>)`.
  3. Read the `checks` and `findings`, and write the commentary from them, never from memory.
  4. Utilities: `sv_find_member`/`sv_members`/`sv_member_info`/`sv_dimensions` for writing specs; `sv_drill` for GL
     lines; `xl_*` for edits; `sv_status`/`sv_restart` for health.
  
  Keep the existing read-only, no-figures-from-memory and one-Excel rules word for word.

- [ ] **Step 3: Rewrite `skills/smartass/SKILL.md`.**
  - Front matter `description`: drop "refreshing a workbook's Smart View sheets"; keep the trigger words
    (FinMax, FinLite, FinPlan, Smart View, EPM, EBITDA, årsverk, un.SY, ...).
  - Sections:
    1. **When to use**.
    2. **The pipeline**: gate → answer (spec v1; worked example: the B02 spec from `scenarios._bus_spec` written
       out) → follow-ups (patch examples: add a month, swap a scenario) → recipes (one line each, with the
       arguments) → checks/findings/commentary.
    3. **Utilities**: member lookup, drill (with the link to `reference/drill.md`), the xl workbench, status/restart.
    4. **Sheets it writes**: `Q<n> Answer`, `Q<n> Lines`, `Sources`, source sheets `Q<n> FinMax`.
    5. **Errors**: link to `reference/operations.md`.
    6. **Rules**: unchanged.
  - Keep it at most about 250 lines.

- [ ] **Step 4: Reference files.**
  - `operations.md`: the errors table only for surviving tools and the new messages: hung Login (hidden: restart;
    yours: "not responding; nothing was closed"); `Empty response from provider` (notice → restart of the hidden
    Excel → next step); `busy` (HTTP 409).
  - `recipes.md`: pnl reconciles "Live read = source sheet"; replace "area (as fin_pnl)" with the area rules
    themselves.
  - `drill.md` and `cubes.md`: replace mentions of `sv_query`/`sv_ws_*` with `sv_answer`.

- [ ] **Step 5: Run the tests**

Run: `uv run pytest -q -p no:cacheprovider tests\test_skill_docs.py tests\test_smoke_stdio.py`
Expected: all pass (or the README/DESIGN test deselected until Task 15, as in Step 1).

- [ ] **Step 6: Commit**

`git add -A; git commit -m "docs(skill): SKILL.md and references around gate -> answer -> follow-ups/recipes -> findings" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 15: README, DESIGN, CHANGELOG, troubleshooting; archive the old docs

**Files:**
- Modify: `README.md`, `docs/DESIGN.md`, `CHANGELOG.md`, `docs/troubleshooting.md`, `engine/tests/live/README.md`
- Move: `docs/specs/2026-10-*`, `docs/plans/2026-10-*`, `docs/audit/2026-10-*` and the spike audit files →
  `docs/archive/specs/`, `docs/archive/plans/`, `docs/archive/audit/`

- [ ] **Step 1: Archive.** From `C:\dev\smartass`:

```powershell
foreach ($d in "specs", "plans", "audit") {
  New-Item -ItemType Directory -Force "docs\archive\$d" | Out-Null
  Get-ChildItem "docs\$d\2026-10-*" | % { git mv $_.FullName "docs\archive\$d\$($_.Name)" }
}
git status --short | Select-String '^R' | Measure-Object | % Count
```

  Expected: a count above 0, with renames only. Then look for spike files dated 2026-11-0x
  (`Get-ChildItem docs\audit\*spike*`) and move them the same way.

- [ ] **Step 2: Fix the links.**

Run: `Select-String -Path README.md, docs\*.md, docs\specs\*.md, docs\plans\*.md, docs\audit\*.md, skills\smartass\*.md, skills\smartass\reference\*.md, engine\tests\live\README.md -Pattern '(specs|plans|audit)/2026-10-'`
Expected: each hit points into the archive: rewrite `docs/specs/2026-10-...` → `docs/archive/specs/2026-10-...`
(relative links: add `archive/` after the right `../`). Re-run the command. Expected: every remaining hit contains
`archive/`.

- [ ] **Step 3: Update the docs.**
  - `README.md`: what smartass does (one pipeline), the tool list (25, grouped as in D1), install, `smartass serve`
    endpoints.
  - `docs/DESIGN.md`: the architecture as built (gate, contract, pipeline, recipes, live P&L facts, excel/workbench
    packages, login hardening). Remove the warehouse section and the old tool table.
  - `CHANGELOG.md` "Unreleased":
    - removed: the 10 tools, the warehouse/DuckDB, Norwegian sheet names (a breaking change: old workbooks' `Svar`/
      `Kilder` sheets are no longer recognised);
    - added: hung-login detection, the `ProviderEmpty` restart, progress notes, HTTP `/gate` and `/answer`;
    - changed: the `LOGIN_GRACE_S` default 30.
    - Delete the Unreleased lines that describe `fin_pnl`/warehouse features as new.
  - `docs/troubleshooting.md`, login section: what the user sees now (notes within about 5 s, the hung-login
    messages, "Empty response" handling) and that the engine never closes the user's Excel.

- [ ] **Step 4: Run the tests**

Run (from `engine`): `uv run pytest -q -p no:cacheprovider tests\test_skill_docs.py`
Expected: 2 passed (no retired tool named in README/DESIGN/skill).

- [ ] **Step 5: Commit**

`git add -A; git commit -m "docs: README, DESIGN, CHANGELOG, troubleshooting for the pipeline; archive pre-rework docs" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`

---

## Task 16: Full verification offline and the line count

- [ ] **Step 1: Full suite and lint**

Run (from `engine`): `uv run pytest -q -p no:cacheprovider; uv run ruff check smartview_mcp tests`
Expected: all pass, 0 failed, 0 errors; `All checks passed!`. Write down the test count.

- [ ] **Step 2: No retired names in the engine**

Run: `Select-String smartview_mcp,tests -Recurse -Include *.py -Pattern 'sv_query|sv_ws_|sv_analyse|sv_concept|fin_pnl"|sv_scope|sv_warehouse|sv_refresh_workbook|duckdb|userexcel|livexl|Kilder|Svar\b|Linjer'`
Expected: only `test_skill_docs.py` (its `GONE` list) and `test_sheetorder.py` (the "Norwegian names are the
user's" cases).

- [ ] **Step 3: Module sizes**

Run: `Get-ChildItem smartview_mcp -Recurse -Filter *.py | % { [pscustomobject]@{n=(Get-Content $_.FullName -Encoding UTF8).Count; f=$_.FullName.Substring((Get-Location).Path.Length+1)} } | ? n -gt 400 | Sort-Object n -Descending | Format-Table -AutoSize`
Expected: only the D9-exempt modules (`styling`, `fin_spec`, `fin_pnl`, `drill_browser`, `catalog`, `vendor/*`)
and `ops.py` ≤ 420.

- [ ] **Step 4: Line count against the baseline.** Commit first (Step 5 of the previous task leaves a clean tree),
  then run the header's measurement script with `origin/main HEAD`.
  Expected: HEAD ≤ 16 718 lines (the D14 gate). Record both numbers, the owner's 10 031 target and the gap. They go
  into the PR body (Task 18) and into `docs/audit/2026-11-02-bench-core-final.md` (Task 17). If HEAD is above
  16 718, go back over Task 7's candidate list: `Select-String` for public functions with one definition hit and no
  caller.

- [ ] **Step 5: Smoke tool count**

Run: `uv run pytest -q -p no:cacheprovider tests\test_smoke_stdio.py`
Expected: passed, with `len(names) == 25`.

(No commit unless Step 4 deleted more code; then commit
`refactor: delete unreferenced code found by the line audit` with the trailer.)

---

## Task 17: Live bench – route core and gate

Needs EPM access and Excel on this PC. The user must not use Excel during the run (the bench polls visible Excel
for timing).

- [ ] **Step 1: Engine status**

Run: `uv run python -c "from smartview_mcp import doctor; print(doctor.as_dict(doctor.run(include_http=False))['status'])"`
Expected: `OK` (or `WARN` with no FAIL item).

- [ ] **Step 2: Run the bench**

Run: `uv run python tests\live\bench.py --out bench-core-final` (about 15–25 min; use `initial_wait` 600 and read
the rest with `read_powershell`).
Expected, per case line: B01–B09 PASS, B12–B15 PASS; B10, B11 and T01–T10 scored by the gate with `gate 12/12`
(suggested outcome = expected, 0 false-go). The oracle pass runs after the server closes and prints `oracle: n/n
cells`. Files: `docs\audit\2026-11-02-bench-core-final.json` and `.md`.

- [ ] **Step 3: If a case fails**, use superpowers:systematic-debugging. Compare its JSON entry with the same case in
  `docs/audit/2026-11-02-bench-core-widen.json`.
  - A changed oracle number (`match: missing`) means the answer moved.
  - A `checks_ok: false` on B01/B15 means the live reconciliation (D5).
  - A login/`Empty response` message means Tasks 10–12.
  
  Fix the cause with a TDD unit test first, commit, and re-run only that case:
  `uv run python tests\live\bench.py --only B01 --out bench-core-final-B01`.

- [ ] **Step 4: Audit note.** Add a "Verdict" section at the top of `docs/audit/2026-11-02-bench-core-final.md`:
  - the pass counts;
  - total_s for B01/B13/B15 against the widen run (the cost of the live pnl facts, D5);
  - any `Empty response`/login notes seen in `server.log` during the run
    (`Select-String $env:LOCALAPPDATA\smartass\logs\server.log -Pattern 'Empty response|ProviderEmpty|stopped responding'`;
    take the log path from `config.LOG_DIR` if it differs);
  - the line counts from Task 16.

- [ ] **Step 5: Commit and push**

`git add -A; git commit -m "docs(audit): final live bench, route core + gate" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"; if ($?) { git push origin rework/pipeline }`
Expected: the push succeeds.

---

## Task 18: Pull request rework/pipeline → main

- [ ] **Step 1: gh auth**

Run: `gh auth status`
If it fails:

```powershell
$cred = "protocol=https`nhost=github.com`n`n" | git credential fill
$env:GH_TOKEN = ($cred | Select-String '^password=(.*)$').Matches[0].Groups[1].Value
gh auth status
```

Expected: logged in to github.com.

- [ ] **Step 2: Write the PR body** to
  `C:\Users\t992115\.copilot\session-state\09673b0c-ab80-492c-a857-e1daf56b6ab1\files\pr-body.md` (outside the repo).
  It contains:
  - a summary (one pipeline, `sv_gate` → `sv_answer`; 35 → 25 tools);
  - **breaking changes** (the 10 retired tools; the warehouse/DuckDB is gone; English-only sheet names; the
    `LOGIN_GRACE_S` default);
  - login hardening;
  - HTTP `/gate` and `/answer`;
  - verification: the unit test count; the bench table (B01–B09, B12–B15 PASS, gate 12/12) with a link to the audit
    file; the line counts (main, HEAD, the owner's target and the gap, D14);
  - the open owner decisions (D3, D5 speed, D9 exemptions, D12, D14).

- [ ] **Step 3: Create the PR**

Run (from `C:\dev\smartass`): `gh pr create --base main --head rework/pipeline --title "Rework: one pipeline (gate -> answer), retire old tools and warehouse, login hardening" --body-file C:\Users\t992115\.copilot\session-state\09673b0c-ab80-492c-a857-e1daf56b6ab1\files\pr-body.md`
Expected: a PR URL is printed. Check it with `gh pr view --json number,state,baseRefName,headRefName`. Expected:
`state: OPEN`, base `main`, head `rework/pipeline`.

---

## Self-review

- **Spec coverage.**
  - Owner decision 1: tools are deleted in Tasks 3–5, with every verdict and reason in D1–D3, dead code in Task 7,
    tests deleted with their code in every task, and old docs archived in Task 15.
  - Decision 2 (warehouse): Task 5 and D5, with the reconciliation kept as "Live read".
  - Decision 3 (English names): Task 6 and D6.
  - Decision 4 (PR): Task 18.
  - Targets: line count (header, D14, Task 16), module sizes (D9, Tasks 8–9), smoke count (35 → 30 → 27 → 25 in
    Tasks 3–5), live bench (Tasks 1 and 17).
  - Login hardening (design §3.2): Tasks 10–12. Hidden Excel: hung → restart, `ProviderEmpty` → note + one
    restart. User's Excel: never restarted; notice, hung error, retrieve retry. Tests use fakes.
  - Skill/docs: Tasks 14–15. httpapi: Task 13.
- **Green after every commit.**
  - The oracle moves before `sv_query` is deleted (Task 1 before Task 4).
  - The reply helpers move before `tools_ws` goes (Task 2 before Task 3).
  - `fin_pnl`'s use of `tools_ws._build_job` is handled in Task 3 Step 3, by deleting the `fin_pnl` tool early if
    needed.
  - English names (Task 6) come before the `userexcel` split (Task 8), so the split moves translated code.
  - Login Task 12 targets the `excel/` modules, so it runs after Task 8.
- **Placeholder scan.** The two places that depend on code read at implementation time are explicit lookups with a
  command and a rule, not open blanks:
  - Task 5's `Row` fields: "copy the real names";
  - Task 7's "deleted:" list: produced by Step 1's grep.
- **Type consistency.**
  - `ProviderEmpty(msg, pid, notify)` is used the same way in `ops`, `host` and the tests.
  - `login_open` returns `list[int]` (still truthy), `wait_login_cleared` gains `clock`/`sleep`, and
    `retrieve_with_retry(do, say, wait)` matches its tests.
  - `make_server(port, secret, health, gate, answer)`.
  - `FactStore` replaces `MemoryWarehouse` everywhere.
- **Risks.**
  - B01/B15 are slower with live P&L facts (measured in Task 17).
  - The line target is not met (D14, owner decision).
  - `/answer` in the `serve` process starts its own hidden Excel when called (D12): it shares nothing with the MCP
    server, so two hidden Excels can run if both are used at once.

