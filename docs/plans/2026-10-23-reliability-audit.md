# Reliability audit — implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to carry this plan out task by task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Run every user workflow of the smartass engine live and score it. Fix the failures in order of impact,
trim the tool list, and commit a repeatable audit.

**Spec:** `docs/specs/2026-10-23-reliability-audit-design.md`. "Done when" 1–6 there is the acceptance gate.

**Architecture:**
- **Harness.** `tests\live\mcpclient.py` (a stdio JSON-RPC client) and `tests\live\matrix.py` (the runner and
  report). `tests\live\scenarios.py` holds S1–S10 as registered functions, and `tests\live\helpers.py` holds the
  injections and Excel helpers.
- **Fixes.** Logic goes into existing engine modules, test-first, with unit tests in `engine\tests`. Live behaviour
  is checked through matrix tags.

**Tech:**
- Python 3.12 and uv.
- pytest. Live tests are excluded by `norecursedirs = ["live"]`.
- pywin32 and psutil for the injections.

**Conventions (every task):**
- Work in `C:\dev\smartass\engine` on branch `impl/reliability-audit`.
- Unit tests: `$env:PYTHONUTF8=1; uv run pytest -q`
- Lint: `uv run ruff check .`
- Guard: `uv run python ..\scripts\check_staged.py --all`
- Commit trailer: `Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>`
- Never use `Get-Content | Set-Content` on source files (it mangles æ/ø/å); use the edit tool.
- Live runs:
  - read-only to EPM;
  - scratch workbooks go in `<analyser>\_matrix\`;
  - never touch a workbook the run did not open.

---

## Phase A — Harness

### Task 1: `mcpclient.py`

**Files:** create `engine\tests\live\mcpclient.py`. `live_relay.py` is left alone until Task 6.

- [ ] **Step 1: Write the client.** Base it on the plumbing in `live_relay.py`: Popen, a reader thread and a queue.

```python
"""Stdio MCP client for live tests: starts run_server.py exactly like Copilot/Claude do."""
from __future__ import annotations

import json
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path

ENGINE = Path(__file__).resolve().parents[2]


class ToolFailed(Exception):
    pass


class McpClient:
    def __init__(self, python: str | None = None, env: dict | None = None):
        self.p = subprocess.Popen([python or sys.executable, "-B", str(ENGINE / "run_server.py")], cwd=ENGINE,
                                  stdin=subprocess.PIPE, stdout=subprocess.PIPE, env=env)
        self.inbox: queue.Queue = queue.Queue()
        self.notes: list[dict] = []
        self._id = 0
        threading.Thread(target=self._read, daemon=True).start()
        self.request("initialize", {"protocolVersion": "2025-06-18", "capabilities": {},
                                    "clientInfo": {"name": "matrix", "version": "1"}}, timeout=120)
        self._send({"jsonrpc": "2.0", "method": "notifications/initialized"})

    def _read(self):
        for line in self.p.stdout:
            if line.strip():
                m = json.loads(line)
                (self.notes.append(m) if "id" not in m else self.inbox.put(m))

    def _send(self, msg: dict):
        self.p.stdin.write((json.dumps(msg) + "\n").encode())
        self.p.stdin.flush()

    def request(self, method: str, params: dict | None = None, timeout: float = 600) -> dict:
        self._id += 1
        i = self._id
        self._send({"jsonrpc": "2.0", "id": i, "method": method, "params": params or {}})
        t0 = time.monotonic()
        while time.monotonic() - t0 < timeout:
            try:
                m = self.inbox.get(timeout=1)
            except queue.Empty:
                continue
            if m.get("id") == i:
                return m
        raise TimeoutError(f"{method} gave no answer within {timeout:.0f}s")

    def call(self, tool: str, args: dict, timeout: float = 600) -> tuple[dict | None, float, str | None]:
        """Returns (body, seconds, error). body is the parsed JSON result, or {"_text": ...} for plain text."""
        t0 = time.monotonic()
        r = self.request("tools/call", {"name": tool, "arguments": args}, timeout)
        secs = time.monotonic() - t0
        if "error" in r:
            return None, secs, json.dumps(r["error"])
        res = r["result"]
        text = res["content"][0]["text"] if res.get("content") else ""
        if res.get("isError"):
            return None, secs, text
        try:
            return json.loads(text), secs, None
        except ValueError:
            return {"_text": text}, secs, None

    def must(self, tool: str, args: dict, timeout: float = 600) -> tuple[dict, float]:
        body, secs, err = self.call(tool, args, timeout)
        if err:
            raise ToolFailed(f"{tool}: {err}")
        return body, secs

    def tools(self) -> list[str]:
        return [t["name"] for t in self.request("tools/list")["result"]["tools"]]

    def close(self):
        try:
            self.p.stdin.close()
            self.p.wait(120)
        except Exception:
            self.p.kill()
```

- [ ] **Step 2: Smoke-test it live.**

```powershell
@'
import sys; sys.path.insert(0, r"C:\dev\smartass\engine\tests\live")
from mcpclient import McpClient
c = McpClient(); print(len(c.tools()), c.must("sv_status", {})[1]); c.close()
'@ | uv run python -
```

Expected: `36 <seconds>`. Before the trim there are 36 tools.

- [ ] **Step 3:** Run ruff, then commit: `test(live): stdio MCP client for the matrix`.

### Task 2: `matrix.py` runner

**Files:**
- Create `engine\tests\live\matrix.py` and `engine\tests\live\scenarios.py` (registry only for now).
- Create `engine\tests\test_matrix_logic.py` to unit-test the pure functions.

Before the code, two placement decisions:
- The pure logic (error-class normalisation, result rules, the before/after merge) lives in
  `engine\smartview_mcp\auditlib.py`. That keeps it importable by unit tests. It is about 60 lines.
- `matrix.py` imports it.

- [ ] **Step 1: Write the failing unit tests** (`tests\test_matrix_logic.py`):

```python
from smartview_mcp.auditlib import error_class, merge_before, verdict


def test_error_class_normalises_digits_and_keeps_first_line():
    a = error_class("HostError: drill page never asked for its data within 45s\nnext: ...")
    b = error_class("HostError: drill page never asked for its data within 47s")
    assert a == b == "HostError: drill page never asked for its data within #s"


def test_verdict_rules():
    assert verdict(ok=True, secs=10, budget=30) == "PASS"
    assert verdict(ok=True, secs=31, budget=30) == "SLOW"
    assert verdict(ok=False, secs=5, budget=30) == "FAIL"
    assert verdict(ok=False, secs=5, budget=30, limit="no FinLite in test env") == "LIMIT"


def test_merge_before_adds_previous_result():
    now = [{"id": "S1.query", "result": "PASS"}]
    before = [{"id": "S1.query", "result": "FAIL"}, {"id": "gone", "result": "PASS"}]
    assert merge_before(now, before)[0]["before"] == "FAIL"
```

- [ ] **Step 2:** Run `uv run pytest -q tests\test_matrix_logic.py`. Expected: ImportError.

- [ ] **Step 3: Implement `smartview_mcp\auditlib.py`.**

```python
"""Pure helpers for the live audit matrix (tests/live/matrix.py): error classes, verdicts, before/after."""
from __future__ import annotations

import re


def error_class(err: str) -> str:
    first = (err or "").strip().splitlines()[0] if err else ""
    return re.sub(r"\d+", "#", first)[:160]


def verdict(ok: bool, secs: float, budget: float, limit: str | None = None) -> str:
    if not ok:
        return "LIMIT" if limit else "FAIL"
    return "SLOW" if secs > budget else "PASS"


def merge_before(now: list[dict], before: list[dict]) -> list[dict]:
    prev = {r["id"]: r["result"] for r in before}
    return [{**r, "before": prev.get(r["id"], "")} for r in now]
```

- [ ] **Step 4:** Run the tests. Expected: pass.

- [ ] **Step 5: Write `scenarios.py` (registry) and `matrix.py` (runner).**

```python
# scenarios.py  (registry part; the scenarios themselves are added in Tasks 3-5)
"""Live scenarios S1-S10. Each is fn(ctx) -> dict(notes=..., budget=...) and raises AssertionError/ToolFailed on failure."""
from __future__ import annotations

from dataclasses import dataclass, field

REGISTRY: dict[str, "Scenario"] = {}


@dataclass
class Scenario:
    id: str
    fn: object
    tags: list[str]
    budget: float            # seconds; over it -> SLOW
    limit: str | None = None # set when the failure is an accepted, documented limit
    fresh: bool = False      # only runs with --fresh


@dataclass
class Ctx:
    client: object           # McpClient
    scratch: object          # Path to <analyser>\_matrix
    keep: bool
    state: dict = field(default_factory=dict)   # shared between scenarios (e.g. S2's workbook for S7)


def scenario(id: str, tags: list[str], budget: float = 60, limit: str | None = None, fresh: bool = False):
    def deco(fn):
        REGISTRY[id] = Scenario(id, fn, tags, budget, limit, fresh)
        return fn
    return deco
```

`matrix.py` does the following:
- parses `--only TAGS` (comma-separated), `--repeat N`, `--fresh`, `--keep` and `--out NAME`;
- imports `scenarios` and `helpers`;
- starts one `McpClient`;
- creates the scratch folder `<analyser>\_matrix` by reading `config.ANALYSER_DIR`, which the implementer finds with
  `grep -n "analyser" smartview_mcp\config.py`;
- runs the selected scenarios in registry order.

For each run:
- Time the run and catch any exception: `ok=False`, `error=str(e)`.
- **Flaky rule:** a scenario that fails is run once more before it counts. If the second run passes, record it as
  `PASS` with `"flaky": true`.
- Record `{id, tags, result, secs, error, error_class, notes, flaky, run}`.

Finally (in a `finally` block):
- Close the client.
- Run `helpers.cleanup(ctx)`, which closes the scratch workbooks unsaved and deletes `_matrix` unless `--keep`.
- Write `docs\audit\<YYYY-MM-DD>-<out or "matrix">.json` and `.md`.
- Merge with the newest earlier `*-matrix*.json` in `docs\audit` through `merge_before`.
- Write a header with the time, the atlas age (from `sv_atlas_status`) and the engine version.
- The `.md` is one table, `| ID | Tags | Result | Before | s | Error class |`, followed by an error-class summary sorted
  by count.
- Exit code 1 if any result is FAIL.

- [ ] **Step 6: Dry run with a dummy scenario.** Add the following temporarily:

```python
@scenario("S0.status", ["smoke"], budget=10)
def s0(ctx):
    b, _ = ctx.client.must("sv_status", {})
    assert b
```

Then run `uv run python tests\live\matrix.py --only smoke --out dry`.
- Expected: `docs\audit\<date>-dry.md` shows `S0.status PASS`.
- Delete the dry-run files. Keep `S0.status`; it is the cheap health check that runs first.

- [ ] **Step 7:** Run ruff and pytest, then commit: `test(live): matrix runner with error classes and before/after`.

### Task 3: Scenarios S1–S7

**Files:** `engine\tests\live\scenarios.py` and `engine\tests\live\helpers.py`.

**Before writing each scenario:**
- Read the tool's signature in `tools_*.py`. Use real parameter names, never guessed ones.
- Reuse the specs and expected numbers in the old scripts:
  - `live_build.py`: the BUS spec, `3 495 317 868`, `88 601 304`;
  - `live_analysis.py`: the TEC templates, `(1 112 834 010)`, WHS `1 994 151 358`;
  - `live_drill*.py`, `verify_gl.py`: FinLite 12881 / un.SY / Sep 26 = −161 233 427.84, 95 lines, gap 0;
  - `live_xl.py`, `live_refresh.py`, `live_sessions.py`, `live_smoke.py`.
- Every scenario targets a new workbook created in `ctx.scratch`, or the workbook S2 left in `ctx.state`.

| ID | Tags | Budget | Calls and assertions |
|---|---|---|---|
| `S1.find` | lookup | 5 | `sv_find_member` for `12881`, `un.SY.2TEC`, a product (`mobil`) and `årsverk`. Each has ≥1 hit and the expected member is first or among the top 3. |
| `S1.query.finmax` | lookup | 5 | `sv_query` FinMax: `ac.EBITDA` × `Sep 26`, entity `un.SY.2WHS`, Accumulated. Equals 1 858 190 823. |
| `S1.query.finlite` | lookup | 5 | `sv_query` FinLite: `12881` × `Sep 26` × `un.SY`. Equals −161 233 427.84. |
| `S1.query.finplan` | lookup | 8 | `sv_query` FinPlan: one EBITDA cell for `BANORWAY` FY26. A number, no error. |
| `S1.bad.balance` | lookup,input | 5 | `sv_query` with dimension `Balance`. Passes with a warning naming `All Balance Types`, or is rejected with a message that names it. |
| `S1.bad.quoted` | lookup,input | 5 | Dimension `"'All Balance Types'"`. Same rule as above. |
| `S1.bad.dictrows` | lookup,input | 5 | `rows={"All Accounts": ["ac.EBITDA"]}`. Passes. |
| `S1.bad.member` | lookup,input | 5 | Member `ac.NOPE`. The error names the member and suggests `sv_find_member`. |
| `S2.build` | build | 30 | The `live_build.py` BUS spec with `target` set to a new workbook under `ctx.scratch`. All `checks` OK, `formula_style.ok`, the known numbers present. Store the workbook in `ctx.state["wb"]`. |
| `S2.patch` | build | 15 | `sv_ws_answer` with `spec_patch` on the same Q. Checks OK, numbers unchanged. |
| `S2.plan` | build | 45 | `sv_ws_build` on the same Q adding a FinPlan pull. No orphan sheets: every data sheet is referenced by the answer. The order is answer before data (`xl_overview`). |
| `S3.<template>.tec` / `.sy` | template | 60 | Four templates × two entities (FinPlan uses `BANORWAY`). Checks OK, no formula errors, TEC P&L contains `(1 112 834 010)`. |
| `S4.scope` | concept | 20 | `sv_scope("hva tjener vi på mobilabonnement")` returns a `fam_*` concept, then `sv_concept` on it and on `splitt`. Revenue + COGS + opex = EBITDA before other items, within 1 NOK. |
| `S5.drill.cold` | drill | 120 | First kill the Edge processes the engine started (`helpers.kill_engine_edge()`). Then `sv_drill` 12881 × un.SY × Sep 26: 95 lines, gap 0, `exact`. |
| `S5.drill.warm` | drill | 60 | The same drill again. |
| `S5.drill.quarter` | drill | 180 | The cell with `Q3 26`. The split reconciles, `exact`. |
| `S5.drill.report` | drill | 90 | `sv_ws_answer` with a `drill` cell on S2's workbook. Builds without error. |
| `S6.xl` | xl | 60 | On a new scratch workbook: `xl_write` → `xl_read` → `xl_format_sheet` → `xl_check` → `xl_trace` → `xl_overview` → `xl_changes` → `xl_revert`. After the revert every value and format equals the snapshot (compare with `xl_read` before and after), and `Saved` is untouched. |
| `S7.refresh` | refresh | 90 | Save S2's workbook to `ctx.scratch`, then `sv_refresh_workbook`. Values are unchanged (compare `xl_read` of the answer range before and after), with no "Call was rejected". |

`helpers.py` (part 1) provides:
- `kill_engine_edge()`: the `msedge` PIDs whose command line contains the engine's drill profile dir (see
  `config.DRILL_PROFILE` or its equivalent).
- `values(client, wb, sheet, rng)`.
- `cleanup(ctx)`.

- [ ] **Step 1:** Write S1 and run `matrix.py --only lookup`. Results are expected to FAIL on `bad.*` until Phase C;
  record them.
- [ ] **Step 2:** Write S2 and S3, then run `--only build,template`.
- [ ] **Step 3:** Write S4 and S5, then run `--only concept,drill`.
- [ ] **Step 4:** Write S6 and S7, then run `--only xl,refresh`.
- [ ] **Step 5:** Commit: `test(live): scenarios S1-S7`.

A scenario that fails here is **not** fixed in this task. It is data for the baseline. Only fix the scenario code when
the scenario itself is wrong: a bad parameter name, or a wrong expected value proven with `sv_query`.

### Task 4: S8 injections

**Files:** `engine\tests\live\helpers.py` (part 2) and `scenarios.py`.

Each injection is a context manager that undoes itself in `finally`. Each scenario ends with a follow-up `sv_query`
(as in `S1.query.finmax`) that must pass. Budget: 120 s per call. A call still running at 120 s is a FAIL with the class
`HANG`, so use `client.call(..., timeout=130)`.

| ID | Injection (helper) | Pass |
|---|---|---|
| `S8.sso` | `expired_sso()`: rename the cached token file (path from `config`) and the engine Edge profile's `Network\Cookies` to `*.matrix-bak`; restore afterwards. | `sv_query` recovers by itself or names re-login as the next step. |
| `S8.dialog` | `modal_dialog(wb)`: a thread runs `xl.ExecuteExcel4Macro('ALERT("matrix")')` in the target Excel; afterwards press Enter on it through `WScript.Shell.AppActivate` + `SendKeys`. | `sv_ws_answer` names "a dialog is open". No hang. |
| `S8.editmode` | `edit_mode(wb)`: activate the scratch sheet, `SendKeys "{F2}x"`; afterwards `SendKeys "{ESC}"`. | Same as above. |
| `S8.parallel` | Two `McpClient`s run `sv_ws_build` at the same time (threads) on different questions into different scratch workbooks. | Both pass. |
| `S8.kill.excel` | During `sv_query` (a thread), kill the engine's hidden Excel PID, read from `sv_status`. | Error naming `sv_restart`, or a retry that succeeds. |
| `S8.kill.edge` | During `sv_drill`, `kill_engine_edge()` after 5 s. | Retry succeeds, or a `DrillError` with a next step. |
| `S8.cold` | Close S2's workbook, then `sv_drill` with that workbook as the target. | Passes within 240 s. |
| `S8.sleep` | `psutil.Process(server_pid).suspend()` for 90 s during `sv_query`, then `resume()`. | Recovers within 120 s after the resume. |

- [ ] **Step 1:** Write the helpers and scenarios.
- [ ] **Step 2:** Run `matrix.py --only inject`.
- [ ] **Step 3:** Check by hand that no `*.matrix-bak` file remains, that no dialog is open, and that Excel/Edge still
  work.
- [ ] **Step 4:** Commit: `test(live): S8 bad-condition injections`.

### Task 5: S9 fresh install and S10 features

**Files:**
- `engine\smartview_mcp\setup_user.py` (the PATH override);
- `engine\tests\test_setup_user.py`;
- `scenarios.py`;
- `helpers.py` (part 3).

**S9 needs one engine change first.** `path_step()` writes the real `HKCU\Environment`. Add an env override.

- [ ] **Step 1: Failing test.**

```python
def test_path_step_honours_registry_override(tmp_path, monkeypatch):
    fake = tmp_path / "path.txt"
    monkeypatch.setenv("SMARTASS_USER_PATH_FILE", str(fake))
    scripts = tmp_path / "Scripts"; scripts.mkdir(); (scripts / "smartass.exe").write_text("")
    out = setup_user.path_step(str(scripts / "python.exe"))
    assert str(scripts) in fake.read_text(encoding="utf-8")
    assert "added" in out or "kept" in out
```

- [ ] **Step 2:** Implement. In `_read_user_path` and `_write_user_path`: when `SMARTASS_USER_PATH_FILE` is set, read
  and write that file (`kind` = `REG_EXPAND_SZ` = 2) instead of the registry, and skip `_broadcast_env_change`. Run
  the test, then commit: `feat(setup): SMARTASS_USER_PATH_FILE override for fresh-profile tests`.

- [ ] **Step 3: `helpers.fresh_profile()`.** Port the simulation from `docs\plans\2026-10-22-smartass-repo.md` Task 16
  Step 2:
  - temporary `LOCALAPPDATA`, `APPDATA`, `USERPROFILE`, `OneDriveCommercial`, `SMARTASS_RELEASE`;
  - `SMARTASS_HTTP_PORT=47899` and `SMARTASS_USER_PATH_FILE`;
  - read `UV_CACHE_DIR` and `UV_PYTHON_INSTALL_DIR` first;
  - a `git archive` of HEAD as the release, plus a copy of the real `atlas.db`.

  It returns an env dict. The install runs in a child `powershell -File installer\install.ps1` with that env, so
  nothing leaks into the parent.

- [ ] **Step 4: S9 scenarios.** Both have `fresh=True` and run only with `--fresh`.
  - `S9.install`: the install succeeds, and `smartass doctor` (from the temp venv, with the same env) prints OK on every
    line.
  - `S9.lookup`: a second `McpClient(python=<temp venv python>, env=env)` runs `S1.find` and `S1.query.finmax`.

- [ ] **Step 5: S10 scenarios.** Tags `feature`; one ID per item in spec S10.

| ID | Check |
|---|---|
| `S10.1.bundle` | `smartass doctor --bundle` → open the zip and search every member for the `sso_token` value and `http.secret` (read from config). Neither is found. |
| `S10.2.setup` | In the fresh env: run `smartass setup` twice. The second output is all `kept`/`unchanged`. `mcp-config.json` has `smartass` and no `smartview`. To test the legacy migration, write a `smartview` entry first and run setup: it becomes `smartass`. |
| `S10.3.reload` | In a scratch client: touch `smartview_mcp\config.py` (`os.utime`). Within 30 s, exactly one `notifications/tools/list_changed` in `client.notes`, and `sv_status` still answers. |
| `S10.4.update` | In the fresh env: `latest.json` 0.2.1 → `smartass update` applies it. A second update within 30 min is refused by the loop guard. `[update] pin = "0.2.0"` keeps 0.2.0. |
| `S10.5.textfixes` | A user `text_fixes.json` with a unique replacement → `sv_ws_build` key points show it. Delete the file afterwards. |
| `S10.6.pack` | `sv_scope` returns produktregnskap knowledge (a `pack` field or a known concept from the pack). |
| `S10.7.style` | S2's result has a `formula_style` field and TN styles in the workbook (`xl_overview` style names contain `TN`). |
| `S10.8.coldexcel` | `sv_restart`, close the user's scratch Excel, `sv_ws_build` into a new workbook: within 240 s, `Visible=True`. |
| `S10.9.options` | Read the Smart View options (`helpers.sv_options()` through the registry/ini the engine uses), run `sv_restart`, read again: unchanged. |
| `S10.10.http` | `GET http://127.0.0.1:<port>/status` without the secret → 401/403; with the `X-Secret` header (name from `httpapi.py`) → 200. |

- [ ] **Step 6:** Run `matrix.py --only feature` and `matrix.py --fresh --only fresh`. Commit:
  `test(live): S9 fresh install and S10 feature checks`.

### Task 6: Consolidate the old live scripts

- [ ] **Step 1:** Delete the scripts that are now covered:
  - `live_smoke.py`, `live_build.py`, `live_analysis.py`, `live_drill.py`, `live_drill_report.py`, `live_xl.py`,
    `live_refresh.py`, `live_sessions.py`, `live_relay.py`, `verify_gl.py`;
  - the one-offs `diag_*`, `probe_finplan.py`, `inspect_*`, `diff_sheet.py`, `render_sheet.py`, `compare_numbers.py`,
    `build_atlas.py`, `close_test_books.py`.

  Before deleting a script, `grep` the scenarios and helpers for anything they still import from it. Move what is
  needed into `helpers.py` first.
- [ ] **Step 2:** In `test_styling_live.py`, replace the `svkit\work` path with `ctx.scratch`, or
  `config.ANALYSER_DIR / "_matrix"`.
- [ ] **Step 3:** Rewrite `tests\live\README.md`: one command, `uv run python tests\live\matrix.py`, the options
  table, and the safety rules.
- [ ] **Step 4:** Run pytest and ruff, then commit: `test(live): consolidate live scripts into the matrix`.

## Phase B — Baseline

### Task 7: Baseline run

- [ ] **Step 1:** Close the scratch workbooks. Run `uv run python tests\live\matrix.py --out matrix-baseline`, then
  `--fresh --only fresh --out matrix-baseline-fresh`.
- [ ] **Step 2: Rank the failures.** Group by `error_class` and rank by count × severity:
  wrong number (5) > hang (4) > failure with no next step (3) > failure with a next step (2) > slow (1).
  Put the ranked list at the top of `docs\audit\<date>-matrix-baseline.md` as "Fix order".
- [ ] **Step 3:** Commit: `audit: baseline matrix`.
- [ ] **Step 4: Tell the owner the fix order** (a short table) before Phase C. Tasks 8–13 are then done in that order.
  Expected fixes that the baseline does **not** reproduce are dropped, with a note in the report.

---

## Phase C — Fixes (test-first; order from Task 7)

### Task 8: Dimension names (quotes and unique aliases)

**Files:** `smartview_mcp\ops.py` (`canon_dim`) and `tests\test_logic.py`.

- [ ] **Step 1: Failing tests.**

```python
from smartview_mcp import ops


def _seed():
    labels = ["All Accounts", "All Time Periods", "All Entities", "All Balance Types", "All Scenarios"]
    ops._aliases["T"] = {**{l: l for l in labels}, **{l.removeprefix("All ").replace(" ", ""): l for l in labels}}


def test_canon_dim_strips_quotes_and_space():
    _seed()
    assert ops.canon_dim("T", "'All Balance Types'") == "All Balance Types"
    assert ops.canon_dim("T", ' "Accounts" ') == "All Accounts"


def test_canon_dim_unique_prefix_word_and_singular():
    _seed()
    assert ops.canon_dim("T", "Balance") == "All Balance Types"
    assert ops.canon_dim("T", "Account") == "All Accounts"
    assert ops.canon_dim("T", "entity") == "All Entities"
    assert ops.canon_dim("T", "Period") == "All Time Periods"


def test_canon_dim_ambiguous_returns_input():
    _seed()
    ops._aliases["T"]["All Accounts Alt"] = "All Accounts Alt"
    ops._aliases["T"]["AccountsAlt"] = "All Accounts Alt"
    assert ops.canon_dim("T", "Acc") == "Acc"
```

- [ ] **Step 2:** Run. Expected: FAIL on quotes and `Balance`.
- [ ] **Step 3: Implement.**

```python
def _dim_key(s: str) -> str:
    s = s.strip().strip("'\"").strip().lower().removeprefix("all ")
    s = s.replace(" ", "")
    return s[:-3] + "y" if s.endswith("ies") else s.rstrip("s")


def canon_dim(conn: str, name: str) -> str:
    """Accept 'All Accounts', 'Accounts', "'All Balance Types'", 'Balance', 'Account', 'entity' ... -> the grid label.
    Falls back to a unique match on the start of the label without 'All '; ambiguous or unknown -> returned as given."""
    amap = _aliases.get(conn, {})
    clean = name.strip().strip("'\"").strip()
    if clean in amap:
        return amap[clean]
    k = clean.removeprefix("All ").replace(" ", "")
    if k in amap:
        return amap[k]
    want = _dim_key(clean)
    labels = sorted(set(amap.values()))
    hits = [lab for lab in labels if _dim_key(lab).startswith(want)]
    if not hits:
        hits = [lab for lab in labels if any(w.lower().rstrip("s").startswith(want)
                                              for w in lab.removeprefix("All ").split())]
    return hits[0] if want and len(hits) == 1 else clean
```

- [ ] **Step 4:** Run the tests until they pass. Then make the correction visible:
  - `validate_axes` already lists the valid dimensions.
  - Add a warning to the result when `canon_dim(x) != x.strip()`, at the two call sites `ops.py:251–253` and
    `userexcel.py:882–884`. The text: `dimension 'Balance' read as 'All Balance Types'`. Use the existing warnings list
    in each function; find it with grep.
- [ ] **Step 5:** Run the full pytest, then `matrix.py --only input`. Commit:
  `fix(query): tolerate quoted and short dimension names`.

### Task 9: Rows/columns given as a dict

**Files:** `smartview_mcp\server.py`, `tools_query.py`, `tools_ws.py` and `tests\test_logic.py`.

- [ ] **Step 1: Failing test.**

```python
from smartview_mcp.server import AxisIn, axes_in


def test_axes_in_accepts_dict_list_and_models():
    d = axes_in({"All Accounts": ["ac.C30"], "All Time Periods": "Sep 26"})
    assert [(a.dimension, a.members) for a in d] == [("All Accounts", ["ac.C30"]), ("All Time Periods", ["Sep 26"])]
    l = axes_in([{"dimension": "All Accounts", "members": ["x"]}, AxisIn(dimension="E", members=["y"])])
    assert [a.dimension for a in l] == ["All Accounts", "E"]
```

- [ ] **Step 2: Implement in `server.py`** under `AxisIn`.

```python
AxesIn = list[AxisIn] | dict[str, list[str] | str]


def axes_in(x) -> list[AxisIn]:
    """rows/columns as [{dimension, members}] (documented) or {dimension: [members]} (what agents often send)."""
    if isinstance(x, dict):
        return [AxisIn(dimension=d, members=[m] if isinstance(m, str) else list(m)) for d, m in x.items()]
    return [a if isinstance(a, AxisIn) else AxisIn(**a) for a in x]
```

- [ ] **Step 3: Use it.**
  - Change the annotations `rows: list[AxisIn]` / `columns: list[AxisIn]` to `AxesIn` in `sv_query`
    (`tools_query.py:28`), `sv_ws_pull` (`tools_ws.py:186`) and `PullIn` (`tools_ws.py:379–380`).
  - Normalise at the first line of the body: `rows, columns = axes_in(rows), axes_in(columns)`.
  - In `PullIn`, add a pydantic `field_validator("rows", "columns", mode="before")` that returns
    `[a.model_dump() for a in axes_in(v)]` when `v` is a dict.
- [ ] **Step 4:** Run `test_smoke_stdio.py`: the tool schema still lists the documented list form first. Run the full
  pytest and ruff, then `matrix.py --only input`. Commit: `fix(query): accept rows/columns as {dimension: members}`.

### Task 10: Spec shape validation

**Files:** `smartview_mcp\specs.py` and `tests\test_specs.py`.

- [ ] **Step 1: Failing tests.**

```python
import pytest

from smartview_mcp.host import InputError
from smartview_mcp.specs import validate_shape


def test_validate_shape_names_the_bad_path():
    spec = {"columns": [{"header": "A"}], "sections": [{"title": "s", "lines": [
        {"key": "a", "label": "A", "cells": [{"pull": "p", "row": ["x"], "col": ["y"]}, "oops"]}]}]}
    with pytest.raises(InputError, match=r"sections\[0\]\.lines\[0\]\.cells\[1\] must be an object or null, got str"):
        validate_shape(spec)


def test_validate_shape_accepts_repeat_and_null_cells():
    spec = {"columns": [{"header": "A"}], "sections": [
        {"title": "s", "repeat": {"over": "months:Jan..Mar", "line": {"key": "m{i}", "label": "{v}", "cells": [None]}},
         "lines": []}]}
    validate_shape(spec)


def test_validate_shape_rejects_string_sections():
    with pytest.raises(InputError, match=r"sections\[0\] must be an object"):
        validate_shape({"columns": [], "sections": ["x"]})
```

- [ ] **Step 2: Implement** (the checks are shape-only; `expand` keeps the semantic checks).

```python
def validate_shape(spec) -> None:
    """Raise InputError naming the first wrongly shaped key (instead of "'str' object has no attribute 'get'")."""
    def need(cond, path, what, got):
        if not cond:
            raise InputError(f"{path} must be {what}, got {type(got).__name__} (fix this key in the spec)")
    need(isinstance(spec, dict), "spec", "an object", spec)
    for k in ("columns", "sections", "checks", "key_points", "notes"):
        if k in spec and spec[k] is not None:
            need(isinstance(spec[k], list), k, "a list", spec[k])
    for i, c in enumerate(spec.get("columns") or []):
        need(isinstance(c, dict), f"columns[{i}]", "an object", c)
    for i, s in enumerate(spec.get("sections") or []):
        need(isinstance(s, dict), f"sections[{i}]", "an object", s)
        lines = list(enumerate(s.get("lines") or []))
        if isinstance(s.get("repeat"), dict) and isinstance(s["repeat"].get("line"), dict):
            lines.append(("repeat.line", s["repeat"]["line"]))
        for j, ln in lines:
            p = f"sections[{i}].lines[{j}]" if isinstance(j, int) else f"sections[{i}].{j}"
            need(isinstance(ln, dict), p, "an object", ln)
            cells = ln.get("cells") or []
            need(isinstance(cells, list), f"{p}.cells", "a list", cells)
            for n, c in enumerate(cells):
                need(c is None or isinstance(c, dict), f"{p}.cells[{n}]", "an object or null", c)
    for i, c in enumerate(spec.get("checks") or []):
        need(isinstance(c, dict), f"checks[{i}]", "an object", c)
```

- [ ] **Step 3:** Call `validate_shape(spec)` as the first line of `prepare` and of `expand`. Find the import of
  `InputError` in `specs.py`; it may already exist.
- [ ] **Step 4:** Run the full pytest. Commit: `fix(answer): validate spec shape and name the bad key`.

### Task 11: Drill

**Files:** `smartview_mcp\ops.py` (`drill_link`), `drill_browser.py` and `tests\test_drill.py`.

- [ ] **Step 1: Failing tests** for the logic that needs no Excel:
  - **(a) Default for a missing `All Entities`.** Extract a pure function `drill_cell_defaults(cell) -> (cell,
    warnings)`:
    - a missing `All Entities` gets the top member, `"All Entities"`, with the warning
      `"All Entities not given - drilled the top member; name an entity to narrow it"`;
    - a missing account or period still raises `InputError`.
  - **(b) No-report wording.** `drill_link`, given a fake host whose `run` raises
    `HostError("... No drill-through reports found ...")`, must now raise
    `InputError("This cell has no drill-through report: it is not a FinLite base-level intersection (use a FinLite
    account and a base entity/period).")`.

    Today it returns `url=None`. Find the caller in `tools_drill.py` and check how it handles `url=None` before
    changing this: if the caller already converts `url=None` into a clear message, test that message instead.
  - **(c) "Never asked" classification.** Extract `classify_stall(final_url: str, page_text: str) -> str`:
    - `"sso"` when the URL contains `login.microsoftonline`, `/oauth2/` or `SIGNIN`;
    - `"blank"` when the page text is empty;
    - `"page"` otherwise.

    Each class maps to a `DrillError` message: SSO → "Sign-in to EPM did not complete — open EPM once in Edge and log
    in, then retry"; blank/page → "The drill page loaded but never requested data — retry; if it repeats, …".
- [ ] **Step 2:** Implement (a)–(c) and wire them in:
  - (a) at the top of `drill_link`, with its warnings added to the tool result;
  - (c) in the capture loop at `drill_browser.py` ~495–527, where "never asked for its data" is raised. Before raising
    for `blank`/`page`, do **one** page reload and wait again (the spec's "reload once"). Use the reload method of the
    page object already used in that loop.
- [ ] **Step 3: Launch retry.** `_launch` already retries `DRILL_LAUNCH_TRIES`. Check that the baseline's "browser
  closed during launch" class still occurs:
  - if it does, add a check for a stale `SingletonLock` in the profile dir before each retry, deleting it if no engine
    Edge PID holds it, plus a unit test on that decision function;
  - if it does not, note "not reproduced" for the report.
- [ ] **Step 4:** Run the full pytest, then `matrix.py --only drill --repeat 5` ("Done when" #2). Commit:
  `fix(drill): entity default, clear no-report and stall errors, reload once`.

### Task 12: COM gaps (baseline-dependent)

Follow this procedure for each COM error class in the baseline ("Call was rejected", RPC failed, `-2147352567`,
`0x800AC472`):

1. Find the stage from `server.log` (`stage > …`) and the matrix error.
2. Check whether that code path goes through the existing retry: `host.RETRY_HRESULTS` (host.py:30, used at ~260 and
   ~391), `userexcel.py:38` or `styling.py:26`.
3. **If it does not:** route it through the same helper. A unit test asserts, with a fake COM object raising
   `pywintypes.com_error(hresult, ...)`, that it retries once and then raises a `HostError` whose text says
   `Excel is busy or a dialog is open - close it and retry`.
4. **If it does but still fails:** the busy check is wrong. Add a "ready" wait (`Application.Ready`, polling for up to
   10 s) before the retry, again with a fake-object test.
5. Check that partial sheets are removed on failure. A test drives the failing build on a fake host and asserts that
   no new sheet remains, or the matrix scenario asserts it with `xl_overview`.

Commit each class separately: `fix(com): <class> retried / named`.

### Task 13: Speed

**Files:** `userexcel.py` (open target), `styling.py` (style check), `timing.py`, `server.py` (progress).

- [ ] **Step 1: Progress notes.**
  - Failing test (`tests\test_logic.py`): with a fake `notify` callback and `timing.PROGRESS_AFTER_S = 0.05`, a
    `stage("slow")` that sleeps 0.1 s emits one progress note naming `slow`.
  - Implement in `timing.py`: `Timer` takes an optional `notify(text)`. `stage()` starts a `threading.Timer` for
    `PROGRESS_AFTER_S = 10`, which calls `notify(f"still working: {full} ({secs:.0f}s)")` and re-arms every 10 s; it is
    cancelled in `finally`.
  - In `server._call` / `_call_user`, pass a notify that schedules `ctx.report_progress(...)` on the event loop
    (`loop.call_soon_threadsafe`). Check that the MCP library's `Context` has `report_progress` (or the equivalent) with
    `grep -rn "def report_progress" .venv\Lib\site-packages\mcp`.
- [ ] **Step 2: Open-target tail.** In the baseline timings, find `open target` at p90. In `userexcel.py` (grep
  `open target`):
  - if the workbook is already open in any Excel, reuse it without `Workbooks.Open`;
  - skip `CalculateFull` when the target has no formulas referring to the new sheets (it is a new sheet).

  Add a unit test for any decision helper you extract.
- [ ] **Step 3: Style cache.** In `styling.py`, cache per workbook FullName: `{name: True}` after the TN styles have
  been confirmed or created once in this server process. Unit test: the second call does not touch `wb.Styles`. Use a
  fake that counts calls.
- [ ] **Step 4:** Run `matrix.py --only build --repeat 3`. Gate: the median of `S2.build` ≤ 30 s ("Done when" #4).
  Commit: `perf: progress notes, open-target reuse, style cache`.

---

## Phase D — Trim, verify, personas, report

### Task 14: Remove the ad hoc tools and `sv_show`

- [ ] **Step 1:** Delete `sv_adhoc_open`, `sv_adhoc_grid`, `sv_adhoc_set_cells`, `sv_adhoc_action` and
  `sv_adhoc_option` (tools_query.py ~166–286), with `_sheet` and `_grid_preview` if nothing else uses them (grep). Also:
  - fix the module docstring ("advanced ad hoc API");
  - delete `sv_show` (tools_status.py ~90) and its mention in that file's docstring;
  - remove the ad hoc ops in `ops.py` that only these tools used (grep each function for callers). **Keep** the VBA
    `B_*` macros and `sv_member_info`.
- [ ] **Step 2:** `grep -rn "sv_adhoc\|sv_show" C:\dev\smartass` (engine, skills, docs, README, `server.py`
  INSTRUCTIONS, `NEXT_STEP`). Remove or reword every hit; the dated `docs/specs` and `docs/plans` are history, so leave
  them.
- [ ] **Step 3:** Add a test in `test_smoke_stdio.py`: `tools/list` has exactly 30 tools, and none starts with
  `sv_adhoc` or equals `sv_show`.
- [ ] **Step 4:** Run the full pytest, ruff and the guard. Commit: `refactor: remove unused ad hoc tools and sv_show (30
  tools)`.

### Task 15: Full verification run

- [ ] **Step 1:** `uv run python tests\live\matrix.py --out matrix-after`. It writes the before/after column against
  the baseline.
- [ ] **Step 2:** `uv run python tests\live\matrix.py --only drill --repeat 5 --out matrix-drill5`.
- [ ] **Step 3:** `uv run python tests\live\matrix.py --fresh --only fresh --out matrix-after-fresh`.
- [ ] **Step 4: Gate.** Every scenario is PASS or LIMIT, and each LIMIT has a reason the owner accepts. Ask the owner
  with `ask_user` for each LIMIT. S2.build median ≤ 30 s, no `HANG`. For anything else, go back to Phase C with a
  failing test first.
- [ ] **Step 5:** Commit: `audit: matrix after fixes`.

### Task 16: Persona pass

- [ ] **Step 1:** Launch four `general-purpose` subagents in parallel. Each prompt contains:
  - only: "Use the smartview skill (`C:\dev\smartass\skills\smartview\SKILL.md`) and the smartass MCP tools";
  - the one Norwegian question from spec Part 4;
  - "answer in a new workbook in `<analyser>\_matrix`";
  - "report: the tools you called, in order, every tool error verbatim, the final numbers".
- [ ] **Step 2: Score each run.**
  - The path against the expected tools: controller `sv_find_member` → `sv_ws_build`; product owner `sv_scope` →
    `sv_concept`; analyst `sv_drill`; planner `sv_analyse actual_vs_plan`.
  - Errors ≤ 3.
  - Time, from the session-store `tool_executions` for those sessions, or the agent's own timings.
  - Correctness against the matrix numbers or a direct `sv_query`.
- [ ] **Step 3:** Fix the unclear skill text or tool docstrings each failure points to, and re-run that persona
  **once**. Commit: `docs(skill): clarify <topic> after persona pass`.

### Task 17: Report, docs, push

- [ ] **Step 1:** Write `docs\audit\2026-10-<dd>-reliability-audit.md`:
  - the 30-day evidence table (from the spec);
  - the matrix before → after (link both `.md` files plus a summary count);
  - the fixes with their commit SHAs (`git log --oneline main..`);
  - the persona results table;
  - accepted limits;
  - the roadmap, at most 10 lines ranked as change / build / remove / speed up. It covers colleague rollout via the
    TelenorNorgeInternal repo transfer, the Excel add-in, Azure OpenAI and anything larger found.
- [ ] **Step 2:** Update `skills\smartview\SKILL.md` and `reference\operations.md` for the removed tools, the new
  input tolerance and the new error messages and next steps.
- [ ] **Step 3:** Run the full pytest, ruff and the guard, then commit: `audit: reliability audit report and roadmap`.
- [ ] **Step 4:** Use superpowers:finishing-a-development-branch: merge to main and push. Then set `GH_TOKEN` (see
  `docs\plans\2026-10-22-smartass-repo.md`) and check that CI is green with `gh run list -R tomchoel/smartass -L 1`.

---

## Spec coverage

| Spec item | Task |
|---|---|
| Part 1: harness, options, safety, consolidation | 1, 2, 3, 6 |
| S1–S7 | 3 |
| S8 | 4 |
| S9, S10 | 5 |
| Part 3: baseline and ranking | 7 |
| Expected fix 1 (drill) | 11 |
| Expected fix 2 (input) | 8, 9, 10 |
| Expected fix 3 (COM) | 12 |
| Expected fix 4 (speed) | 13 |
| Expected fix 5 (trim) | 14 |
| Done when 1–5 | 15 |
| Part 4: personas | 16 |
| Report, docs, CI ("Done when" 6) | 17 |
