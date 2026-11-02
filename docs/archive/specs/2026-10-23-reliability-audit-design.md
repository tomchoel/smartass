# Reliability audit of the smartass engine - design

Date: 2026-10-23 · Status: approved in brainstorming · Owner: tomchoel

## Why

The engine works, but the last 30 days of use (Copilot/Claude session history plus `queries.jsonl` and `server.log`)
show it fails too often in real work:

| Tool | Calls | Failed | p50 | p90 | Main causes |
|---|---|---|---|---|---|
| `sv_drill` | 20 | 15 | 67 s | 133 s | Browser closed during launch (3), "drill page never asked for its data within 45 s" (8), missing dimension in `cell`, "No drill-through reports found", cold Excel open |
| `sv_ws_build` | 20 | 5 | 21 s | 56 s | COM "Exception occurred", `'str' object has no attribute 'get'` (spec shape), unknown dimension `Balance` |
| `sv_ws_answer` | 14 | 2 | 6 s | 24 s | "No pulls for Q<n>", COM errors |
| `sv_query` | 46 | 3 | 1.4 s | 4 s | Dimension names with quotes rejected, `rows` passed as a dict |

- A whole answer takes 30 s at the median and 82 s at the 90th percentile. Opening the target workbook has a 31 s tail (p90).
- 15 of 36 tools were never called.

Several of these failures were fixed during sub-project 1, but nothing proves the fixes end to end, and the
roughly 20 live scripts in `tests\live` are separate one-off checks with no shared result.

## Goal

A reliability round. Every user workflow is run live and scored, the failures are fixed in order of impact,
the tool list is trimmed, and the result is written down in a repeatable audit. The roadmap is only a short
ranked list in the report; this round builds no new features.

## Done when

1. Every scenario in the matrix (S1–S10) passes live, or has a documented and accepted limit.
2. Drill: 5 consecutive live drills succeed, cold and warm, each with reconciliation `exact`.
3. Every failure class in the table above is reproduced, then either fixed or turned into an error that names
   the next step.
4. Speed:
   - A standard answer (S2 Q1) takes 30 s or less at the median.
   - No stage runs over 60 s without a progress note.
   - No call hangs past 120 s.
5. Every fixed bug has an automated test (CI green), and `tests\live\matrix.py` reruns the whole matrix in one command.
6. An audit report is committed: the matrix before and after, the persona results, accepted limits, and the roadmap.

## Part 1 - the harness

- **`tests\live\mcpclient.py`.** Stdio JSON-RPC client, built from `live_relay.py`:
  - starts `run_server.py` exactly as Copilot/Claude do;
  - `call(tool, args, timeout)` returns `(body, seconds, error)`;
  - logs notifications (`tools/list_changed`).
- **`tests\live\matrix.py`.** The runner:
  - **Scenarios** come from a registry of small functions (`@scenario("S5.cold", tags=["drill"])`). Each makes tool
    calls and asserts on the result: numbers, sheet names, `reconciliation`, `checks`, `formula_style.ok`, `timings`.
  - **Result per scenario:** `PASS`, `SLOW` (passed but over its time budget), `FAIL`, or `LIMIT` (a known, accepted limit, with its reason).
  - **Error class:** the first line of the error with digits normalised, so runs can be grouped and compared.
  - **Output:** `docs\audit\<date>-matrix.json` and `.md`. When an earlier JSON exists, it adds a before/after column.
  - **Options:**
    - `--only <tags>`;
    - `--repeat N`;
    - `--fresh` (S9 in a temporary user profile, reusing the fresh-user simulation from Task 16);
    - `--keep` (leave the scratch workbooks open for inspection).
- **Consolidation.** The useful checks in `live_smoke.py`, `live_build.py`, `live_analysis.py`, `live_drill.py`,
  `live_drill_report.py`, `live_xl.py`, `live_refresh.py`, `live_sessions.py`, `live_relay.py` and `verify_gl.py`
  become scenarios. One-off diagnostics (`diag_*`, `probe_finplan.py`, `inspect_*`, `diff_sheet.py`,
  `render_sheet.py`, `compare_numbers.py`, `build_atlas.py`, `close_test_books.py`) are deleted unless a scenario
  needs them as a helper; those move into `tests\live\helpers.py`. `test_styling_live.py` stays as a pytest live test and is
  repointed from `svkit\work` to the smartass folder.
- **Safety.**
  - Read-only to EPM.
  - Scratch workbooks go in `<analyser>\_matrix\` and are deleted at the end unless `--keep`.
  - Never touches workbooks it did not open.
  - Restores the SSO token and closes any dialog it injected, even when a scenario fails (`finally`).
- **Bad-condition injection (S8).** Each one is done on purpose, inside the scenario:
  - SSO expired: rename the cached token and the Edge profile cookie store for the duration.
  - Modal dialog: open a `MsgBox` in the target Excel via COM from a helper thread.
  - Edit mode: `SendKeys F2` into a cell.
  - Two clients: two `mcpclient` instances calling `sv_ws_build` in parallel on different questions.
  - Kill mid-call: terminate the engine's own Excel or Edge PID during `sv_query` / `sv_drill`.
  - Cold OneDrive file: close the session workbook and drill into it, which forces the cold open.
  - Sleep: freeze the engine process with `psutil.suspend` for 90 s during a call, then resume.

## Part 2 - scenarios and pass criteria

| ID | Scenario | Pass |
|---|---|---|
| S1 | Lookups. `sv_find_member` (an account, an entity, a product, "man-years"), `sv_members`, `sv_query` on FinMax, FinLite and FinPlan. Bad input: dimension `Balance`, quoted names, dict `rows`, unknown member. | Numbers are correct. Each call takes 5 s or less. Bad input is either corrected (with a warning) or rejected with a message naming the valid choice. |
| S2 | Free-form build. Q1 `sv_ws_build`: FinMax, two pulls (2026 and 2025 YTD Sep). Then on the same Q: `sv_ws_answer` `spec_patch`, `sv_ws_build` adding a FinPlan pull. | `formula_style.ok`, all `checks` OK, no orphan sheets, Q1 takes 30 s or less, sheet order is correct. |
| S3 | Templates `division_pnl`, `monthly_trend`, `actual_vs_plan`, `cost_base`, on `un.SY.2TEC` and on `un.SY` (or `BANORWAY` for FinPlan). | Built-in checks OK. The bridge or total reconciles to the pull. |
| S4 | `sv_scope("hva tjener vi på mobilabonnement")`, then `sv_concept` on the matched `fam_*` concept and on `splitt`. | Concept found. Blocks reconcile: revenue + COGS + opex = EBITDA before other items. |
| S5 | Drill. FinLite `ac.12881` × `un.SY` × `Sep 26`, cold (browser closed) and warm; a quarter cell `Q3 26` that splits itself; then a report via `sv_ws_answer` with `drill` cells. | 95 lines, gap 0, `exact`. The split reconciles. The report builds. `--repeat 5` all pass. |
| S6 | `xl_*` on a scratch workbook: `xl_read`, `xl_write`, `xl_check`, `xl_trace`, `xl_format_sheet`, `xl_revert`, `xl_overview`, `xl_changes`. | Revert restores every value and format exactly. The file is never saved. |
| S7 | `sv_refresh_workbook` on the saved S2 workbook. | Values unchanged. No "Call was rejected". |
| S8 | The injections from Part 1. | Recovers by itself, or the error names the next step. No hang past 120 s. The engine still works afterwards (an `sv_query` follow-up passes). |
| S9 | `matrix.py --fresh`: install into a clean profile, `smartass doctor`, then S1. | Doctor all OK. S1 passes. |
| S10 | Sub-project 1 features (list below). | Each item passes. |

S10 items:
1. `smartass doctor` and `smartass doctor --bundle`: neither the `sso_token` value nor `http.secret` appears in the zip.
2. Running `smartass setup` twice gives the same result: user PATH, startup shortcuts, MCP entries named `smartass`, and the legacy `smartview` entry migrated.
3. Relay reload: touching `config.py` gives exactly one restart and `tools/list_changed`.
4. Update path against a fake `release.folder` with a newer `latest.json`:
   - the update is applied;
   - the 30-minute loop guard holds;
   - a version pin is honoured.
5. A user `text_fixes.json` overrides the shipped one.
6. The `produktregnskap` pack is loaded (its knowledge is visible in `sv_scope`).
7. The vendored `tn_rapportstil` and `formelsjekk` are used: `formula_style` is present and the TN styles are applied.
8. Cold Excel open within the 240 s budget, with a visible window.
9. Smart View options: the originals are restored after shutdown, with the legacy fallback.
10. The HTTP API returns 401/403 without the secret and 200 with it.

Last comes the **persona pass**, described in Part 4.

## Part 3 - the fix loop

1. Run the full matrix as the baseline and commit `docs\audit\<date>-matrix-baseline.*`.
2. Group the failures by error class and rank them by frequency × severity. Severity: wrong number > hang > failure with no
   next step > failure with a next step > slow.
3. Fix test-first:
   - A unit test where the logic can be tested without Excel (input normalisation, retry decisions,
     error wording, split logic).
   - Otherwise the matrix scenario is the regression check.
4. Re-run the affected tags after each fix, and the full matrix at the end.

Fixes expected from the evidence (the baseline confirms or drops each one):

1. **Drill.**
   - Browser closed during launch: retry once with a fresh profile lock check.
   - "Never asked for its data": detect the SSO redirect or a blank page, reload once, then give an error saying whether the login or the page failed.
   - Missing dimensions in `cell`: fill them with the cube's top member and add a warning.
   - "No drill-through reports": give an error explaining that the cell is not a FinLite base-level intersection.
2. **Input tolerance** in `sv_query`, `sv_ws_pull` and `sv_ws_build`:
   - Strip quotes from dimension names.
   - Resolve aliases (`Balance` → `All Balance Types`, `Account` → `All Accounts`, and so on, per cube from the atlas).
   - Accept `rows` / `columns` given as a dict `{dimension: [members]}`.
   - Validate the spec shape and name the offending key, instead of crashing with `'str' object has no attribute 'get'`.
3. **COM.**
   - Errors: "Call was rejected", RPC failed, `-2147352567` "Exception occurred", OLE 0x800AC472.
   - Wait until Excel is ready, then retry once. Otherwise report the next step (a dialog is open / Excel is busy).
   - Partial sheets are still removed.
4. **Speed.**
   - Opening the target workbook: reuse the open window, and skip the full recalculation when nothing depends on it.
   - Styling: cache the style check per workbook.
   - Progress notes (MCP progress notifications) for stages over 10 s.
5. **Trim.**
   - Remove `sv_adhoc_open`, `sv_adhoc_grid`, `sv_adhoc_set_cells`, `sv_adhoc_action`, `sv_adhoc_option` and `sv_show`, with their code paths, tests and skill/README mentions. That leaves 30 tools.
   - The `xl_*` tools stay.

Out of scope: new tools, new templates, colleague rollout, the Excel add-in, Azure OpenAI. These go on the roadmap.

## Part 4 - persona pass and report

- **Persona pass.** Four subagents. Each gets only the smartview skill and the smartass MCP, plus one question in Norwegian:
  1. controller: "Hva brukte Consumer på konsulenter i Q3 mot i fjor?"
  2. product owner: "Hva tjener vi på mobilabonnement?"
  3. analyst: "Hvilke leverandører ligger bak 12881 i september?"
  4. planner: "Hvordan ligger Technology an mot target hittil i år?"

  Scored on:
  - tools chosen against the expected path;
  - number of tool errors (pass: 3 or fewer);
  - time;
  - whether the answer is correct, checked against the matrix numbers.

  Unclear skill text and tool descriptions found this way get fixed and the run is repeated once.
- **Report `docs\audit\2026-10-<dd>-reliability-audit.md`:**
  - the matrix before and after;
  - the 30-day evidence;
  - the fixes, with their commits;
  - the persona results;
  - accepted limits;
  - **roadmap**: a short ranked list (change / build / remove / speed up), covering colleague rollout, the add-in and any
    larger items found.
- **Docs.** `tests\live\README.md` is reduced to one command, `uv run python tests\live\matrix.py`, plus its options. The skill
  and `operations.md` are updated for the removed tools and the new error messages.

## Order

Harness → baseline run → fixes (in ranked order) → trim → full matrix plus `--repeat 5` drill → persona pass → report →
push, CI green.

## Risks

- **Live EPM and SSO state varies from day to day.** The matrix records the time and the atlas age, and a flaky scenario is re-run once
  before it counts as a FAIL. A scenario that passes only on the re-run is recorded as flaky.
- **Injections touch the real Excel and Edge.** They act only on PIDs the engine started, or on scratch workbooks, and are
  undone in `finally`.
- **Persona runs vary.** Their results guide the documentation fixes; the pass/fail gate is the matrix.
