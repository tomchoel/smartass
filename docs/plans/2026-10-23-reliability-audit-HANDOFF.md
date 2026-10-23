# Handoff: reliability audit (DONE 2026-10-23, see docs/audit/2026-10-23-reliability-audit.md)

**Branch:** `impl/reliability-audit` (not pushed, not merged).

**Plan:** `docs/plans/2026-10-23-reliability-audit.md`.

**Spec:** `docs/specs/2026-10-23-reliability-audit-design.md`. "Done when" 1–6 there is the acceptance gate.

**Baseline:** `docs/audit/2026-10-23-matrix-baseline.md`. Its "Fix order" section was approved by the owner.

**Method:** subagent-driven development: implementer → spec/code review → fixes → re-review. Fixes are test-first. Don't stop between tasks.

## Done
- **Phase A (Tasks 1–6):** the harness `engine/tests/live/{matrix,scenarios,helpers,mcpclient}.py`, scenarios S0–S10, and the old scripts removed.
- **Phase B (Task 7):** the baseline is committed (`87e83e5`). Tasks 8 and 10 were dropped because the baseline didn't reproduce them.
- **Fix order item 1, the Smart View Login/WebView2 dialog:** `9ea7feb`, `f7079dc`, `ef62a6b`. Reviewed and approved. The likely cause is orphaned `msedgewebview2` processes left by another Office app (an old PowerPoint) holding `%APPDATA%\Oracle\SmartView\WebView\EBWebView`. That is in `docs/troubleshooting.md`.
- **Fix order item 2, workbook open in another Excel instance / SaveAs collision:** `da7fe11` and `ffa3754`. Reviewed and approved.
- **Fix order item 3, error codes not noticed in answers (#NoAccess on BANORWAY recon):** `4c9ca14`. Reviewed: 3 Medium issues are open (see Next, item 1).

## Next
1. **Fix the 3 Medium findings from the review of `4c9ca14`,** then re-review. The mapping, `cell_trace`, and the template members checked against the real atlas are all correct.
   - **Case-sensitive member check** (`atlas.py:72-76` `has_member` uses SQLite `=`). Essbase is case-insensitive, so `"un.sy.2tec"` is now wrongly rejected. Fix: `COLLATE NOCASE`, and pass the atlas's canonical spelling into the pull.
   - **Stale-but-usable atlas blocks new members** (`ops.py:250-255`, `atlas.py:65-69`). There is a 30-day window, and FinMax periods only go up to Yr 26, so for example "Jan 27" and "Forecast Q426" would be rejected, with no override. Fix: treat an atlas miss as "unverified". Confirm it live (the `sv_member_info` path) before raising, or raise only for template-generated guesses (the `finmax_entity` mapping) and just warn for members the user gave.
   - **No limit on scan warnings** (`answer.py:177` `error_warnings`): one warning per error cell, possibly thousands for `cost_base`. Fix: group pull-sheet errors per sheet and code ("'cb': 412 cells #NoAccess (first: B7, B9, …)", like `tools_query.py:86`) and cap the answer-line warnings.
2. **Item 4:** `xl_revert` doesn't restore the values. Scenario S6.xl: "after revert differs from the snapshot in ['values']".
3. **Item 5:** after SSO cookie loss, the drill passes 130 s and succeeds on retry (S8.sso, flaky). Plan Task 11.
4. **Item 6:** dict rows are rejected with a raw pydantic error (S1.bad.dictrows). Plan Task 9.
5. **Item 7:** speed. Plan Task 13; the gate is an S2.build median of 30 s or less.
6. **Small fixes:**
   - `xl_read` returns formats for the first cell only (`livexl.py` ~251–260);
   - the `xl_changes` path filter fails on a local path vs a OneDrive URL;
   - S7 refresh copies are left in `smartass\output`;
   - empty leftover Excels after new-workbook routes.
7. **Phase D:**
   - Task 14: remove the `sv_adhoc_*` tools and `sv_show`, leaving 30 tools.
   - Task 15: re-run the full matrix plus `--repeat 5` on the drill, with a before/after comparison against the baseline.
   - Task 16: persona pass.
   - Task 17: report and roadmap, merge, push, CI green.

## Commands
Run from `C:\dev\smartass\engine`:
- `$env:PYTHONUTF8=1; uv run pytest -q`: 256 tests pass.
- `uv run ruff check .`
- `uv run python ..\scripts\check_staged.py --all`
- Live, which drives the owner's real Excel, so ask first: `uv run python tests\live\matrix.py --out <name>`, with `--only <tags or IDs>` and `--fresh --only fresh`.

## Cautions
- PowerShell 5.1, so no `&&`. Never use `Get-Content | Set-Content` on source files.
- Don't commit `engine\Microsoft\` (a stray PowerShell cache).
- Commit trailer: `Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>`.
- Never kill or quit an Excel the engine doesn't own. The owner's PowerPoint and Excel may be in use.
