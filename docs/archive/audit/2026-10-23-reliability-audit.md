# Reliability audit of the smartass engine — report

Date: 2026-10-23 · Branch `impl/reliability-audit` · Spec `docs/specs/2026-10-23-reliability-audit-design.md` ·
Plan `docs/plans/2026-10-23-reliability-audit.md`

## Summary

- **Matrix:** 32 PASS / 3 SLOW / 5 FAIL / 3 LIMIT before, and 32 PASS / 7 SLOW / 1 FAIL / 3 LIMIT after.
  - All five baseline FAILs pass now.
  - The one new FAIL (S6.xl, at a later step than before) was fixed in `f8076aa` and passes in the rerun (3/3).
- **Drill:** cold, warm and quarter drills passed 15/15 in five consecutive rounds, each reconciled `exact`.
- **Hangs:** no call ran past 120 s, including S8.sso, which hung at 130 s in the baseline.
- **Speed:** the S2.build median is **25.8 s**, which meets the ≤ 30 s gate. It was measured after closing a PowerPoint that held Smart View's WebView2 folder. With PowerPoint open, the median was 39.6 s (see Speed).
- **Tools:** 30 (the ad hoc tools and `sv_show` are removed).
- **Personas:** 4/4 answered correctly (2 of them after the environment was fixed). The real findings were fixed in code and in the skill.
- **Tests:** 288 unit tests, ruff clean.

## 30-day evidence (from the spec)

| Tool | Calls | Failed | p50 | p90 | Main causes |
|---|---|---|---|---|---|
| `sv_drill` | 20 | 15 | 67 s | 133 s | Browser closed during launch, "never asked for its data", missing dimension in `cell`, no drill report, cold Excel open |
| `sv_ws_build` | 20 | 5 | 21 s | 56 s | COM "Exception occurred", spec shape, unknown dimension `Balance` |
| `sv_ws_answer` | 14 | 2 | 6 s | 24 s | "No pulls for Q<n>", COM errors |
| `sv_query` | 46 | 3 | 1.4 s | 4 s | Quoted dimension names, `rows` passed as a dict |

## Matrix before → after

Reports:
- before: [baseline](2026-10-23-matrix-baseline.md) and [baseline fresh](2026-10-23-matrix-baseline-fresh.md);
- after: [after](2026-10-23-matrix-after.md), [after fresh](2026-10-23-matrix-after-fresh.md),
  [drill ×5](2026-10-23-matrix-drill5.md), [rerun ×3](2026-10-23-matrix-rerun.md) and
  [speed gate](2026-10-23-matrix-gate.md).

| Scenario | Before | After | Note |
|---|---|---|---|
| S1.bad.dictrows | FAIL | PASS | `rows`/`columns` accepted as `{dimension: members}` |
| S3.actual_vs_plan.sy | FAIL | PASS | FinMax entity mapped via the atlas (`BANORWAY` → `un.SY`); error cells flagged |
| S6.xl | FAIL | PASS (rerun 3/3) | Revert restores typed constants; `add_sheet` revert judges content by values |
| S8.dialog | FAIL | SLOW (passes) | The error names the dialog (workbook found across Excel instances) |
| S8.editmode | FAIL | SLOW (passes) | The same |
| S8.sso | PASS (flaky, HANG 130 s) | PASS 80 s | Bounded sign-in; a stalled SSO redirect is nudged |
| S10.9.options | PASS (flaky) | PASS | Login/WebView2 dialog detected, Excel restarted once |
| S1.query.finmax | SLOW 18 s | SLOW 30 s | First call of a server: starts the hidden Excel plus the SSO login (dialog seen and clicked) |
| S2.build | SLOW 39 s | PASS (median 25.8 s on a clean environment) | See Speed |
| S4.scope | SLOW 45 s | SLOW 81 s | Two concept builds; same slowdown as below |
| S5.drill.report | PASS 69 s | SLOW 99 s | The same |

Fresh profile: S9.install, S10.2.setup and S10.4.update PASS; S9.lookup LIMIT (as before).

## Fixes (in the approved order)

| # | Problem | Commits |
|---|---|---|
| 1 | Smart View Login/WebView2 dialog blocks Excel | `9ea7feb` `f7079dc` `ef62a6b` |
| 2 | Workbook open in another Excel / SaveAs collision | `da7fe11` `ffa3754` |
| 3 | Error codes unnoticed in answers (`#NoAccess` on the BANORWAY recon) | `4c9ca14` `1594cac` `5be7a65` `183a13e` |
| 4 | `xl_revert` does not restore values | `1618307` `29315f4` `f8076aa` |
| 5 | Drill passes 130 s after SSO cookie loss | `3eaa921` `6e3ae41` `28f8db9` |
| 6 | Dict rows rejected with a raw pydantic error | `b2c6012` `de3d089` |
| 7 | Speed: progress notes, open-target timing | `aba1555` `dc94f0e` `8ee0c82` |
| — | `xl_read` formats for the first cell only | `4456c49` |
| — | `xl_changes` path filter (local path vs OneDrive URL) | `0aecb72` |
| — | Empty leftover Excels after a run; S7 copies | `e9602d9` `7f01763` (S7 copies: the cleanup already worked; it now reports a file it cannot delete) |
| — | Trim: ad hoc tools and `sv_show` (30 tools) | `1de5d7c` |
| — | Skill and `operations.md`: new errors and their next steps | `6b82c2e` |
| — | Persona pass: login and no-session errors say who acts | `956abaa` `af3875a` `f8f1b88` |
| — | Persona pass: concept checks blocked by tn-formelstil (`+` chain) | `0455b49` |
| — | Persona pass: `sv_query` names the rows it left out as empty | `c1c6eba` |
| — | Persona pass: skill clarifications | `e1fcfdd` `ff15a06` |

Every fix was test-first, and an independent reviewer went over every commit. Each Medium finding was fixed and re-reviewed. Not reproduced, so not changed: Task 8 (quoted or short dimension names were already rejected with a hint), Task 10 (spec shape), and Task 11 step 3 (the drill browser launch always recovered on its existing retry).

## Speed

| Run | S2.build | Environment |
|---|---|---|
| Baseline 13:04 | 38.6 s (open target 23.6 s) | |
| After / rerun ×3 | 38.5, then 61 / 38 / 40 s | POWERPNT.EXE held Smart View's WebView2 folder, and the SSO watcher's UI Automation calls blocked for 47–118 s |
| Gate ×3 | 51.9 / **25.8** / 18.7 s | PowerPoint and the long-running Excel closed; the first run starts a cold Excel plus the SSO login |

- **`open target`.** The new sub-stages showed where its time goes. On a warm Excel it is now about 2.2 s.
- **Stages this round did not touch were slow too.** With PowerPoint open, the pull stages ran 2–3× slower: `open default grid` 7 s, once 167 s for a persona. Once it was closed they were back to 2–4 s.
- **Progress notes.** A stage over 10 s now sends `still working: <stage>` every 10 s, so "no stage over 60 s without a progress note" holds.
- **Not done.** The plan's style cache: `ensure styles` takes 0.01–0.5 s in every logged build, so there is nothing to gain.

## Persona pass

Four subagents, each with only the skill and the smartass MCP and one Norwegian question. The first pass ran while the SSO environment was degraded (see Speed). The two personas that were blocked were rerun once, after the fixes and on a clean environment.

| Persona | Question | Path | Tool errors | Result |
|---|---|---|---|---|
| Controller | Hva brukte Consumer på konsulenter i Q3 mot i fjor? | `sv_scope`, `sv_find_member`, then `sv_query` (`sv_ws_build` was blocked by the stuck login ×3) | 4 (3 environment, 1 missing `connection` on `sv_member_info`) | Correct handling of the reorg: 2CON is empty and 2SSM + 2EI give 33.5 MNOK against 34.2 MNOK (`ac.C693`). No workbook. |
| Product owner | Hva tjener vi på mobilabonnement? | `sv_scope`, then `sv_concept` ×2 | Pass 1: 2 (environment). **Rerun: 0** | Rerun: direct contribution 7.57 bn NOK YTD Sep 26 (Consumer 5.68 + Business 1.89), +1.0 %. Checks OK, `formula_style` OK. |
| Analyst | Hvilke leverandører ligger bak 12881 i september? | `sv_find_member`, `sv_drill`, then `sv_ws_answer` | Pass 1: 2 (environment; the error named PowerPoint). **Rerun: 0** | Rerun: −161 233 428, 95 lines, `exact`. No suppliers: all lines are asset journals. |
| Planner | Hvordan ligger Technology an mot target hittil i år? | `sv_analyse actual_vs_plan` | 1 (connect rc=-15 "sID empty" while not signed in; the retry worked) | EBITDA −1 114 MNOK against a target of −1 071 MNOK (−42.8 MNOK, −4.0 %). FinMax recon `un.SY.2TEC` (1 112 834 010) matches the reference value. |

**Fixed after the pass:**
- the login error asked the agent to click; it now asks the user and names PowerPoint;
- `rc=-15` had no next step;
- concept checks failed `formula_style` (a `+` chain);
- `sv_query` dropped empty rows silently;
- skill gaps: several matching families, saved files, drill `Entity` = cost centre, "no suppliers".

## Accepted limits

Accepted by the owner on 2026-10-23. They are listed by symptom, with the known cause, in `docs/troubleshooting.md` ("Known limits").

| Item | Reason |
|---|---|
| S9.lookup | A fresh-profile lookup needs a real signed-in Excel. COM Excel inherits the simulated environment, so verify manually on a new machine. |
| S10.5.textfixes | The user override of a shipped key is verified in-process. No current drill line hits a shipped text-fix key. |
| S10.6.pack | The pack has no runtime consumer yet; only the wiring is verified. |
| S10.8.coldexcel | Needs no Excel running. The owner's Excel is open during the run. |
| SSO environment | Another Office app (PowerPoint) holding Smart View's WebView2 folder slows or blocks sign-in. The engine names it, and the user closes the app. Orphaned `crashpad-handler` processes from forced Excel restarts pile up but do not block. |
| S1.query.finmax first call | Starting the hidden Excel plus the hourly SSO login is 15–30 s on a server's first FinMax call. |
| Abandoned user-Excel job | It keeps sending progress notes to its finished request until the hung COM call returns. Clients drop these notes. |
| S5.drill.report in `--only drill` | Needs S2.build's workbook in the same run (it passes in the full run). |

## Roadmap (ranked)

1. **Change:** have `smartass doctor` / `sv_status` detect a WebView2 folder held by another Office app, plus orphaned crashpad handlers, before the user hits the login stall (it cost the persona pass and the speed gate).
2. **Speed up:** reuse one connected grid sheet per cube across the pulls of a build (`open default grid` 2–4 s per pull).
3. **Change:** move the matrix's dependent scenarios (S5.drill.report, S7) onto their own setup step, so `--only` runs stand alone.
4. **Build:** colleague rollout through the TelenorNorgeInternal repo transfer (installer, doctor bundle, first-run checklist from S9).
5. **Build:** the Excel add-in task pane on top of the same engine (progress notes already exist for it to show).
6. **Speed up:** `sv_scope` and `sv_concept`: cache the concept library per server process (S4 81 s).
7. **Change:** give the drill's `All Entities` default a live check. If the top member has no drill report, split from the top instead of returning "no report".
8. **Build:** Azure OpenAI as an alternative model back-end for colleagues without Claude/Copilot.
9. **Remove:** the S10.6 pack wiring, or give it a consumer.
