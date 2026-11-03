# Smart flow: one call, always warm, lean replies, fewer Excel calls (design, 2026-11-03)

Owner goal: the same output, as efficient as possible. Owner decisions 2026-11-03: A + B + C + D below; the hidden
browser stays warm all day; answer page setup once per sheet; the source-sheet read-back check stays; smartass always
uses the user's own Excel (the Smart View startup login is already gone: `CheckOnStartup=0`, PR #4).

Evidence (live, 2026-11-03): an answer costs the AI two round trips (`sv_gate`, then `sv_answer`); the P&L reply was
46 000 chars compact (57 000 as sent) - `table` 8.8k, `key_values` 8.9k (15-digit floats), `review` + `material` +
`flags` 16k (overlapping), `checks` 2.9k, `outline` 2.5k, `timings` 2.5k; the hidden browser closes after 300 s idle,
so the next question pays ~6 s start + sign-in and ~4 s warm drill page; Excel writing makes ~1 000 COM calls per
answer (source sheet ~78 each, `arrange sheets` 82-102, answer styling 136-160, page setup 1.3-1.4 s).

## A. One call per question
- `sv_answer(question=..., answers=None, detail="lean")` with no `spec` and no `followup` runs the gate
  (`tools_core.gate`).
  - `ambiguous` / `infeasible`: returns the gate result as `sv_gate` does today (questions, reason, proxy,
    checklist), nothing pulled.
  - `clear` with a recipe: builds `{version: 1, question, gate: <gate result>, assumptions, recipe: gate.recipe}`
    (assumptions = every checklist item whose source is `convention` or `learned` as "<Item>: <value> -
    <source>.", plus each warning) and answers it exactly as a spec call.
  - `clear` without a recipe (free-form): returns the gate result with `next` = "send a spec (pulls + layout)".
- The two-step flow (`sv_gate` + `sv_answer(spec=...)`) keeps working unchanged.
- The MCP server instructions, the `sv_answer`/`sv_gate` docstrings, `knowledge/advisor/gate.md` and the smartass
  skill text say: for a new question call `sv_answer(question=...)`; on an ambiguous reply ask the questions, then
  `sv_answer(question=..., answers={...})`.

## B. Always warm
- `config.DRILL_KEEP_WARM` (default on, env `SV_DRILL_KEEP_WARM`): the drill browser does not close when idle.
- Keep-alive: every `DRILL_KEEPALIVE_S` (default 600) idle, the worker checks the session (a cheap provider request
  through the existing route, `svhttp` `GetProvisionedDataSources` on the own pod) and signs in again if it is gone
  (same `_ensure` path), so the next question never meets a cold or expired session.
- Warm at server start: the server starts the browser and signs in in the background right after it starts (not
  awaited; a failure is only logged).
- Yield to another smartass process: a process that gets `DrillBusy` writes `<profile>/.wanted` (timestamp); the
  owner's idle worker sees it within one poll and closes its browser (freeing the profile lock); the asker retries
  once after a short wait. So keep-warm never blocks a parallel agent for long.
- Off (`SV_DRILL_KEEP_WARM=0`): today's 300 s idle close.

## C. Lean replies
- `reply.lean(out, detail)` applied at the end of `answer()` (and follow-ups). `detail="full"` returns today's reply.
- Lean keeps: `workbook, question, answer_sheet, sheets, data_sheets, note, gate, period, lines, routes,
  key_points, followups, warnings` unchanged; `key_values` with NOK values rounded to whole numbers and ratios
  (`%` columns or |x| < 10 with a fraction) to 4 decimals; `checks` as `{"ok": n}` plus every check that is not OK
  by label; `review.findings` (deduplicated); `material` as `{line, label, delta, pct}`; `outline` unchanged (needed
  for follow-ups); `timings` = `_total`, `first frame`, `drills` and the 3 largest top-level stages; `formula_style`
  as `{ok, counts, block}`; `style_lint` first 5; the `table` only when it has <= 40 rows; recipe extras (`flags`,
  `missing`, `fills`, `landing`, `freshness`, `headline_kpi`, `forecast`, ...) kept if <= 1 500 chars each, else
  replaced by `"<n> items: ask with detail='full'"`.
- Target: P&L reply <= 12 000 chars, lines_ranking <= 6 000.

## D. Fewer Excel calls
Measured with the per-stage COM counter (`com calls.*`). Same sheets, same look.
- Source sheets (`write source sheet`, ~78 calls): number formats, fonts and fills applied per contiguous range and
  per style group (not per column/row); the provenance panel in one write + one style call; the read-back check
  stays (one bulk read).
- `arrange sheets` on save (82-102 calls): read the current order once; move only sheets that are out of place.
- Answer page setup (1.3-1.4 s): skipped when the sheet already has smartass's page setup (a sheet-level name
  `smartass_pagesetup` = the setup's version), i.e. on follow-up redraws; first draw unchanged.
- Answer styling (136-160 calls): batch the largest groups (cell styles, widths/heights, tab/freeze) into range
  calls.
- Target: Excel writing under ~6 s per answer; COM calls roughly halved.

## Proof
Unit tests per part; live: the same three answers as `docs/audit/2026-11-02-speed-bench.md` (P&L Wholesale YTD,
consultants Technology YTD, revenue trend Telenor Norge) through `sv_answer(question=...)`, all checks OK; reply
sizes, round trips, timings and COM counts into the bench doc.

## Out of scope
Building workbooks without Excel; caching answers or drills; Smart View refresh.

## E. New workbooks without the hidden Excel (owner, 2026-11-03)
Today, with no Excel open, the hidden server Excel (Smart View loaded) creates the file (plain workbook, sheet
'Sources', sensitivity label, SaveAs), then the user's Excel opens it; the hidden Excel took 13.1 s to start and, idle
at 12:18, its Smart View opened a login whose request timed out on the user's screen.
- Seed: the first file made the old way is copied to `%LOCALAPPDATA%\smartass\new_workbook_seed.xlsx` (identical to
  what Excel produced, label included). Later new workbooks = a copy of the seed under the new name; no hidden Excel.
- No seed yet, or the copy fails: the old path (and it writes the seed). A seed older than 30 days is renewed the
  next time the old path runs... simplest: renewed when the label check fails (the copied file must still carry the
  `MSIP_Label_` properties of the seed; if not, old path).
- The hidden Excel then starts only for the HTTP-failure fallback.

## F. Smart View error boxes in smartass's own hidden Excel are closed (owner, 2026-11-03)
While the hidden Excel runs, a watcher (every ~3 s) closes Smart View message boxes in THAT process only (caption
"Oracle Smart View for Office", an OK button), logging their text. The user's Excel is never touched.
