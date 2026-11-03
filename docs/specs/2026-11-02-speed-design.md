# Speed: no Excel and no page per drill (design, 2026-11-02)

Owner goal: **less waiting** - no hard target; cut every known cost, measure after each step, stop when it feels
fast. Baseline: the consultant question (lines_ranking, un.SY, C60080+C60090, YTD Oct 26) ~100-105 s cold.

Decisions (owner, 2026-11-02):
- HTTP is the default everywhere the user waits. **No Smart View refresh feature**: answers are closed-month
  analyses; "refresh = ask smartass again" stays (source and Lines sheets stay values).
- Smart View in Excel stays only as a **logged fallback** when an HTTP drill step fails.
- No caching of drills (earlier decision stands).
- Read-only rule, restated: we send only requests whose content EPM itself produced (the provider XML we build from
  member labels, and the drill query EPM returns). Never `ExecuteDrillThroughReport`, never SQL we write or edit.

Evidence: `docs/audit/2026-11-02-drill-speed-spike.md`, `docs/audit/2026-11-02-drill-api-research.md`.

## Where the time goes and what removes it

| # | Cost (cold run) | Fix | Step |
|---|---|---|---|
| 5 | Drill links through Excel, sequential: 20-30 s | Provider `GetDrillThroughReports` over HTTP (proven, ~0.05 s) | 1 |
| 6 | A drill page per piece: ~5 s alone, ~26 s each in parallel | One warm drill page per session; per piece only the SUI query + getData from that tab (proven, ~1.5-2.3 s) | 1 |
| 7 | Empty child: 60 s capture timeout; too-wide cell: ~20 s before the split | Both come back at once in the SUI answer | 1 |
| 7 | Split child values read through Excel | ExecuteGrid over HTTP (one column) | 1 |
| 4 | EPM "Empty response" from Smart View COM | Gone with 5-7 | 1 |
| 1 | Hidden server Excel start ~7 s (file create, drill links) | Create the workbook from the TN template without Excel; nothing else needs it | 2 |
| 2 | Fixed 12 s login wait after Excel opens | Wait only while an EPM login window is actually showing | 2 |
| 3 | Visible Excel open ~9.5 s | Use the user's running Excel; else start it at once, in parallel with the data | 2 |
| 9 | Page setup 2.5 s | `Application.PrintCommunication = False` around page setup | 3 |
| 10 | Lines sheets / styling / save ~8 s | One block write per sheet, range-wide styling | 3 |

Expected: ~105 s -> roughly 20-35 s cold. Step 1 carries most of it.

## Step 1 - drills without Excel and without a page per piece

### Units
- **`svxml`**: `drill_reports(sid, cube, ords, cells)` request builder and `drill_url(text)` parser
  (`<drillThroughURL ...>` -> unescaped URL or None). `GetDrillThroughReports` added to `ALLOWED`;
  `ExecuteDrillThroughReport` stays refused (test).
- **`svhttp.SvHttp.drill_link(conn, cell) -> {cube_value, url, report}`**: same contract as `ops.drill_link`.
  One ExecuteGrid (account rows x period column, rest POV, Periodic only) gives the value and the Default-alias
  labels; `ords` = the data cell's index in that grid; then `GetDrillThroughReports`. Uses the session lock and the
  reconnect-once rule of `_ask`.
- **`svhttp.SvHttp.child_values(conn, parts)`**: the split's child values as one ExecuteGrid column (replaces
  `drill_split.child_values_on` on the normal path).
- **`drill_direct.py` (new)**: `DirectDrill` on the drill browser's worker thread.
  - `warm(s)`: once per browser session, open one real drill page in a dedicated tab (the first piece's URL), wait
    for its getData request, keep the tab, remember its `X-AppBuilder-Client-Id`.
  - `fetch(url, max_lines) -> {items, total, view, secs}` (same contract as today's drill fetch): from the warm
    tab, GET `<url without DRILLURL>&FORMAT=SUI` with `page_id: aif_drilldown` and the warm client id; then
    - `status 1` / no `QUERY` -> empty result with EPM's `details` as the note (no data / not Data Integration);
    - `EPMFDM-140373` in the answer -> `CellTooWide` (the existing split logic handles it);
    - else page `getData?limit=500&offset=N` with body `{QUERY, PARAMS: PARAMSARRAY, PARAMSTYPE: PARAMSTYPEARRAY}`
      through the existing `page_through` (max_lines, TooManyLines); the view definition from
      `/aif/ui/model/view?page=DRILL_THROUGH&appId=<APPLICATION_ID>&partitionkey=<PARTITIONKEY>` (cached per key).
  - Several pieces run concurrently from the warm tab (bounded, config `DRILL_DIRECT_PARALLEL`, start 4; measured).
  - Guard: getData is only ever sent from the warm tab (a plain tab gets the session signed out - spike).
  - A sign-out / non-JSON answer: re-login, re-warm, retry the piece once; then the old page route for that piece.
- **`core/drillstep.py`**: drill link = HTTP first, `ops.drill_link` through Excel on SvHttpError (logged
  `drill link: HTTP failed (...); using Smart View in Excel`). Lines = `DirectDrill`, today's TabPump/page route as
  fallback. Timers (`drilltime`) for each new stage.

### Tests (TDD, fake transport / fake page)
- svxml: request body shape (golden), URL parse incl. `&amp;`, missing element -> None, refusal of
  ExecuteDrillThroughReport.
- svhttp.drill_link: ords for the 1x1 layout, labels taken from the grid answer, Accumulated refused, reconnect once.
- DirectDrill: SUI status 1 -> empty with note; EPMFDM-140373 -> CellTooWide; paging/hasMore/max_lines; getData
  never sent before warm; sign-out -> re-warm once -> fallback; view cached per (appId, partitionkey).
- drillstep: HTTP link used, Excel only on SvHttpError (logged); empty child skipped without a fetch.
- Live (manual): consultant question timing + every piece's lines sum = cube (existing tie check).

## Step 2 - first frame without the hidden Excel and the fixed wait
- New workbook = copy of the TN template written with openpyxl (sheet `Sources`, sensitivity label "Internal" in
  the file's custom properties - Excel saves silently fail without it), then opened in the user's Excel.
- Prefer the running user Excel; only if none, start Excel immediately when the answer starts (in parallel with the
  gate/data), not after the data.
- `login.wait_login_cleared(..., appear_grace=12)` -> short grace (~2 s) that extends only while a login window is
  present. Short spike first: does a fresh Excel still show the EPM login when no Smart View call is made?
- Done when a cold run never starts the hidden server Excel (log check) and the frame appears < ~10 s with Excel
  already running.

## Step 3 - Excel write
- `PrintCommunication = False` around page setup (restore after, also on error).
- Lines sheets: one `Range.Value` block per sheet; styling per range, not per cell.
- Done when the Excel write stage is < ~6 s on the consultant question.

## Measuring
After each step a live run of the consultant question; stage times from `server.log` / `runs.jsonl` into
`docs/audit/2026-11-02-speed-bench.md` (times only, no figures).

## Out of scope
Smart View refresh of answer sheets; caching; Excel-free workbook writing end to end; a pump redesign beyond what
DirectDrill needs.
