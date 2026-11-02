# sv_drill: drill-through to journal lines (design, 2026-10-15)

Status: **approved 15.10.2026** (rev. 2, after the user reviewed an example workbook). Built on a read-only spike run on 15.10.2026, which is recorded in memory `gl-drill-through`.

**Reference example:** `prodreg2026\analyser\Analyse (Telenor Norge AS) - Produktregnskapet\2026-10-15_1525 Drill example - Wholesale COGS Q3 26.xlsx`. The user said "looks great". The tool's output should look like this, with sections chosen per question.

## 1. Purpose

Let Claude deep dive from any FinLite number down to the ledger lines behind it: supplier, invoice number, voucher, JE source, line text and product tag.

The lines answer questions like these:
- What does product X's COGS contain?
- Who are the suppliers on GL 43100 in Technology?
- What is accrual and what is a real invoice?

Every answer comes with a report that fits the question and reconciles to the cube.

Scope agreed with the user:
- **General tool.** Any account, entity, product, period or other dimension, not only COGS.
- **Size.** Up to about one division-year per deep dive, split automatically.
- **Reports adapt to the question.** No fixed template, unless the user names one.

## 2. Facts the design rests on (spike 15.10.2026)

| Fact | Evidence |
|---|---|
| A FinLite cell has exactly one drill report, "Drill Through to source", of type `post` (Data Integration, integration `DLR_FinLite_SY_HistData`). | `HypGetDrillThroughReports` |
| `HypExecuteDrillThroughReport` opens the default browser (Edge). It ignores the user's "Drill-Through Launch" setting (`m_DTLaunchType=1`). | The spike popped a tab |
| The drill URL's `sso_token` does not authenticate on its own. The IDCS cloud gate redirects to sign-in. | Plain HTTP POST and GET |
| Headless Google Chrome logs in silently. The IT-policy "Windows Accounts" SSO extension handles Entra ID after a click on "Telenor SSO". | Headless runs, no window |
| The drill page only renders after `/HyperionPlanning/` has been opened in the same tab. Otherwise its JS fails with "CDN_HOST null". | Blank vs rendered page |
| The page fetches the lines with `POST /aif/ui/model/drilldown/getData?limit=500&offset=N`. The JSON body carries a server-built `QUERY`. The response is `responsePayload.{items,totalResults,hasMore}`. | Network log |
| Column labels come from `GET /aif/ui/model/view?page=DRILL_THROUGH&appId=…&partitionkey=…`. For example ATTR13 = AP_CUSTOMER_NAME and ATTR11 = AP_INVOICE_NUMBER; `AMOUNTX` is the signed amount. | View definition |
| Parent accounts, parent entities and quarters drill. Σ`AMOUNTX` equals the cube cell exactly in all four tests. | Table below |

| Cell (FinLite, Actual, NOK, Periodic) | Lines | Σ lines = cube |
|---|---|---|
| 40100 × E400 × Sep 26 | 81 | −1 140 790.92 ✓ |
| C40S × 2WHS × Sep 26 | 418 | −8 025 554.03 ✓ |
| C40S × 2TEC × Sep 26 | 120 | −35 008 728.86 ✓ |
| C40S × 2WHS × Q3 26 | 1 231 (3 pages, 10 s) | −24 706 615.55 ✓ |

Timing: 70–130 s per drill including login and the home page. Paging runs at about 120 lines/s.

## 3. Architecture

```
sv_drill(cell)
  ├─ 1. Smart View (our hidden Excel, existing host)
  │     retrieve the cell → cube value
  │     B_DrillReports(sheet, cell) → drill URL with attributes   [new bridge fn, read-only]
  ├─ 2. DrillBrowser (new module drill.py)
  │     one persistent headless Chrome per MCP server, profile %LOCALAPPDATA%\svkit\drill_chrome
  │     login once (click "Telenor SSO"), open /HyperionPlanning/, keep the tab warm
  │     open the drill URL → capture the page's first getData request (url, body, headers)
  │     replay that exact body via in-page fetch, offset += 500 until hasMore = false
  │     fetch the view definition → map ATTRn/UDn to labels
  ├─ 3. Split planner: if totalResults > SPLIT_AT, re-drill per child (month, then cost centre)
  ├─ 4. Store: %LOCALAPPDATA%\svkit\drill.db (SQLite), one table of lines + one row per drill
  └─ 5. Return: reconciliation + a compact profile of the lines (for Claude to choose the report)
```

### Units

| Unit | Does | Depends on |
|---|---|---|
| `modBridge.B_DrillReports` | Returns status, names, URLs, types for one cell. Never executes. | Smart View |
| `drill.DrillBrowser` | Login, warm tab, `fetch_lines(url) -> (columns, items, meta)`, restart on failure | playwright (Python), Google Chrome |
| `drill.plan_split` | Pure function: (cell, totalResults, children from the atlas) → sub-cells | atlas |
| `drill.store` | Write and read drills and lines in SQLite, with the label mapping applied | sqlite3 |
| `drill.profile` | Pure function: lines → counts and sums (net, gross debit, gross credit) by supplier, account, cost centre, month, JE source/category, product tagged vs `No_Product`, top invoices, and **top journals/batches for lines without a supplier**. That last facet is what exposed the +106 M IC allocation in the example. | — |
| `server.sv_drill` | The tool: orchestrates 1–5, enforces limits, logs to `qlog` | all of the above |

### Tool interface

`sv_drill(connection="FinLite", cell={dimension: member, …}, split="auto", max_lines=300000, q=None, target="session")`

- **`cell`.** Same dimension names as `sv_query`'s POV. Unspecified dimensions use their top member. Parents are allowed.
- **Result:**
  - `drill_id`, plus `cube_value`, `lines_sum` and `gap` (cube − lines) with a status of "exact" or "gap";
  - `lines`, `pieces` (the split pieces, each with its own reconciliation) and `secs` per stage;
  - `profile`: top-10 per facet, the share untagged to a product, and the share without a supplier (accruals, journals).
- **Workbook.** Three sheets go to the session workbook. With `target="none"` the lines go only to the cache.
  - `Q<n> Linjer`: an Excel table named `Linjer_Q<n>`, with columns as the drill page shows them, plus a source row. The lines are a **snapshot**, stamped with their time; re-running `sv_drill` refreshes them.
    - Up to `XL_FULL = 50 000` lines: all 46 columns.
    - Above that: only the key columns (period, dates, voucher, entity, account and name, amount, JE source/category, supplier, invoice, line text, product, project, batch, journal). Full detail stays in the cache.
    - Key-like columns are written as text: account, invoice, voucher, project.
  - `Q<n> FinLite`: a **live** Smart View grid of the drilled cell (and of each split piece). The reconciliation reads the cube value from it by label, so a refresh shows whether the ledger has moved since the drill.

**Reports.** Claude reads the `profile` and builds the report with the existing `sv_ws_build` / `sv_ws_answer`, using a new pull kind `{"drill": "<drill_id>"}`. The answer cells are `SUMIFS` over the `Q<n> Linjer` table, keyed on labels, so they follow `tn-formelstil`. Reports are chosen per question, from building blocks:
- a reconciliation block (always included);
- breakdowns by supplier, account, month, cost centre and product;
- invoice vs accrual/reversal/journal;
- top-N invoices;
- untagged COGS.

### Report rules (learned from the example)

- **Reconciliation block, always:** cube value (live), sum of lines, gap (cube − lines), line count, gross debits and gross credits.
- **No "share of total" when a block mixes signs.** The example showed 498 % because purchases (−123 M) and allocations (+106 M) offset each other. Show shares only when every line in the block has the same sign. Otherwise show debit and credit columns.
- **Criteria Excel would coerce.** In `SUMIFS`, date-like keys (`Jul-26`) and digit strings over 15 characters stop matching as text. For those, use the exact form `=SUM(FILTER(Linjer_Q1[AMOUNT],Linjer_Q1[Period]=$B30,0))`; plain `SUMIFS` stays the default. Verify in real Excel.
- **"Other" and "no supplier" rows.**
  - Other = block total − the listed rows (two terms, per `tn-formelstil`).
  - Blank keys use the criterion `""`.
- **Key points** name the drivers found in the profile: the largest suppliers, invoices vs journals, and allocation journals by name. Amounts in running text are full NOK with a minus sign.
- **Styling** goes through the existing answer builder (`styling.py`, open workbook). Long key points must fit on one line in B:E; the checker does not credit merged cells.

### Norwegian letters (text repair)

The pipeline keeps UTF-8 end to end: the server sends `application/json; charset=UTF-8`, and correct æ/ø/å arrive intact (e.g. "Strøm til produksjon"). Two kinds of damage arrive **already broken from EPM**. Raw server bytes were checked on 15.10.2026.

| Damage | Example | Cause | Fix |
|---|---|---|---|
| Replacement characters U+FFFD | `SYPR68 Wholesale H��yhastighet` (PROJECT_DESC, 128 of 1 231 lines in WHS Q3) | Lost at load in EPM. The FinMax member alias of `pj.SY.SY.SYPR68` is broken the same way. | Cannot be computed back. Apply the correction list `knowledge\text_fixes.json` (broken fragment → correct fragment, replaced wherever it occurs in a text field), seeded with `H��yhastighet → Høyhastighet`. |
| Double-encoded UTF-8 (mojibake) | `FjordbÃ¥ndet` (line text, 2 lines in TEC Sep) | UTF-8 read as Latin-1 in the source | Repaired automatically: `s.encode("latin-1").decode("utf-8")`, only when the string matches `Ã[\x80-\xbf]` or `Â[\x80-\xbf]` and the round trip succeeds |

- Repair runs in `drill_lines.normalize`. The result reports `text_repairs: {fixed_mojibake, fixed_by_list, unfixed: [distinct strings with U+FFFD]}`.
- Unfixed strings are kept as they are. They are listed as a warning so Claude or the user can add them to the list; nothing is guessed silently.
- Tell the user once that EPM metadata itself is broken (report to the EPM admins).

## 4. Splitting

- `SPLIT_AT = 20 000` lines per piece (about 3 min of paging).
- The order of split axes is: period (quarter → months, year → months), then the entity's children, then the account's children. Each piece is drilled, reconciled to its own cube cell and stored.
- **Guards.**
  - If the first page shows `totalResults > max_lines`, stop before paging and report the size. Claude then narrows the request.
  - A piece that can't be split further but is still > `SPLIT_AT` is paged anyway.

## 5. Safety and conduct

- **Read-only.** Never call `HypExecuteDrillThroughReport`. Only replay the page's own getData body, unchanged apart from `limit`/`offset`. Never write or alter SQL.
- **No visible windows.** Chrome runs headless, and the profile lives under `svkit`. If headless SSO fails, report it and stop. Never fall back to a visible browser without the user's yes.
- **Process ownership.** Our Chrome processes are recognised by their `--user-data-dir` (the dedicated profile `svkit\drill_chrome`), not by the PID registry, and only those are ever killed. The user's own Chrome and Edge are never touched. (Changed during the build, 15.10.2026.)
- **Data.**
  - Lines stay in `svkit\drill.db` (local, outside OneDrive) and in the user's session workbook (label Internal).
  - Drills older than 90 days are pruned.
  - Chat shows totals and top suppliers, never dumps of lines.

## 6. Errors

| Situation | Behaviour |
|---|---|
| The cell has no drill report (data entered by journal or allocation in EPM) | Return the cube value, `lines=0`, gap = full value, and the reason |
| The IDCS/Entra login does not complete within 60 s | Fail with "SSO did not complete headless". Keep the profile and retry once with a fresh tab |
| The page never sends getData within 45 s (Oracle UI change) | Fail loudly with the console errors attached. Never guess a query |
| The session expires mid-paging (401/redirect) | Re-login, then resume from the last offset |
| Gap ≠ 0 | Report it with the share of the cell, and keep the lines |

## 7. Testing and acceptance

- **Unit tests** (saved spike JSON fixtures): label mapping, paging stop rules (`hasMore`, `totalResults`, empty page), split planner, profile, reconciliation maths.
- **Golden report:** from the Q3 26 Wholesale fixture (1 231 lines) the builder reproduces the reference example's totals per section. Every section total equals −24 706 615.55, the gap is 0, the formula check has 0 BLOCK findings, and no `#` errors appear.
- **Live acceptance:**
  - The four spike cells reproduce the line counts and exact sums above.
  - One division-year drill (2WHS × C40S × YTD Sep 26) splits into months, and every piece reconciles.
  - Regression against the archive: `2026\2 Input\20 Inntekt-varekost\Tech\Avvik pga Drill through i 5BMT avdeling.xlsx`, sheet `2024 VK 4FNP` (4 085 lines). Same counts and sums per account, or each difference explained.
- **No-window check.** During the live tests no new top-level window appears (checked via the window list before and after).

## 8. Dependencies and change list

- `pip install playwright` into the svkit venv. It uses the installed Google Chrome (`channel="chrome"`), so no browser download.
- New: `smartview_mcp/drill.py`, the `B_DrillReports` function in `modBridge.bas`, tool `sv_drill`, pull kind `drill` in `specs.py`/`answer.py`, and tests.
- The skill `smartview` gets a "Drill-through" section and recipe 5, `reference/drill.md`, including the column glossary.
- `DESIGN.md` gets a short drill section that links here.

## 9. To verify during the build (not blockers)

1. Does **FinMax** drill the same way? It is assumed but untested. Until it's verified, FinLite is the default and only connection.
2. On the leaf cell, the page also fired `getData?limit=419&offset=81`, which returned 1 extra line (−1 818). Our loop ignores page-driven calls and pages on `hasMore`. Confirm it isn't a missing line by comparing against the page's own export.
3. Performance on about 50 000 lines (the division-year opex case), both paging and writing the Excel table into an open workbook.
4. The `FILTER` exact-match form against `SUMIFS` on date-like and long numeric keys, in the user's Excel.
5. How long the headless session lasts before a re-login, compared with Smart View's roughly hourly SSO.

## 10. Out of scope

- Drill to source documents (invoice PDFs), even though `ALLOW_DRILL_TO_SOURCE_DOCUMENT` exists.
- Direct ERP access.
- FinPlan, which has no drill.
- Editing anything in EPM.
