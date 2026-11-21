# Ledger lines: lines blocks and fin_drill (v2)

The lines come from FinLite's drill-through in Data Integration (view `DrillFilter`, 45 columns), read over HTTPS
with the user's own EPM sign-in (no browser), never by Smart View's own drill-through. Two ways in, same engine, same
sheets and checks:

- **A `lines` block** in an `fin_analyze` spec, next to the cube blocks:
  `{"type": "lines", "cell": {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"},
  "group_by": "supplier", "top": 10}`. It needs no query. A follow-up patch or `fin_refresh` drills again, so
  lines are always fresh.
- **`fin_drill(cell, group_by, top)`**: the same as a spec with only that lines block.

`cell` needs `All Accounts` and `All Time Periods` (plus `All Entities` for the unit; without it the top member is
drilled, with a warning). Balance is Periodic only. A parent period (`Yr 26`, `Q3 26`) works. `group_by`:
`supplier` (`AP_CUSTOMER_NAME`), `project` (the booked project, `INITIATIVE`) or `account` (`ACCOUNT`).

## What the answer shows

- **Ranking:** top N groups (largest absolute amount first; spellings that differ only in case are one group),
  `Other`, `Lines without a <supplier | booked project | account>`, `Ledger lines total`, `Cube value`, and the
  gap (cube − lines), each with its share of the lines total.
- **Gap:** must be 0 (tolerance 1 NOK). Otherwise report it as a gap, never hide it; ask again to re-drill.
- **Sheets:** `Q<n> Answer` (the ranking, live formulas on the lines table) and `Q<n> Lines <alias>` (Excel table
  `Lines_Q<n>_<alias>`, key columns as text).
- **Chat:** totals and the top groups from `key_values`, never line dumps. Journals, accruals and allocations
  have no supplier: name that line when it is large. Mixed signs (purchases offset by allocation credits) make
  shares misleading: say so instead of quoting a share above 100 %.

## Column glossary (45 columns, in page order)

| # | Label | Meaning |
|---|---|---|
| 1 | Scenario | Scenario of the line (Actual) |
| 2 | Period | Posting period, e.g. `Jul-26`. Text in the table (date-like: use the FILTER form) |
| 3 | GL_POSTED_DATE | Date the line was posted in the GL |
| 4 | GL_VOUCHER_NUMBER | **Key.** GL voucher number (text) |
| 5 | Entity | **Key.** Entity (un.SY cost centre) |
| 6 | CP | Counterparty / intercompany partner |
| 7 | ACCOUNT | **Key.** Account code (text) |
| 8 | AMOUNT | **Key.** Signed amount in NOK, FinMax signs (cost negative). Sums to the cube cell |
| 9 | DEBET_AMOUNT | Debit part of the line |
| 10 | CREDIT_AMOUNT | Credit part of the line |
| 11 | DESCRIPTION | Header description of the journal / document |
| 12 | GL_DESCRIPTION_LINE | **Key.** Line text: what was booked |
| 13 | JE_SOURCE | **Key.** Origin: Payables = AP invoices via Coupa; Spreadsheet = manual journals and accruals; DATABRONN = traffic feeds; Projects; Cost Accounting |
| 14 | ALIAS_EN | English alias of the account |
| 15 | JE_CATEGORY_NAME | **Key.** Journal category (invoice, accrual, reversal, allocation ...) |
| 16 | Products | **Key.** Product tag (pd.SY code). `No_Product` = untagged |
| 17 | Function | Function (unused in Telenor Norge) |
| 18 | PROJECT | **Key.** Project code (text) |
| 19 | SYSTEM | System tag |
| 20 | RESERVE | Reserve tag |
| 21 | AP_INVOICE_NUMBER | **Key.** Supplier invoice number (text) |
| 22 | AP_VOUCHER_NUMBER | AP voucher number |
| 23 | AP_CUSTOMER_NAME | **Key.** Supplier name. Blank = no supplier (journal, accrual, allocation) |
| 24 | FUTURE1 | Spare field |
| 25 | Currency | Currency of the line |
| 26 | TRANS_AMOUNT | Transaction-currency amount |
| 27 | ACCOUNT_NAME | **Key.** Account name |
| 28 | BATCH_NAME | **Key.** Journal batch name |
| 29 | JOURNAL_NAME | **Key.** Journal name. Names like "SF2609 allokering IC VK" are allocations |
| 30 | JOURNAL_SUBLEDGER | Subledger of the journal |
| 31 | LINE | Line number in the journal |
| 32 | PAYABLE_SOURCE | Source of the payable |
| 33 | REVERSED | Reversal flag |
| 34 | REVERSAL_STATUS | Reversal status |
| 35 | GL_TAX | GL tax code |
| 36 | SUBLEDGER_TAX_CODE | Subledger tax code |
| 37 | COUNTRY | Country |
| 38 | BOOKING_DATE | Booking date |
| 39 | Orig_Curr_AMOUNT | Original-currency amount |
| 40 | Integration Name | Data Integration rule that loaded the line |
| 41 | Location | Data Integration location |
| 42 | PRODUCT_DESC | **Key.** Product description |
| 43 | PROJECT_DESC | **Key.** Project description (SYPR68 text is broken in EPM, see text repair) |
| 44 | RESERVE_DESC | Reserve description |
| 45 | CC_DESC | **Key.** Cost centre description |

## Text repair

Applied to every text field of the lines (`epm\drill_lines.normalize`). Text that could not be repaired is listed in the
reply's `text_fixes` (candidates for your own fix list).
- **U+FFFD (letters lost in EPM at load).** Cannot be computed back. Fragment list in `knowledge\text_fixes.json`
  (shipped, replaced by updates) merged with the user's own `%LOCALAPPDATA%\smartass\text_fixes.json` (same
  format `{"fixes": {...}}`, wins on conflicts; broken fragment → correct fragment, replaced wherever it occurs).
  Seed: `H��yhastighet → Høyhastighet`.
- **Mojibake** (`FjordbÃ¥ndet`): repaired exactly with `encode("latin-1").decode("utf-8")`, only when the
  string matches `Ã[\x80-\xbf]` or `Â[\x80-\xbf]` and the round trip succeeds.
- **Unfixed** strings stay as they are, are listed in `text_fixes`, and are never guessed.
  Offer to add them to `%LOCALAPPDATA%\smartass\text_fixes.json` (survives updates).
- The SYPR68 "Høyhastighet" project text is broken in the EPM metadata itself (the FinMax alias of `pj.SY.SY.SYPR68`
  too). Tell the user once; the EPM admins should fix it.

## Timings and limits

- Measured 21.11.2026: 2 839 lines in about 8 s, 7 852 lines in about 23 s, 10 413 lines (22 pieces) in about 11 s.
  EPM's own time dominates; nothing starts or signs in first. At most 4 getData calls are in flight at once.
- Split above 5 000 lines per piece (one getData page): months, then entities, then accounts. Each piece reconciles
  to its own cube cell.
- **Too wide (EPMFDM-140373):** EPM refuses a summary drill when more than 3 dimensions have over 1000 children
  (e.g. `un.SY` with all other dimensions at the top). The cell then splits by child entity (then account) by itself;
  empty (#Missing) children and children without a drill report are skipped. Example: 12881 × `un.SY` × Sep 26 →
  22 children, 4 with data, 95 lines, exact. A cell refused once splits at once on a re-drill (same server).
- **Sign-in:** the user's OAuth token for the pod (`smartass login`; `smartass doctor`, check `oauth`). EPM's drill
  session is taken with it and renewed once for all pieces when EPM drops it.
- Excel gets key columns only above 50 000 lines. Maximum 300 000 lines; above that the tool stops after the first
  page and reports the size, and you narrow the request.
- Each piece has a time budget: 300 s, or 15 s a page once its first page tells its size if that is more, plus
  360 s for one session renewal (660 s up to 20 pages, 1 260 s at 300 000 lines).
- No drill report on the cell (data entered by journal in EPM): the answer is refused with the reason.

## Errors

| Symptom | Action |
|---|---|
| The reply is a sign-in page and a code | No stored sign-in for the pod: show the user the page and the code, let them approve in their own browser, then call again. |
| `EPM access is not set up on this PC …` | `[oauth]` is missing in config.toml: the user (or IT) adds it, then `smartass login`. |
| `refused: no OAuth scope is configured for …` | The drill's pod has no scope under `[oauth.scopes]` in config.toml: add it, then `smartass login`. |
| `EPM answered the drill query with a page, not JSON (signed out?)` / `... the drill view ...` / `getData returned HTTP <n> without data at offset <m>` / `EPM gave no session key for the drill` | EPM's drill session was lost again after renewing it. Retry once; if it repeats, `smartass doctor --bundle` to the maintainer. |
| `The drill of <cell> did not finish within <n> s (…)` | EPM was too slow for the piece's time budget: narrow the cell (a lower entity or account, or one month) and drill again. |
| `this cell has no drill-through report …` | Not a FinLite base-level intersection loaded by Data Integration: use a FinLite account and a lower entity or period. |
| Too wide (EPMFDM-140373) | Splits by itself by child entity, then account. |
