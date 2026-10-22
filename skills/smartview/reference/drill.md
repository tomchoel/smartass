# Drill-through (sv_drill): glossary, report rules, text repair

Spec: `docs\specs\2026-10-15-sv-drill-design.md` in the smartass repo. Plan: `docs\plans\2026-10-15-sv-drill.md`.
The lines come from the FinLite drill page (view `DrillFilter`, 45 columns). **Key** = used in most reports.

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

## Report rules

- **Reconciliation block, always.** Cube value (read live from the `cube_pull` grid), sum of lines, gap (cube − lines),
  line count, gross debits and gross credits. `exact` or explain the gap.
- **No share of total when signs are mixed** (`signs_mixed`). Purchases (−123 M) and allocations (+106 M) offset
  each other and gave 498 %. Shares only when every line in the block has the same sign; otherwise show gross debit
  and gross credit columns.
- **Criteria Excel would coerce.** Default is `SUMIFS` over `Linjer_Q<n>`. Use the exact-match form
  `=SUM(FILTER(Linjer_Q1[AMOUNT],Linjer_Q1[Period]=$B30,0))` (the builder picks it by itself) when the key is:
  - date-like (`Jul-26`);
  - a digit string over 15 characters (rounded to a number);
  - containing wildcards `* ? ~`;
  - starting with `<`, `>` or `=` (read as a comparison);
  - with leading zeros (`0123` becomes 123);
  - a time (`12:30`), a percent (`50%`), `TRUE`/`FALSE`, or a comma decimal (`1,5`);
  - over 255 characters (SUMIFS fails; the FILTER form has no such limit).
  Verify in real Excel. Key-like columns are written as text.
- **"Other" and "no supplier" rows.** Other = block total − the listed rows (two terms). Blank keys use the criterion `""`.
- **Key points** name the drivers from the profile: largest suppliers, invoices vs journals, allocation journals by
  name. Running text uses full NOK with a minus sign (`−24 706 615`).
- **Size.** Excel gets all 45 columns up to 50 000 lines; above that only the key columns (full detail stays in the
  cache `smartass\drill.db`). Chat shows totals and top suppliers, never line dumps.

## Text repair

Applied to every text field in `drill_lines.normalize`. The result carries `text_repairs {fixed_mojibake, fixed_by_list, unfixed}`.
- **U+FFFD (letters lost in EPM at load).** Cannot be computed back. Fragment list in `knowledge\text_fixes.json`
  (broken fragment → correct fragment, replaced wherever it occurs). Seed: `H��yhastighet → Høyhastighet`.
- **Mojibake** (`FjordbÃ¥ndet`): repaired exactly with `encode("latin-1").decode("utf-8")`, only when the
  string matches `Ã[\x80-\xbf]` or `Â[\x80-\xbf]` and the round trip succeeds.
- **Unfixed** strings stay as they are, are listed in `warnings` and `text_repairs.unfixed`, and are never guessed.
  Offer to add them to `text_fixes.json`.
- The SYPR68 "Høyhastighet" project text is broken in the EPM metadata itself (the FinMax alias of `pj.SY.SY.SYPR68`
  too). Tell the user once; the EPM admins should fix it.

## Timings and limits

- About 60–90 s per drill including the silent SSO login (login again after 5 idle minutes). Paging runs at about 120 lines/s (500 per page).
- One drill at a time across sessions (lock).
- Split above 20 000 lines per piece: months, then entities, then accounts. Each piece reconciles to its own cube cell.
- **Too wide (EPMFDM-140373):** EPM refuses a summary drill when more than 3 dimensions have over 1000 children
  (e.g. `un.SY` with all other dimensions at the top). The cell then splits by child entity (then account) by itself;
  empty (#Missing) children and children without a drill report are skipped. Example: 12881 × `un.SY` × Sep 26 →
  22 children, 4 with data, 95 lines, exact. A cell refused once splits at once on a re-drill (same server).
- **Browser:** headless Microsoft Edge first (`config.DRILL_CHANNELS` = msedge, chrome; profile `smartass\drill_edge`).
  A channel with a `UserDataDir` browser policy is skipped (Chrome has one: it forces every launch onto the user's
  own profile); a failed launch is retried once, then the next channel. A browser that died is restarted and the
  drill retried once. `sv_status` → `drill_browser` shows which channels can start. A failed drill page leaves a
  screenshot and HTML in `smartass\logs\drill_fail_*`.
- Excel gets key columns only above 50 000 lines. Maximum 300 000 lines; above that the tool stops after the first
  page and reports the size, and you narrow the request.
- No drill report on the cell (data entered by journal in EPM): cube value returned, `lines=0`, with the reason.

## Worked example: golden report

`mcp\tests\live_drill_report.py`: FinLite `ac.C40S` × `un.SY.2WHS` × `Q3 26` (Wholesale COGS). 1 231 lines, −24 706 615.55,
gap 0, 14/14 checks OK, 0 BLOCK findings from `formelsjekk.py`. Eight sections, each totalling the same amount:
1. Reconciliation. 2. JE source (Payables −123 132 549.19, 359 lines). 3. Journal category. 4. Month (354 / 459 / 418 lines;
`Jul-26` uses the FILTER form). 5. Account. 6. Top 12 suppliers + no supplier (872 lines) + other (SUPPLIER 037 AS
−60 646 654.33, SUPPLIER 020 AS −20 187 679.96). 7. Top 10 products + other. 8. Ten largest invoices.

Business insight recorded from it: the +106.3 M manual-journal credit on account 41130 is +72.4 M "SF26xx allokering IC VK"
plus +33.9 M "KE51 2607 WHS allokering REV og Cogs". These allocations offset the −123 M of purchase invoices, which is why
the supplier shares showed 498 % before the sign rule.
