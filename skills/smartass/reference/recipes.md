# Recipes — Produktregnskapet extracts

Each recipe is the grid the team builds in Smart View, written as a pull of an `sv_answer` spec
(`{id, connection, rows, columns, pov, title?}`; see `SKILL.md` §2.2). Put it in `pulls`, add a layout that reads the
lines the user asked about, and the full grid lands on its own source sheet (`Q<n> FinMax`). "Verified" means the
grid was run against FY2025 and reconciled cell by cell with the archived workbook (2026-10-13, Smart View 25.100).
Every difference was explained; the explanations are the traps listed in `SKILL.md`.

These are extracts, not the server-side recipes of `sv_answer` (`pnl`, `actual_vs_plan`, `monthly_trend`,
`cost_base`, `concept`, `lines_ranking`: see "Server-side recipes" at the end, and [operations.md](operations.md)
"Recipes and follow-ups"). Recipe `pnl` reconciles its live facts with the source sheets ("Live read = source
sheet").

## 1. GL extract — cost base per cost centre (FinMax)

This is the `GL` sheet of the interview and input workbooks (`2 Input\22 Arsverkfordeling\…\<year>_INPUT_Aktivitet_*.xlsm`,
2026 template `3 Modell\Test ny intervjumal\2026 Input Produktregnskap blank.xlsm`). The final cut is taken in
early January: Dec of the model year, **Accumulated**, with FTE.

```json
{"id": "gl", "connection": "FinMax", "title": "GL Dec 25",
 "rows":    [{"dimension": "All Entities", "members": ["un.SY.QT11", "un.SY.K570", "…", "un.SY"]}],
 "columns": [{"dimension": "All Accounts", "members": ["ac.C50S", "ac.C60", "ac.C70", "ac.C50100", "ac.C50200", "…", "ac.STA20Y"]}],
 "pov": {"All Scenarios": "Actual", "All Currencies": "NOK", "All Time Periods": "Dec 25", "All Balance Types": "Accumulated"}}
```

- **Rows.** Take the cost-centre list from the workbook's own `GL` sheet: column A holds the member names. Alternatively use
  `sv_members(FinMax, un.SY, descendants)`. The archive sheet has about 480 entities.
- **Columns.** The account names are in row 6 of the `GL` sheet, shown as labels such as `C50100 Salaries and Holiday Pay (ac)`, which is the member `ac.C50100`.
- **FTE.** FTE is account `ac.STA20Y` ("Total number of man-years - EoP"). It exists only in **FinMax**, not in FinLite.
- **Archive copies.** The archive `GL` sheets hold *pasted values*, not live Smart View grids, so a Smart View Refresh
  changes nothing on them. Rebuild the grid with this pull.
- **Verified** against `2025\2 Input\22 Arsverkfordeling\2IFRS15\2025_INPUT_Aktivitet_G916_CP.xlsm` (extract dated 07.01.2026,
  482 entities × 50 Dec 25 Accumulated accounts). 24,100 cells were compared and **95.5 % are identical**. The retrieve took 11 s.
  - The total `un.SY` agrees within 0.03 %. Pension `C50500` is the exception at 2.8 %: Dec 25 postings were made after 07.01.
  - Every other difference is in parent nodes restated by the 1 June 2026 reorg. `2CON`, `2CWK`, `5BRO_C` and `4IT_T` are now 0, and `2ADM`, `2TEC`, `2IT` and `2IG`
    are redistributed.
  - `un.SY.01_SUM` no longer exists.

## 2. Revenue and COGS per division × product (FinMax)

This is the grid behind `2 Input\20 Inntekt-varekost\<division>\<year>_Inntekt og varekost *.xlsx`.

```json
{"id": "rc", "connection": "FinMax",
 "rows":    [{"dimension": "All Entities", "members": ["un.SY.4REV_W"]},
             {"dimension": "All Products", "members": ["pd.SY.N5Z00", "pd.SY.N6800", "pd.SY.P56"]}],
 "columns": [{"dimension": "All Accounts", "members": ["ac.C30"]},
             {"dimension": "All Time Periods", "members": ["Q1 25", "Apr 25", "May 25", "Jun 25", "Q2 25"]}],
 "pov": {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic"}}
```

- **Accounts.** Revenue is `ac.C30` and COGS is `ac.C40S`, or their detail accounts.
- **Products and entities.** Products are `pd.SY.<code>` (`pd.SY.P56` COLOCATION, `pd.SY.P50` BROADBAND WHS). Find a division's
  product and entity leaves with `sv_members`.
- **Periods.** Quarters (`Q1 25`) and months can be mixed on one axis.
- **Verified** by refreshing `2025\2 Input\20 Inntekt-varekost\Wholesale\2025_Inntekt og varekost WHS_Large Sites.xlsx`.
  **690 of 753 numbers are identical.** The other 63 are all `3WHS_C` rows, which now return 0: the reorg moved `4REV_W` from `3WHS_C`
  to `3WHS_W`. Leaf data is unchanged.

## 3. Opex and depreciation control per division (FinMax)

This is the grid behind `2 Input\21 Costbase\211_Avskrivning\<year>_Avskrivningkost kontroll.xlsx`. Accounts are on the rows, and the divisions × year are on the columns.

```json
{"id": "dep", "connection": "FinMax",
 "rows":    [{"dimension": "All Accounts", "members": ["ac.EBITDA", "ac.60120", "ac.C75130", "ac.60100", "ac.C75110",
                                                      "ac.60300", "ac.C75120", "ac.C75", "ac.C76S", "ac.EBIT"]}],
 "columns": [{"dimension": "All Entities", "members": ["un.SY.2IT", "un.SY.2I16", "un.SY.2TEC", "un.SY.2ADM", "un.SY.2BUS", "un.SY"]},
             {"dimension": "All Time Periods", "members": ["Yr 25"]}],
 "pov": {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic"}}
```

- **IFRS16.** The IFRS16 sheets in the same workbook add counterparts (`cp.*`) or the Towers counterpart (`TE Telenor Towers Norway AS`).
- **Verified** by refreshing the 2025 file. 13,600 numbers were compared and **12,422 are identical**.
  - **The total for `SY Telenor Norge AS` is identical.** All the IFRS16 processing sheets are identical too.
  - The differences are the split between divisions, which is restated to today's organisation. `2CON` went from 10,520 MNOK EBITDA 2024 to 0
    and `2CWK` went to 0. `2BUS`, `2TEC`, `2ADM`, `2IT` and `2IG` were redistributed.
  - The sheet `FinMax opex & avskr 2024` failed with rc=1000. It references members that no longer exist.

## 4. Special costs per cost centre (FinLite)

This is the grid behind `2 Input\22 Arsverkfordeling\<division>\<year>_*særkost*.xlsx`. The base sheets (`Grunnlag …`) are FinLite
grids, and the pivot sheets are built on top of them.

```json
{"id": "sk", "connection": "FinLite",
 "rows":    [{"dimension": "All Entities", "members": ["un.SY.K570", "un.SY.E257", "un.SY.3SEC_C"]}],
 "columns": [{"dimension": "All Accounts", "members": ["ac.C50S", "ac.C60", "ac.C70", "ac.C50100", "ac.C50200", "ac.C50300"]}],
 "pov": {"All Scenarios": "Actual", "All Currencies": "NOK", "All Time Periods": "Yr 25", "All Balance Types": "Periodic"}}
```

- **Why FinLite.** It is lighter and faster than FinMax and has the same dimensions *except Functions, Projects, Reserves and Systems*. It holds no FTE.
- **Verified** by refreshing `2025\2 Input\22 Arsverkfordeling\2CSW_CA_Security_WHS_Kystradio\2025_CWK_SEC_særkostnader_20251117.xlsx`.
  - **All base and pivot sheets are identical** (1,367 of 1,436 numbers).
  - Three `GL` rows differ by about 1.7 %. That sheet uses `Yr 25 / Periodic` and was extracted on 17.11.2025, when `Yr 25` was still year-to-date.

## FinPlan (plan / forecast)

FinPlan has different dimensions:

- **Columns:** `Years` (`FY25`, …) and `Period` (`Jan` … `Dec`).
- **Rows:** `Account` and `Entity`.
- **POV:** `Scenario` (`Forecast Running Version`, `Target`, `Forecast Q425`, …), `Version` (`Working`, `Final`, `Division`),
  `All Products` and `All Balance Types`.

FinPlan's grid shows *aliases*, and the atlas stores those aliases (`BANORWAY`, `4TIP_I Tipo`, `FY26`, `Target 2025`).
**Aliases work as query members**, because Smart View resolves them. The real names still carry prefixes: `BANORWAY` is
`un.BANORWAYP`, `4TIP_I Tipo` is `un.FU.4TIP_I`, and `Wholesale` in Account is `ac.CAP_WHS` (a capex node, not the division).
`sv_member_info` gives the real name.

```json
{"id": "fp", "connection": "FinPlan",
 "rows":    [{"dimension": "Entity", "members": ["BANORWAY"]}],
 "columns": [{"dimension": "Years", "members": ["FY25", "FY26"]}, {"dimension": "Period", "members": ["TotalYear"]}],
 "pov": {"Scenario": "Target 2025", "Version": "Final"}}
```

This was smoke-tested on 2026-10-13: BANORWAY, Target 2025, Final, FY25 TotalYear returned 7,518,989,026.37 at the top Account level.
No FinPlan recipe has been reconciled against an archive file yet.

## Server-side recipes (sv_answer)

### Which recipe the gate names

`sv_gate` returns `recipe: {name, params}` for a clear question: the first matching rule of `core/scope.py`
`_recipe`. Use it as given; anything else is a departure (`deviation`, below).

| # | When | Recipe | Params from the gate |
|---|---|---|---|
| 1 | grain supplier or project, and a cost-type measure with accounts (consultants, contractors, legal, travel …) | `lines_ranking` | `accounts`, `entity`, `period`, `by: supplier`, `projects` (true when projects/initiatives are asked) |
| 2 | grain month, or "trend" / "utvikling" | `monthly_trend` | `entity`, `accounts`, `through` |
| 3 | grain cost centre | `cost_base` | `parent`, `accounts`, `period` |
| 4 | budget or forecast named, and the entity has a FinPlan member | `actual_vs_plan` | `entity` (FinPlan), `month`, `year` |
| 5 | a P&L measure (concept or KPI) | `pnl` | `area`, `period` |
| 6 | a concept from the library | `concept` | `concept`, `period`, `entity` |
| 7 | a measure with accounts only | `cost_base` | `parent`, `accounts`, `period`, `top: 10` |
| — | none of these | `null` | write a free-form spec (`pulls` + `layout`) |

### Finance defaults (each recipe's `DEFAULTS`)

| Recipe | Default | Departs when |
|---|---|---|
| `pnl` | compare with last year (plus budget and the latest forecast when FinPlan has the entity) | `compare` without `ly` |
| `actual_vs_plan` | Target and the latest forecast are always shown | — |
| `monthly_trend` | the same months last year | `compare` without `ly` |
| `cost_base` | top 10 + Other + total + share %, last year | `top` ≠ 10 |
| `concept` | the same period last year | `compare` empty |
| `lines_ranking` | top 10 + Other; last year (`compare: ["ly"]`); group companies separate | `top` ≠ 10, `compare` ≠ `["ly"]`, `group_separate: false` |
| free-form | a ranking (`repeat.over.top`) is the top 10 with an "Other" line | another N, or no "Other" line |

A departure is allowed only on the user's explicit request: `spec.deviation = "<the request>"`. `sv_answer` refuses
it otherwise and lists the departures; with it, the sheet says "Deviation from the standard analysis (user's
request): …" under the assumptions. Budget is compared at total level only, where FinPlan maps the measure (FinPlan
has no supplier or project grain).

### `lines_ranking`

One `sv_answer` call: it draws `Q<n> Answer` (question, assumptions, "Drilling ledger lines …") before any drill,
drills the FinLite ledger lines of each account (d1 = this year, d2 = the same months last year; about 60 s each,
with progress notes), writes `Q<n> Lines d1` / `Q<n> Lines d2`, then renders:

- **Invoiced by supplier:** the top 10 external suppliers by `AP_CUSTOMER_NAME` (costs negative, biggest cost
  first), "Other external suppliers", "Total external suppliers"; columns this year, last year, Change, Change %,
  Share % (of the cube total).
- **Group companies:** own block, never in the top 10 (`knowledge/catalog.toml` `[group_companies]`: a counterparty
  column or name patterns).
- **Reconciliation:** "Accruals / reversals / journals without supplier" (lines without a supplier), "Total (ledger
  lines)", "Unexplained gap (ledger lines vs cube)" = cube − lines, and "Total (cube)" from a FinLite pull. Checks
  "Blocks = lines total" and "Total = cube". A gap over 1 NOK is a Warning finding; there is no automatic re-drill.
- **`projects: true` → `Q<n> By project`**, hyperlinked both ways: booked projects (`PROJECT_DESC` when `PROJECT`
  ≠ 000000) and "Initiative (derived, not booked)" (the cost centre `CC_DESC`, else the start of the line text),
  top 10 + Other each, check "By project = lines total".
- **Fallbacks:** this year's drill fails → the cube total per account only, with a warning; last year's drill fails →
  the last-year columns show the cube total only.
- `key_values` carries "Total (cube)"; `followups` suggest the invoice lines of the largest supplier, the monthly
  trend and contractors.
