# smartass · Presentation, ledger-lines & findings polish (design, 2026-11-25)

Status: **approved by owner on 2026-11-25.** Covers ledger-lines summary layout and codes, table headers and horizontal border hygiene, `Q<n> Data` omission on `lines`-only answers, `Q<n> Lines <alias>` column ordering/widths/POV wrapping, `Key points` coverage and agent priority, and cost-centre `sign_findings` noise reduction.

## Problem

Real controller workbooks surfaced six concrete presentation and signal-to-noise issues:

1. **`lines` tables lacked context in Col B / Col C and missed line counts / mixed-sign splits.** `group_by="account"` showed only account numbers in Col C with Col B empty; `group_by="supplier"` left Col B empty; `group_by="project"` left `PROJECT` out of Col B; `kpi_tiles` omitted `metric_code()`; and table headers left Col B and Col C blank. `"Lines without a account"` had a grammar error, 0-line `"Lines without ..."` rows still rendered, `"Lines without a supplier"` gave no breakdown of which journals/sources created those lines, and mixed-sign drills showed a meaningless `"Share %"` instead of `"Debit | Credit"`.
2. **Horizontal borders bled beyond table spans.** `TN Tittel`, `TN Seksjon`, `TN Kolonne` / `TN Kolonne tall`, and `Sum`/`Hovedsum` styles were painted across `1..last_col`, bleeding navy/dark horizontal borders into Col A, into chart helper columns, and across empty columns on narrow blocks.
3. **`lines`-only answers created an empty `Q<n> Data` sheet.** `Q<n> Data` had 0 facts while `Source:` still pointed to `Data_Q<n>`, `print_area_col` extended past `E` on narrow 2-column `lines`-only sheets, and mixed-cube answers (`FinMax`/`FinPlan` + `LinesBlock`) omitted `FinLite` from the subtitle and `Source:`.
4. **`Q<n> Lines <alias>` sheets put technical columns first with static widths and `F1`/`A2` spillover.** Important analytical columns (`Period`, `Entity`, `CC_DESC`, `ACCOUNT`, `ACCOUNT_NAME`, `AP_CUSTOMER_NAME`, `AMOUNT`, `DEBET_AMOUNT`, `CREDIT_AMOUNT`, `GL_DESCRIPTION_LINE`, `JOURNAL_NAME`, `JE_SOURCE`, `PROJECT`, `PROJECT_DESC`, `PRODUCT_DESC`) sat far to the right, column widths ignored actual cell content, and long `A1`/`A2`/`G1` texts collided or spilled past `F1:K3`.
5. **Cost-centre / small-unit `sign_findings` fired false `"critical"` alerts.** Negative `EBITDA`, `Gross margin`, and `Revenue` on sub-divisions, cost centres, or units with `|Revenue| < 50 000` were flagged `"critical"`.
6. **`Key points` covered only 1 table and could silently drop the agent's points.** Only the first table contributed material lines, and the 5-point cap truncated the agent's `free` points without warning.

## Design

### 1. Column B (`CODE_COL`), Column C (`LABEL_COL`), and table headers

- **`model/blocks/lines.py`**:
  - `group_by="account"`: `Line.code` holds `ACCOUNT`; `Line.label` holds the `ACCOUNT_NAME` with the largest `|amount|` among rows in that account group (falling back to the account code when `ACCOUNT_NAME` is absent/blank). Chart category labels and `Key points` read `sheet_label(ln)` from Col C and therefore show the account name.
  - `group_by="project"`: `Line.code` holds the dominant non-zero `PROJECT` code; Col C shows the project name (`INITIATIVE`).
  - `group_by="supplier"`: `Line.code` holds the dominant `ACCOUNT` (largest `|amount|` in the group), formatted as `"43200"` for a single account or `"43200 +<k>"` when the supplier spans `k + 1` distinct accounts (`k >= 1`).
- **`model/blocks/table.py`**:
  - `build_kpi`: populate `code=metric_code(ctx.resolver, m, q.cube)` (returns `""` for `FinPlan`).
- **`render/layout.py`**:
  - Each table placed by `place(o, r)` writes `"Code"` in `(header_row, CODE_COL)` and a contextual header in `(header_row, LABEL_COL)`:
    - `lines`: `"Account"`, `"Supplier"`, or `"Project"` from `LinesBlock.group_by`.
    - `table` / `chart_data`: `"Unit"` (`rows="entities"`), `"Product"` (`rows="products"`), `"Metric"` (`rows="metrics"`).
    - `ranking`: `"Product"` (`rows="products"`) or `"Unit"` (`rows="entities"`).
    - `matrix`, `drivers`: `"Unit"`.
    - `kpi_tiles`, `bridge`, `concept`: `"Metric"`.
    - `trend`: `"Period"`.
  - Sum/text/FinPlan rows without a code keep Col B empty (`""`) without lint or check errors.

### 2. `lines` summary table on `Q<n> Answer` (`model/blocks/lines.py`, `model/expr.py`, `render/layout.py`, `render/keypoints.py`)

- **Columns**:
  - Uniform signs (all numeric `AMOUNT` values `>= 0` or all `<= 0`): `["Amount", "Share %", "Lines"]` (`cols = ["amount", "share", "lines"]`).
  - Mixed signs (both `< 0` and `> 0` present): `["Amount", "Debit", "Credit", "Lines"]` (`cols = ["amount", "debit", "credit", "lines"]`).
  - `"lines"` uses number kind `"heltall"`; `"debit"` and `"credit"` use `"heltall"`.
- **`LinePick` extensions (`model/expr.py`)**:
  - `agg: Literal["sum", "count", "debit", "credit"] = "sum"` and `extra: tuple[tuple[str, tuple[str, ...]], ...] = ()` (optional additional exact column filters).
  - `to_excel`: emits `SUM`/`ROWS`/`SUMIFS`/`COUNTIFS` (or `SUM(SUMIFS(...))` / `SUM(COUNTIFS(...))`) when criteria are safe for `SUMIFS`, and `FILTER` / `SUM(--(...))` when any criterion fails `safe_for_sumifs`. Extracts `Q<n>` cleanly from either `Data_Q<n>` or `Lines_Q<n>_...`.
- **Blank-group label & suppression**:
  - `"Lines without an account"` (fixing `"a account"`), `"Lines without a supplier"`, `"Lines without a booked project"`.
  - When a `"Lines without ..."` group has 0 matching rows, omit the `"blank"` row (and adjust `"other"` formulas to subtract `named` directly from `total` without referencing `"blank"`).
- **No-supplier journal/source breakdown (`group_by="supplier"`)**:
  - Keep `"Lines without a supplier"` as a `level=0` summary line when it has `> 0` rows.
  - When at least one blank-supplier row has a non-empty `JOURNAL_NAME` or `JE_SOURCE`, group the blank-supplier rows by journal/source group (using `JE_SOURCE` directly when non-empty and not `"Spreadsheet"`, or extracting the normalized journal prefix/identifier from `JOURNAL_NAME` after stripping leading `"Reverses "` when `JE_SOURCE` is `"Spreadsheet"` or blank, e.g. `HG 12`, `ABA`, `MF`, `SF`), rank by `|amount|` descending, and place the top 5 directly underneath as `level=1` indented child lines (`extra=((group_col, raw_values),)`).
  - `_lines_points` in `render/keypoints.py` and `_chart` in `render/layout.py` restrict `lines` groups to `ln.level == 0`, and `_lines_points` reads `Share %` only when `"share"` is in `o.cols` (preventing mixed-sign `Debit` from being formatted as `%`).

### 3. Sheet, border, and `Q<n> Lines <alias>` hygiene

- **Border/style span restriction (`render/layout.py`, `render/look.py`, `xl/styling.py`)**:
  - `layout.py` records `"end_col"` on table `section`, `header`, `line`/`sum`/`grand`, and `check` row specs (for `place(o, r)`: the table's last column `D + len(headers) - 1`; for `_threshold_table`: its last populated column; for `Parameter`: `D`; for `Checks`: `D`).
  - `look.py` and `styling.py` restrict `TN Tittel`, `TN Undertittel`, `TN Note`, and general `TN Seksjon` to `first_label .. text_last` (`B..print_last_col`), and restrict per-table `TN Seksjon`, `TN Kolonne` / `TN Kolonne tall`, and `line`/`sum`/`grand` non-helper cells to `first_label .. end_col`. Col A (`< first_label`) and helper/empty columns outside `end_col` (except helper cells explicitly in `num`) receive `"TN Tom"`.
- **Omit empty `Q<n> Data` for `lines`-only answers (`exceljob.py`, `render/xlsx.py`, `render/layout.py`)**:
  - When `not pre.spec.queries`, `exceljob._build` skips creating `DataSheet` and `Header` (`data = None`), `xlsx.build` accepts `data: DataSheet | None` and omits the `Q<n> Data` sheet, `_record_status` skips `Q<n> Data` when `header is None`, and `_question_parts` omits `Q<q> Data` when the question has `Q<q> Lines ...` sheets and no `Q<q> Data` sheet (while still refusing a missing `Q<q> Data` on non-lines questions).
  - `Source:` on `lines`-only sheets points to `table Lines_Q<n>_b0 (sheet Q<n> Lines b0)` (or all lines tables if multiple).
- **`print_area_col` and mixed-cube subtitle/`Source:` (`render/layout.py`)**:
  - On narrow sheets with a single chart in a row whose table spans `B..E` (`>= 460` pt, e.g. 2-col narrow `lines`-only sheets), clamp the lone chart's width to the table width so `print_area_col` stays within `E` (and on 3-col `lines` sheets `D..F`, `print_area_col == "F"`).
  - Include `"FinLite"` in the cubes list for both `answer_subtitle` and `Source:` whenever `spec.blocks` contains a `LinesBlock`.
- **`Q<n> Lines <alias>` sheet (`render/xlsx.py`)**:
  - When `len(ls.rows) <= config.DRILL_XL_FULL`, order `ANALYTICAL_COLS` (`Period`, `Entity`, `CC_DESC`, `ACCOUNT`, `ACCOUNT_NAME`, `AP_CUSTOMER_NAME`, `AMOUNT`, `DEBET_AMOUNT`, `CREDIT_AMOUNT`, `GL_DESCRIPTION_LINE`, `JOURNAL_NAME`, `JE_SOURCE`, `PROJECT`, `PROJECT_DESC`, `PRODUCT_DESC`) first (those present in `ls.columns`), followed by all remaining columns in original order so structured table references continue to work.
  - Compute each column's width from actual header + formatted cell content (including `C3:D3` and `F1:K3` header labels/values), capped at `LINE_COL_MAX = 50` (and at least `header_width(c)`).
  - Merge `A1:E1`, `A2:E2`, `G1:K1`, and `G2:K2` with `text_wrap=True` (and set row 1 / row 2 heights when wrapped) so `A1`/`A2` never collide with `F1`/`F2` and `G1` POV never spills past `K1`.

### 4. `Key points` & cost-centre `sign_findings`

- **`model/validate.py` (`sign_findings`)**:
  - When `s == "income"` and `v < 0` for `EBITDA`, `EBITDA before other items`, `Gross margin`, or `Revenue` (by metric name, member, or display label): emit `Finding("info", ...)` instead of `"critical"` if `q.entity` is a sub-division or cost-centre unit (not company total `un.SY`/`All Entities` and not a top-level `0/1/2` division code like `2BUS`) **or** if `Revenue` for query `q` is present in `ctx.frame` with `|Revenue| < 50_000`.
  - Cost metrics with positive values (`s == "cost"` and `v > 0`) and negative income metrics on revenue-generating divisions (`|Revenue| >= 50 000`) remain `"critical"`.
- **`render/keypoints.py`, `model/validate.py`, `exceljob.py`**:
  - Raise `MAX_POINTS` and `MAX_KEY_POINTS` from `5` to `7`.
  - In `key_points`, pull material lines from up to 2 main tables (found via `_tables_for_points` on `res.outs` when available, falling back to `[main]`), pooled by `|variance|` descending (keeping `MAX_MATERIAL = 4` cap when there is a single table, up to `avail` across 2 tables).
  - Reserve slots for all of the agent's `free` points first (up to 7) so agent `key_points` are never silently dropped, fill remaining slots (`7 - len(free)`) with the top engine material points (kept in engine-first, agent-last display order), and append a clear warning to `warnings` whenever engine points are truncated to fit `free` points.
