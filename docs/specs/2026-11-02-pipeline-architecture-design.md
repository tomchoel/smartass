# smartass pipeline architecture — design (2026-11-02)

Status: approved (grilling 2026-11-02, four rounds). Supersedes the tool layout in `docs/DESIGN.md` §Tools once the
new core reaches parity; supersedes `2026-10-27-headless-hsgetvalue-design.md` (spike branch, not merged).

## 1. Why

smartass grew feature by feature: 34 MCP tools, six of which overlap on retrieve-and-report (`sv_query`,
`sv_ws_pull`, `sv_ws_build`, `sv_analyse`, `sv_concept`, `fin_pnl`), ~23 k lines in `engine/smartview_mcp`, a
retired `division_pnl`, and a DuckDB warehouse that mainly exists to dodge slow Excel retrieves. A headless spike
built closed `.xlsx` files fast but invisibly, outside the MCP. There is no single main flow.

Measured (docs/audit, 2026-10-26): Wholesale `fin_pnl` 220 s, unit P&L 46 s. Time goes to cold Excel/SSO
(15–27 s) and Smart View ad-hoc retrieves via COM (7–10 s per sheet), not to arithmetic. A one-block COM write of a
whole answer already takes ~0.2 s (`answer.py` `_Cells.flush`).

## 2. What the tool is

An AI finance analyst for Telenor Norge controllers. A user asks a finance question in natural language; the tool
scopes it like an experienced controller would, pulls the numbers from Oracle EPM (FinMax, FinLite, FinPlan) and
builds a traceable answer in the user's open Excel workbook.

Jobs, in build order: (1) ad hoc question → answer sheet; (2) recurring reports as saved, re-runnable answer
specs; (3) Produktregnskapet as a domain pack. Front end: MCP agents (Copilot CLI, Claude) until the backend is
solid; an Office add-in later renders the same contract.

Principles, in priority order:
1. **Correct and traceable** — every number traces to a live EPM source cell; derived numbers are Excel formulas.
2. **Honest** — the tool says when a question cannot be answered as asked, and labels hypotheses.
3. **Visible** — the user sees the answer form in Excel while it is built.
4. **Fast** — quality first, but no needless waiting.

## 3. The pipeline

```
question ─► scope ─────────────► pull ─────────────────► answer ─────────────► check
            fixed checklist       svhttp fetch            answer spec (stored)   reconciliation
            capability map        → source sheet          skeleton first         formula errors
            clear │ ask │ proxy     (value grid +         one-block render       sheet = fetch
            assumptions             provenance header)    formulas only          findings table
            draft answer spec     one-block write         follow-up = patch      variance commentary
```

Four MCP tools: `scope`, `pull`, `answer`, `check`. Utilities stay: `find_member`, `member_info`, `drill`,
`status`, and the `xl_*` workbench. Recipes (P&L, actual-vs-plan, monthly trend, cost base, concepts) are Python
functions that return a draft answer spec; they are not tools. The same four calls are exposed on the local HTTP
API (`httpapi.py`) so the add-in later is only a UI.

### 3.1 scope — the advisor gate

Runs a **fixed checklist on every question** (research: AMBROSIA, NeurIPS 2024 — models do not reliably notice
ambiguity unprompted):

| Check | Example of a gap |
|-------|------------------|
| Period | "this year" — YTD to last closed month, or full-year forecast? |
| Scenario / version | "budget" — FinPlan Target Final, or latest forecast? |
| Entity scope | "Business" — un.SY.2BUS incl. or excl. sub-units? |
| Measure definition | "opex" — which accounts; EBITDA before or after other items? |
| Currency / sign | NOK; costs negative in FinMax |
| Granularity feasibility | the measure exists at the asked level in some cube (capability map) |

Outcomes:
- **clear** — proceed; assumptions are written at the top of the answer sheet.
- **ambiguous** — 1–3 clarifying questions in one round, only for gaps that change the number. No pull.
- **infeasible** — say why (e.g. "opex is not allocated to products in FinMax"), show what is possible, and draft
  the best proxy (e.g. contribution margin = revenue − COGS per product, or Produktregnskapet allocations).

Answers to clarifying questions may be saved as **learned defaults** in a readable `preferences.toml`, so the
gate asks once.

Implemented (phase 4) in `core/scope.py`, exposed as the tool `sv_gate(question)`; the procedure the agent follows is
`knowledge/advisor/gate.md`.
- **Balanced policy:** ask only when the gap changes the number and has no safe convention; otherwise proceed and
  state the convention as an assumption (no period → YTD to the last closed month; no entity → Telenor Norge; no
  scenario → actual, or forecast for a full year; NOK and FinMax signs). At most 3 questions, in checklist order.
- **Infeasible beats ambiguous:** actuals for an open month, a non-NOK currency, or a measure at a grain the
  capability map rejects (e.g. EBITDA per product) → infeasible with a reason and a proxy.
- Warnings, not questions: abolished entities (with successors) and units with no FinPlan member (last year only).
- The closed month comes from the cache or HTTP only; the gate never starts Excel (≤ 0.05 s per call).
- **Learned defaults:** `sv_remember(item, phrase, value)` writes `%LOCALAPPDATA%\smartass\preferences.toml`
  (`paths.preferences`; one file per user, readable, editable by hand; promotion to the team catalog later). Items:
  measure, period ("this year" → ytd / fy), scenario, entity; values are validated against the catalog.
- English questions, assumptions and sheets; Norwegian synonyms are understood.
- Measured on the gate benchmark (B10, B11, T01–T10): 12/12, false-go 0, needless-ask 0
  (`docs/audit/2026-11-02-bench-core-gate.md`).

### 3.2 pull

Fetch through the Smart View HTTP provider (`svhttp.py`; transport = hidden Edge with Telenor SSO). Write each
source sheet into the session workbook in **one block**. Source sheet form, decided after risk spike R1
(`docs/archive/audit/2026-11-02-risk-spikes.md`; HsGetValue formulas written into live Excel show `#NEED_REFRESH` until
Smart View refreshes them, so they cannot be pre-filled from the fetch):
- **C (default):** value grid written from the HTTP fetch (0.02–0.4 s). No Excel Smart View login on the critical
  path. The pull spec is stored in the workbook; "refresh" = `pull` re-fetches over HTTP and rewrites the block.
- **B (on demand):** "make Smart View-refreshable" converts a source sheet to an HsGetValue formula grid (about
  2.3 s) when Excel's own Smart View is signed in; the engine connects the private connection on a hidden helper
  sheet first, because Smart View Refresh does not connect.
- **A (fallback):** Smart View ad-hoc grid via COM, only when svhttp fails.

Each source sheet has a provenance panel: cube, POV, fetch time, route (HTTP / HsGetValue / ad hoc). It sits to
the **right** of the grid (two columns after the last grid column, from row 2), not above it: the grid starts at A1
in Smart View's own shape (row 1 = POV, then column headers, then row labels), so `grid.parse`, `load_pulls` and the
label-keyed lookups of `answer.py` work unchanged on HTTP and ad hoc sheets alike. A header above the grid would
shift every label those lookups read, and extra cells in row 1 would be read as appended POV members.
Implemented in `core/source.py` (walking skeleton, B02).

**Excel login hardening** (applies to B and A): detect a stuck or unresponsive Smart View Login window or the
"Empty response from provider" state within seconds, tell the user, and restart only the engine's own Excel —
never the user's.

### 3.3 answer

1. Create the answer sheet first: question, assumptions, status line ("Fetching FinMax …").
2. Render the layout from the answer spec in one block once pulls land; every derived cell is a formula into source
   sheets; no engine-side maths.
3. Store the answer spec in the workbook (hidden, per answer sheet).
4. A follow-up patches the stored spec (`specs.merge`) and redraws in one block; cells the user edited outside the
   tool-owned range survive.

Workbook model: one session workbook; a new question = a new answer sheet (`Q<n> Answer`); a follow-up edits the
current answer unless the user says "new question"; source sheets are shared across answers. English everywhere.

Walking skeleton (B02): `core/progress.py` writes step 1 (title, question, status line, assumptions) before any
pull and updates the status line per pull; `answer.build(..., suffix="Answer")` then clears the sheet and renders
step 2; `core/pipeline.py` stores the raw spec (step 3). Tool: `sv_answer` (`tools_core.py`), beside the old tools.

Implemented (phase 5, `docs/plans/2026-11-02-widen.md`):
- **Follow-up** = `sv_answer(followup=<patch>, q=None)`; `spec` and `followup` are exclusive. `q` defaults to the
  active `Q<n> Answer`, else the latest stored v1 answer. `contract.patched` (D2): dicts merge and `null` deletes;
  `pulls` merge by `id`, `layout.sections` by `title`, `lines` by `key`, `checks` by `label`; `{..., "drop": true}`
  removes an item, unmatched items are appended; `{"append": [...]}` adds to a list (notes, key points); any other
  list replaces. `recipe.params` cannot be patched (a new question). A follow-up never re-runs the gate.
- **Source reuse** (D3, `core/followup.py`): an unchanged pull (same normalised `contract.Pull`) keeps its sheet,
  read back from `_svmeta`, with no fetch; a changed pull's sheet is parked (`~name`) and fetched fresh, then dropped
  on success or restored on failure; a removed pull's sheet is deleted after success; `layout.expect` of a re-fetched
  pull is dropped.
- **Tool-owned range** (D4, `core/owned.py`): `A1:<last col><last row>` of what the tool drew (table, checks, notes,
  findings) is stored with the spec (`owned`). A redraw clears only that block and deletes only charts named
  `smartass chart*`; rows/columns are inserted first if a taller/wider answer would hit user cells;
  `styling.apply_report` stops at `roles["last_row"]`. Answers drawn before phase 5 have no `owned` and are cleared
  in full as before.
- **Backup** (D5): the answer sheet is copied to a hidden `~Q<n> Answer` before the redraw and restored if the redraw
  fails; follow-up status goes to Excel's status bar only.
- **Recipes** (D6–D8, `core/recipes/`): `spec.recipe = {name, params, overrides}` without `pulls`/`layout`;
  `sv_answer` expands it inside the Excel job (closed month, catalog, member labels, P&L facts in a
  `RecipeContext`), applies `overrides` as a patch, validates, and stores the expanded spec with the `recipe` field.
  Recipes: `pnl` (wraps `fin_spec.build`), `actual_vs_plan`, `monthly_trend`, `cost_base` (top-N ranked after the
  pull through `repeat.over.top`), `concept`. `sv_answer` returns `period` and an `outline` (pull ids, columns,
  section titles, line keys, check labels) to address follow-ups.

### 3.4 check

- Reconciliation checks from the spec (e.g. sum of months = accumulated YTD), formula errors, sheet value = fetched
  value.
- Findings table Critical / Warning / Info (structure vendored from Anthropic `financial-services` `audit-xls`,
  Apache-2.0).
- Variance commentary (structure vendored from `variance-commentary`, Apache-2.0): materiality from `catalog.toml`,
  always-comment list (revenue, EBITDA, FTE), IFRS 8 10 % threshold for per-segment questions. Drivers come from
  pulled data (drill to sub-accounts/cost centres) first; informed hypotheses are allowed but labelled
  "Hypothesis — not verified in data".

Implemented (phase 4) in `core/commentary.py`; vendored text in `knowledge/advisor/` (`variance-commentary.md`,
`findings.md`, `NOTICE.md`):
- **The engine flags, the agent explains.** `pipeline.run` returns `material` (the flagged lines with value,
  comparison, delta, pct); the agent writes `commentary.drivers` (`line`, `text`, `basis` = data | hypothesis), and
  the drivers appear in the answer's key points.
- A line is material when the first `var` cell passes the catalog threshold (pct with a NOK floor, and/or absolute,
  per `rule`). Stock lines (FTE) are tested on pct only; aux lines are skipped; a zero comparison is material only
  on the absolute threshold.
- Findings: a material line without a driver → **Warning**; an always-comment line (revenue, EBITDA, FTE) without a
  driver → **Info**. `Driver.line` must name a layout line (validated in `contract.py`).

Implemented (phase 5):
- **Findings block on the sheet** (D10, `core/findings.py`): after the build and the review, a "Findings" table
  (severity, text; at most 15 lines plus "and N more") is written below the notes, inside the owned range.
- **Recon checks per recipe**: sum of months = YTD (`monthly_trend`), top + other = total (`cost_base`), FinPlan
  EBITDA = FinMax (`actual_vs_plan`, `severity: "Warning"`: the cubes group accounts differently, so a difference is
  a finding to show, not an untrustworthy number), the concept's own checks, and for `pnl` the warehouse = source
  sheet reconciliation carried as `layout.expect` and turned into static checks after the pull (D9).
- `Check.severity` (`Critical` default | `Warning`) decides how a DIFF is graded in the review.

## 4. The contract: answer spec v1

The answer spec is the single contract that the agent writes, recipes produce, follow-ups patch, recurring reports
save and the add-in will render. v1 wraps the proven layout grammar of `specs.py` / `answer.py` (columns, sections,
lines, cell kinds `pull | var | varpct | sum | avg | expr | ref | rowsum | value | drill`, checks, chart) instead of
inventing a new one.

```jsonc
{
  "version": 1,
  "question": "Business Norway EBITDA per month 2026 vs 2025",
  "gate": {"outcome": "clear", "checklist": {"period": "...", "scenario": "...", ...}},
  "assumptions": ["Actual, NOK, periodic months; YTD = Jan–Sep 26"],
  "pulls": [{"id": "m", "connection": "FinMax", "rows": [...], "columns": [...], "pov": {...}}],
  "layout": {"title": "...", "columns": [...], "sections": [...], "checks": [...], "chart": {...}},
  "commentary": {"materiality": {...}, "drivers": [{"text": "...", "basis": "data" | "hypothesis"}]}
}
```

Defined as pydantic models in `engine/smartview_mcp/core/contract.py`; the JSON schema is generated from them.

## 5. Semantic layer

One place for meaning: `knowledge/catalog.toml` (metrics, areas, scenarios, KPIs, materiality) extended with a
capability map — which measures exist at which granularity in which cube, Norwegian + English synonyms, and proxy
rules for infeasible asks.

Implemented (phase 4): `knowledge/capabilities.toml` (measures with EN/NO synonyms, vague measures with their
options, grains with the measures they support and the proxy, currency words) read by `core/semantic.py`; periods
by `core/period.py`; entities, scenarios and materiality stay in `catalog.toml`; concept matches via
`concepts.scored()`. The map encodes these verified cube facts (atlas 2026-10-13, live `sv_members` 2026-11-02):
- No customer dimension in any cube.
- Products: FinMax `All Products` (19 755 members), FinPlan `All Products` (243). Opex is not allocated to
  products; the full cost per product exists only in Produktregnskapet (annual).
- Currency: FinMax `All Currencies` has NOK, Local and FX Rates; no EUR reporting member.
- FinPlan Entity has 337 members (business areas and some units); units below them have no FinPlan member.
- FinMax scenarios: Actual, Target, ForecastQ1–Q4, Act v Tgt, Act Dev Last Yr / Last Per, Act v FCQ1–Q4.
  FinPlan scenarios: Actual, Target, Forecast Qnyy, Forecast Running; version Final / Working.

## 6. What is reused, frozen, dropped

| Part | Fate |
|------|------|
| `svhttp.py`, `svxml.py`, atlas, `catalog.toml`, `grid.py` | reused |
| `answer.py` one-block writer, `styling.py`, `specs.py` | reused |
| `fin_spec.py`, `templates.py`, `concepts.py` | become recipes |
| `hsgetvalue.py` (from spike) | reused for source sheet form B |
| `drill*`, `xl_*`, `changes`/revert, `doctor` | kept as utilities |
| `headless_builder.py`, spike generator scripts | dropped (closed file, invisible) |
| `sv_query`, `sv_ws_pull`, `sv_ws_build`, `sv_ws_answer`, `sv_analyse`, `sv_concept`, `fin_pnl`, `sv_scope` | deleted after parity |
| `warehouse.py` (DuckDB) | frozen; after R2: delete, or hidden closed-month cache; never a sheet source |

Old engine: bug fixes only during the rework. Target ≥ 40 % fewer engine lines at parity; modules ≤ ~400 lines.

Implemented (phase 5): `fin_spec.py`, `templates.py` and `concepts.py` are reused unchanged in logic by the recipes
(`templates.cost_base_members` extracted; concept strings in English). Still on the old tools until the phase 6
parity deletion: `sv_drill` (B12; drill cells in a v1 spec are phase 6) and every tool in the "deleted after parity"
row, which stay registered (35 tools).

## 7. Risks, tested first (go/no-go)

| Spike | Pass criterion | If it fails |
|-------|----------------|-------------|
| R1 HsGetValue | ~500-cell grid written in one block into an open workbook; Smart View refresh < 15 s; values = svhttp exactly; connection resolves without manual fixes | source sheets = form A (ad-hoc grids) |
| R2 svhttp | everything B15 (full Wholesale P&L) needs fetched < 10 s; session survives 1 h idle; parallel fetch safe | keep warehouse as closed-month cache; COM fallback |
| R3 live write | one-block write survives edit mode / Excel busy / another workbook active, with a retry policy; never touches non-session workbooks | write into a session workbook opened by the engine, user switches to it |
| R4 EPM REST (timeboxed) | official REST data export works with corporate SSO | stay on svhttp (undocumented provider protocol) + COM fallback |

## 8. Acceptance: benchmarks

B01–B15 (`docs/specs/2026-11-02-benchmarks.md`) plus ~10 trap questions for the gate. Each run records
correctness (independent `sv_query` oracle, checks OK, no formula errors), time to first visible output, total time
and the gate outcome. Old tools are deleted when the new core passes them all and a one-week trial with 2–3
controllers has turned its failures into benchmark cases.

## 9. Governance

Read-only. The user's own EPM identity (Smart View `#NoAccess` = their permissions). MIP "Internal" label on new
session workbooks. A run log per answer (hidden sheet + local JSONL): question, gate outcome, assumptions, pulls,
timings, check results. No numbers leave the machine beyond the agent's enterprise model.

## 10. Out of scope

The Office add-in, recurring-report scheduling, the Produktregnskapet pack, new cubes, Oracle EPM's own AI agents
(cloud-side; noted for later).
