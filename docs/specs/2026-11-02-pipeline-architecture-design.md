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
            clear │ ask │ proxy     (HsGetValue grid,     one-block render       sheet = fetch
            assumptions             values pre-filled)    formulas only          findings table
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

### 3.2 pull

Fetch through the Smart View HTTP provider (`svhttp.py`; transport = hidden Edge with Telenor SSO). Write each
source sheet into the session workbook in **one block**. Source sheet form, decided by risk spike R1:
- **B (preferred):** HsGetValue formula grid with values pre-filled from the fetch. Every cell carries its exact POV;
  Smart View refresh keeps it live.
- **A (fallback):** Smart View ad-hoc grid via COM.

Each source sheet has a provenance header: cube, POV, fetch time, route. COM retrieve stays the fallback when
svhttp fails.

### 3.3 answer

1. Create the answer sheet first: question, assumptions, status line ("Fetching FinMax …").
2. Render the layout from the answer spec in one block once pulls land; every derived cell is a formula into source
   sheets; no engine-side maths.
3. Store the answer spec in the workbook (hidden, per answer sheet).
4. A follow-up patches the stored spec (`specs.merge`) and redraws in one block; cells the user edited outside the
   tool-owned range survive.

Workbook model: one session workbook; a new question = a new answer sheet (`Q<n> Answer`); a follow-up edits the
current answer unless the user says "new question"; source sheets are shared across answers. English everywhere.

### 3.4 check

- Reconciliation checks from the spec (e.g. sum of months = accumulated YTD), formula errors, sheet value = fetched
  value.
- Findings table Critical / Warning / Info (structure vendored from Anthropic `financial-services` `audit-xls`,
  Apache-2.0).
- Variance commentary (structure vendored from `variance-commentary`, Apache-2.0): materiality from `catalog.toml`,
  always-comment list (revenue, EBITDA, FTE), IFRS 8 10 % threshold for per-segment questions. Drivers come from
  pulled data (drill to sub-accounts/cost centres) first; informed hypotheses are allowed but labelled
  "Hypothesis — not verified in data".

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
