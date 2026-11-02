# "Advisor plans" Implementation Plan (simplified)

> **For agentic workers:** REQUIRED SUB-SKILL: use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task by task. Steps use checkbox (`- [ ]`) syntax for tracking.
> Batches A → B → C run in order; tasks inside a batch run in order. A step marked **LIVE** needs EPM + Excel: the
> owner runs it, and the subagent stops and reports when it reaches one.

**Goal:** every number question ends with live answer sheets built by the engine. The question that failed in the
live test must work end to end: *"Hvilke konsulentselskaper fakturerte mest til Telenor Norge i 2026? og på hvilke
initiativer / prosjekt"*.

In that test the agent:
- ignored the gate's "ambiguous" outcome and picked the accounts itself;
- drilled, did its own sums and put the table only in chat;
- asked "do you want a sheet?";
- on the retry, got a sheet full of where-cell errors (COM 0x800A03EC) and Norwegian assumptions.

**Scope (owner, simplified):**
- **No playbooks and no separate plan object.** `sv_gate` names the **recipe** (+ args) to use, as one more key of
  its existing output. A question with no specific recipe uses the existing free-form spec path.
- **Finance defaults live in each recipe, in code:**
  - same period last year;
  - budget/forecast when FinPlan has the entity;
  - reconciliation to the cube total;
  - rankings = top 10 + Other + total + share %;
  - group companies separate.
  - A departure from a default needs the user's explicit request, and it is written on the sheet.
- **The engine enforces:**
  - no pull while the gate is ambiguous and unanswered;
  - the answer sheet's first frame before any drill;
  - English-only sheet text.
- **New recipe `lines_ranking`:**
  - it drills internally and ranks invoiced amounts;
  - accruals/reversals go on a reconciling line;
  - the initiative comes from the project when booked, else it is "derived, not booked";
  - an extra `Q1 By project` sheet is hyperlinked from `Q1 Answer`.
- **Fix the where-cell bug.**
- **Cost-type measures in `capabilities.toml`.**
- **Read-back:** key values plus ≤ 3 follow-ups in the reply.
- **Docs**, then a **live re-run** that becomes a bench case.
- The vendored Anthropic advisor text stays unchanged.

**Architecture (what changes, nothing else):**
- **`core/scope.py`**
  - `gate()` returns `recipe: {name, params} | None`, chosen by a small rule table in code (`_recipe`, about 40
    lines) from the measure, grain and period it already resolves.
  - `gate()` also accepts `answers={item: value}`, the user's replies to the asked items.
- **`tools_core.answer`**
  - Before any Excel job it re-runs the gate on `spec.question` with `spec.gate.answers`. If that is still
    ambiguous, it refuses to pull.
  - It refuses a recipe other than the gate's, or a recipe param that switches off a finance default, unless
    `spec.deviation` holds the user's request. The deviation is written on the sheet.
  - It runs the English check (`contract.non_english`).
  - It returns `key_values` and `followups`.
- **`core/drillstep.py`**: the async drill half of `tools_drill.py`, moved so a recipe can drill inside
  `sv_answer`.
  - A recipe with drills runs as: Excel job A (first frame) → drills (async, progress notes) → Excel job B (lines,
    cube pull, render, checks).
- **`core/recipes/lines.py`**: the new recipe.
  - It uses `contract.extras` for the topic sheet (`Q<n> <topic>`, rendered by `answer.build(..., suffix=topic)`).
- **The where-cell fix is in `answer_write.py`, `pipeline.py` and `excel/sheets.py`** (root cause below).
- **The advisor text is untouched:** `knowledge/advisor/variance-commentary.md`, `findings.md` and `NOTICE.md`
  are not changed. The recipes' LY/budget columns feed `commentary.material` and `check.review` exactly as today.

**Tech Stack:** Python 3.12, pydantic v2, pywin32 COM (Excel), Playwright (hidden Edge: svhttp + drill), pytest +
anyio, ruff, uv.

**Run everything from `C:\dev\smartass\engine`** (PowerShell 5.1: no `&&`; use `;` and `if ($?) {...}`):
- unit tests: `uv run pytest -q -p no:cacheprovider <files>`;
- full suite: `uv run pytest -q -p no:cacheprovider` (about 80 s; **813 passing** on `2b6c237`);
- lint: `uv run ruff check smartview_mcp tests`;
- live bench (owner): `uv run python tests\live\bench.py --route core --only B16,T11,T12 --out advisor-plans`.

Commit after each task (from `C:\dev\smartass`):
`git add -A; git commit -m "<message>" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"`.
One commit per task. Never edit files with `Get-Content | Set-Content` (it mangles UTF-8); use the edit/create
tools. Do not push.

**Line count:** `engine/smartview_mcp` is **16 718** lines on `2b6c237`. The expected delta is **about +550 to
+700**, landing at ≤ 17 450. Use the script in `docs/plans/2026-11-02-retire-harden.md`, "Line count".

| Change | Lines |
|---|---|
| new `core/recipes/lines.py` | ≈ +300 |
| new `core/drillstep.py`, moved from `tools_drill.py` | net ≈ +40 |
| `scope.py` | ≈ +60 |
| `tools_core.py` | ≈ +90 |
| `contract.py`, incl. the English check | ≈ +70 |
| `reply.py` | ≈ +35 |
| `drill_xl.py`, `answer_write.py`, `answer.py`, `excel/sheets.py`, `pipeline.py`, `progress.py`, `semantic.py`, `catalog.py`, `recipes/cost.py`, `recipes/trend.py` | ≈ +120 together |

No module may pass ~400 lines. `tools_core.py` (238 lines today) must stay ≤ 400; if it would pass that, the
two-job drill flow (Task B3) goes into `core/drillstep.py`.

---

## Root cause of the where-cell failures (COM −2146827284 = 0x800A03EC)

Evidence: `%LOCALAPPDATA%\smartass\logs\queries.jsonl`, `runs.jsonl` and `server.log`, 2026-11-02, 18:15–18:25.

1. **18:15 / 18:16:** `sv_drill` ran without `q`. `tools_drill` uses `q or sheets.next_question(t)`, which gave
   q = 1. It wrote the sheets `Q1 Lines d1` / `Q1 Lines d2` with the tables `Lines_Q1_d1` / `Lines_Q1_d2`.
2. **18:20:54:** the retry called `sv_answer` with a new spec. `next_question` saw the Q1 sheets and gave **Q2**.
   The drill cells (`{"drill": "d2", "where": {"AP_CUSTOMER_NAME": "KPMG AS"}}`) then resolve to
   `drill_xl.table_name(2, "d2")` = `Lines_Q2_d2`. That table did not exist; the q=2 drills only ran at 18:21:47.
3. `answer_write._write` asks `drill_xl.table_columns(wb, "Lines_Q2_d2")`, which returns `None` for a missing
   table. `missing_columns(None, where)` returns `[]` ("cannot check"). So the engine writes
   `=SUMIFS(Lines_Q2_d2[AMOUNT],Lines_Q2_d2[AP_CUSTOMER_NAME],"KPMG AS")`.
4. Excel rejects a structured reference to an unknown table. `Formula2` raises
   `(-2147352567, …, (0, None, None, None, 0, -2146827284), None)`. The batch write fails, and the
   cell-by-cell fallback warns once per cell, two per supplier.
5. The 18:24:37 rerun at Q2 worked, because the Q2 drills existed by then.

Reproduced in a private Excel (`DispatchEx`): the same formula raises the identical tuple when the table is
missing, and works on the existing `Lines_Q1_d1`.

**Fix (Task A1):**
- refuse drill cells whose table is missing, before anything is written, naming the table and the existing
  `Lines_*` tables;
- a definitely-missing table becomes a cell error, never a formula;
- a new spec with drill cells takes the q of a pending drill (Lines sheets without an Answer).

The lines recipe (Batch B) drills inside `sv_answer`, so its q always matches.

---

## Decisions (owner-settled first, then mechanisms chosen here; owner may override the latter)

| # | Decision | Mechanism |
|---|---|---|
| D1 | **Every number question builds answer sheets automatically.** Never ask "do you want a sheet". Several sheets per request are fine. The chat keeps a full summary. | The SKILL, server `INSTRUCTIONS`, gate `next` text and the `sv_drill` reply say so. `test_skill_docs` guards the wording. |
| D2 | **The gate names the recipe.** `sv_gate` returns `recipe: {name, params}` (or `None` → free-form spec). There are **no** playbooks and **no** plan object or schema. | `scope._recipe(found, grains, period, entities)` is a rule table in code (Task A3). This supersedes the earlier "analysis plan" decisions (8, 9, 14). |
| D3 | **Finance defaults are in each recipe, in code:** same period LY; budget/forecast when FinPlan has the entity; reconciliation to the cube total; rankings top 10 + Other + total + share %; group companies separate. | Each recipe module has `DEFAULTS` (a dict). Existing recipes already compare (pnl `COMPARISONS`, concept LY, actual_vs_plan). Task A5 adds LY + top 10/share to `cost_base` and an LY series to `monthly_trend`. |
| D4 | **Deviations are written on the sheet.** Using another recipe than the gate's, or a param that switches off a default (`compare=[]`, `top≠10`, `group_separate=false`, …), needs `spec.deviation` = the user's explicit request. | `tools_core.answer` compares the spec with `scope.gate(...)["recipe"]` and the recipe's `DEFAULTS`. Without a deviation it raises `InputError` listing the departures. With one, the skeleton writes "Deviation from the standard analysis (user's request): …" under the assumptions, and `runs.jsonl` logs it. |
| D5 | **No pull while the gate is ambiguous and unanswered.** | `answer()` re-runs `scope.gate(spec.question, closed, answers=spec.gate.answers)` before any Excel job; it is cheap and local. If the outcome is still ambiguous, it returns the questions and pulls nothing, whatever `spec.gate.outcome` claims. `answers` is new on `contract.Gate` and on `sv_gate`, so the agent resolves asks without rewording the question. |
| D6 | **First frame before the drill: ≤ 2 s after the workbook is open** (owner-accepted; a cold open of 19–23 s is not counted). | A recipe whose draft needs drills runs as job A (`progress.show`: question, assumptions, "Drilling ledger lines (about 60 s) …"), then async drills with `timing.say` notes, then job B. Non-drill recipes draw the frame before `_expand_recipe`. |
| D7 | **English-only sheet text; chat in the user's language.** | `contract.non_english(text)` is a Norwegian-marker heuristic: ≥ 2 Norwegian function words, or æ/ø/å outside quoted names and member tokens. `answer()` checks assumptions, title/subtitle/notes/key points, section titles, column headers, check labels, drivers, extras and the deviation. It raises `InputError` listing the strings, and the agent translates. Exempt: line labels and data names (suppliers, members). The question is shown as "Question (as asked)". gate.md/SKILL.md change from "ask in English" to "ask in the user's language". |
| D8 | **The lines/ranking recipe is one `sv_answer` call:** `lines_ranking` drills internally. | Batch B. `sv_drill` stays as a workbench utility. |
| D9 | **Rankings rank invoiced amounts** (`AP_CUSTOMER_NAME`, costs negative, biggest cost first). Rows without a supplier (accruals, reversals, spreadsheet journals) form the reconciling line "Accruals / reversals / journals without supplier", so the total ties to the cube. | Checks: "Blocks = lines total" and "Total = cube" (which holds by construction, see D20). |
| D10 | **Group companies go in their own block** and are never counted in the top 10. | `catalog.toml [group_companies]`: a lines column naming the counterparty if the drill has one, else regex patterns (initially `^Telenor\b`, `\bTSS\b`, `Shared Services`, `Towers`). They are verified live in Task B4. |
| D11 | **Initiative/project:** the booked `PROJECT_DESC` when `PROJECT` ≠ `000000`; otherwise derived from `CC_DESC` or line text, in a column headed "Initiative (derived, not booked)". | `lines.add_initiative(rows)` adds `INITIATIVE` / `INITIATIVE_DERIVED` key columns to the Lines rows. |
| D12 | **Topic sheets** are named "Q1 Answer" + "Q1 <topic>" (e.g. "Q1 By project") and hyperlinked both ways. | `contract.AnswerSpec.extras: [{topic, layout}]`, merged by topic in follow-ups. `answer.build(..., suffix=topic)`. `sheets._Q` recognises stored topics as ours. Links use `=HYPERLINK("#'Q1 By project'!A1", …)`. |
| D13 | **Chat numbers are read back from the sheet; ≤ 3 follow-ups.** | `sv_answer` returns `key_values` (line label → column → value from `res["values"]` after `wait_calculated`; ≤ 40 lines) and `followups` (each recipe's `FOLLOWUPS`, at most 3, filled in). The SKILL says: quote `key_values`, never your own maths. |
| D14 | **Cost-type measures are added** (consultants = C60080 + C60090, contractors/innleie, legal, travel, …) with English and Norwegian synonyms, without user confirmation. | `Measure.accounts` (a list). The accounts were looked up via atlas while writing this plan (table in Task A2). New grains `supplier` and `project`. A bare current year ("i 2026") = YTD through the closed month, by convention. |
| D15 | **The vendored Anthropic advisor material is unchanged and wired in the check step.** | `test_vendored_advisor_text_unchanged` pins the sha256 of `variance-commentary.md`, `findings.md` and `NOTICE.md` (Task A5). The recipes' LY/budget columns feed `commentary.material` and `check.review` as today. |
| D16 | **The free-form path stays** for questions with no specific recipe (`recipe: None`), under the same enforcement (D5, D7, D13). | Nothing to add beyond D5/D7/D13. |
| D17 | **LY per supplier comes from a second drill** (FinLite, same accounts and entity, previous year). It is filtered to the LY months with a list-valued `where` on `Period` (an OR via `FILTER/ISNUMBER(MATCH())`). If the LY drill fails, LY is shown at total level only (cube pull), with a warning. | Task B2. This keeps the answer whole when the second drill fails (risk R1). |
| D18 | **Budget for rankings is at total level only,** when FinPlan maps the measure; otherwise a note says so. | FinPlan has no supplier or project grain. |
| D19 | **Live re-run of the consultant question, added to the bench** as B16, with gate cases T11/T12. | Batch C. |
| D20 | **No automatic re-drill** (owner). An unexplained difference between the ledger lines and the cube is shown as its own reconciling line, "Unexplained gap (ledger lines vs cube)" = cube − lines total, and it is flagged as a finding. | The recipe passes `check.review(extra=[...])` a **Warning** finding: "Ledger lines differ from the cube by <amount> (<lines> lines); shown as 'Unexplained gap'. Ask again to re-drill if needed." No finding is raised when the gap is \|gap\| ≤ 1 NOK. The status line says "lines do not fully reconcile". |
| D21 | **Legal** (owner): a combined measure `legal` (fees 67410/67420 + disputes C60400) is the default for the word "legal"; `legal_fees` and `legal_disputes` stay as separate measures for the specific phrases. | The longest synonym wins (`semantic._find`): "legal fees" → `legal_fees`, "legal disputes" → `legal_disputes`, "legal" / "legal costs" / "juridiske kostnader" → `legal`. |
| D22 | **Top 10 + Other everywhere** (owner), including `cost_base` (top 25 → 10) and any recipe or free-form ranking: top 10, "Other", total and share %. | `DEFAULTS["top"] = 10` in `cost_base` and `lines_ranking`. Another N is a departure (D4). |

---

## File structure

| File | Action | Why |
|---|---|---|
| `engine/smartview_mcp/answer_write.py`, `core/pipeline.py`, `excel/sheets.py`, `drill_xl.py` | modify | A1 where-cell fix; B2 list-`where`; B3 topic sheets |
| `knowledge/capabilities.toml`, `core/semantic.py` | modify | A2 measures (`accounts`), grains |
| `core/scope.py` | modify | A2 bare year; A3 `recipe`, `answers` |
| `core/contract.py` | modify | A4 `Gate.answers`, `deviation`, `non_english`; B3 `extras` |
| `tools_core.py`, `httpapi.py`, `core/progress.py`, `core/runlog.py` | modify | A3/A4 enforcement, deviation row; A5 read-back |
| `core/reply.py` | modify | A5 `key_values`, `followups` |
| `core/recipes/__init__.py`, `cost.py`, `trend.py` (+ `FOLLOWUPS`/`DEFAULTS` in each recipe) | modify | A5 defaults |
| `core/drillstep.py` | create | B1 async drill, moved from `tools_drill.py` |
| `tools_drill.py` | modify | B1 |
| `core/recipes/lines.py` | create | B2–B3 recipe `lines_ranking` |
| `knowledge/catalog.toml`, `catalog.py` | modify | B2 `[group_companies]` |
| `answer.py` | modify | B3 extras rendering |
| `engine/tests/test_drill.py`, `test_advisor.py`, `test_contract.py`, `test_core_skeleton.py`, `test_recipes.py`, `test_vendor.py`, `test_skill_docs.py`, `test_httpapi.py` | modify | tests per task |
| `engine/tests/test_lines_recipe.py` | create | B2–B4 |
| `engine/tests/live/bench.py` | modify | C2 |
| `skills/smartass/SKILL.md`, `reference/{operations,recipes,drill}.md`, `knowledge/advisor/gate.md`, `server.py` (`INSTRUCTIONS`), `docs/specs/2026-11-02-pipeline-architecture-design.md`, `docs/specs/2026-11-02-benchmarks.md` | modify | C1 |
| `knowledge/advisor/variance-commentary.md`, `findings.md`, `NOTICE.md` | **unchanged** | D15 |

---

## Batch A — engine rules and measures (offline)

### Task A1: fix the where-cell 0x800A03EC error (missing Lines table, wrong q)

**Files:** `engine/smartview_mcp/answer_write.py` (≈ lines 110–145), `drill_xl.py` (`table_columns`, new
`lines_tables`), `core/pipeline.py` (`check_drill_tables`, `question_for`; called in `_new`/`_follow`),
`excel/sheets.py` (`pending_drill_question`); tests in `engine/tests/test_drill.py`. Reuse the `_Sheet`, `_Rect`
and `_Target` fakes and the `answer_write.com` monkeypatch.

- [ ] **Step 1: failing tests**

```python
def test_answer_refuses_drill_cell_whose_table_is_missing_and_names_existing_tables():
    t = _target_with_tables(["Lines_Q1_d1", "Lines_Q1_d2"])
    with pytest.raises(InputError) as e:
        pipeline.check_drill_tables(t, 2, _spec(_drill_cell("d2", {"AP_CUSTOMER_NAME": "KPMG AS"})))
    assert all(s in str(e.value) for s in ("Lines_Q2_d2", "Lines_Q1_d2", "sv_drill", "q=2"))

def test_answer_write_marks_missing_table_as_cell_error_not_formula(monkeypatch):
    written, warnings = _render_drill_cell(tables={})          # find_table -> None
    assert not any("SUMIFS(" in str(v) for v in written) and any("Lines_Q2_d2 not found" in w for w in warnings)

def test_answer_write_keeps_cannot_check_path_when_tables_unreadable(monkeypatch):
    written, _ = _render_drill_cell(tables=RAISES)             # today's behaviour: formula written
    assert any("SUMIFS(" in str(v) for v in written)

def test_new_spec_with_drill_cells_takes_the_pending_drill_q():
    t = _target_with_sheets(["Q1 Lines d1", "Q1 Lines d2"])   # the live case
    assert pipeline.question_for(t, _spec(_drill_cell("d2", {}))) == 1
    assert pipeline.question_for(t, _spec_without_drill()) == 2
    assert sheets.pending_drill_question(_target_with_sheets(["Q1 Answer", "Q1 Lines d1"])) is None
```

- [ ] **Step 2:** `uv run pytest -q -p no:cacheprovider tests\test_drill.py`. Expect FAIL.
- [ ] **Step 3: implement.**
  - `table_columns` returns `None` only when `find_table` finds nothing; exceptions propagate.
  - `check_drill_tables` walks every `drill` cell (lines, sections, extras later). It raises
    `InputError(f"The drill table(s) {missing} do not exist in this workbook (Lines tables here: {have or 'none'}).
    Drill first with sv_drill(q={q}, id='<alias>'), or use the recipe the gate named (lines_ranking drills
    itself).")`. A non-`InputError` exception while reading the tables is skipped ("cannot check", as today).
  - In `answer_write`, `cols is None` writes `"#n/a: <table> not found"` plus a warning instead of the formula.
  - `pending_drill_question` = the highest q that has `Q<q> Lines…` sheets and no `Q<q> Answer`.
  - `question_for` uses it when the spec has drill cells, otherwise `next_question`.
- [ ] **Step 4:** PASS: the same file, plus `tests\test_core_skeleton.py`. The existing
  `test_answer_writes_drill_formulas` and `test_answer_warns_on_where_column_missing_from_table` must stay green.
- [ ] **Step 5:** commit `fix(answer): refuse drill cells whose Lines table is missing; adopt the pending drill's q (0x800A03EC)`.

### Task A2: cost-type measures, supplier/project grains, a bare year

**Files:** `knowledge/capabilities.toml`, `engine/smartview_mcp/core/semantic.py` (`Measure.accounts:
tuple[str, ...] = ()`, `finplan: str | None = None`; `describe` joins the accounts with " + "),
`core/scope.py` (period rule); tests in `engine/tests/test_advisor.py`.

Accounts were looked up on 2026-11-02 with `sv_find_member` and `sv_members`. FinLite and FinMax use the same
names. The tree is `ac.C609` → `ac.C693` (consultancy + temps) and `ac.C694` (other expenses).

| key | accounts | contents | synonyms (EN + NO, whole words) |
|---|---|---|---|
| `consultants` | `ac.C60080`, `ac.C60090` | 67600/67610 IT advisory; 67099, 67100/67110 audit, 67200, 67300, 67410/67420 legal fees, 67500/67501, 67700 | consultants, consultant, consultancy, consultancies, consulting, consulting firms, consulting companies; konsulent, konsulenter, konsulentselskap, konsulentselskaper, konsulentfirma, konsulentfirmaer, konsulenthonorar, rådgivning |
| `consultants_and_temps` | `ac.C693` (the old mapping) | C60080 + C60090 + C60110 + C60095 | consultants and temps, consultancy and temporary workers, konsulenter og innleie |
| `contractors` | `ac.C60110` | 67800/67810 temps, 67910 security, 67920, 67930, 67940 innleie, 67970, 67989, 67990 | contractors, temporary workers, temps, hired staff; innleie, innleid, innleide, vikar, vikarer, fremmedtjenester |
| `legal` (default for "legal", D21) | `ac.67410`, `ac.67420`, `ac.C60400` | fees + disputes | legal, legal costs, legal expenses; juridisk, juridiske kostnader, juridiske utgifter |
| `legal_fees` | `ac.67410`, `ac.67420` | inside C60090 | legal fees, lawyers, legal advice; advokat, advokater, advokatkostnader, juridisk bistand |
| `legal_disputes` | `ac.C60400` | – | legal disputes, litigation; rettstvister, tvister |
| `audit_fees` | `ac.67100`, `ac.67110` | inside C60090 | audit fees, auditor; revisjon, revisor, revisjonshonorar |
| `travel` | `ac.C60130` | 71405, 71442, 71443 … | travel, travel costs, travel expenses, per diem; reise, reiser, reisekostnader, diett |
| `management_fee` | `ac.C60170` | – | management fee, management fees |
| `licence_fees` | `ac.C60180` | – | licence fees, license fees; lisenser, lisenskostnader |
| `energy` | `ac.C60025` | – | energy, electricity; energi, strøm, strømkostnader |
| `shared_services` | `ac.C60910` | – | shared services charges, shared service centre |

The consultants note reads: "C60080 + C60090; temporary contract workers (C60110) are 'contractors'".

New grains:
- `supplier`
  - synonyms: per supplier, by supplier, which suppliers, which companies, who invoiced, invoiced most, top
    suppliers, biggest suppliers; per leverandør, hvilke leverandører, hvilke selskaper, fakturerte mest,
    største leverandører;
  - feasible for the account-list measures above;
  - note "from ledger lines (FinLite drill)".
- `project`
  - synonyms: per project, by project, which projects, initiatives; per prosjekt, hvilke prosjekter, initiativer;
  - feasible for the same measures.

Period rule: a bare current year ("i 2026", "in 2026") with actuals is YTD through the closed month, stated by
convention. "This year" stays an ask (T02 is unchanged).

- [ ] **Step 1: failing tests**

```python
@pytest.mark.parametrize("q,key", [
    ("Hvilke konsulentselskaper fakturerte mest til Telenor Norge i 2026?", "consultants"),
    ("Which consulting firms invoiced us most this year?", "consultants"),
    ("Hvor mye brukte vi på innleie i september?", "contractors"),
    ("Travel costs YTD for Business", "travel"), ("advokatkostnader 2026", "legal_fees"),
    ("Legal costs this year", "legal"), ("legal fees 2026", "legal_fees"), ("legal disputes Q3", "legal_disputes")])
def test_cost_type_measures_match_en_and_no(q, key): assert semantic.measures(q)[0].key == key

def test_consultants_sum_c60080_and_c60090():
    assert semantic.measure("consultants").accounts == ("ac.C60080", "ac.C60090")
```

  Add to `GATE_CASES`:
  - `("T11", "Hvilke konsulentselskaper fakturerte mest til Telenor Norge i 2026? og på hvilke initiativer / prosjekt", "clear")`;
  - `("T12", "Which suppliers invoiced most on travel in 2026?", "clear")`.

  Add `test_gate_supplier_question_states_definition_grain_and_ytd`.
- [ ] **Steps 2–4:** FAIL → implement → PASS (`tests\test_advisor.py tests\test_recipes.py`). T09 ("How much did
  we spend on consultants?") must stay `clear`, and `semantic.problems()` must report no duplicate synonyms.
- [ ] **Step 5:** commit `knowledge: cost-type measures (consultants = C60080+C60090, contractors, legal, travel, ...), supplier/project grains; T11-T12`.

### Task A3: the gate names the recipe; `answers` resolves the asks

**Files:** `engine/smartview_mcp/core/scope.py` (`_recipe`, `gate(..., answers=None)`, `NEXT`),
`tools_core.py` (`sv_gate(question, answers=None)`), `httpapi.py` (`/gate` body `answers`); tests in
`test_advisor.py`, `test_httpapi.py`.

`_recipe(found, grains, period, ents, outcome)` is a first-match table in code. It returns `None` for ambiguous
outcomes.

| when | recipe + params |
|---|---|
| grain supplier/project and the measure has `accounts` | `lines_ranking {accounts, entity, period, by: "supplier", projects: <project grain?>}` |
| grain month / "trend" | `monthly_trend {entity, accounts, through}` |
| grain cost_centre | `cost_base {parent, accounts?, period}` |
| a budget/forecast scenario and a FinPlan entity | `actual_vs_plan {entity, month, year}` |
| a P&L concept measure (revenue, cogs, opex, ebitda, …) | `pnl {area, period}` |
| the measure came from the concept library | `concept {concept, period, entity}` |
| an account-list measure, no grain | `cost_base {parent: entity, accounts, period, top: 10}` |
| else | `None` (free-form spec) |

- `answers[item]` sets that checklist item with `source="user"` and `ask=False`.
- `NEXT["clear"]` becomes: "Call sv_answer with spec {version 1, question, gate (copy), recipe: <the recipe
  above>} now; it builds the sheets. Never ask whether to build a sheet. Quote key_values from its reply."
- `NEXT["ambiguous"]`: "Ask these questions in the user's language, in one message; then call sv_gate again with
  the same question and answers={...}."
- `recipe` is `None` → "write a free-form spec".

- [ ] **Step 1: failing tests**:
  - `test_gate_names_recipe` (parametrized: T11 → `lines_ranking` with `accounts == ["ac.C60080","ac.C60090"]`;
    "EBITDA September" → `pnl`; B07-like trend → `monthly_trend`; B14 → `actual_vs_plan`; B11 → `None`);
  - `test_gate_answers_resolve_the_asks` (B11 + `{"measure": "opex"}` → clear with a recipe);
  - `test_http_gate_accepts_answers`.
- [ ] **Steps 2–4:** FAIL → implement → PASS.
- [ ] **Step 5:** commit `core.scope: the gate names the recipe (+args); answers= resolves an ambiguous gate`.

### Task A4: `sv_answer` enforces: ambiguity, deviations, English

**Files:** `engine/smartview_mcp/core/contract.py` (`Gate.answers: dict[str, str] = {}`, `AnswerSpec.deviation:
str | None` (≤ 300 chars), `non_english(text)`, `sheet_texts(spec)`), `tools_core.py` (`answer()` pre-checks),
`core/progress.py` (deviation row, "Question (as asked)"), `core/runlog.py` (log `deviation`); tests in
`test_core_skeleton.py`, `test_contract.py`.

The pre-checks run in `answer()` for a new question, **before any Excel job**:
1. `g = scope.gate(spec.question, closed_cached, answers=spec.gate.answers)`. Here `closed_cached` is
   `_closed_month()`'s cache, or the previous month.
   - If `g` is ambiguous → return `pipeline.gate_reply` built from `g`'s questions. Nothing is pulled.
2. `departures` = the recipe name differs from `g["recipe"]["name"]`, or a param differs from the recipe's
   `DEFAULTS` in a way that switches a default off (`recipes.departures(name, params)`).
   - If `departures` and no `spec.deviation` → `InputError("This departs from the standard analysis: " + "; ".join(departures)
     + ". Use the recipe and defaults sv_gate named, or pass deviation='<the user's explicit request>' (it is
     written on the sheet).")`.
   - A free-form spec when the gate named a recipe counts as a departure.
3. `bad = [(path, s) for path, s in contract.sheet_texts(spec) if contract.non_english(s)]`.
   - If any → `InputError` listing ≤ 8 `path: text` lines + "Write sheet text in English (the chat stays in the
     user's language); names and member labels may stay as they are."

`non_english`:

```python
_NO = frozenset("og ikke på av til med kun fra som er det den har ble utledet valgt bekreftet kontoene beløp "
                "inkluderer gjelder periode regnskap hittil fjor tall kostnader leverandør".split())
def non_english(s: str) -> bool:
    words = re.findall(r"[A-Za-zÆØÅæøå]+", _QUOTED_OR_MEMBER.sub(" ", s))   # 'quoted', "quoted", un.SY / ac.X
    return sum(w.lower() in _NO for w in words) >= 2 or any(c in "æøåÆØÅ" for w in words for c in w)
```

- [ ] **Step 1: failing tests**
  - `test_answer_refuses_to_pull_when_the_gate_is_still_ambiguous` (spec claims clear; question "How are we doing
    on costs?" → questions returned, the fake host is never called);
  - `test_answer_pulls_when_answers_resolve_the_gate`;
  - `test_answer_refuses_a_departure_without_deviation` (gate says `lines_ranking`, spec `cost_base`);
  - `test_deviation_is_written_on_the_sheet_and_logged`;
  - `test_norwegian_assumption_is_refused_with_its_path` (the live string "Konsulenter = kontoene C60080 og C60090
    (valgt av meg)");
  - `test_english_text_and_quoted_names_pass` ("Supplier 'Bøler Bygg AS'", "un.SY");
  - `test_question_is_shown_as_asked_and_not_checked`.
- [ ] **Steps 2–4:** FAIL → implement → PASS. Then the full suite. Existing tests whose specs disagree with the
  gate get either a matching recipe or a `deviation` (keep the fixes minimal; list them in the commit body).
- [ ] **Step 5:** commit `answer: no pull while the gate is ambiguous; departures need a written deviation; sheet text must be English`.

### Task A5: recipe finance defaults, read-back and follow-ups

**Files:** `core/recipes/__init__.py` (`departures()`, `DEFAULTS`/`FOLLOWUPS` per recipe), `core/recipes/cost.py`,
`core/recipes/trend.py`, `core/reply.py` (`key_values`), `tools_core.py` (return keys); tests in
`test_recipes.py`, `test_core_skeleton.py`, `test_vendor.py`.

- `cost_base`:
  - default `top` 25 → **10** + "Other cost centres" + total (D22: top 10 + Other everywhere);
  - add "Share %" (line / total) and an LY column (the same pull with the LY period in the columns);
  - the existing check "Top + other = total" stays, and a check "Total = cube parent" is added.
  - Deviation: `top ≠ 10`.
- Free-form specs: a `repeat.over.top` section with `top ≠ 10`, or with no "Other" line, is a departure. It is
  refused without `deviation`; with one, the deviation is written on the sheet (D4, D22).
- `monthly_trend`: add an LY series (the same months last year), unless `compare=[]` (a departure).
- `pnl`, `concept`, `actual_vs_plan` already carry LY/plan. Their `DEFAULTS` just declare it (`compare` must
  include LY; budget when FinPlan has the entity, which pnl already handles with the "last year only" warning).
- `key_values(prepared, res, limit=40)` reads label → {column label → value} from `res["values"]` (filled after
  `sheets.wait_calculated` by `answer.build`).
- `sv_answer` returns `key_values`, `followups` (the recipe's `FOLLOWUPS`, at most 3, formatted with the entity
  and period; for free-form specs `[]`) and `sheets` (every sheet written).
- [ ] **Step 1: failing tests**
  - `test_cost_base_defaults_top10_other_share_and_ly`;
  - `test_freeform_ranking_above_top10_without_deviation_is_refused`;
  - `test_trend_has_ly_series_by_default`;
  - `test_departures_flags_switched_off_defaults`;
  - `test_reply_key_values_are_the_sheet_values_after_calc`;
  - `test_reply_has_at_most_three_followups`;
  - `test_vendored_advisor_text_unchanged`. It pins the sha256 of `knowledge/advisor/variance-commentary.md`,
    `findings.md` and `NOTICE.md`, taken from `git show 2b6c237:<file>` at the start of the task.
- [ ] **Steps 2–4:** FAIL → implement → PASS (`tests\test_recipes.py tests\test_core_skeleton.py tests\test_vendor.py`).
- [ ] **Step 5:** commit `recipes: finance defaults in code (LY, top 10 + share), read-back key_values, <=3 follow-ups`.

---

## Batch B — the `lines_ranking` recipe

### Task B1: move the async drill to `core/drillstep.py`

**Files:** create `engine/smartview_mcp/core/drillstep.py`; modify `tools_drill.py`; tests in `test_drill.py`.

- `drillstep.drill(cell, *, say) -> DrillResult(rows, columns, cell_value, reconciliation, profile, source,
  secs)` is today's `_drill_cell` + `_lines_of`, moved.
  - New: one retry with a note on "Empty response from provider" (the first live drill failed on that).
- `drillstep.write(t, q, alias, result, title)` is the Excel half (`drill_xl.write_lines` + log source).
- `sv_drill` = `drill` + one job calling `write`. Its reply `next` says: "To answer a question, use sv_gate →
  sv_answer (recipe lines_ranking drills and builds the ranking)."
- [ ] **Step 1:** move the monkeypatch targets in the tests; add
  `test_drill_retries_once_on_empty_provider_response` (FAIL).
- [ ] **Steps 2–4:** FAIL → move → PASS (`tests\test_drill.py`); `tools_drill.py` ≤ 260 lines.
- [ ] **Step 5:** commit `refactor(drill): core.drillstep so recipes can drill; one retry on an empty provider response`.

### Task B2: recipe `lines_ranking` (ranking, group block, reconciling line, LY) + group companies

**Files:** create `engine/smartview_mcp/core/recipes/lines.py`; modify `core/recipes/__init__.py` (register;
`RecipeContext.lines: dict[str, list[dict]] | None` = drilled rows by alias; a `Draft.result["drills"]` contract),
`drill_xl.py` (list-valued `where`), `core/contract.py` (drill `where` value `str | list[str]`),
`knowledge/catalog.toml` + `catalog.py` (`[group_companies]`); tests in `engine/tests/test_lines_recipe.py`,
`test_drill.py`, `test_catalog.py`.

`lines_ranking(params, ctx)` has two phases, both pure:
- **`drills(params, ctx)`** returns the cells to drill:
  - `d1` = FinLite {accounts, entity, `Yr <yy>`} (filter `Period` to ≤ the closed month);
  - `d2` = the same for the previous year (D17).
- **`draft(params, ctx)`** (with `ctx.lines`) returns the `Draft`:
  - **Pull `cube`:** FinLite accounts × [YTD this year, YTD LY], entity, Accumulated. Plus a FinPlan total if the
    measure has `finplan` (D18).
  - **Keys** (Python), with the values still drill formulas:
    - suppliers = non-empty `AP_CUSTOMER_NAME`, not group, ranked by `sum(AMOUNT)` ascending (costs negative);
    - top `top`;
    - the group suppliers;
    - an empty supplier = reconciling.
  - **Section "Invoiced by supplier — YTD <period>"**, columns: YTD yy, YTD yy-1, Change, Change %, Share %.
    - `s1..s10` = `{"drill":"d1","where":{"AP_CUSTOMER_NAME":name,"Period":[months]}}`, and LY with `d2` and the
      LY months;
    - `other` = expr `ext - s*`;
    - `ext` = "Total external suppliers".
  - **Section "Group companies"**: one line per group supplier, and "Total group companies".
  - **Section "Reconciliation"**:
    - "Accruals / reversals / journals without supplier" (where `AP_CUSTOMER_NAME` = "");
    - "Total (ledger lines)";
    - "Unexplained gap (ledger lines vs cube)" = cube − lines total (D20; no automatic re-drill);
    - "Total (cube)" = the pull.
  - **Checks:**
    - "Blocks = lines total" (ext + group + no-supplier − lines);
    - "Total = cube" (lines + gap − cube; it holds by construction).
  - **Finding:** when |gap| > 1 NOK, `Draft.result["findings"]` carries a Warning that `check.review(extra=...)`
    reports: "Ledger lines differ from the cube by <amount> (<n> lines); shown as 'Unexplained gap'". The status
    line then says "lines do not fully reconcile".
  - **Assumptions** (English):
    - the accounts + entity;
    - "Ranking = invoiced amounts per supplier (AP lines); accruals, reversals and spreadsheet journals have no
      supplier and form a reconciling line";
    - "Group companies shown separately";
    - "NOK, EPM signs (costs negative)";
    - "Last year = the same months of <yy-1>".
  - **Fallback:** if `ctx.lines` has no `d2` (the LY drill failed), the LY columns are the cube LY total only, and
    a warning is added.
- `DEFAULTS = {"top": 10, "compare": ["ly"], "group_separate": True}`. `FOLLOWUPS`:
  - "Show the invoice lines for {top1}";
  - "Monthly trend of {measure} this year vs last year";
  - "Top suppliers on contractors (C60110)".

`drill_xl.drill_formula`: a list value becomes `ISNUMBER(MATCH(T[col], {"Jan-26","Feb-26",…}, 0))` in the FILTER
form. An empty list is an `InputError`.

`[group_companies]` in `catalog.toml`:
- `column = ""` (a counterparty column, if the drill has one);
- `patterns = ['^Telenor\b', '\bTSS\b', 'Shared Services', 'Towers']`.

- [ ] **Step 1: failing tests** use fixture rows that mirror the live data: KPMG, Accenture, Sopra Steria, TSS,
  Telenor ASA, Towers, and empty-supplier journals.
  - `test_ranking_top10_excludes_group_and_orders_by_cost`
  - `test_ranking_has_other_external_group_and_reconciling_lines`
  - `test_ranking_checks_tie_blocks_to_lines_and_lines_to_cube`
  - `test_ranking_unexplained_gap_is_a_reconciling_line_and_a_warning_finding` (fixture cube 11.2 M above the
    lines → the gap line is 11.2 M, one Warning finding, no re-drill call)
  - `test_ranking_no_gap_finding_within_1_nok`
  - `test_ranking_ly_cells_use_a_list_where_on_ly_months`
  - `test_ranking_without_ly_drill_falls_back_to_total_ly`
  - `test_ranking_assumptions_pass_the_english_check`
  - `test_drill_formula_list_value_is_an_or`
  - `test_group_company_patterns_match_telenor_tss_towers_not_kpmg`
- [ ] **Steps 2–4:** FAIL → implement → PASS (`tests\test_lines_recipe.py tests\test_drill.py tests\test_catalog.py tests\test_recipes.py`).
  `test_every_data_recipe_has_a_check` covers the new recipe.
- [ ] **Step 5:** commit `recipe lines_ranking: invoiced ranking top 10 + other, group block, reconciling line, LY, checks`.

### Task B3: the "By project" topic sheet, derived initiative, links

**Files:** `core/contract.py` (`extras: list[Extra(topic, layout)]`, `_IDS["extras"]="topic"`, topic regex
`^[A-Z][A-Za-z0-9 &/-]{1,24}$`), `answer.py` (render extras with `suffix=topic`), `excel/sheets.py` (`_Q` and
`rollback`/`next_question` know the stored topics), `core/progress.py` (link rows), `core/recipes/lines.py`
(`add_initiative`, the extra), `drill_lines.py` (`KEY_LABELS` += `INITIATIVE`, `INITIATIVE_DERIVED`); tests in
`test_lines_recipe.py`, `test_contract.py`, `test_core_skeleton.py`.

- `add_initiative(rows)` sets:
  - `INITIATIVE` = `PROJECT_DESC` (or `PROJECT`) when `PROJECT` is filled and ≠ `000000`;
  - otherwise `INITIATIVE_DERIVED` = `CC_DESC`, else the first 40 chars of `GL_DESCRIPTION_LINE` with digits-only
    tokens removed.
- Extra `By project` (only when the gate's grain has `project`, `params.projects=True`):
  - section "Booked projects": top suppliers × `INITIATIVE`;
  - section "Initiative (derived, not booked)": × `INITIATIVE_DERIVED`;
  - check "By project = lines total".
- Row 4 of `Q<n> Answer` holds `=HYPERLINK("#'Q<n> By project'!A1","Q<n> By project →")`, and each topic sheet
  links back.
- [ ] **Step 1: failing tests**
  - `test_initiative_is_booked_project_else_derived_and_marked`
  - `test_by_project_sheet_sums_to_lines_total`
  - `test_answer_links_to_topic_sheets_and_back`
  - `test_topic_sheet_is_ours_for_rollback_and_next_question`
  - `test_extras_merge_by_topic_in_followups`
- [ ] **Steps 2–4:** FAIL → implement → PASS.
- [ ] **Step 5:** commit `lines_ranking: 'Q<n> By project' sheet with booked/derived initiative, hyperlinked`.

### Task B4: first frame before the drill; end-to-end offline; **LIVE** smoke

**Files:** `tools_core.py` (`answer()`: two-job flow when the recipe declares drills; if `tools_core.py` would pass
400 lines, put the flow in `core/drillstep.run_with_drills`), `core/pipeline.py` (`run(..., q=, frame_done=True)`
skips `progress.show`); tests in `test_lines_recipe.py`, `test_core_skeleton.py`.

Flow:
1. **Job A:**
   - open the target and take `q = pipeline.question_for(t, spec)`;
   - `progress.show` with the question, assumptions, deviation and the status "Drilling ledger lines for <accounts>
     (about 60 s) …".
2. **Async:** `drillstep.drill` d1 then d2, with `timing.say` notes ("Drilled 1 834 lines for 2026 YTD (58 s); now
   last year …").
3. **Job B:**
   - `drillstep.write` d1/d2 as `Q<q> Lines d1/d2`;
   - `_expand_recipe` with `ctx.lines`;
   - `pipeline.run(..., q=q, frame_done=True)`.
   - If d1 failed: status "Drill failed: <reason>; cube totals only". A cube-only draft is rendered with a
     warning, so the sheet is never left as a bare frame.

Non-drill recipes: one job, with `progress.show` called before `_expand_recipe`.

- [ ] **Step 1: failing tests**
  - `test_drill_recipe_draws_the_frame_before_any_drill` (the fake host records the order: `show:Q1 Answer`
    before `drill:d1`);
  - `test_failed_drill_leaves_a_cube_only_answer`;
  - `test_frame_before_recipe_expansion_for_non_drill_recipes`;
  - `test_consultant_question_end_to_end_offline`:
    - `gate(T11)` → `answer(spec with gate's recipe)`, with a fake drill and a fake cube;
    - the sheets are Q1 Answer, Q1 By project, Q1 Lines d1/d2;
    - the checks pass and `key_values["Total (cube)"]` equals the fixture;
    - the group block holds TSS / Telenor ASA / Towers;
    - the reconciling line equals the sum of the empty-supplier rows;
    - there are ≤ 3 follow-ups.
- [ ] **Steps 2–4:** FAIL → implement → PASS; then the full suite + ruff.
- [ ] **Step 5:** commit `answer: first frame before drills; lines_ranking end to end in one sv_answer call`.
- [ ] **Step 6 (LIVE, owner):**
  1. `sv_gate(<T11 question>)` → `clear` with `recipe.name == "lines_ranking"`.
  2. `sv_answer(spec={version:1, question, gate, recipe}, target="new")`.
  3. Record:
     - the first-frame time and the total time;
     - the line counts, and the unexplained gap (it should be 0; otherwise a gap line + a Warning finding, D20);
     - whether the `Yr 25` drill works (R1);
     - the group-company supplier strings, and whether a counterparty column exists. Edit `[group_companies]` to
       match; that is a data-only commit.
  4. Compare with the known facts:
     - TSS −121.4 MNOK, Telenor ASA −15.2, Towers −2.9;
     - spreadsheet journals +25.4 MNOK without supplier;
     - Oct 26 +12.7 MNOK accrual reversals.
  5. Any defect: add a failing offline test first, then fix.

---

## Batch C — docs, bench, live

### Task C1: skill, gate text, references, server instructions, design spec, benchmarks

**Files:** `skills/smartass/SKILL.md`, `skills/smartass/reference/{operations,recipes,drill}.md`,
`knowledge/advisor/gate.md`, `engine/smartview_mcp/server.py` (`INSTRUCTIONS`),
`docs/specs/2026-11-02-pipeline-architecture-design.md` (Implemented notes: §3.1 the gate names the recipe +
`answers`; §3.3 enforcement, first frame, English, read-back; §3.4 drill inside `lines_ranking`; §5 account-list
measures and grains), `docs/specs/2026-11-02-benchmarks.md` (B16, T11, T12); tests in `test_skill_docs.py`.

- **SKILL.md:**
  - The flow: `sv_gate`; if it is ambiguous, ask in the user's language, then call `sv_gate(..., answers=)`; then
    `sv_answer` with the gate's recipe.
  - Always build sheets, and never ask whether to. Several sheets are fine.
  - A departure only on the user's explicit request, via `deviation=`.
  - Quote `key_values` and never compute numbers yourself.
  - The chat is in the user's language, with a full summary and the ≤ 3 `followups`. The sheets are English.
  - Line-level questions use `lines_ranking`, not `sv_drill` + your own sums.
- **operations.md:** the new errors and their fixes:
  - "still ambiguous";
  - "departs from the standard analysis";
  - "Write sheet text in English";
  - "drill table(s) … do not exist" (the q mismatch explained).
- **recipes.md:** `lines_ranking`, the gate's recipe table and each recipe's defaults.
- **drill.md:** `sv_drill` is a workbench utility.
- [ ] **Step 1: failing tests**
  - `test_skill_says_always_build_sheets_and_never_asks`
  - `test_skill_says_quote_key_values`
  - `test_skill_and_gate_say_chat_in_user_language_sheets_english`
  - `test_instructions_lead_with_gate_then_answer` (kept/updated)
- [ ] **Steps 2–4:** FAIL → edit → PASS.
- [ ] **Step 5:** commit `docs: gate names the recipe, always build sheets, quote read-back, chat in user's language; B16/T11/T12`.

### Task C2: bench B16 + T11/T12; verification; line count

**Files:** `engine/tests/live/bench.py`; pure-helper test in `engine/tests/test_live_helpers.py`.

- T11 and T12 join the trap loop.
- B16 = the Norwegian question. The run calls `sv_gate`, then `sv_answer` with the gate's recipe on
  `target="new"`.
  - Oracle = the FinLite cube, YTD C60080 + C60090, un.SY.
  - Problems are flagged when:
    - the sheets Answer or By project are missing;
    - a check fails;
    - "Total (cube)" is missing from `key_values`;
    - `first_visible_s > 2.5`.
- `_core` (the existing cases) passes `deviation="bench pins recipe <name>"` only when the gate names another
  recipe. The notes record this.
- [ ] **Step 1:** failing test `test_b16_problems_flag_missing_sheets_and_slow_frame` → implement → PASS.
- [ ] **Step 2:** the full suite, `uv run ruff check smartview_mcp tests`, and the line count (expected ≤ 17 450).
  Check that every touched module is ≤ ~400 lines.
- [ ] **Step 3:** commit `bench: B16 consultant ranking, T11-T12; verification + line count`.

### Task C3: **LIVE** — the consultant re-run and the bench (owner)

- [ ] `uv run python tests\live\bench.py --route core --only B16,T09,T11,T12 --out advisor-plans`
- [ ] `uv run python tests\live\bench.py --route core --out advisor-plans-full`. Earlier PASS results must stay
  PASS.
- [ ] Ask the original question in a real chat. Expect:
  - a reply in Norwegian;
  - no "do you want a sheet";
  - `Q<n> Answer` + `Q<n> By project` (+ Lines);
  - the numbers quoted from `key_values`;
  - ≤ 3 follow-ups;
  - English sheet text.
- [ ] Commit `docs/audit/2026-11-02-bench-advisor-plans.{md,json}`: `docs(audit): live bench advisor plans (B16, T11-T12)`.

---

## Self-review: scope → tasks

| Scope item | Task(s) |
|---|---|
| (1) the gate names the recipe + args | A3 |
| (2) finance defaults in recipe code; free-form fallback | A5, B2 (D3, D16) |
| (3) enforcement: ambiguity, first frame, English | A4, B4 |
| (4) lines/ranking recipe, reconciling line, initiative, topic sheet + links | B1–B4 |
| (5) where-cell 0x800A03EC fix | A1 |
| (6) cost-type measures | A2 |
| (7) read-back + ≤ 3 follow-ups | A5 |
| (8) skill/docs | C1 |
| (9) live re-run + bench case | B4 step 6, C2, C3 |
| deviations written on the sheet | A4 |
| vendored advisor text unchanged | A5 (sha256 test) |
| owner decisions 2026-11-02: gap line + finding, combined legal, top 10 everywhere | B2 (D20), A2 (D21), A5 (D22) |

**12 tasks:** A 5, B 4, C 3. LIVE steps: B4 step 6 and C3. Everything else is offline.

## Risks and open questions

- **R1 — LY drill.** It is unverified whether FinLite drills the previous year for these accounts. The fallback is
  LY at total level (D17). This is measured in B4.
- **R2 — gate dodging.** The re-gate in `answer()` uses `spec.question`, so the agent could reword the question.
  `answers=` removes the reason to do so, and the runs log question + answers. The engine never sees the user's
  message.
- **R3 — group-company names.** The patterns are a first guess, fixed in B4 step 6.
- **R4 — B06 sheet changes.** The `cost_base` top 10 (D22) changes B06's sheet. The bench oracle is on the total,
  so it should still pass. Older answers re-drawn by a follow-up keep their stored spec.

Settled by the owner on 2026-11-02 (no longer open):
- first frame ≤ 2 s after the workbook is open (D6);
- no auto re-drill; the gap is a reconciling line plus a finding (D20);
- `consultants` = C60080 + C60090, with C693 kept as `consultants_and_temps` (A2);
- a combined `legal` measure plus fees and disputes (D21);
- top 10 + Other everywhere (D22).
