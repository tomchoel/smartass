# Phase 4 — semantic layer + advisor gate: implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:executing-plans (inline) or subagent-driven-development.
> Steps use checkbox (`- [ ]`) syntax. Run everything from `C:\dev\smartass\engine` with `uv run`.

**Goal:** every question passes a deterministic, testable scope gate (design §3.1) before anything is pulled; material
variances on an answer must carry a labelled driver (design §3.4).

**Architecture:** one semantic layer (`knowledge/catalog.toml` + new `knowledge/capabilities.toml`, loaded by
`core/semantic.py`), a period parser (`core/period.py`), per-user learned defaults (`core/prefs.py`) and the gate
(`core/scope.py`). The engine *suggests* the outcome with the fixed checklist; the agent makes the final call and
writes the answer spec. New tools `sv_gate` and `sv_remember`; `check` gains material-variance findings
(`core/commentary.py`). Advisor text vendored from Anthropic financial-services (Apache-2.0) in `knowledge/advisor/`.

**Tech stack:** Python 3.12, tomllib, pydantic v2, pytest (+anyio), FastMCP (`mcp.server.mcpserver`).

**Owner decisions (2026-11-02, round 5):**
| # | Decision |
|---|----------|
| Ask policy | **Balanced:** ask only when a gap changes the number AND no learned default or standard convention covers it; otherwise proceed and state the assumption |
| Learned defaults | Per-user `preferences.toml` (in `config.SVKIT`); an entry can later be promoted to the team catalog by hand |
| Language | Clarifying questions and commentary in **English** (sheets already English) |
| Commentary | Engine flags material variances (catalog thresholds + always-comment list); agent writes drivers (data-backed with evidence, or labelled hypothesis); an uncommented material line is a Warning finding |

**Verified cube facts (atlas 2026-10-13, live sv_members 2026-11-02):**
- No customer dimension in any cube.
- Products: FinMax `All Products` (19 755 members), FinPlan `All Products` (243).
  - Opex is not allocated to products; the full cost per product exists only in Produktregnskapet (annual).
- Currency: FinMax `All Currencies` has NOK, Local and FX Rates. There is no EUR reporting member.
- FinPlan Entity has 337 members (business areas and some units). Units below them have no FinPlan member (B13).
- FinMax scenarios:
  - Actual, Target, ForecastQ1–Q4
  - Act v Tgt
  - Act Dev Last Yr / Last Per
  - Act v FCQ1–Q4
- FinPlan scenarios: Actual, Target, Forecast Qnyy and Forecast Running.
  - Version: Final / Working.
  - Budget = Target Final.
- Consultants: `ac.C693` Consultancy and Temporary Contract workers (business map).
  - Includes temporary workers, so it must be stated.

---

## File structure

| File | Responsibility |
|------|----------------|
| `knowledge/capabilities.toml` (new) | measure synonyms (EN + NO) → catalog concept/KPI/account; grains (product, customer, cost centre, month, unit) with feasible measures + proxy; currencies |
| `engine/smartview_mcp/core/semantic.py` (new) | load + validate capabilities against the catalog; `measures(text)`, `grains(text)`, `feasible(measure, grain)`, `currency(text)` |
| `engine/smartview_mcp/core/period.py` (new) | `parse(text, closed)` → `Period` (kind, label, gap, open-month flag) |
| `engine/smartview_mcp/core/prefs.py` (new) | per-user learned defaults: `load()`, `lookup(item, phrase)`, `remember(item, phrase, value)` |
| `engine/smartview_mcp/core/scope.py` (new) | the gate: `gate(question, closed, prefs)` → checklist, suggested outcome, questions, proxy, warnings |
| `engine/smartview_mcp/core/commentary.py` (new) | `material(prepared, values, cat)` → flagged lines; `findings(flagged, drivers)`; driver lines for the sheet |
| `engine/smartview_mcp/core/contract.py` | `Driver.line` (which answer line it explains), validated against layout keys |
| `engine/smartview_mcp/core/pipeline.py` | render drivers as key points; pass commentary findings to `check.review` |
| `engine/smartview_mcp/core/check.py` | accept extra findings |
| `engine/smartview_mcp/tools_core.py` | `sv_gate`, `sv_remember` |
| `knowledge/advisor/` (new) | `NOTICE.md` (Apache-2.0 attribution), `gate.md` (procedure the agent follows), `variance-commentary.md`, `findings.md` |
| `engine/tests/test_advisor.py` (new) | unit tests incl. the gate benchmark table (B09–B14, T01–T10) |
| `engine/tests/live/bench.py` | route core: gate cases scored by `sv_gate` outcome; false-go / needless-ask counts |

---

### Task 1: capability map + semantic layer

**Files:** create `knowledge/capabilities.toml`, `engine/smartview_mcp/core/semantic.py`; test `tests/test_advisor.py`.

- [ ] **Step 1: failing tests**

```python
from smartview_mcp.core import semantic

def test_measure_synonyms_resolve_english_and_norwegian():
    assert [m.key for m in semantic.measures("EBITDA for Business")] == ["ebitda"]
    assert [m.key for m in semantic.measures("driftskostnader i IT")] == ["opex"]
    assert [m.key for m in semantic.measures("Cost per FTE in IT")] == ["opex_per_fte"]   # longest match wins
    vague = semantic.measures("How are we doing on costs?")
    assert [m.key for m in vague] == ["costs"] and vague[0].vague

def test_grain_feasibility_and_proxy():
    g = semantic.grains("EBITDA per product")[0]
    ok, proxy = semantic.feasible("ebitda", g)
    assert g.key == "product" and not ok and "contribution margin" in proxy
    assert semantic.feasible("revenue", g) == (True, None)
    assert not semantic.feasible("ebitda_margin", semantic.grains("per customer")[0])[0]

def test_currency_only_nok():
    assert semantic.currency("Revenue in EUR for Technology") == ("EUR", False)
    assert semantic.currency("Revenue for Technology") == ("NOK", True)

def test_capabilities_reference_only_known_concepts():
    assert semantic.problems() == []
```

- [ ] **Step 2:** run `uv run pytest tests/test_advisor.py -q` → FAIL (no module).
- [ ] **Step 3: implement.**

`capabilities.toml` shape:

```toml
# Capability map (design §5): what can be answered at which grain. Hand-maintained; semantic.problems() checks every
# concept/kpi reference against catalog.toml. Synonyms are matched as whole words, case-insensitive, longest first.
[[measure]]
key = "ebitda"; concept = "ebitda"; synonyms = ["ebitda", "driftsresultat før avskrivninger"]
[[measure]]
key = "opex_per_fte"; kpi = "opex_per_fte"; synonyms = ["cost per fte", "opex per fte", "kostnad per årsverk"]
[[measure]]
key = "consultants"; account = "ac.C693"; label = "Consultancy and temporary contract workers (ac.C693)"
synonyms = ["consultants", "consultancy", "konsulenter", "konsulent", "innleie"]
note = "C693 includes temporary contract workers: state it"
[[measure]]
key = "costs"; vague = true; synonyms = ["costs", "cost", "kostnader", "spending"]
options = ["opex (C70)", "opex + COGS", "personnel cost (C50S)"]
[[measure]]
key = "numbers"; vague = true; synonyms = ["numbers", "tall", "results", "resultater", "how are we doing"]
options = ["P&L to EBITDA", "revenue and EBITDA", "opex"]
# ... revenue, cogs, gross_margin, personnel, other_opex, opex, ebitda_boi, other_items, da, fte, ebitda_margin,
#     gross_margin_pct, personnel_per_fte, profit (synonyms "profit", "resultat", "lønnsomhet": concept ebitda,
#     note = "profit read as EBITDA (the cubes' bottom line for plan comparisons)")

[[grain]]
key = "product"; synonyms = ["per product", "by product", "per produkt", "product level"]
feasible = ["revenue", "cogs", "gross_margin"]
proxy = "contribution margin (revenue - COGS) per product; opex is not allocated to products in FinMax (full cost per product only in Produktregnskapet, annual)"
[[grain]]
key = "customer"; synonyms = ["per customer", "by customer", "per kunde"]; feasible = []
proxy = "no customer dimension in FinMax/FinPlan; answer per business area / segment"
[[grain]]
key = "cost_centre"; synonyms = ["per cost centre", "per cost center", "per kostnadssted", "by cost centre"]
feasible = ["*"]; note = "FinMax only; FinPlan has business areas, not cost centres"
[[grain]]
key = "month"; synonyms = ["per month", "monthly", "per måned", "månedlig", "trend"]; feasible = ["*"]

[currency]
reporting = ["NOK"]
known = ["NOK", "EUR", "USD", "SEK", "DKK"]
proxy = "values in NOK (FinMax reports NOK and Local only; no EUR translation)"
```

`semantic.py`:

```python
"""One semantic layer (design §5): catalog.toml (what lines/KPIs/scenarios mean) + capabilities.toml (synonyms,
grains, currency) -> measures/grains/feasibility for the scope gate."""
from __future__ import annotations
import functools, re, tomllib
from dataclasses import dataclass, field
from .. import catalog, config

@dataclass(frozen=True)
class Measure:
    key: str; synonyms: tuple[str, ...]; concept: str | None = None; kpi: str | None = None
    account: str | None = None; label: str | None = None; vague: bool = False
    options: tuple[str, ...] = (); note: str | None = None

@dataclass(frozen=True)
class Grain:
    key: str; synonyms: tuple[str, ...]; feasible: tuple[str, ...]; proxy: str | None = None; note: str | None = None

@functools.cache
def _caps() -> dict: return tomllib.loads(config.knowledge_file("capabilities.toml").read_text(encoding="utf-8"))

def _rx(words): return re.compile(r"(?<![\w])(" + "|".join(re.escape(w) for w in sorted(words, key=len, reverse=True)) + r")(?![\w])", re.I)

def _find(items, text):   # longest-match-first, non-overlapping, in order of appearance
    ...
def measures(text) -> list[Measure]: ...
def grains(text) -> list[Grain]: ...
def feasible(measure_key, grain) -> tuple[bool, str | None]:
    ok = "*" in grain.feasible or measure_key in grain.feasible
    return (True, None) if ok else (False, grain.proxy)
def currency(text) -> tuple[str, bool]: ...
def problems() -> list[str]:  # every concept/kpi reference exists in the catalog; synonyms unique across measures
    ...
```

- [ ] **Step 4:** tests pass. **Step 5:** commit `core.semantic + capabilities.toml: measures, grains, currency (design §5)`.

### Task 2: period parser

**Files:** create `engine/smartview_mcp/core/period.py`; tests in `tests/test_advisor.py`.

- [ ] **Step 1: failing tests** (closed month `"Sep 26"`):

```python
from smartview_mcp.core import period

@pytest.mark.parametrize("text,kind,label,gap", [
    ("Business EBITDA Sep 26 vs Sep 25", "month", "Sep 26", None),
    ("EBITDA YTD September 2026", "ytd", "YTD Sep 26", None),
    ("Q3 numbers for Consumer", "quarter", "Q3 26", None),
    ("full-year EBITDA latest forecast", "fy", "FY 26", None),
    ("Opex for Wholesale this year", "this_year", None, "YTD to Sep 26, or full year (actual + forecast)?"),
    ("Cost per FTE in IT", "none", "YTD Sep 26", None),           # convention: YTD to the last closed month
    ("Forecast for October 26 actuals", "month", "Oct 26", None),
])
def test_period_parse(text, kind, label, gap):
    p = period.parse(text, "Sep 26")
    assert (p.kind, p.label, p.gap) == (kind, label, gap)

def test_open_month_is_flagged():
    assert period.parse("Oct 26 actuals", "Sep 26").open and not period.parse("Sep 26", "Sep 26").open
```

- [ ] **Step 2:** FAIL. **Step 3: implement** `Period(kind, label, gap, open, convention)`:
  - Month names: EN and NO short and long forms (`jan`, `januar`, `mai`, `okt`, `des` …). Year as `26`, `'26` or `2026`.
  - `ytd` / `year to date` / `hittil i år` / `so far this year` → `ytd` at the named month, or else at the closed month.
  - `q1`–`q4`, with an optional year → `quarter`. The year defaults to the closed month's year.
  - `full year` / `full-year` / `fy26` / `helår` / `årsprognose` → `fy`.
  - `this year` / `i år` alone, with no ytd/fy word → `this_year`, with the gap text from the test.
  - Nothing found → `none`. The label is the YTD convention, and `convention=True` so the gate states it.
  - `open` = the parsed month (or the quarter's last month) is after the closed month.
- [ ] **Step 4:** pass. **Step 5:** commit `core.period: EN/NO period phrases, closed-month conventions`.

### Task 3: learned defaults

**Files:** create `engine/smartview_mcp/core/prefs.py`; add `PREFERENCES = SVKIT / "preferences.toml"` to `config.py`.

- [ ] **Step 1: failing tests**

```python
def test_remember_then_lookup_roundtrip(tmp_path, monkeypatch):
    from smartview_mcp.core import prefs
    monkeypatch.setattr(prefs.config, "PREFERENCES", tmp_path / "p.toml")
    assert prefs.lookup("measure", "costs") is None
    prefs.remember("measure", "Costs", "opex")
    assert prefs.lookup("measure", "costs") == "opex"
    prefs.remember("measure", "costs", "personnel")          # replaces, never duplicates
    assert prefs.lookup("measure", "COSTS") == "personnel"
    assert (tmp_path / "p.toml").read_text(encoding="utf-8").count("[[default]]") == 1

def test_remember_rejects_unknown_item_or_value(tmp_path, monkeypatch):
    from smartview_mcp.core import prefs
    monkeypatch.setattr(prefs.config, "PREFERENCES", tmp_path / "p.toml")
    with pytest.raises(InputError): prefs.remember("colour", "x", "y")
    with pytest.raises(InputError): prefs.remember("measure", "costs", "not_a_measure")
    with pytest.raises(InputError): prefs.remember("period", "this year", "sometimes")
```

- [ ] **Step 3: implement.** Write the file as readable TOML by hand: `[[default]]` with item, phrase, value and saved (ISO date). The allowed values per item are:
  - `measure`: a non-vague measure key from semantic.
  - `period`: `ytd` or `fy`.
  - `scenario`: a catalog scenario key.
  - `entity`: a catalog entity key.
- [ ] **Step 5:** commit `core.prefs: per-user learned defaults (preferences.toml)`.

### Task 4: the gate

**Files:** create `engine/smartview_mcp/core/scope.py`; tests in `tests/test_advisor.py`.

`gate(question, closed, learned=prefs.lookup) -> dict` with keys:
- `question` and `suggested_outcome` (`clear`, `ambiguous` or `infeasible`).
- `checklist`: one entry per CHECKLIST item, each `{value, source: question|convention|learned|catalog, gap, ask}`.
- `questions` (at most 3), `reason`, `proxy`, `warnings`.
- `entities` (key, label, finmax, finplan) and `measures` (key, label/concept/kpi/account, note).
- `concepts`: top concept ids from `concepts.find`.
- `next`: instruction text for the agent.

Rules, applied in order:
1. **Entity**
   - Use catalog `resolve_entities`, or a literal `un.SY…` member token.
   - If nothing is named, use the convention "Telenor Norge (un.SY)".
   - An abolished entity adds a warning plus its successors.
   - A literal sub-unit with no FinPlan member adds the warning "no FinPlan member: plan/budget comparison not possible, last year only".
2. **Measure**
   - Use `semantic.measures`. If none matches, fall back to concepts (`concepts.find` score ≥ 0.75) as source `catalog`.
   - If there is still none, or the measure is vague, it is a gap with `ask=True`, unless a learned default exists for the matched phrase.
   - The question text lists the options.
3. **Period**
   - Use `period.parse`.
   - `this_year` is a gap with `ask=True` unless learned (`period`, `this year`).
   - `none` resolves by convention to YTD of the closed month.
4. **Scenario**
   - Words map to catalog scenarios: budget/budsjett/target/mål → `budget`; forecast/prognose → `forecast`; actual/faktisk/regnskap → `actual`; `last year`/`i fjor`/`vs <same month LY>` → actual LY comparison.
   - With none, the convention is `actual`.
   - The phrase "budget" alone is the convention Target Final (catalog), stated, not asked.
5. **Currency/sign**
   - Use `semantic.currency`. A non-NOK currency makes the outcome infeasible with the currency proxy.
   - Otherwise "NOK; FinMax signs (costs negative)" by convention.
6. **Granularity**
   - For each grain × measure, call `semantic.feasible`. Any False → infeasible with the grain's proxy.
   - An open month with scenario `actual` → infeasible: "actuals for {month} do not exist yet (last closed month {closed})", with proxy "latest forecast for {month}".
7. **Outcome**
   - infeasible if any of the above makes it infeasible.
   - Otherwise ambiguous if any item has `ask=True`, with one question per such item (at most 3, in checklist order).
   - Otherwise clear.

- [ ] **Step 1: failing tests** — the gate benchmark table (expected = `docs/specs/2026-11-02-benchmarks.md`):

```python
GATE_CASES = [
    ("B09", "What do we earn on mobile subscriptions?", "clear"),
    ("B10", "EBITDA per product for Mobile Consumer", "infeasible"),
    ("B11", "How are we doing on costs?", "ambiguous"),
    ("B13", "EBITDA YTD for unit un.SY.5OK_W", "clear"),
    ("B14", "Business full-year EBITDA: latest forecast vs budget", "clear"),
    ("T01", "What's our budget for Business?", "ambiguous"),
    ("T02", "Opex for Wholesale this year", "ambiguous"),
    ("T03", "EBITDA margin per customer", "infeasible"),
    ("T04", "Revenue in EUR for Technology", "infeasible"),
    ("T05", "Profit per product", "infeasible"),
    ("T06", "Business EBITDA Sep 26 vs Sep 25", "clear"),
    ("T07", "Q3 numbers for Consumer", "ambiguous"),
    ("T08", "Cost per FTE in IT", "clear"),
    ("T09", "How much did we spend on consultants?", "clear"),
    ("T10", "Forecast for October 26 actuals", "infeasible"),
]

@pytest.mark.parametrize("cid,q,want", GATE_CASES, ids=[c[0] for c in GATE_CASES])
def test_gate_outcome(cid, q, want, no_prefs):
    assert scope.gate(q, "Sep 26")["suggested_outcome"] == want

def test_gate_learned_default_turns_an_ask_into_a_stated_assumption(no_prefs):
    no_prefs["measure", "costs"] = "opex"; no_prefs["period", "this year"] = "ytd"
    g = scope.gate("How are we doing on costs this year?", "Sep 26", learned=lambda i, p: no_prefs.get((i, p)))
    assert g["suggested_outcome"] == "clear" and g["checklist"]["measure"]["source"] == "learned"

def test_gate_details():
    g = scope.gate("Opex for Wholesale this year", "Sep 26", learned=lambda i, p: None)
    assert len(g["questions"]) == 1 and "full year" in g["questions"][0]
    assert scope.gate("EBITDA YTD for unit un.SY.5OK_W", "Sep 26", learned=lambda i, p: None)["warnings"]
    t10 = scope.gate("Forecast for October 26 actuals", "Sep 26", learned=lambda i, p: None)
    assert "Oct 26" in t10["reason"] and "forecast" in t10["proxy"]
    t08 = scope.gate("Cost per FTE in IT", "Sep 26", learned=lambda i, p: None)
    assert t08["checklist"]["measure"]["value"].startswith("Opex per FTE")     # the definition is stated
```

(`no_prefs` fixture: a dict, with `learned=lambda i, p: d.get((i, p))` passed by the parametrized test.)

- [ ] **Steps 2–4:** FAIL → implement → PASS. The table is the advisor benchmark: if a case fails, fix the rule or
  the capability map, never the expected outcome (that is an owner decision in the benchmarks doc).
- [ ] **Step 5:** commit `core.scope: the advisor gate (fixed checklist, balanced ask policy); B09-B14 + T01-T10 as tests`.

### Task 5: tools `sv_gate` and `sv_remember`

**Files:** modify `engine/smartview_mcp/tools_core.py`; smoke test tool count (`tests/test_server_smoke.py` or wherever
`33` tools is asserted) → 35.

- `sv_gate(question: str) -> dict`, annotated `LOCAL`.
  - Gets `closed` via `svhttp.shared().closed_month("FinMax")` inside `_call` with a 30 s budget. If that fails it uses the cached closed month, or the previous month of today, and adds a warning.
  - Then calls `scope.gate`.
  - Docstring: call it FIRST for every finance question. Follow `knowledge/advisor/gate.md`. Ambiguous → ask the questions in ONE message, in English. Clear or infeasible → write the answer spec with the gate, checklist and assumptions copied from the result.
- `sv_remember(item: Literal["measure","period","scenario","entity"], phrase: str, value: str) -> dict`, annotated `LOCAL`.
  - Docstring: only after the user agreed to save their answer to a clarifying question.
  - Returns the saved entry and the file path.
- [ ] Tests:
  - `sv_gate` with `closed_month` monkeypatched returns `suggested_outcome`.
  - `sv_remember` round-trip with a tmp `PREFERENCES`.
  - The smoke test count.
- [ ] Commit `tools: sv_gate + sv_remember`.

### Task 6: commentary — material variances need a driver

**Files:**
- Create `engine/smartview_mcp/core/commentary.py`.
- Modify `contract.py` (`Driver.line`), `pipeline.py` and `check.py`.

- [ ] **Step 1: failing tests**

```python
from smartview_mcp import catalog
from smartview_mcp.core import commentary

PREP = {"columns": [{"header": "Sep 26"}, {"header": "Sep 25"}, {"header": "Δ"}],
        "sections": [{"lines": [
            {"key": "rev", "label": "Revenue", "cells": [{"pull": "m"}, {"pull": "m"}, {"var": [0, 1]}]},
            {"key": "small", "label": "Small", "cells": [{"pull": "m"}, {"pull": "m"}, {"var": [0, 1]}]},
            {"key": "big", "label": "Big cost", "cells": [{"pull": "m"}, {"pull": "m"}, {"var": [0, 1]}]}]}]}
VALUES = {"rev": [100.0, 99.0, 1.0], "small": [1_000.0, 2_000.0, -1_000.0], "big": [-20e6, -10e6, -10e6]}

def test_material_uses_catalog_thresholds_and_always_list():
    flagged = commentary.material(PREP, VALUES, catalog.load())
    assert [f["line"] for f in flagged] == ["rev", "big"]     # rev: always list; big: >= NOK 5m; small: under the 500k floor
    assert flagged[1]["delta"] == -10e6 and flagged[1]["pct"] == pytest.approx(-1.0)

def test_material_line_without_driver_is_a_warning():
    flagged = commentary.material(PREP, VALUES, catalog.load())
    f = commentary.findings(flagged, [{"line": "big", "text": "x", "basis": "hypothesis"}])
    assert [x["severity"] for x in f] == ["Warning"] and "Revenue" in f[0]["text"]

def test_driver_lines_label_hypotheses():
    out = commentary.sheet_lines([{"line": "big", "text": "More licences", "basis": "data", "evidence": "Q1 FinMax C60"},
                                  {"line": "rev", "text": "Price rise", "basis": "hypothesis"}])
    assert out == ["Driver: More licences (data: Q1 FinMax C60)", "Hypothesis — not verified in data: Price rise"]

def test_contract_driver_line_must_exist():
    ...   # a v1 spec whose commentary.drivers[0].line names no layout line -> ValidationError
```

- [ ] **Step 3: implement.**
  - `material()`:
    - Walk the lines of the prepared spec. For the first `var` cell `[a, b]`, take `cur = values[key][a]` and `cmp = values[key][b]`, both numeric; `delta = cur - cmp`; `pct = delta / abs(cmp)` if `cmp` else None.
    - A line is material by `Materiality`:
      - `pct_hit = pct is not None and abs(pct) >= m.pct and abs(delta) >= m.pct_floor_nok`
      - `abs_hit = abs(delta) >= m.abs_nok`
      - `rule == "either"` → `pct_hit or abs_hit`; otherwise both.
    - The always list matches the line key, or the line's `concept`, against `m.always` and the concepts of those keys (FinMax lines).
    - Skip lines with no var cell, lines whose values are not numeric, and stock lines (`fte`, or `measure == stock`) from the abs-NOK test.
  - `findings(flagged, drivers)`:
    - Warning "material variance on '{label}' ({delta:+,.0f}, {pct:+.0%}) has no driver — write one (data-backed with evidence, or labelled hypothesis) or 'driver unclear — flag for controller'".
    - Info "always-commented line '{label}' has no driver" for always-list lines under the threshold.
  - `pipeline.run`:
    - `answer_layout` appends `commentary.sheet_lines(spec.commentary drivers)` to `key_points`.
    - After `answer.build`, call `commentary.material(prepared, res["values"], catalog.load())`.
    - Pass it to `check.review(res, sources, extra=commentary.findings(...))`.
    - Return `flagged` in the `sv_answer` output as `material`.
  - `contract.Driver`:
    - Add `line: str | None = None`.
    - `AnswerSpec._shape` errors if a driver line is not a layout line key.
- [ ] Commit `core.commentary: material variances flagged, drivers rendered and enforced (design §3.4)`.

### Task 7: vendored advisor text

**Files:** create `knowledge/advisor/NOTICE.md`, `gate.md`, `variance-commentary.md`, `findings.md`.
- `NOTICE.md`:
  - Says that `variance-commentary.md` and `findings.md` are adapted from anthropics/financial-services-plugins (`month-end-closer/skills/variance-commentary`, `financial-analysis/skills/audit-xls`).
  - Apache License 2.0, copyright Anthropic.
  - Lists the changes: Telenor materiality from catalog.toml, a NOK 500k floor, always-comment revenue/personnel/EBITDA, the hypothesis label, and the restriction to answer-sheet checks.
  - Includes a link to the licence text.
- `gate.md`:
  - The agent procedure: the `sv_gate` call, the three outcomes, the balanced policy, the English-only rule.
  - Offer `sv_remember` after a clarifying answer: "Save 'costs = opex' as your default?"
  - Never pull on ambiguous.
  - For infeasible: say why, show what is possible, and build the proxy with a gate reason and proxy.
- `variance-commentary.md`: threshold (catalog), the always list, driver = *why*, not *what*; data first (drill); otherwise "Hypothesis — not verified in data"; or "driver unclear — flag for controller"; a 3–5 sentence summary.
- `findings.md`: Critical / Warning / Info definitions as used by `core/check.py`.
- `sv_gate` and `sv_answer` docstrings name these files.
- [ ] Commit `knowledge/advisor: gate procedure + vendored commentary/findings (Apache-2.0)`.

### Task 8: benchmark gate cases

**Files:** modify `engine/tests/live/bench.py`.
- With route `core`, cases with an expected gate outcome (`Case.gate`, already in the benchmarks table: B10, B11, T01–T10) call `sv_gate`.
  - The result is PASS when `suggested_outcome == expected`.
  - The report adds a line: `gate: N/M right; false-go K (asked→should ask, answered instead); needless-ask L`.
    - false-go = expected ambiguous/infeasible but the result was clear.
    - needless-ask = expected clear but the result was ambiguous.
  - Baseline keeps `AGENT`.
- The `T01`–`T10` cases must exist in `CASES`. Add them (`kind="trap"`, `agent=True`, `gate=...`) if bench.py lacks them.
- [ ] Unit test: `_bench().gate_score([...])` counts false-go / needless-ask correctly.
- [ ] Live: `uv run python tests\live\bench.py --route core --only B10,B11,T01,...,T10 --out bench-core-gate` → all PASS.
- [ ] Commit `bench: gate cases scored on route core (false-go / needless-ask)`.

### Task 9: docs + push

- Design §3.1:
  - Implemented in `core/scope.py` + `sv_gate`; the balanced policy; English.
  - Learned defaults are in `%SVKIT%\preferences.toml` via `sv_remember`.
- §5: point to `capabilities.toml` and the verified cube facts above.
- §3.4: the commentary rules.
- `docs/operations.md`: a row for `preferences.toml`.
- Skill (`skills/smartass/SKILL.md`): a short "new pipeline: sv_gate → sv_answer" note. The full rewrite is phase 7.
- Full suite + ruff, then push `rework/pipeline`.
- Mark the SQL todo `semantic-advisor` done.

## Self-review
- Spec coverage:

  | Spec section | Covered by |
  |---|---|
  | §3.1 checklist, outcomes, learned defaults | Tasks 2–5 |
  | §3.4 materiality, always-comment, labelled hypotheses, vendored structure | Tasks 6–7 |
  | §5 semantic layer | Task 1 |
  | Recommendation 7, measure the advisor | Tasks 4 and 8 |

- IFRS 8 10 % per-segment rule:
  - The catalog `pct = 0.10` already applies to every line.
  - The per-segment case adds nothing new, so it is not a separate rule.
- Types:
  - `Measure.key` is used by prefs and scope.
  - `semantic.feasible(measure_key, grain)` is used in Task 4.
  - `commentary.material/findings/sheet_lines` are used in `pipeline`.
