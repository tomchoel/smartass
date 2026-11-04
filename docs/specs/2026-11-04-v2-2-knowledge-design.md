# smartass v2 · sub-spec 2: knowledge

Status: draft for review (2026-11-04). Parent: [overview](2026-11-04-v2-overview.md) (D4, D5, D13–D15).
Covers what the finance agent knows, where it lives, how it reaches the agent, how it stays fresh every month,
and how the plan → build → critic loop uses it.

## 1. Principles

1. **Knowledge by type**, each with its own store and retrieval (table below).
2. **Executable, not prose.** Definitions the engine enforces live in typed files, not in skill text.
3. **Just-in-time.** Caveats arrive in the tool reply that touches the member or metric; the agent does not
   have to remember them.
4. **Time-valid.** Every fact and mapping has `valid_from` / `valid_to`.
5. **Client-neutral.** All knowledge is served by the MCP server; the Claude skill and the Copilot instructions
   are thin adapters generated from one source.
6. **Eval-gated.** A knowledge change runs the test set; docs are generated from the registries.
7. **No manual entry month to month.** People approve new or changed items once; nothing else is typed in.
8. **No vector store.** Knowledge is structured and must be exact: exact + fuzzy lookup over aliases (SQLite FTS5).

| Type | Examples | Source of truth | Reaches the agent via |
|---|---|---|---|
| Structure | Dimension trees, members, aliases | EPM (atlas build) | `fin_lookup`, validation |
| Semantics | Metrics, FinMax↔FinPlan mapping, signs, P&L chain, materiality, sensitivity | `knowledge/semantic/metrics/*.toml` | `fin_lookup`, resolve, guards |
| Facts and traps | Reorg 1 Jun 2026, FinPlan groups differently, KPIs read Periodic | `knowledge/semantic/facts/*.toml` | `caveats` in every reply that touches the scope |
| Method | Monthly pack, variance commentary, GL extract, Prodreg input | `knowledge/playbooks/*.md` | MCP prompts + resources |
| Past answers | Earlier questions, plans, key values | `answers.db` (per user, automatic) | `fin_lookup(kind="answers")` |

## 2. Semantic layer

One small TOML file per metric. Example `knowledge/semantic/metrics/opex.toml`:

```toml
name = "Opex"
aliases = ["opex", "driftskostnader", "operating expenses", "C70"]
description = "Total operating expenses: personnel (C50S) + other opex (C60). Negative in FinMax."
unit = "NOK"
sign = "cost"                    # costs negative; variance positive = better
balance = "Periodic"
materiality = { abs = 5_000_000, pct = 0.05 }
status = "approved"              # draft | approved
approved = { by = "tom", on = "2026-11-10" }

[finmax]
members = ["ac.C70"]             # = ac.C50S + ac.C60 (checked by the build)

[finplan]
members = ["OPEX"]
differs = "FinPlan groups some revenue-related costs under COGS; expect a definitional gap vs FinMax."

[[valid]]
from = "2020-01"
```

- **Seeded, not hand-written:** the first set is generated from v1's `catalog.toml`, `capabilities.toml`,
  `concepts.json` (including `splitt` and the 22 `fam_*` families) and the four business maps, as `draft`; Tom
  approves each once.
- **Member attributes** (division per period, product tagging, cost-base flags, sensitivity) are generated from
  the atlas by the v1 `build_*_map.py` logic into tables in `knowledge.db`; they are build output, not files in git.
- **Facts** (`facts/*.toml`): `text`, `scope` (members/metrics it applies to), `valid_from/to`, `severity`
  (`info` / `warn` / `block`), `source`. Example: `2con-abolished.toml`, scope `un.SY.2CON`, from `2026-06`,
  text "Abolished 1 Jun 2026; successors 2SSM + 2EI; history restated".
- **Sensitivity**: metrics or members flagged `sensitive` (salary, FTE below division) make `fin_analyze` warn.

## 3. Build: `knowledge.db`

`smartass knowledge build` compiles atlas + semantic files + generated attribute tables + an FTS5 alias index into
one SQLite file. Checks during the build (a failure stops the build):

- every metric member exists in the atlas for its valid period;
- metric identities hold where declared (`C70 = C50S + C60` against the atlas tree);
- every fact scope resolves;
- aliases are unique across metrics.

`knowledge.db` ships with each release; the TOML files are the reviewable source in git (small, diffable). The
2 MB JSON maps leave git.

## 4. Delivery to the agent

- **`fin_lookup(query, kind?)`** (`kind`: member, metric, fact, answer, any): ranked hits with label, cube, path,
  attributes and **caveats**.
- **Caveats everywhere:** `fin_analyze` and `fin_drill` replies include `caveats` for every member and metric the
  answer touched (facts in scope and valid for the periods used), deduplicated, `block` first.
- **MCP prompts:** `plan` (how to write `plan`), `critic` (rubric and procedure), `monthly-pack`,
  `variance-commentary`, `gl-extract`, `prodreg-input`. **MCP resources:** conventions (signs, formats, cube
  choice), block catalogue, playbooks.
- **Thin adapters:** `skills/smartass/SKILL.md` (Claude) and `copilot-instructions.md` (Copilot) are generated from
  one template, under 2k tokens: the loop (plan → build → critic), when to ask the user, and pointers to the
  prompts. Server `instructions` hold the same short text.

## 5. Finance agent loop

1. **Clarify** only when scope is missing after `fin_lookup` (period, entity, measure): at most 3 questions in one
   message. No regex gate.
2. **Plan**: purpose, hypotheses, audience, depth; blocks chosen from the catalogue; past answers looked up.
3. **Build**: `fin_analyze(spec)`; hard guards may refuse with the fix.
4. **Explain**: read `key_values`, `drivers` evidence and caveats; write `commentary` (basis `data` with the
   evidence cell, or `hypothesis`) by patch.
5. **Critic** (depth `standard` or `deep`, and every pack): a separate pass with the `critic` prompt scores the
   answer 1–5 on the rubric — right question answered, numbers reconciled, comparison basis, drivers backed by
   data, caveats respected, clear for a controller — and lists fixes; the agent applies them by patch, at most two
   rounds. The score is stored with the answer.

## 6. Playbooks (start with four)

Each 1–2 pages, method and judgement only (the deterministic part is blocks and guards):
`monthly-pack`, `variance-commentary`, `gl-extract`, `prodreg-input`. A new playbook needs a recurring job with a
standard and at least one test-set question.

## 7. Answer memory

Every `fin_analyze` run stores question, plan, spec, key values, metrics, periods, as-of time and critic score in
`answers.db` under `%LOCALAPPDATA%\smartass` (per user; never in git, never shared). Searchable through
`fin_lookup(kind="answer")`; kept 24 months.

## 8. Monthly refresh (D14)

Runs automatically when the closed month advances: detected on the first call after the change, with a Windows
scheduled task on Tom's PC (working day 6, 07:00) as the backstop.

1. Rebuild the atlas (about 3 minutes).
2. Diff against last month: members added, removed, moved; renamed aliases.
3. Rebuild `knowledge.db`; build checks run.
4. **Impact report** (`knowledge/reports/<yyyy-mm>.md`, outside the release): metrics whose members changed,
   facts whose scope changed, new members with no metric or attribute, and drafted proposals (new members mapped
   by their parent's metric; new facts for moves).
5. Run the offline test set (recorded facts) and the knowledge checks.
6. `fin_status` shows "knowledge refresh: N items await approval"; Tom approves with
   `smartass knowledge approve <item>` (or all with a comment). Unapproved changes keep the last approved state
   and add a `warn` caveat where they apply.
7. A release with the new `knowledge.db` goes out to colleagues (sub-spec 4).

## 9. Testing

- Build checks (§3) in CI on the TOML sources.
- Golden lookups: 50+ phrases → expected metric/member (Norwegian and English), in CI.
- Caveat tests: given members and periods, the expected facts are returned.
- Generated adapters equal the template output (drift test).
- Critic calibration: the approved golden answers must score ≥ 4; seeded bad answers must score ≤ 2.

## 10. Migration from v1

`catalog.toml`, `capabilities.toml`, `concepts.json`, `knowledge/advisor/*.md` and the SKILL references are
converted to metrics, facts, playbooks and prompts; the four `Forretningskart` workbooks stay as human reading,
generated from `knowledge.db`. v1's knowledge files are deleted with v1.
