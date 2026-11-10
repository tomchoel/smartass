# smartass · knowledge layer: one source, behaviour cards, inherited caveats, a learning loop

Status: phase A done 10.11.2026 (v0.2.22). Design approved by Tom 10.11.2026 (grill rounds Q1–Q10, 06–10.11). Builds on sub-spec 2
(`2026-11-04-v2-2-knowledge-design.md`); where the two differ, this spec wins.

## 0. Why

The model must know FinMax, FinLite and FinPlan and all their quirks. Today it only partly does:

- Only `knowledge.db` (11 metrics, 37 facts, 24 concepts) and `atlas.db` (names, labels, parents) are read at
  runtime. `catalog.toml`, `capabilities.toml` and the four `*_map.json` (KPIs, account bundles, entity synonyms,
  vague words, grain, currency rules, group companies, always-comment list) are approved but never read (only
  `seed.py` touches them).
- A fact matches only the exact member it names: a question on a division misses its reorg caveats; C40S misses
  the C40510 trap.
- No member properties are stored: `AtlasTree.weight()` is always 1.0; time balance, account type and storage are
  unknown.
- Missing: EBIT, net income, EBITDA margin, capex, balance sheet, cash flow, FinPlan volumes/ARPU/FTE budget,
  FinPlan members for personnel/other opex/D&A, statistics accounts, counterparts, projects, reserves, the
  Target/Forecast scenarios in FinMax/FinLite.
- Contradictions: FinPlan entity level (catalog 2BUS, atlas 1BUS under BANORWAY), scenario rules in both catalog
  and `series.py`, materiality in two places, `fte.toml` materiality 5 000 000.

Best practice the design follows (research 06.11): a deterministic semantic layer behind tools, fail loudly;
caveats attached to members and returned in tool output; per-account behaviour cards; value search by tool;
verified questions with who/when; a result-based eval with several phrasings; corrections become regression tests.

## 1. Decisions

| # | Decision |
|---|---|
| Q1 | Full scope, built in three phases A → B → C (§8). |
| Q2 | One source: TOML in `knowledge/semantic/` (git) → `knowledge.db`. Catalog, capabilities and maps are folded in and deleted. |
| Q3 | Approval is automatic, with the guards of §5. |
| Q4 | A question bank of 50–100 questions, 2–3 phrasings each, judged on the result. |
| Q5 | Member properties from Smart View (read-only) where it gives them; otherwise derived from data or set by hand. No metadata export (not available to us). |
| Q6 | Caveats inherit down, up or both, per caveat; caveats can also scope a metric. |
| Q7 | New coverage: P&L gaps, capex + balance + cash flow, FinPlan volume/ARPU/FTE + scenarios, statistics/counterpart/project. |
| Q8 | A short always-on rule set; everything else just in time in tool output and `fin_lookup`. |
| Q9 | Stale or contradictory knowledge: the build fails on a contradiction; a stale atlas warns loudly, answers still run. |
| Q10 | The offline bank runs in every test run; the live bank on every version bump and before every automatic intake. |

## 2. One source

New semantic kinds next to `metrics/`, `facts/`, `concepts/` (each a pydantic model in `semantic.py`, strict,
`extra="forbid"`):

| Folder | Holds | From |
|---|---|---|
| `rules/` | cube choice, scenario rules (actual/plan per cube, open month), sign convention, P&L chain, materiality defaults, currency words, group-company patterns, checks | `catalog.toml` [pnl], [checks], [scenarios], [materiality], [group_companies]; `capabilities.toml` [currency]; `series.py` scenario constants |
| `entities/` | an entity: id, cube members, level, synonyms, successors, valid_from/to | `catalog.toml` [[entity]], `org_map.json` |
| `kpis/` | a FinPlan KPI (volume, ARPU, FTE): members, unit, balance, scenarios that hold it | `catalog.toml` [[kpi]], `finplan_map.json` |
| `metrics/` (bundles) | an account bundle is a plain metric: aliases, FinMax/FinLite members | `capabilities.toml` [[measure]] accounts |
| `notes/` | per member: meaning and booking pattern (234 FinMax accounts, 1 478 entities, 654 FinPlan accounts) | `accounts_map.json`, `org_map.json`, `finplan_map.json` |
| `vocab/` | vague words and grain ("costs", "the business", "per month") → the options to offer | `capabilities.toml` [[grain]], vague-word lists |
| `cards/` | behaviour cards (§3), generated + hand-set | new |

- `metrics/` gains EBIT, D&A (FinPlan members added), net income, EBITDA margin (ratio of two metrics), capex,
  balance-sheet totals, cash flow, statistics accounts, counterparts, projects, reserves; materiality fixed
  (`fte.toml`).
- `products_map.json` stays the input of `build_family_concepts.py` (it generates `concepts/`), but is no longer
  read at runtime by anything else.
- After migration: `catalog.toml`, `capabilities.toml`, `accounts_map.json`, `org_map.json`, `finplan_map.json`
  are deleted; `seed.py` loses its readers; `series.py` reads scenario rules from `knowledge.db`.
- Memory files and the skill point to `fin_lookup` / the semantic folder, never restate a rule.

Every item carries `status` and `approved = {by, on}`; `by` is `tom`, a colleague id, or `auto` (§5). Items also
carry `verified = {by, on}` when a live check confirmed them.

## 3. Behaviour cards

One card per member that matters (all accounts the metrics, bundles and KPIs name, their ancestors, plus every
member an answer has touched). Fields:

```toml
member = "ac.C40510"
cube = "FinMax"
time_balance = "flow"          # flow | balance_last | balance_first | average
account_type = "expense"       # revenue | expense | asset | liability | equity | statistic
variance = "expense"           # expense: lower is better; non_expense: higher is better
operator = "+"                 # +, -, ~ (ignored) towards its parent
storage = "stored"             # stored | dynamic_calc | shared | label_only
[provenance]                   # per field: smartview | derived | manual, with the date
```

Sources, in order of trust:

1. **Smart View** (read-only): try `req_GetMemberInformation` / `req_GetMembersInformation` and EnumMembers with
   `getAtts=1` on a few accounts first (probe, no code change). If they return the properties, add them to
   `svxml.ALLOWED` and fill the cards on every `smartass atlas build` (accounts and entities; not the 140 000
   projects).
2. **Derived from data** (ordinary ExecuteGrid reads, during `smartass knowledge build --cards`):
   - operator: parent vs the sum of its children for one closed year; a child that must be subtracted gives `-`,
     one that is left out gives `~`;
   - time balance: year total vs the sum of months vs December;
   - shared: a member that appears twice in the atlas;
   - account type and variance: from the P&L chain and the sign of the totals.
3. **Manual**: anything the two cannot settle, set in the TOML by hand or by an automatic proposal (§5).

`AtlasTree.weight()` reads `operator` from the card (default `+` when there is no card). A card field that two
sources disagree on fails the build (§6).

## 4. Caveats that inherit

`Fact` gains `inherit: "self" | "down" | "up" | "both"` (default `self`; every fact sets it explicitly) and keeps
`scope_metrics`. `up` climbs only to the nearest ancestor a metric or an entity names (C40510 stops at C40S),
so one account trap does not ride into every EBITDA answer.

- **down:** the fact applies to the member and every descendant (a reorg on 2CON reaches its cost centres).
- **up:** the fact applies to every ancestor ("contains X": C40S carries the C40510 trap; a division carries a
  quirk of one of its units).
- **metric:** a fact scoped to a metric applies wherever the metric is used.

The build expands the scopes against the atlas into `fact_scope` rows (with `via` = the member it came from), so
the resolver stays one indexed lookup. An expanded caveat says where it came from ("via ac.C40510").

## 5. Learning loop (automatic, with guards)

Sources of proposals: a member `fin_analyze` rejected and the model then found; a correction Tom gives in chat
that the model records with a new tool action `fin_lookup(kind="propose", ...)`; a derived card field.

A proposal is a TOML item written to the local overlay `%LOCALAPPDATA%\smartass\knowledge\auto\` (never to git
automatically). Intake:

1. Build a candidate `knowledge.db` with the overlay.
2. Run the offline bank; for metrics, mappings and rules also the live bank (§6).
3. Green: the candidate replaces `knowledge.db`; the item is marked `approved = {by="auto", on=<date>}`.
   Red: the item stays a draft and `fin_status` says why.
4. `fin_status` lists auto items since the last release; `smartass knowledge rollback <id|since DATE>` removes
   them and rebuilds.
5. On a release, auto items Tom has not rolled back are moved into `knowledge/semantic/` and committed.

Fasit answers in the question bank and the goldens are never changed automatically: they need Tom's yes.

## 6. Checks and the question bank

**Build checks (fail the build, keep the old db):** every member exists in the atlas; an entity's level matches
the atlas; a rule is defined once; card sources agree; no two items claim the same alias.

**Runtime freshness (warn loudly, never stop):** atlas older than 30 days, knowledge older than the atlas, cards
missing for a touched member. The warning goes into `warnings` of the answer and into `fin_status`.

**Question bank** (`engine/smartass/testset/bank/*.toml`): 50–100 questions, each with 2–3 phrasings and a fasit:
cube, metrics, entity, periods, and the expected key values (tolerance). Seeded with about 30 from `answers.db`
and the analyser workbooks; fasit approved by Tom. Two runs:

- offline (every test run): phrasing → spec → resolution, compared with the fasit spec;
- live (version bump, before an automatic intake): the full answer, compared with the fasit values.

The 9 existing goldens stay and become part of the bank.

## 7. Delivery to the model

- **Always on** (skill + server instructions, kept short): cube choice, signs and variance, P&L chain, "quote only
  key_values", "use fin_lookup when unsure".
- **Just in time:** `fin_lookup` and `fin_analyze` return for each touched member and metric its card (compact)
  and its caveats (with `via`); `fin_lookup(kind="rule")` returns a rule; vague words return their options.

## 8. Phases

| Phase | Builds | Done when |
|---|---|---|
| A | §2 (rules, entities, kpis, bundles, vocab; deletion of catalog/capabilities/maps), §4 inheritance, §6 build checks + freshness warnings, §7 delivery | 1252+ tests and 9/9 goldens green; nothing reads the deleted files; C40S and a division query return the inherited caveats |
| B | §3 cards (Smart View probe first, then derive, then manual), new metrics/KPIs of Q7, `weight()` from cards | cards for every member a metric/bundle/KPI names; a parent with a `-` child sums right |
| C | §6 question bank (offline + live), §5 learning loop | ≥50 questions green offline and live; a proposal goes through intake and rollback |

Each phase gets its own plan in `docs/plans/`.

## 9. Out of scope

Write access to EPM; a metadata export from the EPM owners; cards for the project dimension; the GitHub/colleague
roll-out (fields are ready for it).
