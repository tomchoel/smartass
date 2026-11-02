# The scope gate — agent procedure

Every finance question goes through the gate before anything is pulled (design §3.1). The engine runs a fixed
checklist; the agent follows the suggested outcome and may overrule it only with a stated reason.

## 1. Call `sv_gate(question)` first

Pass the user's question as asked. The result has `suggested_outcome`, a `checklist` with one entry per item
(`period`, `scenario`, `entity`, `measure`, `currency_sign`, `granularity`), each with `value`, `source`
(`question` | `convention` | `learned` | `catalog`), `gap` and `ask`, plus `questions`, `reason`, `proxy`,
`warnings`, `entities`, `measures`, `concepts` and `next`.

## 2. Follow the outcome

| Outcome | Do |
|---------|----|
| **clear** | Proceed. Every checklist item with source `convention` or `learned`, and every warning, becomes a stated assumption on the answer sheet. Build the answer spec and call `sv_answer`. |
| **ambiguous** | Pull nothing. Ask the returned questions (max 3) in **one** message. When answered, call `sv_gate` again with the answers folded into the question. |
| **infeasible** | Do not pull what was asked. Say why in one or two sentences (`reason`), show what is possible, and offer the `proxy`. When the user agrees, build it with gate `infeasible`, `reason` and `proxy` in the spec — the sheet says what is answered instead. |

Copy the gate result into the answer spec: `gate.outcome`, `gate.checklist` (as text), `gate.questions` /
`reason` / `proxy`, and `assumptions`.

## 3. Ask policy: balanced

Ask only when the gap changes the numbers and no convention covers it: a vague measure ("costs", "numbers"),
"this year" (YTD or full year), a measure the catalog cannot map. Everything else is a stated convention:

- No period → YTD to the last closed month.
- No scenario → Actual (FinMax ledger). "Budget" → Target, version Final (FinPlan). Full year without a scenario →
  latest forecast.
- No entity → Telenor Norge (un.SY).
- NOK, FinMax signs (revenue positive, costs negative).

## 4. Learned defaults

After the user answers a clarifying question, offer once: *"Save 'costs = opex' as your default?"* Only on yes,
call `sv_remember(item, phrase, value)`. Next time the gate states it (source `learned`) instead of asking.
Defaults live in the user's `preferences.toml`; good ones are later promoted to `knowledge/capabilities.toml`.

## 5. Language

Questions, assumptions, sheets and commentary are in English, also when the user writes Norwegian.
