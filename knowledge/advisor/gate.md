# The scope gate — agent procedure

Every finance question goes through the gate before anything is pulled (design §3.1). The engine runs a fixed
checklist; the agent follows the suggested outcome and may overrule it only with a stated reason.

## 1. Call `sv_answer(question=...)` first

Pass the user's question as asked. `sv_answer(question=...)` runs the gate itself and answers a clear question with a
recipe in the same call; `ambiguous` / `infeasible` return the gate result below and pull nothing. `sv_gate(question)`
alone only inspects the scope. The gate result has `suggested_outcome`, a `checklist` with one entry per item
(`period`, `scenario`, `entity`, `measure`, `currency_sign`, `granularity`), each with `value`, `source`
(`question` | `convention` | `learned` | `catalog` | `user`), `gap` and `ask`, plus `questions`, `reason`, `proxy`,
`warnings`, `entities`, `measures`, `concepts`, `recipe` and `next`.

`recipe` is the standard analysis for a clear question, `{name, params}` (first matching rule in
`core/scope.py` `_recipe`), or `null` when no recipe fits (write a free-form spec).

## 2. Follow the outcome

| Outcome | Do |
|---------|----|
| **clear** | Proceed. Every checklist item with source `convention` or `learned`, and every warning, becomes a stated assumption on the answer sheet. `sv_answer(question=...)` has already built the sheets when the gate named a recipe; when `recipe` is `null` it asks for a free-form spec (`sv_answer(spec=...)`). Never ask whether to build a sheet. |
| **ambiguous** | Pull nothing. Ask the returned questions (max 3) in **one** message, in the user's language. When answered, call `sv_answer(question=<same question>, answers={item: answer})` (e.g. `{"measure": "opex"}`); each answer resolves its item with source `user`, and the answer is built in that call. A hand-written `spec` passes the same `answers` as `spec.gate.answers`: `sv_answer` re-runs the gate and pulls nothing while it is still ambiguous. |
| **infeasible** | Do not pull what was asked. Say why in one or two sentences (`reason`), show what is possible, and offer the `proxy`. When the user agrees, build it with gate `infeasible`, `reason` and `proxy` in the spec — the sheet says what is answered instead. |

Copy the gate result into the answer spec: `gate.outcome`, `gate.checklist` (as text), `gate.questions` /
`reason` / `proxy`, `gate.answers`, and `assumptions`. A different recipe, or a param that switches off a recipe's
finance default, needs `deviation` = the user's explicit request (written on the sheet).

## 3. Ask policy: balanced

Ask only when the gap changes the numbers and no convention covers it: a vague measure ("costs", "numbers"),
"this year" (YTD or full year), a measure the catalog cannot map. Everything else is a stated convention:

- No period → YTD to the last closed month. A bare current year ("i 2026") with actuals → YTD to the last closed
  month.
- No scenario → Actual (FinMax ledger). "Budget" → Target, version Final (FinPlan). Full year without a scenario →
  latest forecast.
- No entity → Telenor Norge (un.SY).
- NOK, FinMax signs (revenue positive, costs negative).
- Cost types map to fixed accounts (`knowledge/capabilities.toml`), e.g. consultants = C60080 + C60090,
  consultants and temps = C693, contractors = C60110, legal = 67410 + 67420 + C60400, travel = C60130.

## 4. Learned defaults

After the user answers a clarifying question, offer once: *"Save 'costs = opex' as your default?"* Only on yes,
call `sv_remember(item, phrase, value)`. Next time the gate states it (source `learned`) instead of asking.
Defaults live in the user's `preferences.toml`; good ones are later promoted to `knowledge/capabilities.toml`.

## 5. Language

Ask the questions and write the chat in the user's language. Sheet text is English (assumptions, titles, notes,
headers, check labels, drivers); `sv_answer` refuses Norwegian sheet text. Norwegian synonyms are understood.
