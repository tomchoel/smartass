# smartass v2 · Plan 2B: critic, playbooks, answer memory

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Insight at controller level, as set out in sub-spec 2 §5–§7:
- **Critic.** A scored rubric the agent applies to every standard or deep answer and every pack, fed by a compact
  brief the engine returns.
- **Playbooks.** The four recurring jobs as MCP prompts and resources, so Claude and Copilot get the same method.
- **Answer memory.** Every run is stored locally and is searchable, so the agent can write continuity ("last time
  the driver was X").

**Architecture:**
- **Content.** `knowledge/playbooks/*.md` and `knowledge/agent/{critic,plan}.md` are Markdown in git.
- **Exposure.** The server exposes them as MCP prompts (`@mcp.prompt`, available in the mcp SDK 2.3) and resources
  (`@mcp.resource`).
- **Critic brief.** `fin_analyze` and `fin_refresh` replies gain `critic_brief`:
  - the plan
  - the material lines with whether a driver backs each of them
  - caveats, findings and checks
  - which comparisons are present
  - the commentary drivers and their basis
- **Critic tool.** `fin_critic(q)` returns the brief and the rubric for a stored answer.
- **Scores.** The agent scores the answer and stores the score with `fin_critic(q, score, notes)`.
- **Answer memory.** It lives in `%LOCALAPPDATA%\smartass\answers.db` (SQLite: run id, time, workbook, q,
  question, purpose, spec JSON, key_values JSON, metrics, entities, periods, critic score). It is searched with
  `fin_lookup(kind="answer")`. Rows are kept 24 months.

**Out of scope:** product concepts (`splitt`, `fam_*`), which need the Product dimension. That is Plan 2C.

**Conventions:** as Plan 2A.
- Branch `v2/2b` from `main`. TDD.
- Gates after every task: `uv run pytest -q`, `uv run ruff check .` and the test set at 8/8.
- Never run python through a stdin heredoc.
- Stage files explicitly.
- Commits end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

### Task 0: Branch

`cd /c/dev/smartass && git switch main && git pull --ff-only && git switch -c v2/2b`

### Task 1: Rubric and plan guide (content)

`knowledge/agent/critic.md` (English, under 1 page) holds the rubric. Each criterion is scored 1–5, with what
earns a 2 and what earns a 5:
1. **Right question.** The answer covers the plan's purpose; no unasked analysis dominates.
2. **Numbers reconciled.** Checks OK, Sheet = model OK, gaps explained.
3. **Comparison basis.** Last year for P&L and cost; budget or forecast in FinPlan when the question asks for
   plan; like-for-like periods.
4. **Drivers backed by data.** Every material line has a driver; drivers carry `data` evidence (drivers block,
   lines block, cell keys). Hypotheses are labelled.
5. **Caveats respected.** Every warn and block caveat is reflected in the commentary or the notes.
6. **Clear for a controller.** Key points first, 3–5 of them; NOK format; no jargon.

The score is the rounded mean. It must be at least 4 for standard or deep answers and packs. Below 4 the agent
fixes the answer by patch, at most 2 rounds.

`knowledge/agent/plan.md` is the guide for writing `plan`:
- purpose, hypotheses, audience and depth;
- when to use which block;
- the full EBITDA bridge;
- drivers for material lines;
- lines blocks for supplier questions;
- how to read caveats.

**Test:** both files exist and are under 4000 characters, and the critic file names all 6 criteria.

### Task 2: Critic brief in the reply

`model/critic.py` provides `brief(spec, result, caveats, checks) -> dict`:

```text
{"purpose", "depth", "hypotheses",
 "material": [{"key", "line", "variance", "backed": bool, "evidence": [...]}],
 "comparisons": ["ly", "budget", ...],
 "caveats": [...warn/block only],
 "checks": {...}, "findings": [...],
 "commentary": {"key_points": n, "drivers": [{"line", "basis"}]},
 "gaps": ["material line 'X' has no driver", "warn caveat 'Y' not mentioned", ...]}
```

- **Backed.** A material line counts as backed when either of these holds:
  - a drivers or lines block in the same answer covers the same metric and entity;
  - a commentary driver names the line with basis `data`.
- **Gaps.** Material lines that are not backed, warn or block caveats whose text is not referenced in the
  commentary (match on the caveat id or its first 30 characters), and a missing last-year comparison on a P&L
  query.
- **Wiring.** `pipeline.reply` adds `critic_brief` when depth is standard or deep. `fin_refresh` adds one per
  question.
- **Tests:** the brief for a spec with a drivers block marks the Opex line backed; no commentary gives a gap per
  warn caveat; a quick depth gives no brief.

### Task 3: `fin_critic` and score storage

- **`answers.py`**
  - `store(run)` is called at the end of `excel_job` and `redraw_job` with the reply plus spec data.
  - `search(text, limit)` searches `question`, `purpose`, `metrics` and `entities` (FTS5).
  - `get(workbook, q)`.
  - `set_score(workbook, q, score, notes)`.
  - It purges rows older than 24 months.
- **`fin_critic(q=None, score=None, notes="")`**
  - Without a score, it returns `{"brief", "rubric"}` for the stored answer: the active Q, otherwise the latest,
    found the same way follow-ups find it. The rubric is the text of `critic.md`.
  - With a score (1–5, a float), it stores it and returns `{"stored": true}`.
  - Register it with `_register(kit.LOCAL)`.
- **`fin_lookup(kind="answer")`** returns the matching past answers: workbook, q, question, when, critic score and
  the top key values (at most 5 lines).
- **Tests:**
  - a store and search round trip, on a temporary home;
  - storing a score;
  - `fin_critic` without a score returns the rubric text;
  - `TOOL_NAMES` includes `fin_critic`, and the skill-tools test is updated.

### Task 4: Playbooks as MCP prompts and resources

Create `knowledge/playbooks/{monthly-pack,variance-commentary,gl-extract,prodreg-input}.md`. Each is 1–2 pages,
method and judgement only:
- **variance-commentary:** based on `knowledge/advisor/variance-commentary.md`, which is vendored (Apache-2.0, see
  NOTICE.md). Put the vendored text in a quoted section with attribution and add our rules: drivers basis, caveats
  and the FinMax/FinPlan rule.
- **monthly-pack:**
  1. build the pack questions once with relative periods;
  2. run `fin_refresh` each month;
  3. run `fin_critic` per question;
  4. the commentary patch.
- **gl-extract:** use `fin_drill` and the lines block per account and cost centre, the Yr period plus a month
  filter, and the reconciliation that must be exact.
- **prodreg-input:** use the drill per division node, sum to YTD, and keep the interview support workbook rule.
  Point to the produktregnskapet skill's `reference/intervjuforberedelse.md`.

**Server wiring:**
- `@mcp.prompt` for `plan`, `critic` and the four playbooks. Each returns the Markdown with a one-line preface.
- `@mcp.resource("smartass://playbooks/<name>")` and `smartass://conventions`. The conventions are finance
  format, signs and cube choice, taken from the skill.

**Tests:**
- the prompts and resources are listed (use the SDK's `list_prompts` and `list_resources` on the server object,
  through anyio);
- every playbook file is under 6000 characters;
- the vendored attribution is present.

### Task 5: Skill, docs and instructions

- **`SKILL.md`** (stay under the size test):
  - the loop now ends with the critic for standard and deep answers (`fin_critic` → score → fix → store);
  - the playbooks are MCP prompts;
  - `fin_lookup(kind="answer")` gives continuity;
  - the tool list now has 9 tools.
- **INSTRUCTIONS:** one sentence on the critic and the prompts, keeping the server-instructions length test green.
- **Other:** `docs/DESIGN.md` and the CHANGELOG.

### Task 6: Live and finish (controller)

1. Run `fin_status(restart=True)`.
2. Run `fin_analyze` with a standard-depth Business opex question: kpi tiles, table, drivers and a commentary with
   one data driver. The reply must carry `critic_brief`.
3. Run `fin_critic`, score the answer against the rubric, and store the score.
4. Run `fin_lookup(kind="answer", query="Business opex")`. The answer must be found with its score.
5. The MCP prompt list must show the 6 prompts. Use the `ListMcpResourcesTool` / prompts if available, or a stdio
   smoke check.
6. Merge and push.

---

## Live result (2026-11-04)

- **Restart:** `fin_status(restart=True)` brought in the new code. The doctor warned that knowledge.db was older
  than the semantic files; `smartass knowledge build` fixed that (11 metrics, 37 facts, 0 errors).
- **Critic brief on a standard answer:** Business Norway opex YTD, with kpi tiles, a table, drivers and
  commentary. The brief flagged two real gaps: Personnel cost and Other opex had no driver.
- **Score:** `fin_critic(q=4, score=4, notes)` stored it. The per-criterion scores were 5/5/5/3/5/3, mean 4.33.
- **Answer memory:** `fin_lookup(kind="answer", "Business Norway opex")` found the answer with its score and notes.
- **Prompts:** the six MCP prompts were verified over stdio by the implementer.
- Merged into main and pushed.

## Carried forward

- **Auto-build:** the knowledge build is stale after a branch checkout that touches `knowledge/semantic`. Make
  `run_prepare` rebuild automatically when the semantic hash differs, not only when the db is missing.
- **SKILL.md size:** it is at 6983/7000 characters. Move detail to `reference/` before adding anything.
