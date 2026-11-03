# Smart Flow Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Same answers with less waiting: one AI call per question, a hidden browser that is always signed in, replies a quarter of their size, and about half the Excel calls.

**Architecture:** `tools_core.answer()` gains a question-only path that runs the gate and builds the recipe spec; a `reply.lean()` pass trims the reply; `DrillBrowser`'s worker stays alive (keep-alive check, yield flag for other processes) and the server signs in at start; Excel writers batch their COM calls (measured by the `com calls.*` counter).

**Tech Stack:** Python 3.12, pywin32 COM, Playwright (sync), pytest. Run tests from `C:\dev\smartass\engine`: `.venv/Scripts/python.exe -m pytest -q tests` (never plain `python`); lint `.venv/Scripts/python.exe -m ruff check smartview_mcp tests`. A hook blocks `python -` (stdin): write .py files. Never run anything against live EPM / browser / Excel in these tasks (Task 6 is the live check, done by the controller).

Spec: `docs/specs/2026-11-03-smart-flow-design.md` (read it first). Branch `feat/smart-flow`. Commits end with:
```
Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01MW9ANsye1x7As56S36PZbu
```
Keep each file's existing line endings; don't reformat untouched lines.

---

### Task 1: Lean replies (spec C)

**Files:** Modify `engine/smartview_mcp/core/reply.py` (+ `lean`), `engine/smartview_mcp/tools_core.py` (`sv_answer`/`answer` get `detail: str = "lean"`; every return of a built answer passes through `reply.lean(out, detail)`). Test `engine/tests/test_reply_lean.py` (new).

- [ ] **Step 1: Failing tests** in `tests/test_reply_lean.py`:

```python
from smartview_mcp.core import reply


def _big():
    return {
        "workbook": "w.xlsx", "question": "Q1", "answer_sheet": "Q1 Answer", "sheets": ["Q1 Answer"],
        "data_sheets": [], "note": "n", "gate": "clear", "period": "Oct 26", "routes": {"Q1 FinMax": "http"},
        "key_points": ["k"], "followups": ["f"], "warnings": ["w"], "lines": {"d1": 3},
        "key_values": {"Revenue": {"Oct 26": 2247931399.1299996, "Change %": -0.5236831758915402}},
        "checks": {"a": "OK", "b": "OK", "c": "MISMATCH 12"},
        "review": {"ok": True, "findings": [{"severity": "Warning", "text": "x"}, {"severity": "Warning", "text": "x"}]},
        "material": [{"line": "s1", "label": "L", "cur": 1.0, "cmp": 2.0, "delta": -1.0, "pct": -0.5, "material": True,
                      "always": False}],
        "outline": {"recipe": "pnl", "pulls": []},
        "timings": {"_total": 10.0, "answer": 3.0, "answer.style": 2.0, "write source sheet": 2.5, "log sources": 1.0,
                    "open target": 0.1, "com calls.answer": 150, "first frame": 0.8, "drills": 8.0},
        "formula_style": {"ok": True, "checked": 9, "counts": {"LONG": 2}, "block": [], "warn": [{"cell": "A1"}] * 20},
        "style_lint": [f"l{i}" for i in range(9)],
        "table": "| Line |\n" + "\n".join(f"| r{i} |" for i in range(60)),
        "flags": [{"x": "y" * 100}] * 40, "landing": {"a": 1},
    }


def test_lean_keeps_what_the_answer_needs_and_trims_the_rest():
    out = reply.lean(_big(), "lean")
    assert out["key_values"]["Revenue"] == {"Oct 26": 2247931399, "Change %": -0.5237}
    assert out["checks"] == {"ok": 2, "c": "MISMATCH 12"}
    assert out["review"]["findings"] == [{"severity": "Warning", "text": "x"}]
    assert out["material"] == [{"line": "s1", "label": "L", "delta": -1.0, "pct": -0.5}]
    assert set(out["timings"]) == {"_total", "first frame", "drills", "answer", "write source sheet", "log sources"}
    assert out["formula_style"] == {"ok": True, "counts": {"LONG": 2}, "block": []}
    assert out["style_lint"] == [f"l{i}" for i in range(5)]
    assert "table" not in out                                  # 60 rows > 40: the numbers are in key_values
    assert out["flags"] == "40 items: ask with detail='full'" and out["landing"] == {"a": 1}
    for k in ("workbook", "question", "answer_sheet", "sheets", "routes", "key_points", "followups", "warnings",
              "outline", "gate", "period", "lines", "note"):
        assert out[k] == _big()[k]


def test_full_detail_and_small_tables_are_kept():
    assert reply.lean(_big(), "full") == _big()
    small = {**_big(), "table": "| a |\n| b |"}
    assert reply.lean(small, "lean")["table"] == "| a |\n| b |"
    assert reply.lean({"gate": "ambiguous", "questions": ["q"]}, "lean") == {"gate": "ambiguous", "questions": ["q"]}
```

- [ ] **Step 2:** run, see FAIL (`AttributeError: lean`).
- [ ] **Step 3: Implement** `reply.lean(out: dict, detail: str = "lean") -> dict` in `core/reply.py` exactly per spec C (round NOK to int; a value counts as a ratio when its column label contains `%` → 4 decimals; any other float → int; dedupe findings by `(severity, text)` keeping order; `material` keeps `line, label, delta, pct`; timings = `_total`, `first frame`, `drills` (when present) + the 3 largest keys without `.` and not starting with `com calls.` or `_`; table kept when it has ≤ 40 lines; unknown extra keys kept when `len(json.dumps(v)) <= 1500`, a list longer → `"<len> items: ask with detail='full'"`, a dict longer → `"<len> keys: ask with detail='full'"`). `detail == "full"` returns `out` unchanged. Replies without `key_values` (gate replies, errors) pass through unchanged except the generic rules.
- [ ] **Step 4:** in `tools_core.py` add `detail: str = "lean"` to `sv_answer` (document it: "detail='full' returns the whole table and every recipe extra") and to `answer`; wrap every built-answer return (`return out` in `job`, the drill path's return) with `reply.lean(..., detail)`. Check `http_api`/`POST /answer` callers keep working (default lean is fine).
- [ ] **Step 5:** tests pass; full suite; ruff. **Commit** `feat(answer): lean replies - key values rounded, checks summarised, big extras on request (detail='full')`.

---

### Task 2: One call per question (spec A)

**Files:** Modify `engine/smartview_mcp/tools_core.py` (`sv_answer`, `answer`), `engine/smartview_mcp/server.py` (instructions text), `knowledge/advisor/gate.md`, `skills/smartass/SKILL.md`, `skills/smartass/reference/operations.md` (and `recipes.md` if it describes the two-step call). Test `engine/tests/test_answer_one_call.py` (new).

- [ ] **Step 1: Failing tests** (mock the gate and the spec path; read `tools_core.answer` first):

```python
import asyncio

import pytest

from smartview_mcp import tools_core


def _gate(outcome, recipe=None, checklist=None, warnings=()):
    return {"question": "q", "suggested_outcome": outcome, "questions": ["Which measure?"] if outcome == "ambiguous" else [],
            "reason": None, "proxy": None, "warnings": list(warnings), "recipe": recipe, "closed_month": "Oct 26",
            "checklist": checklist or {"period": {"value": "YTD Oct 26", "source": "convention"},
                                       "entity": {"value": "Wholesale", "source": "question"},
                                       "currency_sign": {"value": "NOK", "source": "learned"}},
            "next": "..."}


def test_a_clear_question_is_answered_in_one_call(monkeypatch):
    seen = {}

    async def gate(question, answers=None):
        seen["answers"] = answers
        return _gate("clear", {"name": "pnl", "params": {"area": "un.SY.2WHS"}}, warnings=["note w"])

    async def spec_answer(spec, followup, q, target, allow_archive, detail="lean"):
        seen["spec"] = spec
        return {"ok": True}
    monkeypatch.setattr(tools_core, "gate", gate)
    monkeypatch.setattr(tools_core, "_answer_spec", spec_answer)
    out = asyncio.run(tools_core.answer(question="q", answers={"measure": "EBITDA"}))
    assert out == {"ok": True} and seen["answers"] == {"measure": "EBITDA"}
    spec = seen["spec"]
    assert spec["recipe"] == {"name": "pnl", "params": {"area": "un.SY.2WHS"}} and spec["question"] == "q"
    assert spec["assumptions"] == ["Period: YTD Oct 26 - convention.", "Currency sign: NOK - learned.", "note w"]


@pytest.mark.parametrize("outcome", ["ambiguous", "infeasible"])
def test_an_unclear_question_returns_the_gate_and_pulls_nothing(monkeypatch, outcome):
    async def gate(question, answers=None):
        return _gate(outcome)

    async def never(*a, **k):
        raise AssertionError("must not answer")
    monkeypatch.setattr(tools_core, "gate", gate)
    monkeypatch.setattr(tools_core, "_answer_spec", never)
    out = asyncio.run(tools_core.answer(question="q"))
    assert out["suggested_outcome"] == outcome


def test_clear_without_a_recipe_asks_for_a_spec(monkeypatch):
    async def gate(question, answers=None):
        return _gate("clear", None)
    monkeypatch.setattr(tools_core, "gate", gate)
    out = asyncio.run(tools_core.answer(question="q"))
    assert "spec" in out["next"]


def test_exactly_one_of_question_spec_followup():
    from smartview_mcp.host import InputError
    with pytest.raises(InputError):
        asyncio.run(tools_core.answer(question="q", spec={"version": 1}))
    with pytest.raises(InputError):
        asyncio.run(tools_core.answer())
```

- [ ] **Step 2:** run, see FAIL.
- [ ] **Step 3: Implement.** Rename today's body of `answer()` to `_answer_spec(spec, followup, q, target, allow_archive, detail="lean")` (unchanged logic, plus Task 1's lean). New `answer(spec=None, followup=None, q=None, target="session", allow_archive=False, question=None, answers=None, detail="lean")`: exactly one of `question`, `spec`, `followup` (InputError otherwise, message naming the three). With `question`: `g = await gate(question, answers)`; if `g["suggested_outcome"] != "clear"` return `g`; if not `g.get("recipe")`: return `{**g, "next": "Clear, but no recipe fits: send sv_answer(spec=...) with pulls and layout."}`; else build the spec: `{"version": 1, "question": question, "gate": g, "assumptions": _assumptions(g), "recipe": g["recipe"]}` and `return await _answer_spec(spec, None, None, target, allow_archive, detail)`. `_assumptions(g)`: for each checklist item in order whose `source` is `convention` or `learned`: `f"{label}: {value} - {source}."` with label = the item key with `_` → space, first letter capitalised (`currency_sign` → "Currency sign"); then each warning as is. Passing the gate result as `spec.gate` already works (`contract.normalize_gate`; check `answers` reach the ambiguity re-check — today the ambiguous path needs `gate.answers`: set `g["answers"] = answers` when answers were given, and verify with the existing gate tests).
- [ ] **Step 4:** `sv_answer` MCP tool: add parameters `question: str | None = None`, `answers: dict[str, str] | None = None`, `detail: str = "lean"`; docstring first lines: "Answer ONE finance question. New question: sv_answer(question=...) - it runs the gate itself; ambiguous -> ask its questions in ONE message, then sv_answer(question=..., answers={item: answer}). Follow-up: sv_answer(followup=...). A hand-written spec (spec=...) still works." Keep the rest of the docstring accurate. `sv_gate`'s docstring: "Optional: sv_answer(question=...) runs the gate itself; call sv_gate alone only to inspect the scope."
- [ ] **Step 5: Docs.** `server.py` instructions: step 1/2 become "sv_answer(question=...) (it runs the gate; ambiguous -> ask, then sv_answer(question, answers=...))". Update `knowledge/advisor/gate.md`, `skills/smartass/SKILL.md`, `skills/smartass/reference/operations.md` the same way (short, keep their style; keep the rule that a departure needs spec.deviation, which now means writing a spec).
- [ ] **Step 6:** tests pass; full suite; ruff. **Commit** `feat(answer): one call per question - sv_answer(question=...) runs the gate and the recipe`.

---

### Task 3: Always warm (spec B)

**Files:** Modify `engine/smartview_mcp/config.py`, `engine/smartview_mcp/drill_browser.py` (`_loop`, DrillBusy path, + keep-alive), the server start-up (find where the MCP server starts: `server.py` / `child_server.py`; add a background warm-up), `engine/tests/conftest.py` (keep-warm off in tests unless a test turns it on). Test `engine/tests/test_drill_keepwarm.py` (new).

- [ ] **Step 1: Config:** `DRILL_KEEP_WARM = os.environ.get("SV_DRILL_KEEP_WARM", "1") != "0"`, `DRILL_KEEPALIVE_S = int(os.environ.get("SV_DRILL_KEEPALIVE_S", "600"))`. In conftest's autouse fixture: `monkeypatch.setattr(config, "DRILL_KEEP_WARM", False)`.
- [ ] **Step 2: Failing tests** with the real `_loop` on a thread and fakes (see `tests/test_drill.py::test_idle_worker_closes_browser_and_frees_lock` for the pattern):
  1. keep-warm on: after several idle timeouts the browser is NOT closed, and a keep-alive callable is invoked once `DRILL_KEEPALIVE_S` (patched to 0.05) has passed;
  2. keep-warm on + a `.wanted` file newer than the browser start in the profile dir: the idle worker closes the browser (lock freed) and deletes the flag;
  3. keep-warm off: today's idle close (existing test keeps passing).
  4. `DrillBrowser._start` raising `DrillBusy` writes `<profile>/.wanted`.
- [ ] **Step 3: Implement.**
  - `_loop(q, s, idle_s, keepalive=None)`: when idle (queue.Empty, no inflight): if `config.DRILL_KEEP_WARM` and `s.ctx is not None`: if the profile's `.wanted` exists and is newer than `s.started_at` → `_close(s)`, delete it; elif `time.monotonic() - s.last_alive >= config.DRILL_KEEPALIVE_S` → call `keepalive(s)` (exceptions logged, never raised), set `s.last_alive`. Else (keep-warm off) today's close. Use a short idle poll (e.g. `min(idle_s, 30)`) when keep-warm is on so the flag and keep-alive are noticed; `_Session` gets `started_at`, `last_alive`.
  - Keep-alive callable (`DrillBrowser._keepalive(s)`): `self._ensure(s)` (restarts a dead browser / re-signs-in) then one cheap check that the EPM session is alive: `s.page.evaluate("1")` plus a GET of `/HyperionPlanning/` HEAD-like fetch through `_fire` is fine; simplest robust choice: `self._ensure(s)` and a `_fire` GET of `self.base + "/HyperionPlanning/"` with `{"method": "GET"}`, wait on it inside the job with `_collect` polling until done or 30 s; a non-200 or a sign-in URL → `s.ready = False; self._ensure(s)`. Set `s.deadline = time.monotonic() + config.DRILL_SETUP_S` first.
  - DrillBusy in `_start`: before raising, write `profile / ".wanted"` (`Path.write_text(str(time.time()))`, ignore OSError). The caller paths that get DrillBusy (svhttp → Excel fallback; drill) need no change; optional: `svhttp` retries once after 3 s on DrillBusy — only if trivial.
  - Server start: after the MCP server is up, submit a warm-up job to `drill_browser.shared_browser()` (`lambda s: b._ensure(s)` with `s.deadline` set), not awaited, wrapped in try/except with a log line; only when `config.DRILL_KEEP_WARM`.
- [ ] **Step 4:** tests pass; full suite; ruff. **Commit** `feat(drill_browser): always warm - no idle close, keep-alive sign-in check, yields to another smartass process, warm-up at start`.

---

### Task 4: Fewer Excel calls (spec D)

**Files:** Modify `engine/smartview_mcp/core/source.py` (`write`), `engine/smartview_mcp/styling.py` (`apply_data_sheet`, page setup, answer styling groups), the save/arrange code (find `"arrange sheets"` stage: `grep -rn "arrange sheets" engine/smartview_mcp`). Tests next to the existing ones for those modules (`grep -rln "apply_data_sheet\|arrange" engine/tests`).

The goal is fewer `jobs.com` calls with **identical results**. Method for each part:
- [ ] **Step 1: A call-count test harness.** Use the existing COM fakes in the tests for these writers (find them first). For each of: `source.write` on a 5x14 grid, the arrange-sheets function on 8 sheets already in order and on 8 with one out of place, page setup on a sheet that already has it: write a test asserting the call count **after** your change (count `jobs.com` calls via monkeypatch) AND that the resulting fake workbook state (values, formats, fonts, sheet order, page setup fields) equals the state from before your change (capture it first, before changing code, and hard-code it in the test).
- [ ] **Step 2: Source sheet.** In `apply_data_sheet` and `source.write`: group number formats/fonts/fills by identical setting and apply each to one multi-area or contiguous range (`ws.Range("B3:M3,B5:M5")` style union or contiguous blocks) instead of per row/column; provenance panel: one `Value` write + one font call (Color + Italic via the same `Font` object, two attribute sets inside ONE `com(lambda: ...)` call). Target ≤ 25 calls per source sheet.
- [ ] **Step 3: Arrange sheets.** Read all sheet names in order with one call (`[s.Name for s in wb.Worksheets]` inside one `com`), compute the target order, move only the sheets whose position differs (each move one call). Target: 1 call when already ordered.
- [ ] **Step 4: Page setup once per sheet.** At the end of the answer page setup, add a sheet-level name `smartass_pagesetup` with value `="v1"` (`ws.Names.Add`); before page setup, if the sheet has that name with `v1`, skip the page setup (one call to check). Bump the version string if the page setup ever changes.
- [ ] **Step 5: Answer styling.** From `com calls.answer.style (tn-rapportstil).*` (B cell styles 20-38, C widths+heights 24-36, F tab+freeze 20): batch per-cell/per-column calls into range calls where the result is identical. Keep `fast_excel` usage.
- [ ] **Step 6:** tests pass (existing styling/answer tests unchanged = output unchanged); full suite; ruff. **Commit** per part (`perf(excel): ...`).

---

### Task 5: Review (controller dispatches a reviewer over `git diff main..HEAD -- engine/`), fix findings.

### Task 6: Live proof (controller)
The server reloads by itself. Run the three answers with `sv_answer(question=...)` (P&L Wholesale YTD vs last year, answers {"measure": "the P&L to EBITDA"}; "Hvilke konsulentselskaper fakturerte mest til Technology hittil i år?"; "Telenor Norge monthly revenue trend this year vs last year", answers {"period": "YTD to Oct 26"}); record reply sizes, round trips, `_total`, stage times and `com calls.*` vs the 2026-11-03 bench; all checks OK; append to `docs/audit/2026-11-02-speed-bench.md`; PR, CI, merge.
