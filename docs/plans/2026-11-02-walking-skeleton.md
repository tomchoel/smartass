# Walking skeleton (B02 end to end) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** One question (benchmark B02, Business Norway EBITDA per month 2026 vs 2025, reconciled to YTD) answered
end to end on the new core: scope → pull (form C) → answer (skeleton sheet first) → check, with a run log, through
one new MCP tool `sv_answer` that takes an answer spec v1.

**Architecture:** A thin new core in `engine/smartview_mcp/core/` beside the old engine (design
`docs/specs/2026-11-02-pipeline-architecture-design.md` §3). Pulls go over the Smart View HTTP provider
(`svhttp.fetch`) and are written as one block in Smart View's own grid shape at A1, so the proven label-keyed
formulas of `answer.py` work unchanged; the ad hoc COM pull (`userexcel.pull`) is the fallback. The answer sheet
`Q<n> Answer` appears first (question, assumptions, status line) and is then rendered by `answer.build` in one block.

**Tech Stack:** Python 3.12, pywin32 COM, pydantic v2, pytest (+ anyio), uv, ruff.

**Commands** (always from `C:\dev\smartass\engine`): tests `uv run pytest <file> -q`; lint `uv run ruff check
smartview_mcp tests`. Live benchmark: `uv run python tests\live\bench.py --only B02 --route core --out bench-core-b02`.

**Decisions taken in this plan**
- Provenance "header" (design §3.2) is a **panel to the right of the grid** (one blank column gap, from row 2). A
  header above the grid would move every label the lookups key on, and row 1 is Smart View's POV row (`grid.parse`
  reads extra row-1 cells as appended POV members). Design §3.2 is updated accordingly (Task 9).
- The gate verdict arrives in the spec (the agent applies the checklist); an `ambiguous` spec returns its questions
  without touching Excel. The full advisor is phase 4.
- New questions only; follow-ups (`spec_patch`) are phase 5. The full v1 spec is stored in `_svmeta` as `SPEC Q<n>`.
- On failure the partial sheets are rolled back, as the old tools do.

---

## File structure

| File | Responsibility |
|------|----------------|
| `smartview_mcp/svhttp.py` (modify) | `Fetched` + `SvHttp.fetch()` returning the raw provider grid; `retrieve()` wraps it |
| `smartview_mcp/sheetorder.py`, `userexcel.py` (modify) | Recognise `Q<n> Answer`; `fte_rows()` helper; `save_meta(extra=)` |
| `smartview_mcp/answer.py` (modify) | `suffix` parameter: `Q<n> Answer` for the core, `Q<n> Svar` for the old tools |
| `smartview_mcp/core/source.py` (new) | Form-C source sheet: one-block write, provenance panel, sheet = fetch check, ad hoc fallback |
| `smartview_mcp/core/progress.py` (new) | Answer skeleton sheet and status line |
| `smartview_mcp/core/check.py` (new) | Findings Critical / Warning / Info from the build result and the sources |
| `smartview_mcp/core/runlog.py` (new) | Run log: JSONL (`config.RUN_LOG`) + hidden `_runlog` sheet |
| `smartview_mcp/core/pipeline.py` (new) | Orchestration of one question; rollback |
| `smartview_mcp/tools_core.py` (new) | MCP tool `sv_answer` |
| `smartview_mcp/config.py`, `server.py` (modify) | `RUN_LOG`; register `tools_core` |
| `tests/test_core_skeleton.py` (new) | Unit tests for all of the above |
| `tests/live/bench.py` (modify) | Route `core` with B02 |

---

### Task 1: `svhttp.fetch` returns the raw grid

**Files:** Modify `engine/smartview_mcp/svhttp.py:122-157`; Test `engine/tests/test_svhttp.py`

- [ ] **Step 1: Failing test** (append to `tests/test_svhttp.py`)

```python
def test_fetch_returns_the_raw_grid_a_source_sheet_is_written_from(db_path):
    f = _sv(FakeTransport(answer(GRID)), db_path).fetch(
        "FinMax", {"Entities": "un.SY.2WHS", "All Scenarios": "Actual"},
        [Axis("All Accounts", ["ac.C30", "ac.EBITDA"])], [Axis("All Time Periods", ["Aug 26", "Sep 26"])])
    assert [list(r[:3]) for r in f.values] == GRID
    assert f.layout.pov == [("All Entities", "un.SY.2WHS"), ("All Scenarios", "Actual")] and f.seconds >= 0
```

- [ ] **Step 2:** `uv run pytest tests/test_svhttp.py -q -k fetch_returns` → FAIL (`AttributeError: fetch`).
- [ ] **Step 3: Implement.** Add `from dataclasses import dataclass` and above `class SvHttp`:

```python
@dataclass
class Fetched:
    """A provider grid as fetched: Smart View's own shape from A1 (POV row, column headers, row labels, data)."""
    layout: Layout
    values: list[list]
    seconds: float
```

Rename the body of `retrieve` to `fetch(self, conn, pov, rows, cols) -> Fetched` ending with
`return Fetched(lay, values, time.monotonic() - t0)` (all validation, `_check_grid`, the POV and label checks stay),
and make `retrieve` a wrapper:

```python
    def retrieve(self, conn: str, pov: dict | None, rows: list[Axis], cols: list[Axis],
                 suppress_missing: bool, suppress_zeros: bool):
        f = self.fetch(conn, pov, rows, cols)
        return f.layout, parse(f.layout, f.values, suppress_missing, suppress_zeros), f.seconds
```

- [ ] **Step 4:** `uv run pytest tests/test_svhttp.py -q` → all PASS.
- [ ] **Step 5:** `git commit -am "svhttp.fetch: the raw provider grid for source sheets"`

### Task 2: `Q<n> Answer` sheets

**Files:** Modify `sheetorder.py:22,43`, `userexcel.py:954,967`, `answer.py:47-60,76-80,239-256,307`;
Test `tests/test_sheetorder.py`

- [ ] **Step 1: Failing tests** (append to `tests/test_sheetorder.py`)

```python
def test_answer_sheets_rank_like_svar_and_count_as_questions():
    from smartview_mcp.userexcel import _Q
    names = ["Q2 FinMax", "Q2 Answer", "Q1 Svar"]
    assert sheet_order(names, set(names)) == ["Q1 Svar", "Q2 Answer", "Q2 FinMax"]
    assert _Q.match("Q3 Answer").group(1) == "3"
```

- [ ] **Step 2:** `uv run pytest tests/test_sheetorder.py -q` → FAIL.
- [ ] **Step 3: Implement.**
  - `sheetorder._OURS` and `userexcel._Q`: `(Svar|Answer|FinMax|FinLite|FinPlan|Linjer)`;
    `sheetorder.classify`: `if kind in ("Svar", "Answer"):`; `userexcel.park_question_data`:
    `m.group(2) not in ("Svar", "Answer", "Linjer")`.
  - `answer.answer_sheet(t, q, suffix: str = "Svar")` with `name = f"Q{q} {suffix}"`;
    `build(t, q, spec, pulls, suffix: str = "Svar")` passes it to `_build(t, q, spec, pulls, suffix)`, which calls
    `answer_sheet(t, q, suffix)`; `_build`'s status line uses `f"formaterer 'Q{q} {suffix}' ..."`, and `_build`
    passes `suffix` on to `_write(t, q, spec, pulls, ws, suffix="Svar")` for its status line (not `ws.Name`: the
    test fakes have no Name).
- [ ] **Step 4:** `uv run pytest tests/test_sheetorder.py tests/test_answer_write.py -q` → PASS.
- [ ] **Step 5:** `git commit -am "Q<n> Answer sheets alongside Q<n> Svar"`

### Task 3: `userexcel.fte_rows` and `save_meta(extra=)`

**Files:** Modify `userexcel.py` (`save_meta` ~1055, `_pull` ~1222); Test `tests/test_core_skeleton.py` (new)

- [ ] **Step 1: Failing test** (`tests/test_core_skeleton.py`)

```python
"""The new core's walking skeleton (docs/plans/2026-11-02-walking-skeleton.md)."""
from types import SimpleNamespace

from smartview_mcp import userexcel
from smartview_mcp.grid import Axis, Layout


def test_fte_rows_are_the_sheet_rows_of_man_year_lines():
    lay = Layout([], [Axis("All Accounts", ["ac.C30", "ac.STA20Y"])], [Axis("All Time Periods", ["Sep 26"])])
    parsed = SimpleNamespace(row_labels=[["C30 REVENUES (ac)"], ["STA20Y Man-years (ac)"]])
    assert userexcel.fte_rows(lay, parsed) == [4]
```

- [ ] **Step 2:** `uv run pytest tests/test_core_skeleton.py -q` → FAIL.
- [ ] **Step 3: Implement** in `userexcel.py` (above `_make_pull`) and use it in `_pull` instead of the inline list:

```python
def fte_rows(lay: Layout, parsed) -> list[int]:
    """Sheet rows (1-based) of FTE / man-year lines: they are styled as decimals, not NOK."""
    return [2 + lay.nc + i for i, labs in enumerate(parsed.row_labels)
            if any(isinstance(x, str) and ("STA" in x or "man-years" in x.lower()) for x in labs)]
```

`save_meta(t, p, extra: dict | None = None)`: add `**(extra or {})` as the last entry of the JSON dict
(`load_pulls` ignores unknown keys).
- [ ] **Step 4:** `uv run pytest tests/test_core_skeleton.py -q` → PASS.
- [ ] **Step 5:** `git commit -am "userexcel: fte_rows helper, save_meta extra"`

### Task 4: `core/source.py` (pull, form C)

**Files:** Create `engine/smartview_mcp/core/source.py`; Test `tests/test_core_skeleton.py`

- [ ] **Step 1: Failing tests** (append)

```python
import datetime as dt

import pytest

from smartview_mcp import svhttp
from smartview_mcp.core import contract, source
from smartview_mcp.host import InputError

PIN = contract.Pull.model_validate({"id": "m", "connection": "FinMax",
                                    "rows": [{"dimension": "All Accounts", "members": ["ac.EBITDA"]}],
                                    "columns": [{"dimension": "All Time Periods", "members": ["Sep 26"]}]})


def test_cells_keep_text_as_text_and_blanks_empty():
    assert source.cells([[None, "Jan 26"], ["EBITDA (ac)", 1.5, "x"]]) == (("", "'Jan 26", ""), ("'EBITDA (ac)", 1.5, "'x"))


def test_provenance_names_cube_pov_time_and_route():
    lay = Layout([("All Entities", "un.SY.2BUS"), ("All Scenarios", "All Scenarios")], [], [])
    p = dict(source.provenance("FinMax", lay, "http", dt.datetime(2026, 11, 2, 10, 5)))
    assert p["Source"] == "Oracle EPM Cloud · FinMax" and p["POV"] == "un.SY.2BUS"
    assert p["Fetched"] == "2026-11-02 10:05" and "re-fetch" in p["Route"]


def test_mismatches_compare_only_the_fetched_numbers():
    assert source.mismatches([[None, "a"], ["x", 1.0]], ((None, "a"), ("x", 1.0, "panel"))) == []
    assert source.mismatches([["x", 1.0]], (("x", 2.0),)) == ["row 1 column 2: fetched 1.0, sheet shows 2.0"]


def _adhoc(monkeypatch, seen):
    pull = SimpleNamespace(sheet="Q1 FinMax")
    monkeypatch.setattr(source.userexcel, "pull", lambda h, t, q, conn, pov, rows, cols, title: seen.append(title) or pull)
    return pull


def test_an_http_failure_falls_back_to_the_adhoc_grid_with_a_note(monkeypatch):
    def fail(*a):
        raise svhttp.SvHttpError("token expired")
    monkeypatch.setattr(source.svhttp, "shared", lambda: SimpleNamespace(fetch=fail))
    seen = []
    pull = _adhoc(monkeypatch, seen)
    s = source.pull(None, None, 1, 0, PIN)
    assert s.pull is pull and s.route == "adhoc" and "token expired" in s.note and seen == ["Q1 FinMax"]


def test_an_input_error_is_not_hidden_by_the_fallback(monkeypatch):
    def bad(*a):
        raise InputError("Invalid query: unknown dimension")
    monkeypatch.setattr(source.svhttp, "shared", lambda: SimpleNamespace(fetch=bad))
    with pytest.raises(InputError):
        source.pull(None, None, 1, 0, PIN)


def test_http_success_writes_the_sheet_and_names_later_pulls_by_id(monkeypatch):
    f = svhttp.Fetched(Layout([], [], []), [], 0.1)
    monkeypatch.setattr(source.svhttp, "shared", lambda: SimpleNamespace(fetch=lambda *a: f))
    monkeypatch.setattr(source, "write", lambda t, q, title, pin, fetched: (title, fetched))
    assert source.pull(None, None, 2, 1, PIN) == ("Q2 FinMax m", f)
```

- [ ] **Step 2:** `uv run pytest tests/test_core_skeleton.py -q` → FAIL (no module `source`).
- [ ] **Step 3: Implement** `core/source.py`:

```python
"""pull, form C (design §3.2): fetch a grid over Smart View's HTTP provider and write it as a source sheet in ONE
block, with a provenance panel and the pull spec stored in _svmeta; a Smart View ad hoc grid in Excel (form A) when
the HTTP route fails. The grid starts at A1 in Smart View's own shape (POV row, column headers, row labels), so the
label-keyed lookups of answer.py work unchanged - which is why provenance is a panel to the RIGHT of the grid: a
header above it would move every label, and row 1 is the POV row."""
from __future__ import annotations

import datetime as dt
import math
from dataclasses import dataclass

from .. import config, svhttp, userexcel
from ..grid import Axis, Layout, parse
from ..timing import stage
from .contract import Pull as PullSpec

GREY = 0x808080
ROUTES = {"http": "HTTP value grid; refresh = ask smartass to fetch again",
          "adhoc": "Smart View ad hoc grid; refresh in Smart View"}


@dataclass
class Source:
    pull: userexcel.Pull
    route: str                          # http | adhoc
    mismatches: list[str]               # sheet cells that differ from the fetch (must be empty)
    note: str | None = None             # why the fallback was used


def cells(values: list[list]) -> tuple[tuple, ...]:
    """The fetched grid as Excel values: text stays text ('Jan 26' would become a date), blanks stay empty."""
    def one(v):
        if v is None or v == "":
            return ""
        return "'" + v if isinstance(v, str) else v
    w = max(len(r) for r in values)
    return tuple(tuple(one(r[c]) if c < len(r) else "" for c in range(w)) for r in values)


def provenance(connection: str, lay: Layout, route: str, fetched: dt.datetime) -> list[tuple[str, str]]:
    pov = " · ".join(m for d, m in lay.pov if m != d) or "top members"
    return [("Source", f"Oracle EPM Cloud · {connection}"), ("POV", pov),
            ("Fetched", fetched.strftime("%Y-%m-%d %H:%M")), ("Route", ROUTES[route])]


def mismatches(values: list[list], read, limit: int = 5) -> list[str]:
    """Every fetched number must be on the sheet unchanged (design §3.4: sheet value = fetched value)."""
    out = []
    for r, row in enumerate(values):
        for c, v in enumerate(row):
            if isinstance(v, bool) or not isinstance(v, (int, float)):
                continue
            got = read[r][c] if r < len(read) and c < len(read[r]) else None
            if isinstance(got, bool) or not isinstance(got, (int, float)) or abs(got - v) > 1e-6 * max(1.0, abs(v)):
                out.append(f"row {r + 1} column {c + 1}: fetched {v!r}, sheet shows {got!r}")
                if len(out) == limit:
                    break
    return out


def sheet_title(q: int, i: int, pin: PullSpec) -> str:
    return pin.title or f"Q{q} {pin.connection}" + (f" {pin.id}" if i else "")


def write(t, q: int, title: str, pin: PullSpec, f: svhttp.Fetched) -> Source:
    from .. import styling
    from ..ops import _used_values
    com, lay = userexcel.com, f.layout
    name = userexcel.unique_sheet_name(t, title)
    ws = userexcel.add_sheet(t, name, after_q=q, color=0xA6A6A6)
    grid = cells(f.values)
    h, w = len(grid), len(grid[0])
    parsed = parse(lay, f.values, False, False)
    p = userexcel._make_pull(name, pin.connection, lay, parsed, f.seconds)
    now = dt.datetime.now()
    with userexcel.fast_excel(t.app):
        com(lambda: setattr(ws.Range(ws.Cells(1, 1), ws.Cells(h, w)), "Value", grid))
        styling.apply_data_sheet(ws, header_rows=1 + lay.nc, label_cols=lay.nr, fte_rows=userexcel.fte_rows(lay, parsed))
        panel = provenance(pin.connection, lay, "http", now)
        prng = ws.Range(ws.Cells(2, w + 2), ws.Cells(1 + len(panel), w + 3))
        com(lambda: setattr(prng, "Value", tuple((k, "'" + v) for k, v in panel)))
        com(lambda: setattr(prng.Font, "Color", GREY))
        com(lambda: setattr(prng.Font, "Italic", True))
        userexcel.save_meta(t, p, extra={"route": "http", "fetched": now.isoformat(timespec="seconds"),
                                         "pull": pin.model_dump()})
    return Source(p, "http", mismatches(f.values, com(lambda: _used_values(ws))))


def pull(h, t, q: int, i: int, pin: PullSpec) -> Source:
    """Pull i of question q: HTTP first, the ad hoc grid if the HTTP route fails or the grid is too big for one
    request. A wrong query (InputError) is never retried through Excel."""
    title = sheet_title(q, i, pin)
    rows = [Axis(a.dimension, a.members) for a in pin.rows]
    cols = [Axis(a.dimension, a.members) for a in pin.columns]
    n = math.prod(len(a.members) for a in rows + cols)
    if n / 200 > config.SVHTTP_TIMEOUT_S:
        note = f"{n} cells is too big for one HTTP request; fetched through Smart View in Excel."
    else:
        try:
            with stage("fetch (HTTP)"):
                f = svhttp.shared().fetch(pin.connection, dict(pin.pov), rows, cols)
        except svhttp.SvHttpError as e:
            note = f"HTTP route failed ({e}); fetched through Smart View in Excel instead."
        else:
            with stage("write source sheet"):
                return write(t, q, title, pin, f)
    with stage("pull (Smart View ad hoc)"):
        p = userexcel.pull(h, t, q, pin.connection, dict(pin.pov), rows, cols, title=title)
    return Source(p, "adhoc", [], note)
```

- [ ] **Step 4:** `uv run pytest tests/test_core_skeleton.py -q` → PASS.
- [ ] **Step 5:** `git add smartview_mcp/core/source.py tests/test_core_skeleton.py; git commit -m "core.source: form-C source sheets"`

### Task 5: `core/progress.py` (answer skeleton first)

**Files:** Create `engine/smartview_mcp/core/progress.py`; Test `tests/test_core_skeleton.py`

- [ ] **Step 1: Failing test** (append; `SPEC` is reused by later tasks)

```python
from smartview_mcp.core import progress

SPEC = {"version": 1, "question": "EBITDA Sep 26 for Business", "gate": {"outcome": "clear"},
        "assumptions": ["Actual, NOK"],
        "pulls": [{"id": "m", "connection": "FinMax",
                   "rows": [{"dimension": "All Accounts", "members": ["ac.EBITDA"]}],
                   "columns": [{"dimension": "All Time Periods", "members": ["Sep 26"]}]}],
        "layout": {"title": "Business — EBITDA", "columns": [{"header": "Sep 26"}],
                   "sections": [{"lines": [{"key": "e", "label": "EBITDA",
                                            "cells": [{"pull": "m", "row": ["ac.EBITDA"], "col": ["Sep 26"]}]}]}]}}


def test_skeleton_shows_title_question_status_and_assumptions():
    rows = progress.skeleton(contract.validate(SPEC), "fetching 1 source grid ...")
    assert rows[0] == "Business — EBITDA" and rows[1] == "Question: EBITDA Sep 26 for Business"
    assert rows[progress.STATUS_ROW - 1] == "Status: fetching 1 source grid ..."
    assert rows[-2:] == ["Assumptions", "•  Actual, NOK"]


def test_skeleton_of_a_proxy_answer_says_what_is_answered_instead():
    spec = contract.validate({**SPEC, "gate": {"outcome": "infeasible", "reason": "no product dimension",
                                               "proxy": "contribution margin per segment"}})
    assert progress.skeleton(spec, "x")[-1] == "Answered instead: contribution margin per segment"
```

- [ ] **Step 2:** run → FAIL.
- [ ] **Step 3: Implement** `core/progress.py`:

```python
"""The answer sheet appears first (design §3.3): title, question, status line and assumptions, before any pull.
answer.build later clears it and renders the full answer in one block."""
from __future__ import annotations

from .. import answer as answer_mod
from ..userexcel import com
from .contract import AnswerSpec

ANSWER = "Answer"            # sheet 'Q<n> Answer'
STATUS_ROW = 3


def skeleton(spec: AnswerSpec, status: str) -> list[str]:
    rows = [spec.layout.title if spec.layout else spec.question, f"Question: {spec.question}", f"Status: {status}"]
    if spec.assumptions:
        rows += ["", "Assumptions"] + [f"•  {a}" for a in spec.assumptions]
    if spec.gate.outcome == "infeasible":
        rows += ["", f"Not answerable as asked: {spec.gate.reason}", f"Answered instead: {spec.gate.proxy}"]
    return rows


def show(t, q: int, spec: AnswerSpec, status: str) -> str:
    ws = answer_mod.answer_sheet(t, q, suffix=ANSWER)
    rows = skeleton(spec, status)
    rng = ws.Range(ws.Cells(1, 2), ws.Cells(len(rows), 2))
    com(lambda: setattr(rng, "Value", tuple((("'" + r) if r else "",) for r in rows)))
    com(lambda: setattr(ws.Cells(1, 2).Font, "Bold", True))
    com(lambda: setattr(ws.Cells(1, 2).Font, "Size", 14))
    com(lambda: setattr(ws.Cells(STATUS_ROW, 2).Font, "Italic", True))
    com(t.wb.Activate)
    com(ws.Activate)
    t.status(status)
    return ws.Name


def set_status(t, q: int, text: str) -> None:
    ws = com(lambda: t.wb.Worksheets(f"Q{q} {ANSWER}"))
    com(lambda: setattr(ws.Cells(STATUS_ROW, 2), "Value", "'Status: " + text))
    t.status(text)
```

- [ ] **Step 4:** run → PASS. **Step 5:** commit `"core.progress: answer skeleton sheet"`.

### Task 6: `core/check.py`

**Files:** Create `engine/smartview_mcp/core/check.py`; Test `tests/test_core_skeleton.py`

- [ ] **Step 1: Failing tests** (append)

```python
from smartview_mcp.core import check


def _src(mismatches=(), note=None, pov=(), unknown=()):
    parsed = SimpleNamespace(pov_problems=list(pov), unknown_members=list(unknown))
    return SimpleNamespace(pull=SimpleNamespace(sheet="Q1 FinMax", parsed=parsed), mismatches=list(mismatches),
                           note=note, route="http")


def test_all_green_is_ok_without_findings():
    assert check.review({"checks": {"Sum = YTD": "OK"}, "warnings": [], "lint": []}, [_src()]) == \
        {"ok": True, "findings": []}


def test_a_failed_check_or_a_sheet_that_differs_from_the_fetch_is_critical():
    r = check.review({"checks": {"Sum = YTD": "DIFF"}}, [_src(mismatches=["row 3 column 2: fetched 1.0, sheet shows 2.0"])])
    assert not r["ok"] and [f["severity"] for f in r["findings"]] == ["Critical", "Critical"]


def test_fallback_rejected_members_and_missing_checks_are_warnings():
    r = check.review({"checks": {}}, [_src(note="HTTP route failed (x)", pov=["Smart View rejected y"])])
    assert r["ok"] and [f["severity"] for f in r["findings"]] == ["Warning"] * 3


def test_formula_errors_are_critical_and_style_lint_is_info():
    r = check.review({"checks": {"c": "OK"}, "formula_errors": ["ytd"], "lint": ["narrow column"]}, [])
    assert [(f["severity"], f["text"]) for f in r["findings"]] == [
        ("Critical", "formula error on line 'ytd'"), ("Info", "narrow column")]
```

- [ ] **Step 2:** run → FAIL. **Step 3: Implement** `core/check.py`:

```python
"""check (design §3.4), skeleton: reconciliation checks, formula errors and sheet = fetch, as a findings table
Critical / Warning / Info (structure after Anthropic financial-services audit-xls, Apache-2.0). Critical = a number
on the answer cannot be trusted; Warning = trustworthy but the user must know; Info = cosmetic."""
from __future__ import annotations

ORDER = {"Critical": 0, "Warning": 1, "Info": 2}


def review(res: dict, sources: list) -> dict:
    f: list[tuple[str, str]] = []
    checks = res.get("checks") or {}
    f += [("Critical", f"check '{k}' is {v}") for k, v in checks.items() if v != "OK"]
    f += [("Critical", w) for w in (res.get("error_warnings")
                                    or [f"formula error on line '{k}'" for k in res.get("formula_errors") or []])]
    for s in sources:
        f += [("Critical", f"{s.pull.sheet}: the sheet differs from the fetch ({m})") for m in s.mismatches]
        if s.note:
            f.append(("Warning", f"{s.pull.sheet}: {s.note}"))
        f += [("Warning", f"{s.pull.sheet}: {w}") for w in s.pull.parsed.pov_problems]
        if s.pull.parsed.unknown_members:
            f.append(("Warning", f"{s.pull.sheet}: unknown members {', '.join(s.pull.parsed.unknown_members)}"))
    if not checks:
        f.append(("Warning", "the answer has no reconciliation checks"))
    f += [("Warning", f"formula style: {b.get('message')} ({b.get('cell')})")
          for b in (res.get("formula_style") or {}).get("block") or []]
    f += [("Warning", str(w)) for w in res.get("warnings") or []]
    f += [("Info", str(x)) for x in res.get("lint") or []]
    f.sort(key=lambda x: ORDER[x[0]])
    return {"ok": not any(s == "Critical" for s, _ in f), "findings": [{"severity": s, "text": t} for s, t in f]}
```

- [ ] **Step 4:** run → PASS. **Step 5:** commit `"core.check: findings table"`.

### Task 7: `core/runlog.py` + `config.RUN_LOG`

**Files:** Create `engine/smartview_mcp/core/runlog.py`; Modify `config.py:138`; Test `tests/test_core_skeleton.py`

- [ ] **Step 1: Failing tests** (append)

```python
import json

from smartview_mcp import config
from smartview_mcp.core import runlog


def test_entry_records_gate_pulls_routes_and_result(monkeypatch, tmp_path):
    spec = contract.validate(SPEC)
    src = SimpleNamespace(pull=SimpleNamespace(sheet="Q1 FinMax", seconds=0.234), route="http")
    e = runlog.entry(spec, 1, [src], {"ok": True, "findings": []}, {"_total": 3.1}, workbook="C:/x/a.xlsx")
    assert e["pulls"] == [{"id": "m", "cube": "FinMax", "sheet": "Q1 FinMax", "route": "http", "seconds": 0.23}]
    assert (e["q"], e["gate"], e["ok"], e["error"]) == ("Q1", "clear", True, None)
    row = runlog.sheet_row(e)
    assert row[1:6] == ("Q1", "EBITDA Sep 26 for Business", "clear", "m=FinMax (http, 0.23 s)", "OK")
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "logs" / "runs.jsonl")
    runlog.append(e)
    assert json.loads((tmp_path / "logs" / "runs.jsonl").read_text(encoding="utf-8"))["q"] == "Q1"
```

- [ ] **Step 2:** run → FAIL. **Step 3: Implement.** `config.py` after `QUERY_LOG`:

```python
RUN_LOG = _path(_pick("SV_RUN_LOG", "paths.run_log", HOME / "logs" / "runs.jsonl"))   # one line per answer (core)
```

`core/runlog.py`:

```python
"""Run log per answer (design recommendation 6): one JSON line per run in config.RUN_LOG and one row in the
workbook's hidden '_runlog' sheet - question, gate, assumptions, pulls (cube, route, seconds), timings, result."""
from __future__ import annotations

import datetime as dt
import json

from .. import config, userexcel

SHEET = "_runlog"
HEADER = ("When", "Question no", "Question", "Gate", "Pulls", "Result", "Details")


def entry(spec, q: int | None, sources: list, review: dict | None, timings: dict, error: str | None = None,
          workbook: str | None = None) -> dict:
    return {"ts": dt.datetime.now().isoformat(timespec="seconds"), "workbook": workbook,
            "q": f"Q{q}" if q else None, "question": spec.question, "gate": spec.gate.outcome,
            "assumptions": list(spec.assumptions),
            "pulls": [{"id": pin.id, "cube": pin.connection, "sheet": s.pull.sheet, "route": s.route,
                       "seconds": round(s.pull.seconds, 2)} for pin, s in zip(spec.pulls, sources)],
            "ok": None if review is None else review["ok"],
            "findings": [] if review is None else [f"{x['severity']}: {x['text']}" for x in review["findings"]],
            "timings": timings, "error": error}


def sheet_row(e: dict) -> tuple:
    result = "error" if e["error"] else {True: "OK", False: "check findings"}.get(e["ok"], e["gate"])
    pulls = "; ".join(f"{p['id']}={p['cube']} ({p['route']}, {p['seconds']} s)" for p in e["pulls"])
    return (e["ts"], e["q"] or "", e["question"], e["gate"], pulls, result,
            json.dumps(e, ensure_ascii=False, default=str)[:30000])


def append(e: dict) -> None:
    try:
        config.RUN_LOG.parent.mkdir(parents=True, exist_ok=True)
        with open(config.RUN_LOG, "a", encoding="utf-8") as f:
            f.write(json.dumps(e, ensure_ascii=False, default=str) + "\n")
    except OSError:
        pass                      # the log must never break an answer (OneDrive lock etc.)


def write_sheet(t, e: dict) -> None:
    com = userexcel.com
    if not userexcel._sheet_exists(t, SHEET):
        ws = com(lambda: t.wb.Worksheets.Add(None, t.wb.Worksheets(t.wb.Worksheets.Count)))
        com(lambda: setattr(ws, "Name", SHEET))
        com(lambda: setattr(ws, "Visible", 0))                 # xlSheetHidden
        com(lambda: setattr(ws.Range("A1:G1"), "Value", (HEADER,)))
    ws = com(lambda: t.wb.Worksheets(SHEET))
    r = com(lambda: ws.Cells(ws.Rows.Count, 1).End(-4162).Row) + 1
    com(lambda: setattr(ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)), "Value",
                        (tuple("'" + str(x) for x in sheet_row(e)),)))
```

- [ ] **Step 4:** run → PASS. **Step 5:** commit `"core.runlog: JSONL + hidden sheet"`.

### Task 8: `core/pipeline.py` + tool `sv_answer`

**Files:** Create `core/pipeline.py`, `tools_core.py`; Modify `server.py:269`, `tests/test_smoke_stdio.py:6`;
Test `tests/test_core_skeleton.py`

- [ ] **Step 1: Failing tests** (append)

```python
from mcp.server.mcpserver.exceptions import ToolError

from smartview_mcp import tools_core
from smartview_mcp.core import pipeline
from smartview_mcp.host import HostError


def test_assumptions_and_the_question_lead_the_notes_and_the_raw_layout_is_untouched():
    raw = {**SPEC, "layout": {**SPEC["layout"], "notes": ["n"]}}
    out = pipeline.answer_layout(raw["layout"], contract.validate(raw))
    assert out["notes"] == ["Question: EBITDA Sep 26 for Business", "Assumption: Actual, NOK", "n"]
    assert raw["layout"]["notes"] == ["n"]


def _fake_run(monkeypatch, calls, pull_fails=False):
    t = SimpleNamespace(path="C:/x/a.xlsx")
    pull = SimpleNamespace(sheet="Q3 FinMax", seconds=0.2, parsed=SimpleNamespace(pov_problems=[], unknown_members=[]))
    monkeypatch.setattr(pipeline.userexcel, "resolve_target", lambda h, target, allow, title: t)
    monkeypatch.setattr(pipeline.userexcel, "next_question", lambda t: 3)
    monkeypatch.setattr(pipeline.userexcel, "rollback", lambda t: ["Q3 Answer"])
    monkeypatch.setattr(pipeline.userexcel, "save_spec", lambda t, q, raw, ids: calls.append(("save_spec", q)))
    monkeypatch.setattr(pipeline.progress, "show", lambda t, q, spec, status: calls.append(("show", q)))
    monkeypatch.setattr(pipeline.progress, "set_status", lambda t, q, text: calls.append(("status", text)))

    def pull_(h, t, q, i, pin):
        calls.append(("pull", pin.id))
        if pull_fails:
            raise HostError("Retrieve failed")
        return source.Source(pull, "http", [])
    monkeypatch.setattr(pipeline.source, "pull", pull_)
    monkeypatch.setattr(pipeline.specs, "prepare", lambda layout, ids, sess: calls.append(("prepare", ids)) or layout)
    monkeypatch.setattr(pipeline.answer_mod, "build", lambda t, q, spec, sess, suffix: calls.append(("build", suffix))
                        or {"sheet": "Q3 Answer", "checks": {"c": "OK"}})


def test_run_shows_the_answer_sheet_first_then_pulls_then_renders(monkeypatch):
    calls = []
    _fake_run(monkeypatch, calls)
    r = pipeline.run(None, SPEC)
    assert [c[0] for c in calls] == ["show", "status", "pull", "status", "prepare", "build", "save_spec"]
    assert ("prepare", {"m": "Q3 FinMax"}) in calls and ("build", "Answer") in calls
    assert r.q == 3 and r.review == {"ok": True, "findings": []}


def test_a_failed_pull_removes_the_partial_sheets(monkeypatch):
    calls = []
    _fake_run(monkeypatch, calls, pull_fails=True)
    with pytest.raises(HostError, match="removed the partial sheets: Q3 Answer"):
        pipeline.run(None, SPEC)


@pytest.mark.anyio
async def test_an_ambiguous_question_returns_its_questions_and_touches_no_excel(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")
    async def no_excel(*a, **k):
        raise AssertionError("Excel was touched")
    monkeypatch.setattr(tools_core, "_call_user", no_excel)
    spec = {"version": 1, "question": "How are we doing on costs?",
            "gate": {"outcome": "ambiguous", "questions": ["Which unit?", "Which period?"]}}
    out = await tools_core.sv_answer(spec)
    assert out["gate"] == "ambiguous" and out["questions"] == ["Which unit?", "Which period?"]
    assert json.loads((tmp_path / "runs.jsonl").read_text(encoding="utf-8"))["gate"] == "ambiguous"


@pytest.mark.anyio
async def test_an_invalid_spec_is_a_readable_error():
    with pytest.raises(ToolError, match="not a valid answer spec v1"):
        await tools_core.sv_answer({"question": "x"})
```

- [ ] **Step 2:** run → FAIL. **Step 3: Implement** `core/pipeline.py`:

```python
"""One question through the new core (design §3): scope -> pull -> answer -> check. The scope gate's verdict
arrives in the answer spec (the agent applies the checklist); ambiguous questions stop in tools_core before any
Excel work. Walking skeleton: new questions only (follow-ups are phase 5)."""
from __future__ import annotations

import copy
from dataclasses import dataclass

from pydantic import ValidationError

from .. import answer as answer_mod
from .. import specs, userexcel
from ..host import InputError, rewrap
from ..timing import stage
from . import check, contract, progress, source


@dataclass
class Run:
    t: object
    q: int
    spec: contract.AnswerSpec
    prepared: dict
    res: dict
    sources: list
    review: dict


def parse_spec(raw: dict) -> contract.AnswerSpec:
    try:
        return contract.validate(raw)
    except ValidationError as e:
        raise InputError("This is not a valid answer spec v1: " + "; ".join(
            f"{'.'.join(map(str, x['loc'])) or 'spec'}: {x['msg']}" for x in e.errors()[:8])) from None


def gate_reply(spec: contract.AnswerSpec) -> dict | None:
    if spec.gate.outcome != "ambiguous":
        return None
    return {"gate": "ambiguous", "questions": spec.gate.questions, "checklist": spec.gate.checklist,
            "note": "Nothing was pulled. Ask the user these questions in ONE message, then call sv_answer again "
                    "with the answers applied (gate 'clear', assumptions stated)."}


def answer_layout(raw_layout: dict, spec: contract.AnswerSpec) -> dict:
    """The layout answer.build renders, with the question, assumptions and (for a proxy) what is answered instead
    leading the notes. The raw layout is not changed: it is what save_spec stores."""
    lay = copy.deepcopy(raw_layout)
    lead = [f"Question: {spec.question}"] + [f"Assumption: {a}" for a in spec.assumptions]
    if spec.gate.outcome == "infeasible":
        lead.append(f"Not answerable as asked ({spec.gate.reason}); answered instead: {spec.gate.proxy}")
    lay["notes"] = lead + list(lay.get("notes") or [])
    return lay


def run(h, raw: dict, target: str = "session", allow_archive: bool = False) -> Run:
    spec = parse_spec(raw)
    if spec.gate.outcome == "ambiguous":
        raise InputError("An ambiguous question pulls nothing: ask its questions first.")
    with stage("open target"):
        t = userexcel.resolve_target(h, target, allow_archive, title=spec.question[:50])
        q = userexcel.next_question(t)
    sources: list[source.Source] = []
    try:
        with stage("answer skeleton"):
            progress.show(t, q, spec, f"fetching {len(spec.pulls)} source grid(s) ...")
        for i, pin in enumerate(spec.pulls):
            progress.set_status(t, q, f"fetching {pin.connection} '{pin.id}' ({i + 1}/{len(spec.pulls)}) ...")
            sources.append(source.pull(h, t, q, i, pin))
        progress.set_status(t, q, "building the answer ...")
        ids = {pin.id: s.pull.sheet for pin, s in zip(spec.pulls, sources)}
        sess = {s.pull.sheet: s.pull for s in sources}
        with stage("prepare spec"):
            prepared = specs.prepare(answer_layout(raw["layout"], spec), ids, sess)
        with stage("answer"):
            res = answer_mod.build(t, q, prepared, sess, suffix=progress.ANSWER)
    except Exception as e:
        removed = userexcel.rollback(t)
        raise rewrap(e, f"{e} (the answer was not completed; removed the partial sheets: "
                     f"{', '.join(removed) or 'none'})") from None
    review = check.review(res, sources)
    try:
        userexcel.save_spec(t, q, raw, ids)
    except Exception:   # noqa: BLE001 - only needed for follow-ups
        pass
    return Run(t, q, spec, prepared, res, sources, review)
```

`tools_core.py`:

```python
"""sv_answer: one question through the new pipeline (design 2026-11-02 §3) - the walking skeleton beside the old
tools, which stay until the benchmarks reach parity."""
from __future__ import annotations

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError

from . import tools_ws, userexcel
from .core import pipeline, runlog
from .host import HostError, InputError
from .server import WRITE, _call_user, _err, _is_our_analysis_file, tool
from .timing import Timer, stage


@tool(annotations=WRITE)
async def sv_answer(spec: dict, target: str = "session", allow_archive: bool = False,
                    ctx: Context | None = None) -> dict:
    """NEW PIPELINE (walking skeleton): answer ONE question from an answer spec v1 (core/contract.py):
    {version: 1, question, gate: {outcome: clear|ambiguous|infeasible, checklist: {period, scenario, entity, measure,
    currency_sign, granularity}, questions (ambiguous, 1-3), reason + proxy (infeasible)}, assumptions: [...],
    pulls: [{id, connection, rows, columns, pov, title?}], layout: {title, columns, sections, checks, ...} (the
    sv_ws_build spec grammar; pull cells use the pull ids)}.
    ambiguous -> returns the questions, pulls nothing. Otherwise: the answer sheet 'Q<n> Answer' appears first with
    the question, assumptions and a status line; each pull is fetched over HTTP and written as a source sheet in one
    block (Smart View ad hoc grid if HTTP fails); then the answer is rendered (live formulas) and checked.
    Returns the answer table, checks, review findings (Critical/Warning/Info), routes and timings."""
    try:
        parsed = pipeline.parse_spec(spec)
    except InputError as e:
        raise _err(e) from None
    early = pipeline.gate_reply(parsed)
    if early is not None:
        runlog.append(runlog.entry(parsed, None, [], None, {}))
        return early

    def job(h):
        tm = Timer.start()
        try:
            r = pipeline.run(h, spec, target, allow_archive)
        except Exception as e:
            runlog.append(runlog.entry(parsed, None, [], None, tm.summary(), error=str(e)[:500]))
            raise
        r.t.created = r.t.created or _is_our_analysis_file(r.t.path)
        with stage("log sources"):
            for s in r.sources:
                try:
                    userexcel.log_source(r.t, r.q, parsed.question, s.pull, notes=s.note or s.route)
                except Exception as e:   # noqa: BLE001 - the source log must never undo a finished answer
                    r.res["warnings"].append(f"could not log {s.pull.sheet} in 'Kilder': {e}")
        e = runlog.entry(parsed, r.q, r.sources, r.review, tm.summary(), workbook=r.t.path)
        try:
            runlog.write_sheet(r.t, e)
        except Exception as ex:   # noqa: BLE001
            r.res["warnings"].append(f"run log sheet not written: {ex}")
        runlog.append(e)
        with stage("save + activate"):
            tools_ws._finish(r.t, r.q, r.prepared, r.res, saved_new=r.t.created)
        out = tools_ws._result(r.t, r.q, r.prepared, r.res, [s.pull for s in r.sources], [])
        return {**out, "gate": parsed.gate.outcome, "routes": {s.pull.sheet: s.route for s in r.sources},
                "review": r.review, "timings": tm.summary()}
    try:
        return await _call_user(job, what="sv_answer")
    except (HostError, ToolError) as e:
        raise _err(e)
```

`server.py`: add `from . import tools_core  # noqa: E402,F401` after `tools_xl`. `tests/test_smoke_stdio.py`:
`len(names) == 33` and add `"sv_answer"` to the required set.
- [ ] **Step 4:** `uv run pytest tests/test_core_skeleton.py tests/test_smoke_stdio.py -q` → PASS.
- [ ] **Step 5:** full suite `uv run pytest -q` + `uv run ruff check smartview_mcp tests` → green; commit
  `"core.pipeline + sv_answer: B02 walking skeleton"`.

### Task 9: Benchmark route `core`, live run, docs

**Files:** Modify `tests/live/bench.py`, `tests/live/README.md`, design spec §3.2 + §3.3.

- [ ] **Step 1:** `bench.py`: `Case.core: Callable | None = None`; decorator

```python
def core_route(id: str):
    def deco(fn):
        CASES[id].core = fn
        return fn
    return deco
```

`run_case`: `fn = c.baseline if route == "baseline" else c.core`; when `fn is None` →
`UNSUPPORTED` with note `f"no {route} route yet"`; call `fn(ctx)`. `--route` choices `["baseline", "core"]`;
default `--out` = `f"bench-{route}"`. Add after `b02`:

```python
def b02_v1() -> dict:
    pulls = S._bus_pulls()
    pulls[1].pop("title", None)                     # sheet names come from the pipeline
    return {"version": 1, "question": CASES["B02"].question,
            "gate": {"outcome": "clear", "checklist": {
                "period": "Jan-Dec 2026 and 2025 (closed months), YTD Sep", "scenario": "Actual",
                "entity": "un.SY.2BUS Business Norway", "measure": "EBITDA (ac.EBITDA)",
                "currency_sign": "NOK, FinMax signs", "granularity": "month (Periodic) + Accumulated YTD"}},
            "assumptions": ["Actual, NOK; months are Periodic; the reconciliation uses Accumulated YTD Sep"],
            "pulls": pulls, "layout": S._bus_spec()}


@core_route("B02")
def b02_core(ctx):
    out, _ = ctx.client.must("sv_answer", {"spec": b02_v1(), "target": "new"})
    ctx.state.update(b02_wb=out["workbook"], b02_q=int(str(out["question"]).lstrip("Q")))
    problems = [f"{x['severity']}: {x['text']}" for x in out["review"]["findings"] if x["severity"] == "Critical"]
    return _run(out, problems=problems, notes=f"routes {out['routes']}")
```

- [ ] **Step 2:** unit test in `tests/test_matrix_logic.py`: `bench.b02_v1()` passes `contract.validate`.
- [ ] **Step 3 (live; needs the HTTP route free):** `uv run python tests\live\bench.py --only B02 --route core
  --out bench-core-b02` → B02 PASS, routes all `http`; compare first-visible / total with the baseline B02.
- [ ] **Step 4:** docs — design §3.2: provenance is a panel right of the grid (why); §3.3: skeleton implemented in
  `core/progress.py`; `tests/live/README.md`: route `core`. Commit + push.
