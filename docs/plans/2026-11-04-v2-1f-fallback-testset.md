# smartass v2 · Plan 1F: ad hoc fallback, recorder and the test set

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:**
- **Ad hoc fallback.** When the Smart View HTTP route fails for a grid, v2 fetches that grid as a Smart View ad hoc
  grid in the user's Excel, turns it into facts and removes the ad hoc sheet. The answer is built as usual, with
  the route reported.
- **Recorder and replay.** Every fetch and drill of a run can be recorded locally and replayed without EPM.
- **Test set.** About 8 benchmark questions as v2 specs. `record` runs them live and stores facts plus candidate
  key values, `approve` freezes the candidates as golden numbers after Tom's review, and `run` replays offline and
  compares within ±1 NOK.
- **Three small items from 1E:** text-fix candidates listed in the reply, long notes no longer spill, and several
  lines blocks drilled at once.

**Architecture:**
- **One parser.** `data/provider.py` gets one pure function, `facts_from_parsed(req, parsed, labels)`, used by both
  the HTTP route and the ad hoc route.
- **Failures are recorded, not raised.** `prepare` records failed grid requests (`Prepared.failed`) instead of
  raising.
- **The fallback runs as its own job.** The server runs one Excel job, `adhoc_fetch`, for the failed requests, then
  prepares again with a `ComboProvider` (fixed results for the failed requests, live for the rest).
- **Recorder, replay and runner.** `data/recorder.py` (`RecordingProvider`, `ReplayProvider`) and
  `smartass/testset/` (cases as JSON in git, a CLI).
- **No data in git.** Recordings, candidates and goldens stay under `%LOCALAPPDATA%\smartass\v2\testset\`.

**Conventions:** as Plan 1E.
- Branch `v2/1f` from `main`. Git Bash from `/c/dev/smartass/engine`.
- Tests go in `engine/tests/v2/test_v2_*.py`.
- Never run python through a stdin heredoc. Bare python only with `< /dev/null`.
- Stage files explicitly.
- No live Excel, EPM, browser or drill. Task 8 waits for Tom.
- Never edit `engine/smartview_mcp` or `knowledge/`.
- Commit messages end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- Read the current v2 code before each task; follow the code where a snippet differs and say so.

**Facts from v1:**
- `excel/adhoc.pull(h, t, q, connection, pov, rows, cols, title=None) -> sheets.Pull`. `Pull` has `.sheet`,
  `.layout` and `.parsed` (the same `grid.parse` output as the HTTP route, with `columns` and `rows`).
- `pull` runs in a user-Excel job and leaves a sheet `Q<q> <connection>` behind. Remove that sheet with
  `sheets.drop_parked`, as Plan 1C does.
- v1 also chooses ad hoc when a grid is too big for HTTP: `core/source.too_big` (cells / 200 >
  `config.SVHTTP_TIMEOUT_S`).

---

### Task 0: Branch

- [ ] `cd /c/dev/smartass && git switch main && git switch -c v2/1f`

---

### Task 1: One parser, and failures recorded instead of raised

**Files:** `engine/smartass/data/provider.py`, `engine/smartass/pipeline.py`. Test:
`engine/tests/v2/test_v2_failed_grids.py`.

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_failed_grids.py
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass.data.provider import facts_from_parsed
from smartass.model.requests import GridRequest
from smartass.pipeline import prepare
from smartview_mcp.svhttp import SvHttpError

REQ = GridRequest("fm", "FinMax", ("ac.C30",), ("E1",), ("Sep 26",), "Actual", "", "Periodic")


class Parsed:
    columns = ["All Accounts", "All Entities", "Sep 26"]
    rows = [["ac.C30", "E1", 5.0], ["ac.C30", "E2", None]]
    unknown_members = []


def test_facts_from_parsed_is_shared():
    out, warnings = facts_from_parsed(REQ, Parsed(), {"Sep 26": "Sep 26"})
    assert [(f.entity, f.value) for f in out] == [("E1", 5.0)] and warnings == []


class FailingProvider(FakeProvider):
    def fetch(self, req):
        if req.periods == ("Sep 25",):
            raise SvHttpError("Smart View provider at x: lost the request (TimeoutError)")
        return super().fetch(req)


def test_a_failed_grid_is_recorded_not_raised():
    pre = prepare(spec_dict(), resolver(), FailingProvider(facts()))
    assert [r.periods for r in pre.failed] == [("Sep 25",)]
    assert pre.failed_errors and "lost the request" in pre.failed_errors[0]
```

- [ ] **Step 2:** Run it and see it fail with `ImportError: facts_from_parsed`.

- [ ] **Step 3: Implement**
  - **`provider.py`.** Extract the row loop of `fetch` into
    `facts_from_parsed(req, parsed, labels) -> (list[Fact], list[str])`, which collects the unknown-member and
    error-cell warnings. `fetch` calls it per grid.
  - **`pipeline.py`.**
    - `Prepared` gets `failed: list[GridRequest] = field(default_factory=list)` and
      `failed_errors: list[str] = field(default_factory=list)`.
    - In `prepare`, the parallel fetch runs a wrapper. It catches `SvHttpError`, which the HTTP route raises, and
      returns `None` for that request plus the error text. It never catches `InputError`, because a wrong member
      is wrong on every route.
    - Failed requests go to `failed`.
    - Compute still runs. A missing comparison is `None`, which is honest.
    - The reply's `warnings` say "N grid(s) could not be fetched over HTTP".

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): one grid parser; failed HTTP grids recorded instead of raised`.

---

### Task 2: The ad hoc fetch job

**Files:** Create `engine/smartass/data/adhoc.py`. Test: `engine/tests/v2/test_v2_adhoc.py`.

- [ ] **Step 1: Write the failing test** with a fake `pull`. This is COM glue, so the test checks the calls and the
  facts.

```python
# engine/tests/v2/test_v2_adhoc.py
from smartass.data import adhoc
from smartass.model.requests import GridRequest

REQ = GridRequest("fm", "FinMax", ("ac.C30",), ("E1",), ("Sep 26",), "Actual", "", "Periodic")


class Parsed:
    columns = ["All Accounts", "All Entities", "Sep 26"]
    rows = [["ac.C30", "E1", 7.0]]
    unknown_members = []


class Pull:
    sheet = "Q4 FinMax"
    parsed = Parsed()


def test_adhoc_fetch_pulls_parses_and_removes_the_sheet(monkeypatch):
    log = []
    monkeypatch.setattr(adhoc, "_target", lambda h, target: "T")
    monkeypatch.setattr(adhoc, "_next_q", lambda t: 4)
    monkeypatch.setattr(adhoc, "_pull", lambda h, t, q, conn, pov, rows, cols: log.append(("pull", conn)) or Pull())
    monkeypatch.setattr(adhoc, "_drop", lambda t, name: log.append(("drop", name)))
    got = adhoc.adhoc_fetch(None, [REQ], "session")
    assert [(f.value, f.period) for f in got[REQ].facts] == [(7.0, "Sep 26")]
    assert got[REQ].route == "adhoc"
    assert log == [("pull", "FinMax"), ("drop", "Q4 FinMax")]


def test_one_failed_adhoc_grid_does_not_stop_the_rest(monkeypatch):
    monkeypatch.setattr(adhoc, "_target", lambda h, target: "T")
    monkeypatch.setattr(adhoc, "_next_q", lambda t: 4)

    def boom(*a):
        raise RuntimeError("Smart View not installed")
    monkeypatch.setattr(adhoc, "_pull", boom)
    monkeypatch.setattr(adhoc, "_drop", lambda t, name: None)
    got = adhoc.adhoc_fetch(None, [REQ], "session")
    assert got[REQ].facts == [] and "Smart View not installed" in got[REQ].warnings[0]
```

- [ ] **Step 2:** Run it and see it fail with `ModuleNotFoundError`.

- [ ] **Step 3: Implement**

```python
# engine/smartass/data/adhoc.py
"""The ad hoc fallback: grids the HTTP route could not fetch are pulled as Smart View ad hoc grids in the user's
Excel (v1 excel/adhoc.pull, which needs Smart View and the bridge), parsed by the same facts_from_parsed and the
temporary sheet removed. Runs as one user-Excel job; a grid that fails here is a warning, never a crash."""
from __future__ import annotations

from .provider import FetchResult, _grids, facts_from_parsed


def _target(h, target):
    from smartview_mcp.excel import target as xltarget
    return xltarget.resolve_target(h, target, allow_archive=False, title="Ad hoc fetch")


def _next_q(t):
    from smartview_mcp.excel import sheets
    return sheets.next_question(t)


def _pull(h, t, q, conn, pov, rows, cols):
    from smartview_mcp.excel import adhoc as v1adhoc
    return v1adhoc.pull(h, t, q, conn, pov, rows, cols)


def _drop(t, name):
    from smartview_mcp.excel import sheets
    sheets.drop_parked(t, {name: name})


def adhoc_fetch(h, reqs, target: str) -> dict:
    t = _target(h, target)
    q = _next_q(t)
    out: dict = {}
    for req in reqs:
        res = FetchResult(req, route="adhoc")
        for pov, rows, cols, labels in _grids(req):
            try:
                p = _pull(h, t, q, req.cube, pov, rows, cols)
            except Exception as e:   # noqa: BLE001 - one grid failing must not stop the others
                res.warnings.append(f"{req.cube}: ad hoc fallback failed - {e}")
                continue
            try:
                facts, warns = facts_from_parsed(req, p.parsed, labels)
                res.facts += facts
                res.warnings += warns
            finally:
                _drop(t, p.sheet)
        out[req] = res
    return out
```

  Check `sheets.drop_parked`'s real signature. It takes the parked mapping produced by `park_sheets`, so adapt
  `_drop` to delete one sheet by name the way `pipeline._drop_backup` does. Check also that `GridRequest` is
  hashable (a frozen dataclass), because it is a dict key.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): ad hoc fallback job - failed grids pulled in Excel, parsed, sheet removed`.

---

### Task 3: Fallback wiring and `ComboProvider`

**Files:** `engine/smartass/data/provider.py` (`ComboProvider`), `engine/smartass/server.py` (`_prepare`). Test:
`engine/tests/v2/test_v2_fallback_flow.py`.

- [ ] **Step 1: Write the failing tests**
  - `ComboProvider(primary, fixed)`. `fetch(req)` returns `fixed[req]` when present, otherwise
    `primary.fetch(req)`. `closed_months` delegates to `primary`.
  - **`server._prepare` fallback.**
    1. The first `run_prepare` returns `pre.failed = [req]`.
    2. `_prepare` awaits `_user(lambda h: adhoc_fetch(h, pre.failed, target), "ad hoc fallback")`.
    3. It runs `run_prepare` again with `ComboProvider(provider, fixed)`.
    4. The reply's `routes` count `adhoc: 1`.
    - Monkeypatch `run_prepare`, `_user` and `adhoc_fetch`.
  - **No failures.** No ad hoc job runs.

- [ ] **Step 2:** Run them and see them fail.

- [ ] **Step 3: Implement.**
  - `_prepare(raw, provider, spec=None, target="session")` gets the target and does the fallback round above at
    most once.
  - When the second round still has failures, they stay in the reply's warnings.
  - The `routes` in `Prepared` come from the `FetchResult.route` values, so they count `adhoc` automatically.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): failed HTTP grids fall back to ad hoc in Excel, once`.

---

### Task 4: Recorder and replay

**Files:** Create `engine/smartass/data/recorder.py`. Test: `engine/tests/v2/test_v2_recorder.py`.

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_recorder.py
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass.data.recorder import RecordingProvider, ReplayProvider, load_lines, save_lines
from smartass.model.lines import LineSet
from smartass.pipeline import prepare


def test_record_then_replay_gives_the_same_answer(tmp_path):
    rec = RecordingProvider(FakeProvider(facts()))
    live = prepare(spec_dict(), resolver(), rec)
    rec.save(tmp_path / "b02.json")
    replay = prepare(spec_dict(), resolver(), ReplayProvider(tmp_path / "b02.json"))
    assert replay.result.key_values == live.result.key_values


def test_replay_of_an_unrecorded_request_is_an_error(tmp_path):
    RecordingProvider(FakeProvider(facts())).save(tmp_path / "empty.json")
    import pytest
    with pytest.raises(KeyError, match="not recorded"):
        prepare(spec_dict(), resolver(), ReplayProvider(tmp_path / "empty.json"))


def test_line_sets_round_trip(tmp_path):
    ls = {"b0": LineSet("b0", [{"AMOUNT": -1.0, "AP_CUSTOMER_NAME": "X"}], ["AMOUNT", "AP_CUSTOMER_NAME"],
                        cube_value=-1.0, cell={"All Accounts": "ac.C60"})}
    save_lines(tmp_path / "l.json", ls)
    back = load_lines(tmp_path / "l.json")
    assert back["b0"].rows == ls["b0"].rows and back["b0"].cube_value == -1.0
```

- [ ] **Step 2:** Run it and see it fail with `ModuleNotFoundError`.

- [ ] **Step 3: Implement**

```python
# engine/smartass/data/recorder.py
"""Record a run's grids (and drills) to a local JSON file and replay them without EPM: the test set's offline
mode. Files live outside git (%LOCALAPPDATA%\\smartass\\v2\\testset): they hold real financial data."""
from __future__ import annotations

import dataclasses
import json
from pathlib import Path

from ..model.frame import Fact
from ..model.lines import LineSet
from ..model.requests import GridRequest
from .provider import FetchResult


def _key(req: GridRequest) -> str:
    return json.dumps(dataclasses.asdict(req), sort_keys=True)


class RecordingProvider:
    def __init__(self, primary):
        self.primary = primary
        self.grids: dict[str, list[dict]] = {}
        self.closed = None

    def closed_months(self):
        self.closed = self.primary.closed_months()
        return self.closed

    def fetch(self, req):
        res = self.primary.fetch(req)
        self.grids[_key(req)] = [dataclasses.asdict(f) for f in res.facts]
        return res

    def save(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        closed = {k: list(v) for k, v in (self.closed or {}).items()}
        path.write_text(json.dumps({"closed": closed, "grids": self.grids}, ensure_ascii=False), encoding="utf-8")


class ReplayProvider:
    def __init__(self, path: Path):
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        self.closed = {k: tuple(v) for k, v in data["closed"].items()}
        self.grids = data["grids"]

    def closed_months(self):
        return self.closed

    def fetch(self, req):
        k = _key(req)
        if k not in self.grids:
            raise KeyError(f"grid not recorded: {k} - record the case again")
        return FetchResult(req, [Fact(**f) for f in self.grids[k]], route="replay")


def save_lines(path: Path, sets: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({k: dataclasses.asdict(v) for k, v in sets.items()}, ensure_ascii=False, default=str),
                    encoding="utf-8")


def load_lines(path: Path) -> dict:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    return {k: LineSet(**v) for k, v in data.items()}
```

  Notes:
  - **Recording before the first fetch.** `RecordingProvider.closed_months` is called once by `run_prepare`. An
    empty recording (no closed months) must still load. Handle `closed: null` by saving `{}`, and make
    `ReplayProvider.closed_months` raise `KeyError` "not recorded" when it is empty, so the second test still
    raises `KeyError`.
  - **`prepare` and `KeyError`.** Check that the Task 1 wrapper only catches `SvHttpError`. A `KeyError` from
    replay must not be recorded as a failed grid.
  - **Dates in drill rows.** Drill rows can hold dates. `default=str` stores them as strings, which is enough for
    `LinePick` (it only reads `AMOUNT` and the group column).

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): record and replay grids and line sets (offline test mode)`.

---

### Task 5: Test set: cases, CLI and offline comparison

**Files:**
- Create `engine/smartass/testset/__init__.py`, `engine/smartass/testset/runner.py`,
  `engine/smartass/testset/__main__.py`.
- Create the cases `engine/smartass/testset/cases/*.json` (specs only, no numbers).
- Test: `engine/tests/v2/test_v2_testset.py`.

**Cases** (v2 specs; the live record step validates the members; the ids follow the v1 bench):

| File | Question | Spec in short |
|---|---|---|
| `b01.json` | Wholesale P&L YTD vs LY | FinMax un.SY.2WHS, Revenue/COGS/Opex/Other items/EBITDA, closed-ytd, ly; kpi + table + bridge |
| `b02.json` | Business EBITDA YTD Sep 26 vs LY | FinMax un.SY.2BUS, EBITDA, "Jan 26..Sep 26", ly; kpi (golden 3 496 062 985 / 89 346 421) |
| `b05.json` | Technology actual vs budget and forecast YTD | FinPlan "1TEC Technology", Revenue/Opex/EBITDA, closed-ytd, budget + forecast; table [value, budget, var, forecast, var] |
| `b06.json` | Top cost centres by opex under Technology | FinMax un.SY.2TEC, Opex, closed-ytd, ly; ranking top 15 |
| `b07.json` | Telenor Norge monthly revenue trend | FinMax un.SY, Revenue, closed-ytd, ly; trend |
| `b13.json` | EBITDA YTD for un.SY.5OK_W | FinMax un.SY.5OK_W, EBITDA, closed-ytd, ly; kpi |
| `b15.json` | Business opex drivers | FinMax un.SY.2BUS, Opex, closed-ytd, ly; drivers depth 2 |
| `b16.json` | Technology other opex lines by supplier, Sep 26 | lines block cell ac.C60 / un.SY.2TEC / Sep 26, supplier |

  `b02` pins absolute periods, so its golden values never move. The others use relative periods, so their goldens
  belong to the closed month they were recorded in. The runner stores the recorded closed month and compares only
  against replays of the same recording.

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_testset.py
import json

from test_v2_pipeline import FakeProvider
from v2util import facts, spec_dict

from smartass.testset import runner


def test_cases_are_valid_specs():
    from smartass.model.spec import AnalysisSpec
    cases = runner.load_cases()
    assert {c["id"] for c in cases} >= {"b01", "b02", "b05", "b06", "b07", "b13", "b15", "b16"}
    for c in cases:
        AnalysisSpec.model_validate(c["spec"])


def test_compare_within_one_nok():
    golden = {"K · EBITDA": {"Actual": 100.0, "Variance %": 0.1}}
    assert runner.compare(golden, {"K · EBITDA": {"Actual": 100.4, "Variance %": 0.1}}) == []
    diffs = runner.compare(golden, {"K · EBITDA": {"Actual": 102.0, "Variance %": 0.1}})
    assert diffs and "Actual" in diffs[0]
    assert runner.compare(golden, {}) == ["K · EBITDA: missing"]


def test_record_approve_run_offline(tmp_path, monkeypatch):
    monkeypatch.setattr(runner, "HOME", tmp_path)
    case = {"id": "t1", "spec": spec_dict()}
    from v2util import resolver
    monkeypatch.setattr(runner, "_resolver_for", lambda closed: resolver())
    runner.record(case, FakeProvider(facts()), lines={})
    assert (tmp_path / "t1" / "candidate.json").exists()
    runner.approve("t1")
    report = runner.run_offline([case])
    assert report == [{"id": "t1", "ok": True, "diffs": []}]
    gold = json.loads((tmp_path / "t1" / "golden.json").read_text(encoding="utf-8"))
    assert "Key figures · EBITDA" in gold["key_values"]
```

- [ ] **Step 2:** Run it and see it fail.

- [ ] **Step 3: Implement**
  - **`runner.py`.**
    - `HOME = Path(%LOCALAPPDATA%)/smartass/v2/testset`. Use v1 `config.SVKIT` / "v2" / "testset", so that tests'
      `SMARTASS_HOME` applies.
    - `load_cases()` reads `cases/*.json`, each `{"id", "question", "spec"}`.
    - `_resolver_for(closed)` opens the atlas and builds an `AtlasResolver`, as `run_prepare` does.
    - `record(case, provider, lines)`:
      1. Wraps the provider in a `RecordingProvider`.
      2. Calls `prepare(case["spec"], _resolver_for(closed), rec, closed=closed, lines=lines)`.
      3. Saves `<id>/grids.json`, `<id>/lines.json` and `<id>/candidate.json`
         (`{"closed", "key_values", "recorded_at"}`).
    - `approve(id)` copies the candidate to `golden.json`.
    - `run_offline(cases)` does, for every case with a golden: `prepare` on a `ReplayProvider` with
      `load_lines(...)`, then `compare(golden.key_values, new.key_values)`, and returns `[{id, ok, diffs}]`.
      Cases without a golden are reported `{"ok": None, "diffs": ["no golden"]}`.
    - `compare(golden, new, tol_abs=1.0, tol_frac=1e-6)` returns `"line: missing"` for a missing line. A header
      differs when `|a-b| > tol`, where `tol` is `tol_abs` for amounts with `|a| >= 1`, else `tol_frac`. Treat
      `None` as `None` (both None is equal).
  - **`__main__.py`.** A CLI: `uv run python -m smartass.testset record [ids] | approve ids | run`.
    - `record` uses the live `SvHttpProvider(svhttp.shared())` and the server's async `_line_sets` for lines
      cases, through `asyncio.run`. It needs EPM, so it is only run on Tom's go.
    - `run` is offline and prints a table. Its exit code is 1 when any case fails.
  - **The case JSON files**, per the table above. Write them as valid `AnalysisSpec`s: plan `quick` or `standard`
    with a purpose, English text.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): test set - cases, record / approve / run offline`.

---

### Task 6: Three small items from 1E

**Files:** `engine/smartass/pipeline.py` (reply), `engine/smartass/render/layout.py` (notes),
`engine/smartass/server.py` (`_line_sets`). Tests: extend the existing test files.

- [ ] **Step 1: Write the failing tests**
  - **Text-fix candidates.** Warnings that start with "Text broken in EPM" are moved out of `warnings` into
    `reply["text_fixes"]` as a list of the broken texts. Today they arrive through the drill warnings.
  - **Long notes.** A note row longer than the answer's text width is split at word boundaries into several note
    rows of at most 78 characters, the width styling's lint uses. The test runs the layout with a 200-character
    source note and checks that every note cell is at most 78 characters long.
  - **Drills in parallel.** With two lines blocks, `_line_sets` drills them concurrently. The fake `_drill`
    records overlapping start and end times; use `asyncio.gather`.
- [ ] **Step 2:** Run them and see them fail.
- [ ] **Step 3:** Implement each item minimally.
- [ ] **Step 4:** `uv run pytest -q && uv run ruff check .` are both green.
- [ ] **Step 5:** Add a CHANGELOG line: `v2: ad hoc fallback for failed HTTP grids, record/replay, test set
  (record / approve / run offline), text-fix candidates in the reply, wrapped notes, parallel drills.`
- [ ] **Step 6: Commit.** Message: `fix(v2): text-fix candidates listed, notes wrapped, drills in parallel`.

---

### Task 7: Live (only on Tom's go)

1. **Record the set.** `uv run python -m smartass.testset record`. It uses the hidden browser for HTTP and
   drill, not Excel. Report each case's key values and any refusals; fix a case's members if a guard refuses.
2. **Tom reviews the candidates**, with B02 compared to v1. The controller then runs
   `uv run python -m smartass.testset approve b01 b02 ...`.
3. **Run offline.** `uv run python -m smartass.testset run` should show every case `ok`.
4. **Ad hoc fallback.** Only if Tom wants to see it: force one HTTP failure, for example with
   `SV_HTTP_TIMEOUT=0.001` for one `fin_analyze` call. Check that the answer is built with `routes.adhoc`, Sheet =
   model OK, and that no `Q<n> FinMax` ad hoc sheet is left behind.
