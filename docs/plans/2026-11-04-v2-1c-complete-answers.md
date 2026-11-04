# smartass v2 · Plan 1C: complete answers

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make v2 answers complete and living: charts on trend and bridge blocks, follow-ups that change an existing
answer in place, monthly packs refreshed with one call (`fin_refresh`), and the three live findings from Plan 1B.

**Architecture:**
- **Charts.** The layout records a chart anchor per block (pure). `render/charts.py` draws them through COM in v1's
  style: a line chart for trends, a waterfall for bridges, named "smartass chart".
- **Follow-ups and refresh.** These run as three steps:
  1. A short Excel job reads the stored v2 spec.
  2. `prepare` runs on the patched spec over HTTP.
  3. A redraw job backs up `Q<n> Answer` and `Q<n> Data`, clears the old answer's range and charts, rebuilds both
     sheets, and restores the backup on failure.
- **`fin_refresh`.** It does the same with no patch, for every v2 question in the workbook. Relative periods
  (`closed-ytd`) roll forward to the new closed month.

**Tech Stack:** Python 3.12, pydantic 2.14, mcp SDK (v1's `MCPServer`), pywin32 via v1, pytest, ruff.

**Scope split (Tom, 2026-11-04):**
- 1C: complete answers.
- 1D:
  - drill with ledger lines (`fin_drill` and the `lines` block)
  - `xl_snapshot`, `xl_revert`, `xl_tidy`
  - the ad hoc Excel fallback
  - the runtime fixes
  - the test-set runner

**Browser slots (live finding 1): no code here.** Every server keeps its own warm, signed-in browser. A shared
browser would need a CDP attach, and whether the SSO cookies survive that is unproven. v2 is registered with
`SV_DRILL_PROFILE_SLOTS=4` and `SV_DRILL_KEEP_WARM=0`. 1D revisits this.

**Conventions for every task:**
- Branch `v2/1c`, created from `main` in Task 0.
- Git Bash from `/c/dev/smartass/engine`.
- Tests go in `engine/tests/v2/test_v2_*.py`. Import fixtures with `from v2util import ...`. v1 test helpers
  import directly.
- Never run python through a stdin heredoc. Bare python only with `< /dev/null`.
- Stage files explicitly.
- No live Excel, EPM or browser runs. Task 9 waits for Tom.
- Never modify `engine/smartview_mcp` or `knowledge/`. Runtime hooks go through `smartass/legacy.py`.
- Commit messages end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

**Before each task, read the current v2 code it touches.** Plans 1A and 1B review fixes changed several
signatures:
- `prepare(raw, resolver, provider, closed=None)`
- `run_prepare` in `server.py`
- `_arrange(t, sheet)`
- `col_number` in `render/excel.py`
- the `_svmeta` creation in `pipeline`

Where a snippet here differs from the code, follow the code and say so.

## File structure

| File | Change |
|---|---|
| `engine/smartass/model/blocks/drivers.py` | Skip single-child pass-through nodes |
| `engine/smartass/legacy.py` | `configure_v2()` also teaches v1's `sheetorder` that `Q<n> Data` is our data sheet |
| `engine/smartass/render/layout.py` | `ChartSpec` and `SheetPlan.charts` anchors (pure) |
| `engine/smartass/render/charts.py` | New: `write_charts(ws, plan)` through COM |
| `engine/smartass/followup.py` | New: `merge_patch`, `v2_questions`, `find_question` (pure helpers over stored specs) |
| `engine/smartass/pipeline.py` | `owned` saved with the spec; `read_stored`, `redraw_job`; charts in `excel_job` |
| `engine/smartass/server.py` | `fin_analyze(spec \| patch, q)`, `fin_refresh`; INSTRUCTIONS updated |

---

### Task 0: Branch

- [ ] `cd /c/dev/smartass && git switch main && git switch -c v2/1c`

---

### Task 1: Drivers skip pass-through nodes

**Files:** Modify `engine/smartass/model/blocks/drivers.py`. Test: `engine/tests/v2/test_v2_drivers_passthrough.py`.

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_drivers_passthrough.py
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.drivers import build_drivers
from smartass.model.context import Ctx
from smartass.model.frame import Fact, Frame
from smartass.model.spec import AnalysisSpec

# 2BUS -> PASS (one child) -> 4BUS_B, G910 ; 4BUS_B -> B1, B2
TREE = {"un.SY.2BUS": ("un.SY.PASS",), "un.SY.PASS": ("un.SY.4BUS_B", "un.SY.G910"),
        "un.SY.4BUS_B": ("un.SY.B1", "un.SY.B2")}


def passthrough_frame() -> Frame:
    out = []
    for f in facts().facts:
        out.append(f)
        if f.entity == "un.SY.2BUS":                       # PASS carries the same numbers as its only parent
            out.append(Fact(f.cube, f.query, f.account, "un.SY.PASS", f.period, f.scenario, f.version, f.value,
                            f.balance))
    return Frame(out)


def test_single_child_node_is_skipped_and_costs_no_depth():
    r = resolver()
    r.tree.clear()
    r.tree.update(TREE)
    r.__post_init__()
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "drivers", "query": "fm", "metric": "Opex",
                                                       "depth": 2}]))
    out = build_drivers(s.blocks[0], Ctx(s, r, CLOSED, passthrough_frame()), "b0")
    labels = [ln.label for ln in out.lines]
    assert "un.SY.PASS" not in labels
    assert labels[:2] == ["2BUS Business Norway", "un.SY.4BUS_B"]
    assert "un.SY.B1" in labels                            # depth 2 still reaches the grandchildren
```

- [ ] **Step 2:** Run `uv run pytest tests/v2/test_v2_drivers_passthrough.py -q` and see it fail: PASS appears, or
  B1 is missing.

- [ ] **Step 3: Implement.** In `walk()` in `drivers.py`, after the children list is computed (use the same
  filtered list the code uses for weights, e.g. `counted(...)`, if present), add:

```python
        if len(kids) == 1 and level < b.depth:            # a pass-through node: same numbers, no split - skip it
            walk(kids[0], level, prefix)
            return
```

  The recursion keeps `level` unchanged, so the node costs no depth.

  Also check the `DrillRequest` leaf logic. A pass-through leaf (one child that has no children) must still emit
  its drill request for the real leaf.

- [ ] **Step 4:** Run `uv run pytest tests/v2 -q` and see it pass. Existing driver tests stay green, because their
  tree has no single-child node.

- [ ] **Step 5: Commit**

```bash
git add engine/smartass/model/blocks/drivers.py engine/tests/v2/test_v2_drivers_passthrough.py
git commit -m "fix(v2): drivers skip single-child pass-through nodes

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Data sheets sort with our data

**Files:** Modify `engine/smartass/legacy.py`. Test: `engine/tests/v2/test_v2_legacy_sheetorder.py`.

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_legacy_sheetorder.py
from smartview_mcp import sheetorder

from smartass import legacy


def test_data_sheet_is_classified_as_ours(monkeypatch):
    monkeypatch.setattr(legacy, "_configured", False)
    monkeypatch.setattr(sheetorder, "_OURS", sheetorder._OURS)
    monkeypatch.setattr(sheetorder, "_RANK", dict(sheetorder._RANK))
    assert sheetorder.classify("Q3 Data")[0] == "other"
    legacy.teach_sheetorder()
    kind = sheetorder.classify("Q3 Data")[0]
    assert kind != "other"
    order = sheetorder.sheet_order(["Sources", "Q3 Data", "Mine", "Q3 Answer"])
    assert order.index("Q3 Answer") < order.index("Mine") < order.index("Q3 Data") < order.index("Sources")
```

  Read `sheetorder.py` first. If `classify`, `sheet_order` or the return shapes differ (for example
  `sheet_order` taking extra arguments), adapt the test to the real API while keeping the assertion: Answer
  first, the user's sheet next, then Data, and `Sources` last.

- [ ] **Step 2:** Run it and see it fail with `AttributeError: ... teach_sheetorder`.

- [ ] **Step 3: Implement.** In `legacy.py`:

```python
def teach_sheetorder() -> None:
    """v1's sheetorder only knows Q<n> Answer/FinMax/FinLite/FinPlan/Lines; add Data (ranked like a source sheet,
    before FinMax). Module globals are read at call time, so this needs no v1 edit."""
    import re

    from smartview_mcp import sheetorder
    if "Data" in sheetorder._RANK:
        return
    sheetorder._OURS = re.compile(r"^Q(\d+) (Answer|FinMax|FinLite|FinPlan|Lines|Data)(?: (.+))?$")
    sheetorder._RANK["Data"] = min(sheetorder._RANK.values()) - 1
```

  Then call `teach_sheetorder()` inside `configure_v2()`, before `_configured = True`.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `fix(v2): Q<n> Data sorts with the data sheets`.

---

### Task 3: Chart anchors in the layout (pure)

**Files:** Modify `engine/smartass/render/layout.py`. Test: `engine/tests/v2/test_v2_layout_charts.py`.

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_layout_charts.py
from v2util import CLOSED, facts, monthly_facts, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.spec import AnalysisSpec
from smartass.render.layout import layout


def run(blocks, frame, periods=None):
    d = spec_dict(blocks=blocks)
    if periods:
        d["queries"][0]["periods"] = periods
    s = AnalysisSpec.model_validate(d)
    return layout(1, s, compute(s, resolver(), CLOSED, frame), stock=set(), subtitle="x")


def test_bridge_gets_a_waterfall_anchor():
    p = run([{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}], facts())
    (c,) = p.charts
    assert c.type == "waterfall" and c.title == "EBITDA bridge"
    assert c.series_cols == ["C"] and len(c.rows) == 5 and c.totals == [1, 5]
    assert c.rows == list(range(c.rows[0], c.rows[0] + 5))
    assert c.header_row == c.rows[0] - 1


def test_trend_gets_a_line_anchor_without_total_row():
    p = run([{"type": "trend", "query": "fm", "metrics": ["Opex"]}], monthly_facts(), "Jul 26..Sep 26")
    (c,) = p.charts
    assert c.type == "line" and c.series_cols == ["C", "D"] and len(c.rows) == 3


def test_no_chart_without_chart_blocks_and_anchor_right_of_table():
    p = run([{"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA"]}], facts())
    assert p.charts == []
    p = run([{"type": "trend", "query": "fm", "metrics": ["Opex"], "chart": False}], monthly_facts(),
            "Jul 26..Sep 26")
    assert p.charts == []
```

- [ ] **Step 2:** Run it and see it fail with `AttributeError: 'SheetPlan' object has no attribute 'charts'`.

- [ ] **Step 3: Implement.** In `layout.py`:

```python
@dataclass
class ChartSpec:
    type: str                 # "line" | "waterfall"
    title: str
    header_row: int
    rows: list[int]           # data rows (trend: months only, no Total; bridge: start..end)
    series_cols: list[str]    # one series per column
    totals: list[int]         # waterfall points (1-based) that are totals
    anchor_col: str           # top-left cell of the chart: anchor_col + header_row
```

  Add `charts: list[ChartSpec] = field(default_factory=list)` to `SheetPlan`.

  In the block loop, remember the header row. After the lines are placed, when `o.chart` is set:
  - `rows` is every placed line row whose `kind` is not `"grand"` or `"check"` for a trend. For a bridge it is
    every line row.
  - `series_cols` is every value column letter for a trend, and only the first value column for a bridge.
  - `totals` is `[1, len(rows)]` for a waterfall, otherwise `[]`.

  Set `anchor_col` after the loop, when `widest` is known, to `col_letter(FIRST_COL + widest + 1)`, so every
  chart sits right of the widest table at its own block's header row. Use the block title from `o.title`.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): chart anchors for trend and bridge blocks`.

---

### Task 4: Chart writer (COM)

**Files:** Create `engine/smartass/render/charts.py`. Test: `engine/tests/v2/test_v2_charts.py`.

- [ ] **Step 1: Write the failing test** with a recording fake. Test only what our code decides: chart type
  codes, the source ranges, names and titles. That is COM glue.

```python
# engine/tests/v2/test_v2_charts.py
from smartass.render.charts import CHART_NAME, XL_LINE, XL_WATERFALL, write_charts
from smartass.render.layout import ChartSpec, SheetPlan


class Obj:
    """Records attribute sets and calls; any attribute is another Obj."""
    def __init__(self, log, name):
        object.__setattr__(self, "_log", log)
        object.__setattr__(self, "_name", name)

    def __getattr__(self, k):
        child = Obj(self._log, f"{self._name}.{k}")
        object.__setattr__(self, k, child)
        return child

    def __setattr__(self, k, v):
        self._log.append(("set", f"{self._name}.{k}", v))

    def __call__(self, *a):
        self._log.append(("call", self._name, a))
        return Obj(self._log, f"{self._name}()")


class WS(Obj):
    def Range(self, a):
        self._log.append(("range", a))
        return Obj(self._log, f"Range({a})")

    def Cells(self, r, c):
        return Obj(self._log, f"Cells({r},{c})")


def test_line_and_waterfall_are_drawn_and_named():
    log = []
    ws = WS(log, "ws")
    p = SheetPlan(charts=[
        ChartSpec("line", "Trend", 10, [11, 12, 13], ["C", "D"], [], "G"),
        ChartSpec("waterfall", "EBITDA bridge", 20, [21, 22, 23, 24, 25], ["C"], [1, 5], "G")])
    warnings = write_charts(ws, p)
    assert warnings == []
    adds = [e for e in log if e[0] == "call" and e[1].endswith("Shapes.AddChart2")]
    assert [a[2][1] for a in adds] == [XL_LINE, XL_WATERFALL]
    assert ("range", "C11:C13") in log and ("range", "B11:B13") in log and ("range", "D11:D13") in log
    assert ("range", "B21:C25") in log
    assert any(e[0] == "set" and e[1].endswith(".Name") and e[2] == CHART_NAME for e in log)
```

- [ ] **Step 2:** Run it and see it fail with `ModuleNotFoundError`.

- [ ] **Step 3: Implement**

```python
# engine/smartass/render/charts.py
"""Charts for v2 answers, in v1's style (answer_write._chart / _multi_chart): a line chart per trend block (one
series per value column, months from column B), a waterfall per bridge block (column B labels + column C amounts,
start and end as totals; falls back to a column chart where Excel has no waterfall). Every chart is named
'smartass chart' so a follow-up can drop it. A chart that fails is a warning, never a failed answer."""
from __future__ import annotations

from smartview_mcp.excel.jobs import com

from .layout import SheetPlan

CHART_NAME = "smartass chart"
XL_LINE = 4
XL_WATERFALL = 119
XL_COLUMN = 51
WIDTH = 480.0


def _place(ws, spec, xl_type):
    cell = ws.Range(f"{spec.anchor_col}{spec.header_row}")
    left, top = com(lambda: cell.Left), com(lambda: cell.Top)
    height = max(180.0, 15.0 * (len(spec.rows) + 3))
    shape = com(lambda: ws.Shapes.AddChart2(-1, xl_type, left, top, WIDTH, height))
    com(lambda: setattr(shape, "Name", CHART_NAME))
    return shape, com(lambda: shape.Chart)


def _title(chart, text):
    com(lambda: setattr(chart, "HasTitle", True))
    com(lambda: setattr(chart.ChartTitle, "Text", text))


def _millions(chart):
    try:
        ax = com(lambda: chart.Axes(2))
        com(lambda: setattr(ax, "DisplayUnit", -6))          # xlMillions
        com(lambda: setattr(ax, "HasDisplayUnitLabel", False))
        com(lambda: setattr(ax.TickLabels, "NumberFormat", '0" MNOK"'))
    except Exception:   # noqa: BLE001 - cosmetic
        pass


def _line(ws, spec):
    shape, chart = _place(ws, spec, XL_LINE)
    r1, r2 = spec.rows[0], spec.rows[-1]
    for col in spec.series_cols:
        s = com(lambda: chart.SeriesCollection().NewSeries())
        com(lambda: setattr(s, "Values", ws.Range(f"{col}{r1}:{col}{r2}")))
        com(lambda: setattr(s, "XValues", ws.Range(f"B{r1}:B{r2}")))
        com(lambda: setattr(s, "Name", f"={ws.Name}!${col}${spec.header_row}"))
    _title(chart, spec.title)
    com(lambda: setattr(chart, "HasLegend", len(spec.series_cols) > 1))
    _millions(chart)


def _waterfall(ws, spec):
    r1, r2 = spec.rows[0], spec.rows[-1]
    col = spec.series_cols[0]
    try:
        shape, chart = _place(ws, spec, XL_WATERFALL)
        com(lambda: chart.SetSourceData(ws.Range(f"B{r1}:{col}{r2}")))
        series = com(lambda: chart.FullSeriesCollection(1))
        for idx in spec.totals:
            com(lambda: setattr(series.Points(idx), "IsTotal", True))
    except Exception:   # noqa: BLE001 - no waterfall in this Excel: a column chart says the same
        shape, chart = _place(ws, spec, XL_COLUMN)
        com(lambda: chart.SetSourceData(ws.Range(f"B{r1}:{col}{r2}")))
    _title(chart, spec.title)
    com(lambda: setattr(chart, "HasLegend", False))
    _millions(chart)


def write_charts(ws, plan: SheetPlan) -> list[str]:
    warnings = []
    for spec in plan.charts:
        try:
            (_waterfall if spec.type == "waterfall" else _line)(ws, spec)
        except Exception as e:   # noqa: BLE001 - a chart never fails the answer
            warnings.append(f"chart '{spec.title}' not drawn: {e}")
    return warnings
```

  If `com(fn)` in v1 needs `fn` to take no arguments, it already fits: every call above is a zero-argument
  lambda. If the recording fake makes `com` misbehave, monkeypatch `smartass.render.charts.com` to
  `lambda fn, **k: fn()` in the test, and keep the assertions.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): line and waterfall charts, named for follow-ups`.

---

### Task 5: Charts and the owned range in `excel_job`

**Files:** Modify `engine/smartass/pipeline.py`. Test: extend `engine/tests/v2/test_v2_pipeline_job.py`.

- [ ] **Step 1: Write the failing test.** In `wire()` add recorders for `write_charts` (returns `[]`). Make
  `_save_spec` record its arguments. Then add:

```python
def test_job_draws_charts_and_saves_owned_range(monkeypatch):
    log, saved = [], {}
    wire(monkeypatch, log)
    monkeypatch.setattr(pipeline, "_save_spec", lambda t, q, raw, owned: saved.update(q=q, owned=owned))
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    pipeline.excel_job(None, pre, target="session")
    assert log.index("write_answer") < log.index("write_charts") < log.index("calculate")
    assert saved["q"] == 4 and saved["owned"]["rows"] > 10 and saved["owned"]["cols"] >= 3
```

  Update `test_job_order`'s expected list to include `"write_charts"` after `"write_answer"`.

- [ ] **Step 2:** Run it and see it fail.

- [ ] **Step 3: Implement.** In `pipeline.py`:
  - Import `write_charts` from `.render.charts`.
  - In `excel_job`, after `write_answer`, add
    `lint = list(lint or []) + write_charts(ws_answer, plan)`.
  - Change `_save_spec(t, q, raw)` to `_save_spec(t, q, raw, owned)` and pass
    `owned={"rows": plan.last_row, "cols": col_number(plan.last_col)}`. Use `col_number` from `render/excel.py`.
  - Inside `_save_spec`, pass `owned` to `sheets.save_spec(..., owned=owned)`.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): charts in the answer job; owned range saved with the spec`.

---

### Task 6: Follow-up helpers (pure)

**Files:** Create `engine/smartass/followup.py`. Test: `engine/tests/v2/test_v2_followup.py`.

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_followup.py
import pytest

from smartass.followup import find_question, merge_patch, v2_questions
from smartass.model.validate import GuardError

STORED = {1: {"spec": {"version": 1, "question": "old v1"}, "ids": {}, "owned": None},
          2: {"spec": {"v2": {"version": 2, "question": "a"}}, "ids": {}, "owned": {"rows": 30, "cols": 6}},
          5: {"spec": {"v2": {"version": 2, "question": "b"}}, "ids": {}, "owned": {"rows": 40, "cols": 7}}}


def test_merge_patch_is_rfc7386():
    base = {"plan": {"depth": "quick", "purpose": "x"}, "blocks": [1, 2], "question": "q"}
    out = merge_patch(base, {"plan": {"purpose": "y"}, "blocks": [3], "question": None})
    assert out == {"plan": {"depth": "quick", "purpose": "y"}, "blocks": [3]}
    assert base["plan"]["purpose"] == "x"                        # input untouched


def test_v2_questions_skip_v1():
    assert v2_questions(STORED) == [2, 5]


def test_find_question_explicit_active_latest():
    assert find_question(STORED, q=2, active_sheet="Q5 Answer")[0] == 2
    assert find_question(STORED, q=None, active_sheet="Q2 Answer")[0] == 2
    assert find_question(STORED, q=None, active_sheet="Mine")[0] == 5
    q, raw, owned = find_question(STORED, q=None, active_sheet=None)
    assert (q, raw["question"], owned["rows"]) == (5, "b", 40)
    with pytest.raises(GuardError, match="Q1 is not a v2 answer"):
        find_question(STORED, q=1, active_sheet=None)
    with pytest.raises(GuardError, match="no v2 answer"):
        find_question({}, q=None, active_sheet=None)
```

- [ ] **Step 2:** Run it and see it fail with `ModuleNotFoundError`.

- [ ] **Step 3: Implement**

```python
# engine/smartass/followup.py
"""Follow-ups on a stored v2 answer. A patch is a JSON merge patch (RFC 7386) on the stored spec: objects merge,
lists (queries, blocks) are replaced whole, null deletes a key. The question is the explicit q, else the active
'Q<n> Answer' sheet when it holds a v2 answer, else the latest v2 answer in the workbook."""
from __future__ import annotations

import copy
import re

from .model.validate import GuardError

_ANSWER = re.compile(r"^Q(\d+) Answer$")


def merge_patch(target, patch):
    if not isinstance(patch, dict):
        return copy.deepcopy(patch)
    out = copy.deepcopy(target) if isinstance(target, dict) else {}
    for k, v in patch.items():
        if v is None:
            out.pop(k, None)
        else:
            out[k] = merge_patch(out.get(k), v)
    return out


def v2_questions(stored: dict[int, dict]) -> list[int]:
    return sorted(q for q, a in stored.items() if isinstance(a.get("spec"), dict) and "v2" in a["spec"])


def find_question(stored: dict[int, dict], q: int | None, active_sheet: str | None) -> tuple[int, dict, dict | None]:
    """(q, raw v2 spec, owned range) of the answer a follow-up changes."""
    ours = v2_questions(stored)
    if q is not None:
        if q not in ours:
            raise GuardError(f"Q{q} is not a v2 answer in this workbook",
                             f"use one of {', '.join(f'Q{x}' for x in ours) or 'none'}, or ask a new question")
    else:
        m = _ANSWER.match(active_sheet or "")
        if m and int(m.group(1)) in ours:
            q = int(m.group(1))
        elif ours:
            q = ours[-1]
        else:
            raise GuardError("there is no v2 answer in this workbook to follow up",
                             "ask a new question with fin_analyze(spec)")
    a = stored[q]
    return q, a["spec"]["v2"], a.get("owned")
```

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): follow-up helpers - merge patch, find the v2 question`.

---

### Task 7: Read and redraw jobs (COM glue)

**Files:** Modify `engine/smartass/pipeline.py`. Test: `engine/tests/v2/test_v2_redraw.py`.

- [ ] **Step 1: Write the failing test** in the same recorder style as `test_v2_pipeline_job.py`:

```python
# engine/tests/v2/test_v2_redraw.py
import pytest
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass import pipeline


class T:
    path = "C:/a/v2/x.xlsx"
    created = True


def wire(monkeypatch, log, fail_at=None):
    def rec(name, ret=None):
        def f(*a, **k):
            log.append(name)
            if name == fail_at:
                raise RuntimeError("boom")
            return ret
        return f
    monkeypatch.setattr(pipeline, "_target", rec("target", T()))
    monkeypatch.setattr(pipeline, "_backup", lambda t, name: log.append(f"backup {name}") or f"~{name}")
    monkeypatch.setattr(pipeline, "_sheet", rec("sheet", object()))
    monkeypatch.setattr(pipeline, "_clear", rec("clear"))
    monkeypatch.setattr(pipeline, "_delete_sheet", rec("delete_data"))
    monkeypatch.setattr(pipeline, "_add_sheet", rec("add_sheet", object()))
    monkeypatch.setattr(pipeline, "write_data", rec("write_data"))
    monkeypatch.setattr(pipeline, "write_answer", rec("write_answer", []))
    monkeypatch.setattr(pipeline, "write_charts", rec("write_charts", []))
    monkeypatch.setattr(pipeline, "_calculate", rec("calculate"))
    monkeypatch.setattr(pipeline, "read_block", rec("read_block", ()))
    monkeypatch.setattr(pipeline, "compare", rec("compare", []))
    monkeypatch.setattr(pipeline, "_save_spec", rec("save_spec"))
    monkeypatch.setattr(pipeline, "_arrange", rec("arrange", []))
    monkeypatch.setattr(pipeline, "_drop_backup", lambda t, b: log.append(f"drop {b}"))
    monkeypatch.setattr(pipeline, "_restore", lambda t, name, b: log.append(f"restore {name}"))
    monkeypatch.setattr(pipeline, "_rollback", rec("rollback"))
    monkeypatch.setattr(pipeline, "_save", rec("save"))


def test_redraw_order(monkeypatch):
    log = []
    wire(monkeypatch, log)
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    out = pipeline.redraw_job(None, pre, q=2, owned={"rows": 30, "cols": 6}, target="session")
    assert log[:4] == ["target", "backup Q2 Answer", "backup Q2 Data", "sheet"]
    assert log.index("clear") < log.index("delete_data") < log.index("add_sheet") < log.index("write_data")
    assert log[-3:] == ["drop ~Q2 Answer", "drop ~Q2 Data", "save"]
    assert out["q"] == 2 and out["followup"] is True


def test_redraw_restores_both_sheets_on_failure(monkeypatch):
    log = []
    wire(monkeypatch, log, fail_at="write_answer")
    pre = pipeline.prepare(spec_dict(), resolver(), FakeProvider(facts()))
    with pytest.raises(RuntimeError):
        pipeline.redraw_job(None, pre, q=2, owned=None, target="session")
    assert log[-3:] == ["rollback", "restore Q2 Answer", "restore Q2 Data"]
```

- [ ] **Step 2:** Run it and see it fail with `AttributeError: redraw_job`.

- [ ] **Step 3: Implement** in `pipeline.py`. Use v1's `core/owned.py` (`backup`, `restore`, `clear`,
  `drop_charts`) and `excel/sheets.py` (`_stored_specs`) through small wrappers, so tests can replace them:

```python
def _backup(t, name: str) -> str:
    from smartview_mcp.core import owned
    return owned.backup(t, name)


def _restore(t, name: str, bname: str) -> None:
    from smartview_mcp.core import owned
    owned.restore(t, name, bname)


def _drop_backup(t, bname: str) -> None:
    from smartview_mcp.excel.jobs import com
    ws = com(lambda: t.wb.Worksheets(bname))
    app = t.wb.Application
    com(lambda: setattr(app, "DisplayAlerts", False))
    try:
        com(lambda: ws.Delete())
    finally:
        com(lambda: setattr(app, "DisplayAlerts", True))


def _sheet(t, name: str):
    from smartview_mcp.excel.jobs import com
    return com(lambda: t.wb.Worksheets(name))


def _clear(ws, owned: dict | None) -> None:
    from smartview_mcp.core import owned as owned_mod
    from smartview_mcp.excel.jobs import com
    owned_mod.clear(ws, owned)
    owned_mod.drop_charts(ws, com)


def _delete_sheet(t, name: str) -> None:
    _drop_backup(t, name)                      # same COM dance: delete a sheet without the confirm dialog


def read_stored(h, target: str = "session") -> tuple[dict, str | None]:
    """(stored specs by q, active sheet name): runs as a short user-Excel job."""
    from smartview_mcp.excel import sheets
    from smartview_mcp.excel.jobs import com
    t = _target(h, target, "Follow-up")
    active = com(lambda: t.wb.ActiveSheet.Name)
    return sheets._stored_specs(t), active


def redraw_job(h, pre: Prepared, q: int, owned: dict | None, target: str = "session") -> dict:
    t = _target(h, target, pre.spec.question[:60])
    a_name, d_name = f"Q{q} Answer", f"Q{q} Data"
    b_answer, b_data = _backup(t, a_name), _backup(t, d_name)
    try:
        ws_answer = _sheet(t, a_name)
        _clear(ws_answer, owned)
        _delete_sheet(t, d_name)
        ws_data = _add_sheet(t, d_name, q, DATA_TAB)
        out = _draw(t, q, pre, ws_answer, ws_data)
    except Exception:
        _rollback(t)
        _restore(t, a_name, b_answer)
        _restore(t, d_name, b_data)
        raise
    _drop_backup(t, b_answer)
    _drop_backup(t, b_data)
    _save(t)
    out["followup"] = True
    return out
```

  Refactor `excel_job` so that the drawing part (data, layout, answer, charts, calculate, readback, save spec
  with `owned`, arrange) is a shared `_draw(t, q, pre, ws_answer, ws_data) -> dict`. Both jobs call it, and it
  returns the `reply(...)` dict. `excel_job` keeps its own `_save` after `_draw`. Keep `test_v2_pipeline_job.py`
  green: same observable order.

  `_restore` handles the case where the sheet being restored was deleted: v1's `owned.restore` deletes the
  broken sheet, un-hides the backup and renames it. If the Data sheet was already deleted, catch that and only
  rename the backup. Read `owned.restore` and handle it in our wrapper, not in v1.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): follow-up redraw job - backup, rebuild, restore on failure`.

---

### Task 8: `fin_analyze` follow-ups, `fin_refresh`, instructions

**Files:** Modify `engine/smartass/server.py`. Test: extend `engine/tests/v2/test_v2_server.py`.

- [ ] **Step 1: Write the failing tests**

```python
def test_four_tools_registered():
    assert server.TOOL_NAMES == ("fin_analyze", "fin_lookup", "fin_status", "fin_refresh")


def test_patched_spec_is_validated_before_any_excel_work():
    from smartass.followup import merge_patch
    from smartass.model.spec import AnalysisSpec
    from v2util import spec_dict
    raw = merge_patch(spec_dict(), {"plan": {"purpose": "New purpose"}})
    assert AnalysisSpec.model_validate(raw).plan.purpose == "New purpose"


def test_instructions_mention_followup_refresh_and_bridge_steps():
    text = server.INSTRUCTIONS
    assert "patch" in text and "fin_refresh" in text and "COGS" in text
    assert len(text) < 2500
```

  Replace the old `test_three_tools_registered`.

- [ ] **Step 2:** Run them and see them fail.

- [ ] **Step 3: Implement** in `server.py`:
  - `TOOL_NAMES = ("fin_analyze", "fin_lookup", "fin_status", "fin_refresh")`.
  - Change `fin_analyze` to
    `(spec: AnalysisSpec | None = None, patch: dict | None = None, q: int | None = None, target: str = "session",
    ctx=None)`:
    - If `spec` is given and `patch` is not: the 1B path, a new question.
    - If `patch` is given (`spec` must be absent):
      1. `stored, active = await v1._call_user(lambda h: read_stored(h, target), what="fin_analyze follow-up")`.
      2. `q2, raw, owned = find_question(stored, q, active)`.
      3. `raw2 = merge_patch(raw, patch)`.
      4. `pre = await asyncio.to_thread(run_prepare, raw2, provider)`.
      5. `return await v1._call_user(lambda h: redraw_job(h, pre, q2, owned, target), what="fin_analyze follow-up")`.
    - Neither, or both: raise `GuardError("give spec (a new question) or patch (a follow-up), not both", ...)`.
  - Add `fin_refresh(target: str = "session", ctx=None) -> dict`:
    1. Read the stored specs once.
    2. For each q in `v2_questions(stored)`: run `run_prepare(raw, provider)`, then `redraw_job` (no patch).
    3. Collect per question `{"q", "ok", "closed", "error"?}`. One question failing does not stop the others.
    4. Return `{"refreshed": [...], "workbook": ...}`.
  - Register `fin_refresh` with `_register(v1.WRITE)`.
  - Replace `INSTRUCTIONS` (keep it under 2500 characters):

```python
INSTRUCTIONS = """smartass2 (v2, read-only towards Oracle EPM: FinMax/FinLite actuals, FinPlan plan/forecast).
1) fin_lookup: metric names and members (entities 'un.SY.<code>' in FinMax, '<code> <name>' in FinPlan).
2) New question: fin_analyze(spec) with an AnalysisSpec v2 - plan (depth, purpose, hypotheses), queries (cube,
metrics, entity, periods e.g. 'closed-ytd', compare ly | budget | forecast; budget/forecast only in FinPlan),
blocks (table, kpi_tiles, trend, matrix, bridge, ranking, drivers, commentary). An EBITDA bridge uses every P&L
step: Revenue, COGS, Opex, Other items (FinMax) - otherwise 'Other / unexplained' is large.
3) Follow-up: fin_analyze(patch=..., q=n) - a JSON merge patch on the stored spec (lists like blocks are replaced
whole); default q = the active Q<n> Answer, else the latest. 4) Monthly packs: fin_refresh() re-runs every v2
answer in the workbook to the latest closed month (relative periods roll forward), no AI needed.
5) Quote numbers only from key_values; explain material lines from drivers evidence or label them a hypothesis.
Sheet text is English; answer the user in their language. The user's own workbooks are never saved."""
```

- [ ] **Step 4:** `uv run pytest tests/v2 -q && uv run pytest -q && uv run ruff check .` are all green.
- [ ] **Step 5:** Add a CHANGELOG line under `## Unreleased` / `### Added`:
  `v2 complete answers: charts (trend, bridge), follow-ups via fin_analyze(patch, q), fin_refresh for monthly
  packs, drivers skip pass-through nodes, Q<n> Data sorted with data sheets.`
- [ ] **Step 6: Commit.** Message: `feat(v2): follow-ups, fin_refresh, instructions for full P&L bridges`.

---

### Task 9: Live check (only on Tom's go)

After Tom reconnects `smartass2` (`/mcp`), the controller runs these in the session workbook:
1. **Q1 again with charts:** the 1B spec with the bridge steps Revenue, COGS, Opex, Other items. Expect a line or
   waterfall chart, Sheet = model OK, and a small "Other / unexplained".
2. **Follow-up:** `fin_analyze(patch={"plan": {"purpose": "Business Norway EBITDA YTD vs last year - revised"}})`.
   Expect the same Q, a new title, and the charts redrawn exactly once.
3. **Refresh:** `fin_refresh()`. Expect every v2 Q `ok`.

Whatever fails becomes a 1D task with the observed error.

---

## Live result (2026-11-04, Tom's go)

- Q3 (full EBITDA bridge, opex trend and drivers): built; the bridge closes (Other ≈ 0); drivers skip 50BUS_B
  and show 5ADM_B, 5MS_B and 5NWW_B. The first try timed out while the cold browser signed in; the retry passed.
- `fin_refresh` failed live at first:
  - Cause: writing Value2 over a table's header and body as one rectangle deletes the table, so the formulas
    became #REF!.
  - OneDrive AutoSave stored the broken state.
  - Fixed in 778a295 (header and body written separately; a missing table is refused, never recreated). After
    the fix, `fin_refresh` repaired all three answers.
- Follow-up `fin_analyze(patch, q=3)`: rebuilt in place, new title, Sheet = model OK (near-zero readback fix
  df10cbd).

## Carried to Plan 1D

1. Turn off OneDrive AutoSave for v2's own files during writes, so a broken intermediate state is never stored.
2. Keep the session workbook across a `/mcp` reconnect. The relay pid changes, so the v2 session is lost and
   follow-ups need an explicit path.
3. Cold browser: the first request after the idle close times out during the SSO sign-in. Warm the browser up
   before the first grid, or retry automatically once the sign-in completes.
4. Claude Code keeps the old tool schema after a reload: `fin_analyze` showed no `patch` until the server was
   reconnected, and only ToolSearch was refreshed. Note this in the skill and the instructions.
