# Semantic catalogue + Smart View HTTP source: implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task by task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:**
- `sv_query` reads cube grids over Smart View's HTTP protocol, with no Excel or COM: about 1 s instead of 25–50 s.
  It falls back to the COM route when that fails.
- A hand-maintained semantic catalogue (`knowledge\catalog.toml`) becomes the single source for P&L lines, identity
  checks, business areas, scenarios, KPIs and materiality.

**Architecture:**
- `svxml.py` is pure: it builds and parses provider XML (`/HyperionPlanning/SmartView`).
- `svhttp.py` holds the sessions per connection and returns exactly what `ops._query_once` returns:
  `(Layout, Parsed, seconds)`. It reuses `grid.Layout`/`grid.parse`.
- Its transport posts from inside the drill's hidden Edge (Telenor SSO, one tab per EPM pod).
- `catalog.py` loads, structurally checks and atlas-validates the TOML file. `templates.py` takes its lists from it.

**Tech stack:** Python 3.12, `tomllib` (stdlib), Playwright (already a dependency), pytest + anyio, ruff.

**Background:**
- `docs\specs\2026-10-25-finance-expert-design.md` (stages)
- `docs\specs\2026-10-25-stage0-epm-access-probe.md` (protocol)
- `scripts\probe_sv_http.py`

**Protocol facts proven live on 2026-10-25** (the code below depends on them):
- **Request order:** `ConnectToProvider` → `sID`; then `OpenApplication(srv=app=FinMax|FinPlan)` → a new `sID`; then
  `OpenCube(cube)`. Without OpenCube, a retrieve fails with "adhoc mode or metadata mode should be set".
- **Retrieve:** `ExecuteGrid` with `<action>retrieve</action><ords/>`, a free-form grid, and the same layout as
  `grid.Layout.matrix()`: row 0 = POV, then column headers, then row headers. `vals enc="1"` (pipe-separated,
  row-major from cell 0), `types enc="0"` (7 blank, 0 member, 2 data).
- **Labels:** with `<aliasTableName val="Default"/>` the response labels equal Smart View in Excel
  (`2WHS Wholesale (un.SY)`, `C30 REVENUES (ac)`).
- **POV:** top members as the atlas names them (`All Counterparts`, `Account`, `Entity` …) are accepted. The server
  reorders the POV row in its answer; `grid.parse` only counts it.
- **Missing and no-access cells:** missing data comes back as an empty value. No access comes back as `#noaccess`, in
  lower case.
- **Unknown members:** an unknown member makes the whole request fail (`Grid is Invalid ...`), so the COM route,
  which finds and drops bad members by bisection, is the fallback.
- **Pods:**
  - FinMax and FinLite run on the `plan1` pod (cube `FinMax` / `FinLite`).
  - FinPlan runs on `stratmodel` (cube `FinPlan`). It needs its own tab, signed in on that origin (silent SSO,
    2–9 s); wait for network idle before posting.
  - `GetProvisionedDataSources` returns the pod URL per app.
- **Writes:** `freeformsubmit`, `WriteBack`, `SetVals` and so on write data. They must never be sent: there is an
  allowlist.

**Test commands** (from `C:\dev\smartass\engine`):
- unit tests: `.venv\Scripts\python.exe -m pytest -q`
- lint: `.venv\Scripts\python.exe -m ruff check .`
- before every commit, from the repo root: `python scripts\check_staged.py`

There are 295 tests passing today. Never commit real figures.

**Not in this plan (next plans):**
- The DuckDB warehouse and its extract jobs (Stage 2).
- Time rules: the closed month, LY and YTD stay in `templates.closed_month` / `same_period_ly` until the calc engine
  (Stage 3).
- The `sv_ws_*` workbooks keep building live Smart View sheets through COM: those are the refreshable sources.
- A nightly protocol probe in the live matrix.

---

## File structure

| File | Responsibility |
|---|---|
| Create `engine\smartview_mcp\svxml.py` | Pure provider XML: request bodies, read-only allowlist, response parsing to a dense 2D list |
| Create `engine\smartview_mcp\svhttp.py` | `SvHttp` (sessions per connection, retrieve contract = `ops._query_once`), `BrowserTransport`, `shared()` |
| Modify `engine\smartview_mcp\atlas.py` | Add `dimensions(db, connection)` |
| Modify `engine\smartview_mcp\drill_browser.py` | `_login_once(..., base=None)` for other pods; `shared_browser()` singleton |
| Modify `engine\smartview_mcp\tools_drill.py:53-58` | `browser()` delegates to `shared_browser()` |
| Modify `engine\smartview_mcp\config.py` | `SVHTTP_TIMEOUT_S` |
| Modify `engine\smartview_mcp\tools_query.py` | `_retrieve()` (HTTP first, COM fallback); `sv_query(source=...)` |
| Create `knowledge\catalog.toml` | The semantic catalogue (hand-maintained) |
| Create `engine\smartview_mcp\catalog.py` | Load + structural checks, `resolve_entity`, `validate` against the atlas |
| Modify `engine\smartview_mcp\templates.py:195-202,405-409` | `PNL`, `PNL_CHECKS`, `FP_PNL`, `FP_CHECKS` from the catalogue |
| Modify `engine\smartview_mcp\doctor.py` | `catalog` check |
| Modify `engine\tests\live\scenarios.py:737` | `DOCTOR_CHECKS` gains `"catalog"` |
| Create `engine\tests\test_svxml.py`, `test_svhttp.py`, `test_catalog.py`, `test_retrieve_route.py` | Unit tests |
| Create `engine\tests\live\svhttp_vs_com.py` | Live check: the HTTP route equals the COM route cell for cell |
| Modify `skills\smartass\reference\operations.md`, `docs\specs\2026-10-25-stage0-epm-access-probe.md` | Docs |

---

## Part A: Smart View HTTP source

### Task 1: `atlas.dimensions`

**Files:**
- Modify: `engine\smartview_mcp\atlas.py` (after `has_member`, line ~88)
- Test: `engine\tests\test_svhttp.py` (new)

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/test_svhttp.py
from smartview_mcp import atlas


def _store(db, conn, dims):
    for d in dims:
        atlas.store(db, conn, d, [{"name": d, "label": d, "parent": None, "depth": 0, "order": 0}], 0.1)


def test_atlas_dimensions_sorted_per_connection(tmp_path):
    db = atlas.connect(tmp_path / "a.db")
    _store(db, "FinMax", ["All Time Periods", "All Accounts"])
    _store(db, "FinPlan", ["Account"])
    assert atlas.dimensions(db, "FinMax") == ["All Accounts", "All Time Periods"]
    assert atlas.dimensions(db, "FinLite") == []
```

- [ ] **Step 2: Run it and check it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_svhttp.py -q`
Expected: FAIL, `AttributeError: module 'smartview_mcp.atlas' has no attribute 'dimensions'`

- [ ] **Step 3: Implement**

Add to `engine\smartview_mcp\atlas.py` after `has_member`:

```python
def dimensions(db, connection: str) -> list[str]:
    """The connection's dimension names (= their top members) that the atlas holds, sorted."""
    return [r[0] for r in db.execute("SELECT DISTINCT dimension FROM member WHERE connection=? ORDER BY dimension",
                                     (connection,))]
```

- [ ] **Step 4: Run the test again**

Run: `.venv\Scripts\python.exe -m pytest tests/test_svhttp.py -q`. Expected: `1 passed`.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/atlas.py engine/tests/test_svhttp.py
git commit -m "atlas: dimensions() per connection"
```

### Task 2: `svxml.py`, pure provider XML

**Files:**
- Create: `engine\smartview_mcp\svxml.py`
- Create: `engine\tests\svxml_util.py` (test helper, imported like `stdio_util`; `tests` is not a package)
- Test: `engine\tests\test_svxml.py`

- [ ] **Step 1: Write the test helper and the failing tests**

```python
# engine/tests/svxml_util.py
"""Provider answers for tests - shapes copied from the live probe of 2026-10-25; every figure is invented."""


def answer(matrix):
    """A res_ExecuteGrid like the provider's, for a dense matrix (float = data cell)."""
    h, w = len(matrix), len(matrix[0])
    cells = [v for row in matrix for v in row]
    vals = "|".join("" if v is None else (repr(v) if isinstance(v, float) else v) for v in cells)
    types = "|".join("2" if isinstance(v, float) else ("7" if v is None else "0") for v in cells)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><res_ExecuteGrid><grid><slices>'
            f'<slice rows="{h}" cols="{w}"><data><generations/><range start="0" end="{h * w - 1}">'
            f'<vals enc="1">{vals}</vals><types enc="0">{types}</types></range></data></slice></slices></grid>'
            '<stats><Total_Provider_Time>9</Total_Provider_Time></stats></res_ExecuteGrid>')
```

```python
# engine/tests/test_svxml.py
import pytest
from svxml_util import answer

from smartview_mcp import svxml
from smartview_mcp.grid import Axis, Layout, parse


def _lay():
    return Layout([("All Scenarios", "Actual")], [Axis("All Accounts", ["ac.A", "ac.B"])],
                  [Axis("All Time Periods", ["Aug 26", "Sep 26"])])


def test_grid_cells_mark_pov_headers_and_data():
    vals, types = svxml.grid_cells(_lay())
    assert vals == ["", "Actual", "", "", "Aug 26", "Sep 26", "ac.A", "", "", "ac.B", "", ""]
    assert types == ["7", "0", "7", "7", "0", "0", "0", "2", "2", "0", "2", "2"]


def test_grid_cells_escape_xml_and_refuse_pipes():
    lay = Layout([("All Entities", "un.SY.5P&C_P&C")], [Axis("All Accounts", ["ac.A"])],
                 [Axis("All Time Periods", ["Sep 26"])])
    assert "un.SY.5P&amp;C_P&amp;C" in svxml.grid_cells(lay)[0]
    bad = Layout([], [Axis("All Accounts", ["a|b"])], [Axis("All Time Periods", ["Sep 26"])])
    with pytest.raises(ValueError, match="separator"):
        svxml.grid_cells(bad)


def test_execute_grid_is_a_retrieve_with_default_aliases():
    body = svxml.execute_grid("S1", "FinMax", _lay())
    svxml.check_read_only(body)
    assert body.startswith("<req_ExecuteGrid><sID>S1</sID><action>retrieve</action>")
    assert '<aliasTableName val="Default"/>' in body and "<cube>FinMax</cube>" in body
    assert '<slice rows="4" cols="3">' in body and '<range start="0" end="11">' in body


@pytest.mark.parametrize("body", [
    "<req_WriteBack><sID>x</sID></req_WriteBack>",
    "<req_SetVals/>",
    "<req_ExecuteGrid><sID>x</sID><action>freeformsubmit</action></req_ExecuteGrid>",
    "<req_ExecuteGrid><sID>x</sID></req_ExecuteGrid>",
    "not xml",
])
def test_check_read_only_refuses_everything_else(body):
    with pytest.raises(ValueError, match="refused"):
        svxml.check_read_only(body)


def test_session_requests_pass_the_allowlist():
    for body in (svxml.connect(), svxml.provisioned_data_sources(), svxml.open_application("s", "FinMax"),
                 svxml.open_cube("s", "FinMax", "FinLite")):
        svxml.check_read_only(body)


def test_sid_and_exceptions():
    assert svxml.sid_of("<res_OpenApplication><sID>-4f:1a</sID></res_OpenApplication>") == "-4f:1a"
    with pytest.raises(svxml.SvXmlError) as e:
        svxml.sid_of('<exception errcode="1000" type="error"><desc>Grid is &lt;bad&gt;</desc></exception>')
    assert e.value.code == 1000 and str(e.value) == "Grid is <bad>"
    with pytest.raises(svxml.SvXmlError, match="not a Smart View provider answer"):
        svxml.sid_of("<html>401 Authorization Required</html>")


def test_data_sources_give_the_pod_per_app():
    text = ('<res_GetProvisionedDataSources><Product id="HP"><Server name="FinMax" '
            'context="https://plan1.example/HyperionPlanning/SmartView" appName="FinMax"/></Product>'
            '<Product id="HP"><Server name="FinPlan_for_DataLoad" '
            'context="https://strat.example/HyperionPlanning/SmartView" appName="FinPlan"/></Product>'
            '</res_GetProvisionedDataSources>')
    assert svxml.data_sources(text) == {"FinMax": "https://plan1.example", "FinPlan": "https://strat.example"}


def test_grid_values_round_trip_through_grid_parse():
    m = [[None, "Actual", None], [None, "Aug 26", "Sep 26"],
         ["A account (ac)", 100.5, 110.0], ["B account (ac)", "#noaccess", None]]
    values = svxml.grid_values(answer(m))
    assert values[2] == ["A account (ac)", 100.5, 110.0] and values[3][1] == "#NoAccess"
    p = parse(_lay(), values, suppress_missing=True, suppress_zeros=False)
    assert p.rows == [["ac.A", 100.5, 110.0], ["ac.B", "#NoAccess", None]]
    assert p.row_labels == [["A account (ac)"], ["B account (ac)"]] and p.no_access_cells == 1


def test_grid_values_raise_on_provider_exception():
    with pytest.raises(svxml.SvXmlError, match="at least 2 members"):
        svxml.grid_values('<exception errcode="1000" type="error"><desc>Grid is Invalid. It should have at '
                          'least 2 members valid for selected plan type.</desc></exception>')
```

- [ ] **Step 2: Run them and check they fail**

Run: `.venv\Scripts\python.exe -m pytest tests/test_svxml.py -q`
Expected: FAIL, `ImportError: cannot import name 'svxml'`

- [ ] **Step 3: Implement**

```python
# engine/smartview_mcp/svxml.py
"""Smart View provider XML for EPM Cloud Planning (`/HyperionPlanning/SmartView`), learnt by probing on 2026-10-25
(docs/specs/2026-10-25-stage0-epm-access-probe.md). Pure functions: request bodies and response parsing.

Read-only by construction: check_read_only() admits only the requests built here, and ExecuteGrid only with the
'retrieve' action. svhttp.BrowserTransport passes every body through it before sending."""
from __future__ import annotations

import html
import re

from .grid import Layout

ALLOWED = frozenset({"ConnectToProvider", "GetProvisionedDataSources", "OpenApplication", "OpenCube", "ListCubes",
                     "GetDefaultGrid", "ExecuteGrid"})
CLIENT_XML_VERSION = "4.2.5.7.3"
_SPECIAL = {"#noaccess": "#NoAccess", "#no access": "#NoAccess", "#missing": "#Missing", "#invalid": "#Invalid"}


class SvXmlError(Exception):
    """The provider answered with <exception>, or with something that is not provider XML."""

    def __init__(self, desc: str, code: int | None = None):
        super().__init__(desc)
        self.code = code


def check_read_only(body: str) -> None:
    m = re.match(r"\s*<req_(\w+)>", body)
    name = m.group(1) if m else body[:40]
    if not m or name not in ALLOWED:
        raise ValueError(f"refused: {name!r} is not a read-only Smart View request")
    if name == "ExecuteGrid" and re.findall(r"<action>(.*?)</action>", body) != ["retrieve"]:
        raise ValueError("refused: ExecuteGrid is only sent with the 'retrieve' action")


def _x(v) -> str:
    return html.escape(str(v), quote=False)


def connect() -> str:
    return (f'<req_ConnectToProvider><ClientXMLVersion>{CLIENT_XML_VERSION}</ClientXMLVersion><lngs enc="0">en_US'
            '</lngs><usr></usr><pwd></pwd><sso></sso></req_ConnectToProvider>')


def provisioned_data_sources() -> str:
    return '<req_GetProvisionedDataSources><usr></usr><pwd></pwd><filters></filters></req_GetProvisionedDataSources>'


def open_application(sid: str, app: str) -> str:
    return (f'<req_OpenApplication><sID>{_x(sid)}</sID><srv>{_x(app)}</srv><app>{_x(app)}</app><type></type>'
            '<url></url><sessionInfo></sessionInfo></req_OpenApplication>')


def open_cube(sid: str, app: str, cube: str) -> str:
    return (f'<req_OpenCube><sID>{_x(sid)}</sID><srv>{_x(app)}</srv><app>{_x(app)}</app><cube>{_x(cube)}</cube>'
            '<type></type><url></url></req_OpenCube>')


def grid_cells(lay: Layout) -> tuple[list[str], list[str]]:
    """Row-major cell values and types of the free-form layout: 7 = blank, 0 = member, 2 = data cell."""
    m = lay.matrix()
    data_rows = range(1 + lay.nc, len(m))
    data_cols = range(lay.nr, lay.nr + len(lay.col_tuples))
    vals, types = [], []
    for r, row in enumerate(m):
        for c, v in enumerate(row):
            text = "" if v is None else str(v)
            if "|" in text:
                raise ValueError(f"member {text!r} contains '|', which the provider uses as separator")
            vals.append(_x(text))
            types.append("2" if r in data_rows and c in data_cols else ("7" if v is None else "0"))
    return vals, types


def execute_grid(sid: str, cube: str, lay: Layout, alias: str = "Default") -> str:
    """Free-form ad hoc retrieve. Alias table 'Default' returns the labels Smart View shows in Excel."""
    m = lay.matrix()
    h, w = len(m), len(m[0])
    vals, types = grid_cells(lay)
    prefs = (f'<preferences><aliasTableName val="{_x(alias)}"/><missingLabelText val="#Missing"/>'
             '<noAccessText val="#NoAccess"/><repeatMemberLabels val="1"/></preferences>')
    return (f'<req_ExecuteGrid><sID>{_x(sid)}</sID><action>retrieve</action><ords></ords>{prefs}<grid>'
            f'<cube>{_x(cube)}</cube><gridType>1</gridType><dims/><slices><slice rows="{h}" cols="{w}"><data>'
            f'<range start="0" end="{h * w - 1}"><vals enc="1">{"|".join(vals)}</vals>'
            f'<types enc="0">{"|".join(types)}</types></range></data></slice></slices></grid></req_ExecuteGrid>')


def raise_for_exception(text: str) -> None:
    m = re.search(r'<exception\s+errcode="(-?\d+)"[^>]*>\s*<desc>(.*?)</desc>', text, re.S)
    if m:
        raise SvXmlError(html.unescape(m.group(2)).strip(), int(m.group(1)))
    if "<res_" not in text:
        raise SvXmlError("not a Smart View provider answer: " + re.sub(r"\s+", " ", text)[:120])


def sid_of(text: str) -> str:
    raise_for_exception(text)
    m = re.search(r"<sID>(.*?)</sID>", text, re.S)
    if not m:
        raise SvXmlError("no session id (sID) in the provider answer")
    return m.group(1)


def data_sources(text: str) -> dict[str, str]:
    """{app: pod base URL} from res_GetProvisionedDataSources."""
    raise_for_exception(text)
    out = {}
    for m in re.finditer(r'<Server\b[^>]*\bcontext="([^"]+)"[^>]*\bappName="([^"]+)"', text):
        out[html.unescape(m.group(2))] = html.unescape(m.group(1)).split("/HyperionPlanning/")[0]
    return out


def _cell(v: str, t: str):
    if v == "":
        return None
    special = _SPECIAL.get(v.strip().casefold())
    if special:
        return special
    if t == "2":
        try:
            return float(v)
        except ValueError:
            return v
    return v


def grid_values(text: str) -> list[list]:
    """The returned slice as a dense 2D list (like Excel's UsedRange.Value2): floats for data, None for blanks."""
    raise_for_exception(text)
    s = re.search(r'<slice\s+rows="(\d+)"\s+cols="(\d+)"', text)
    r = re.search(r'<range\s+start="(\d+)"\s+end="(\d+)"', text)
    v = re.search(r"<vals[^>]*>(.*?)</vals>", text, re.S)
    t = re.search(r"<types[^>]*>(.*?)</types>", text, re.S)
    if not (s and r and v and t):
        raise SvXmlError("no grid in the provider answer")
    rows, cols, start = int(s.group(1)), int(s.group(2)), int(r.group(1))
    vals = [html.unescape(x) for x in v.group(1).split("|")]
    types = t.group(1).split("|")
    out: list[list] = [[None] * cols for _ in range(rows)]
    for i, (x, ty) in enumerate(zip(vals, types)):
        k = start + i
        out[k // cols][k % cols] = _cell(x, ty)
    return out
```

- [ ] **Step 4: Run the tests again**

Run: `.venv\Scripts\python.exe -m pytest tests/test_svxml.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/svxml.py engine/tests/svxml_util.py engine/tests/test_svxml.py
git commit -m "svxml: read-only Smart View provider XML (build, allowlist, parse)"
```

### Task 3: drill browser, sign in on another pod and share one browser

**Files:**
- Modify: `engine\smartview_mcp\drill_browser.py:436-466` (`_login_once`), append `shared_browser()` at the end
- Modify: `engine\smartview_mcp\tools_drill.py:53-58`
- Test: `engine\tests\test_drill.py` (append)

- [ ] **Step 1: Write the failing tests** (append to `engine\tests\test_drill.py`)

```python
class _FakeLoginPage:
    def __init__(self):
        self.url = "about:blank"

    def goto(self, url, wait_until=None, timeout=None):
        self.url = url

    def wait_for_load_state(self, *a, **k):
        pass

    def wait_for_timeout(self, ms):
        pass


def test_login_once_signs_in_on_another_pod(tmp_path):
    from smartview_mcp import drill_browser as db
    b = db.DrillBrowser(base="https://plan1.example", profile=tmp_path / "p")
    p = _FakeLoginPage()
    b._login_once(p, float("inf"), base="https://strat.example/")
    assert p.url == "https://strat.example/HyperionPlanning/"
    b._login_once(p, float("inf"))
    assert p.url == "https://plan1.example/HyperionPlanning/"


def test_shared_browser_is_one_instance(monkeypatch):
    from smartview_mcp import drill_browser as db, tools_drill
    monkeypatch.setattr(db, "_shared", None)
    made = []
    monkeypatch.setattr(db, "DrillBrowser", lambda: made.append(1) or object())
    first = db.shared_browser()
    assert db.shared_browser() is first and tools_drill.browser() is first and made == [1]
```

- [ ] **Step 2: Run them and check they fail**

Run: `.venv\Scripts\python.exe -m pytest tests/test_drill.py -q -k "another_pod or shared_browser"`
Expected: FAIL (`_login_once() got an unexpected keyword argument 'base'`; `no attribute '_shared'`)

- [ ] **Step 3: Implement**

In `drill_browser.py`, change `_login_once` (lines 436-466). Only the signature, `home` and the `login_action` call change:

```python
    def _login_once(self, p, deadline: float, base: str | None = None):
        """Open the EPM home of `base` (default: this browser's pod; signs in via Telenor SSO if needed). The drill
        page needs its session state; the Smart View HTTP route signs in on other pods (FinPlan: stratmodel)."""
        base = (base or self.base).rstrip("/")
        home = base + "/HyperionPlanning/"
        self._goto(p, home, deadline)
        t0, clicked, nudged = time.monotonic(), False, False
        last, changed = None, t0
        end = min(t0 + config.DRILL_LOGIN_TIMEOUT_S, deadline)
        while time.monotonic() < end:
            url, now = p.url, time.monotonic()
            if url != last:
                last, changed = url, now
            act = login_action(url, base, now - t0, now - changed, clicked, nudged, deadline - now)
```

The rest of the method stays as it is. Then append at the end of `drill_browser.py`:

```python
_shared: DrillBrowser | None = None
_shared_lock = threading.Lock()


def shared_browser() -> DrillBrowser:
    """The process's one drill browser: sv_drill and the Smart View HTTP route (svhttp) share it, its worker thread
    and its profile lock."""
    global _shared
    with _shared_lock:
        if _shared is None:
            _shared = DrillBrowser()
        return _shared
```

In `tools_drill.py`, replace lines 53-58 (`def browser(): ...`) and remove the now-unused `_browser = None` on line 48:

```python
def browser():
    from .drill_browser import shared_browser
    return shared_browser()
```

- [ ] **Step 4: Run the drill tests**

Run: `.venv\Scripts\python.exe -m pytest tests/test_drill.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/drill_browser.py engine/smartview_mcp/tools_drill.py engine/tests/test_drill.py
git commit -m "drill browser: sign in on other pods; one shared browser per process"
```

### Task 4: `svhttp.SvHttp`, sessions and the retrieve contract

**Files:**
- Create: `engine\smartview_mcp\svhttp.py`
- Modify: `engine\smartview_mcp\config.py` (after `DRILL_LOGIN_TIMEOUT_S`, line ~222)
- Test: `engine\tests\test_svhttp.py` (extend)

- [ ] **Step 1: Write the failing tests**

Replace the import line at the top of `engine\tests\test_svhttp.py` with the block below, then append the tests.

```python
import re

import pytest
from svxml_util import answer

from smartview_mcp import atlas, svhttp, svxml
from smartview_mcp.grid import Axis
from smartview_mcp.host import InputError

SOURCES = ('<res_GetProvisionedDataSources><Product id="HP"><Server name="FinMax" '
           'context="https://plan1.example/HyperionPlanning/SmartView" appName="FinMax"/></Product>'
           '<Product id="HP"><Server name="FinPlan_x" context="https://strat.example/HyperionPlanning/SmartView" '
           'appName="FinPlan"/></Product></res_GetProvisionedDataSources>')
EXPIRED = '<exception errcode="1000" type="error"><desc>session expired</desc></exception>'


class FakeTransport:
    """Answers like the provider (shapes from the live probe; figures invented)."""

    def __init__(self, *grid_answers):
        self.sent, self.answers = [], list(grid_answers)

    def post(self, base, body):
        svxml.check_read_only(body)
        tag = re.match(r"<req_(\w+)>", body).group(1)
        self.sent.append((base, tag))
        if tag == "GetProvisionedDataSources":
            return SOURCES
        if tag in ("ConnectToProvider", "OpenApplication"):
            return f"<res_{tag}><sID>sid{len(self.sent)}</sID></res_{tag}>"
        if tag == "OpenCube":
            return "<res_OpenCube></res_OpenCube>"
        return self.answers.pop(0)

    def tags(self):
        return [t for _, t in self.sent]


GRID = [[None, "2WHS Wholesale (un.SY)", "Actual"], [None, "Aug 26", "Sep 26"],
        ["C30 REVENUES (ac)", 100.0, 110.0], ["EBITDA EBITDA (ac)", 50.0, None]]


@pytest.fixture
def db_path(tmp_path):
    db = atlas.connect(tmp_path / "a.db")
    _store(db, "FinMax", ["All Accounts", "All Entities", "All Scenarios", "All Time Periods"])
    _store(db, "FinPlan", ["Account", "Entity", "Period", "Scenario", "Version", "Years"])
    return tmp_path / "a.db"


def _q(sv, conn="FinMax"):
    return sv.retrieve(conn, {"Entities": "un.SY.2WHS", "All Scenarios": "Actual"},
                       [Axis("All Accounts", ["ac.C30", "ac.EBITDA"])], [Axis("All Time Periods", ["Aug 26", "Sep 26"])],
                       True, False)


def test_retrieve_returns_the_com_contract(db_path):
    t = FakeTransport(answer(GRID))
    lay, p, secs = _q(svhttp.SvHttp(t, home="https://plan1.example", db_path=db_path))
    assert lay.pov == [("All Entities", "un.SY.2WHS"), ("All Scenarios", "Actual")]
    assert p.rows == [["ac.C30", 100.0, 110.0], ["ac.EBITDA", 50.0, None]]
    assert p.row_labels == [["C30 REVENUES (ac)"], ["EBITDA EBITDA (ac)"]] and secs >= 0
    assert t.tags() == ["GetProvisionedDataSources", "ConnectToProvider", "OpenApplication", "OpenCube", "ExecuteGrid"]
    assert t.sent[-1][0] == "https://plan1.example"


def test_session_is_reused_and_finplan_goes_to_its_pod(db_path):
    t = FakeTransport(answer(GRID), answer(GRID), answer([[None, "x"], [None, "Sep"], ["Revenue", 1.0]]))
    sv = svhttp.SvHttp(t, home="https://plan1.example", db_path=db_path)
    _q(sv)
    _q(sv)
    assert t.tags().count("ConnectToProvider") == 1
    sv.retrieve("FinPlan", {"Entity": "2WHS Wholesale"}, [Axis("Account", ["Revenue"])], [Axis("Period", ["Sep"])],
                True, False)
    assert t.sent[-1] == ("https://strat.example", "ExecuteGrid") and t.tags().count("ConnectToProvider") == 2


def test_one_reconnect_then_a_clear_error(db_path):
    t = FakeTransport(EXPIRED, answer(GRID))
    _q(svhttp.SvHttp(t, home="https://plan1.example", db_path=db_path))
    assert t.tags().count("ConnectToProvider") == 2
    t2 = FakeTransport(EXPIRED, EXPIRED)
    with pytest.raises(svhttp.SvHttpError, match="session expired"):
        _q(svhttp.SvHttp(t2, home="https://plan1.example", db_path=db_path))


def test_input_errors_and_missing_atlas(db_path, tmp_path):
    sv = svhttp.SvHttp(FakeTransport(), home="https://plan1.example", db_path=db_path)
    with pytest.raises(InputError, match="unknown dimension 'All Bogus'"):
        sv.retrieve("FinMax", {"All Bogus": "x"}, [Axis("All Accounts", ["a"])], [Axis("All Time Periods", ["p"])],
                    True, False)
    with pytest.raises(svhttp.SvHttpError, match="no HTTP route"):
        sv.retrieve("Other", {}, [], [], True, False)
    empty = svhttp.SvHttp(FakeTransport(), home="https://plan1.example", db_path=tmp_path / "empty.db")
    with pytest.raises(svhttp.SvHttpError, match="atlas has no dimensions"):
        _q(empty, "FinLite")
```

- [ ] **Step 2: Run them and check they fail**

Run: `.venv\Scripts\python.exe -m pytest tests/test_svhttp.py -q`
Expected: FAIL, `ImportError: cannot import name 'svhttp'`

- [ ] **Step 3: Implement**

Add to `config.py` after `DRILL_LOGIN_TIMEOUT_S`:

```python
SVHTTP_TIMEOUT_S = float(os.environ.get("SV_HTTP_TIMEOUT", "120"))   # one Smart View HTTP request incl. a sign-in
```

Create `engine\smartview_mcp\svhttp.py`:

```python
"""Grids over the Smart View provider's HTTP protocol - no Excel, no COM (Stage 0, 2026-10-25:
docs/specs/2026-10-25-stage0-epm-access-probe.md). retrieve() returns what ops._query_once returns,
(Layout, Parsed, seconds), so sv_query can use either route; the COM route stays the fallback.

The transport is BrowserTransport: it posts from inside the drill's hidden Edge (Telenor SSO), one tab per EPM pod,
on the drill browser's worker thread. Tests use a fake transport."""
from __future__ import annotations

import logging
import threading
import time
from pathlib import Path
from typing import Protocol

from . import atlas, config, svxml
from .grid import Axis, Layout, full_pov, parse, validate_axes
from .host import InputError

log = logging.getLogger(__name__)

CUBES = {"FinMax": ("FinMax", "FinMax"), "FinLite": ("FinMax", "FinLite"), "FinPlan": ("FinPlan", "FinPlan")}


class SvHttpError(Exception):
    """The HTTP route could not answer; sv_query(source='auto') then uses Smart View in Excel."""


class Transport(Protocol):
    def post(self, base: str, body: str) -> str: ...


def _norm_dim(name: str) -> str:
    return name.removeprefix("All ").replace(" ", "").casefold()


class SvHttp:
    def __init__(self, transport: Transport, home: str | None = None, db_path: Path | None = None):
        self.t = transport
        self.home = (home or config.EPM_BASE).rstrip("/")
        self.db_path = db_path or config.ATLAS_DB
        self._bases: dict[str, str] | None = None
        self._sids: dict[str, str] = {}
        self._lock = threading.Lock()

    def base_for(self, app: str) -> str:
        if self._bases is None:
            self._bases = svxml.data_sources(self.t.post(self.home, svxml.provisioned_data_sources()))
        if app not in self._bases:
            raise SvHttpError(f"EPM lists no Smart View data source for app {app} "
                              f"(has: {', '.join(sorted(self._bases))})")
        return self._bases[app]

    def _session(self, conn: str) -> tuple[str, str, str]:
        app, cube = CUBES[conn]
        base = self.base_for(app)
        if conn not in self._sids:
            sid = svxml.sid_of(self.t.post(base, svxml.connect()))
            sid = svxml.sid_of(self.t.post(base, svxml.open_application(sid, app)))
            svxml.raise_for_exception(self.t.post(base, svxml.open_cube(sid, app, cube)))
            self._sids[conn] = sid
        return base, cube, self._sids[conn]

    def dimensions(self, conn: str) -> list[str]:
        dims = atlas.dimensions(atlas.connect(self.db_path), conn)
        if not dims:
            raise SvHttpError(f"the atlas has no dimensions for {conn}: build it with sv_atlas_build")
        return dims

    def retrieve(self, conn: str, pov: dict | None, rows: list[Axis], cols: list[Axis],
                 suppress_missing: bool, suppress_zeros: bool):
        if conn not in CUBES:
            raise SvHttpError(f"no HTTP route for connection {conn}")
        t0 = time.monotonic()
        dims = self.dimensions(conn)
        by_norm = {_norm_dim(d): d for d in dims}

        def canon(d: str) -> str:
            return by_norm.get(_norm_dim(d), d)
        pov = {canon(d): m for d, m in (pov or {}).items()}
        rows = [Axis(canon(a.dimension), a.members) for a in rows]
        cols = [Axis(canon(a.dimension), a.members) for a in cols]
        problems = validate_axes(dims, pov, rows, cols)
        if problems:
            raise InputError("Invalid query: " + "; ".join(problems))
        lay = Layout(full_pov(dims, pov, rows, cols), rows, cols)
        with self._lock:
            values = self._execute(conn, lay)
        return lay, parse(lay, values, suppress_missing, suppress_zeros), time.monotonic() - t0

    def _execute(self, conn: str, lay: Layout) -> list[list]:
        for attempt in (1, 2):
            try:
                base, cube, sid = self._session(conn)
                return svxml.grid_values(self.t.post(base, svxml.execute_grid(sid, cube, lay)))
            except svxml.SvXmlError as e:
                self._sids.pop(conn, None)                   # maybe an expired session: reconnect once
                if attempt == 2:
                    raise SvHttpError(f"Smart View provider: {e}") from None
                log.info("svhttp %s: %s - reconnecting once", conn, e)
        raise AssertionError("unreachable")
```

- [ ] **Step 4: Run the tests again**

Run: `.venv\Scripts\python.exe -m pytest tests/test_svhttp.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/svhttp.py engine/smartview_mcp/config.py engine/tests/test_svhttp.py
git commit -m "svhttp: Smart View HTTP retrieve with the ops.query contract"
```

### Task 5: `BrowserTransport` and `shared()`

**Files:**
- Modify: `engine\smartview_mcp\svhttp.py` (append)
- Test: `engine\tests\test_svhttp.py` (append; the real browser is exercised in Task 7)

- [ ] **Step 1: Write the failing tests**

```python
class _Page:
    def __init__(self, replies):
        self.replies, self.bodies = list(replies), []

    def is_closed(self):
        return False

    def evaluate(self, js, body):
        self.bodies.append(body)
        return self.replies.pop(0)


class _Browser:
    base = "https://plan1.example"

    def __init__(self, page):
        self.s = type("S", (), {"page": page, "ctx": object(), "ready": True, "deadline": 0})()
        self.logins = 0

    def submit(self, fn):
        import concurrent.futures as cf
        f = cf.Future()
        f.set_result(fn(self.s))
        return f

    def _ensure(self, s):
        if not s.ready:
            self.logins += 1
            s.ready = True


def test_browser_transport_refuses_writes_and_signs_in_again_once():
    page = _Page([[401, "<html>login</html>"], [200, "<res_OpenCube></res_OpenCube>"]])
    b = _Browser(page)
    t = svhttp.BrowserTransport(b)
    with pytest.raises(ValueError, match="refused"):
        t.post(b.base, "<req_WriteBack/>")
    assert page.bodies == []
    assert t.post(b.base, svxml.open_cube("s", "FinMax", "FinMax")) == "<res_OpenCube></res_OpenCube>"
    assert b.logins == 1


def test_browser_transport_gives_up_after_two_bad_answers():
    b = _Browser(_Page([[401, "x"], [401, "x"]]))
    with pytest.raises(svhttp.SvHttpError, match="HTTP 401 twice"):
        svhttp.BrowserTransport(b).post(b.base, svxml.connect())


def test_shared_is_one_instance(monkeypatch):
    from smartview_mcp import drill_browser
    monkeypatch.setattr(svhttp, "_shared", None)
    monkeypatch.setattr(drill_browser, "shared_browser", lambda: _Browser(_Page([])))
    assert svhttp.shared() is svhttp.shared()
```

- [ ] **Step 2: Run them and check they fail**

Run: `.venv\Scripts\python.exe -m pytest tests/test_svhttp.py -q -k "transport or shared"`
Expected: FAIL, `AttributeError: module 'smartview_mcp.svhttp' has no attribute 'BrowserTransport'`

- [ ] **Step 3: Implement** (append to `svhttp.py`)

```python
_JS = """async (b) => { const r = await fetch('/HyperionPlanning/SmartView', {method: 'POST', credentials: 'include',
    body: b, headers: {'Content-Type': 'text/xml'}}); return [r.status, await r.text()]; }"""


class BrowserTransport:
    """Posts provider XML from a signed-in tab of the drill browser: its own tab for its pod (plan1) and one extra
    tab per other pod (FinPlan: stratmodel), signed in once with the same SSO. All page work runs on the browser's
    worker thread via submit(); _tabs is only touched there."""

    def __init__(self, browser):
        self.b = browser
        self._tabs: dict[str, tuple[object, object]] = {}       # pod base -> (browser context, page)

    def post(self, base: str, body: str) -> str:
        svxml.check_read_only(body)
        return self.b.submit(lambda s: self._post(s, base.rstrip("/"), body)).result(timeout=config.SVHTTP_TIMEOUT_S)

    def _tab(self, s, base: str):
        if base == self.b.base.rstrip("/"):
            return s.page
        ctx, page = self._tabs.get(base, (None, None))
        if ctx is not s.ctx or page is None or page.is_closed():
            page = s.ctx.new_page()
            self.b._login_once(page, s.deadline, base=base)
            self._tabs[base] = (s.ctx, page)
        return page

    def _post(self, s, base: str, body: str) -> str:
        s.deadline = time.monotonic() + config.SVHTTP_TIMEOUT_S
        status = None
        for attempt in (1, 2):
            self.b._ensure(s)
            try:
                status, text = self._tab(s, base).evaluate(_JS, body)
            except Exception as e:   # noqa: BLE001 - the tab navigated or the browser died: one fresh try
                if attempt == 2:
                    raise SvHttpError(f"browser error at {base}: {str(e).splitlines()[0][:200]}") from None
                s.ready = False
                self._tabs.pop(base, None)
                continue
            if status == 200 and ("<res_" in text or "<exception" in text):
                return text
            log.warning("svhttp: HTTP %s from %s; signing in again", status, base)
            s.ready = False
            self._tabs.pop(base, None)
        raise SvHttpError(f"Smart View provider at {base} answered HTTP {status} twice (sign-in failed?)")


_shared: SvHttp | None = None
_shared_lock = threading.Lock()


def shared() -> SvHttp:
    """The process's one HTTP route (on the shared drill browser)."""
    global _shared
    with _shared_lock:
        if _shared is None:
            from .drill_browser import shared_browser
            _shared = SvHttp(BrowserTransport(shared_browser()))
        return _shared
```

- [ ] **Step 4: Run all svhttp tests**

Run: `.venv\Scripts\python.exe -m pytest tests/test_svhttp.py tests/test_svxml.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/svhttp.py engine/tests/test_svhttp.py
git commit -m "svhttp: browser transport (one tab per pod, one re-sign-in) and shared route"
```

### Task 6: `sv_query` HTTP first, COM fallback

**Files:**
- Modify: `engine\smartview_mcp\tools_query.py` (imports; new `_retrieve` above `sv_query`; `sv_query` body lines 29-107)
- Test: `engine\tests\test_retrieve_route.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/test_retrieve_route.py
import pytest
from mcp.server.mcpserver.exceptions import ToolError

from smartview_mcp import svhttp, tools_query
from smartview_mcp.host import InputError


class _Http:
    def __init__(self, result=None, exc=None):
        self.result, self.exc = result, exc

    def retrieve(self, *a):
        if self.exc:
            raise self.exc
        return self.result


def _route(monkeypatch, http, com=("L", "P", 1.0)):
    calls = []

    async def fake_call(fn, timeout=None, what="call"):
        calls.append(what)
        return com
    monkeypatch.setattr(svhttp, "shared", lambda: http)
    monkeypatch.setattr(tools_query, "_call", fake_call)
    return calls


@pytest.mark.anyio
async def test_http_first(monkeypatch):
    calls = _route(monkeypatch, _Http(result=("l", "p", 0.2)))
    assert await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1) == ("l", "p", 0.2, "http", None)
    assert calls == []


@pytest.mark.anyio
async def test_auto_falls_back_to_com(monkeypatch):
    calls = _route(monkeypatch, _Http(exc=svhttp.SvHttpError("provider down")))
    lay, p, secs, used, note = await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1)
    assert (lay, used, calls) == ("L", "com", ["sv_query"]) and "provider down" in note


@pytest.mark.anyio
async def test_http_only_never_falls_back(monkeypatch):
    calls = _route(monkeypatch, _Http(exc=svhttp.SvHttpError("provider down")))
    with pytest.raises(ToolError, match="source='com'"):
        await tools_query._retrieve("FinMax", {}, [], [], True, False, "http", 1)
    assert calls == []


@pytest.mark.anyio
async def test_input_errors_are_not_retried_in_excel(monkeypatch):
    calls = _route(monkeypatch, _Http(exc=InputError("Invalid query: unknown dimension 'x'")))
    with pytest.raises(ToolError, match="unknown dimension"):
        await tools_query._retrieve("FinMax", {}, [], [], True, False, "auto", 1)
    assert calls == []


@pytest.mark.anyio
async def test_com_skips_http(monkeypatch):
    calls = _route(monkeypatch, _Http(exc=AssertionError("HTTP must not run")))
    assert (await tools_query._retrieve("FinMax", {}, [], [], True, False, "com", 1))[3] == "com"
    assert calls == ["sv_query"]
```

- [ ] **Step 2: Run them and check they fail**

Run: `.venv\Scripts\python.exe -m pytest tests/test_retrieve_route.py -q`
Expected: FAIL, `AttributeError: module 'smartview_mcp.tools_query' has no attribute '_retrieve'`

- [ ] **Step 3: Implement**

In `tools_query.py`:
- add `import logging` and `import anyio.to_thread` to the imports;
- change `from . import atlas, config, ops, qlog` to `from . import atlas, config, logsetup, ops, qlog`;
- add `from .host import InputError` next to `from .host import HostError`;
- add `log = logging.getLogger(__name__)` after the imports.

Then add above `sv_query`:

```python
async def _retrieve(connection, pov, rows_ax, cols_ax, suppress_missing, suppress_zeros, source: str, timeout: float):
    """(layout, parsed, seconds, route, note). Smart View's HTTP protocol first (no Excel) unless source='com'; with
    source='auto' an HTTP failure that is not an input error falls back to Smart View in Excel (note says why)."""
    note = None
    if source != "com":
        from . import svhttp
        try:
            lay, parsed, secs = await anyio.to_thread.run_sync(lambda: svhttp.shared().retrieve(
                connection, pov, rows_ax, cols_ax, suppress_missing, suppress_zeros))
            return lay, parsed, secs, "http", None
        except InputError as e:
            raise _err(e) from None
        except Exception as e:   # noqa: BLE001 - any HTTP-route failure: the Excel route, or a clear error for 'http'
            msg = f"{type(e).__name__}: {logsetup.redact(e)}"
            if source == "http":
                raise ToolError(f"HTTP route failed - {msg}. Next step: call sv_query with source='com'.") from None
            log.warning("svhttp failed, using Smart View in Excel: %s", msg)
            note = f"HTTP route failed ({msg}); the grid was read through Smart View in Excel instead."
    lay, parsed, secs = await _call(
        lambda h: ops.query(h, connection, pov, rows_ax, cols_ax, suppress_missing, suppress_zeros),
        timeout=timeout, what="sv_query")
    return lay, parsed, secs, "com", note
```

Change `sv_query`:

1. **Signature.** Add the last parameter `source: Literal["auto", "http", "com"] = "auto"`.

2. **Docstring.** Append:
   ```
   source: 'auto' (default) reads over Smart View's HTTP protocol (no Excel, ~1 s; signs in through the hidden
   browser) and falls back to Smart View in Excel; 'com' forces Excel; 'http' never falls back.
   ```

3. **Retrieve.** Replace the `try: lay, parsed, secs = await _call(...)` block (lines 57-64) with:

```python
    try:
        lay, parsed, secs, route, http_note = await _retrieve(
            connection, pov, rows_ax, cols_ax, suppress_missing, suppress_zeros, source,
            timeout=max(config.CALL_TIMEOUT_S, n_cells / 200))
    except (HostError, ToolError) as e:
        qlog.write({"tool": "sv_query", "connection": connection, "pov": pov, "rows": [r.model_dump() for r in rows],
                    "columns": [c.model_dump() for c in columns], "source": source, "error": str(e)})
        raise _err(e)
```

4. **Query log.** In the success `qlog.write({...})`, add `"route": route,`.

5. **Warnings.** After the `warnings = ...` / `if parsed.unknown_members:` block, add:
   ```python
       if http_note:
           warnings.insert(0, http_note)
   ```

6. **Return value.** Add `"route": route,` to the returned dict.

- [ ] **Step 4: Run the full suite**

Run: `.venv\Scripts\python.exe -m pytest -q`. Expected: all passed (295 + new).
Run: `.venv\Scripts\python.exe -m ruff check .`. Expected: `All checks passed!`

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/tools_query.py engine/tests/test_retrieve_route.py
git commit -m "sv_query: Smart View HTTP route first, Excel as fallback (source=auto|http|com)"
```

### Task 7: live check, HTTP equals COM cell for cell

**Files:**
- Create: `engine\tests\live\svhttp_vs_com.py`
- Modify: `docs\specs\2026-10-25-stage0-epm-access-probe.md` (results), `skills\smartass\reference\operations.md`

- [ ] **Step 1: Write the live script**

```python
# engine/tests/live/svhttp_vs_com.py
"""Live: every grid read over Smart View's HTTP protocol equals the COM route (hidden Excel) cell for cell.
Run from engine\\: .venv\\Scripts\\python.exe tests\\live\\svhttp_vs_com.py  (needs SSO; prints no figures)."""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from smartview_mcp import ops, svhttp  # noqa: E402
from smartview_mcp.grid import Axis  # noqa: E402
from smartview_mcp.host import ExcelHost  # noqa: E402

PNL = ["ac.C30", "ac.C40S", "ac.C48", "ac.C50S", "ac.C60", "ac.C70", "ac.EBITDA_BOI", "ac.C73", "ac.EBITDA", "ac.C75",
       "ac.STA20Y"]
BA = ["un.SY.2WHS", "un.SY.2BUS", "un.SY.2SSM", "un.SY.2EI", "un.SY.2TEC", "un.SY.2IT", "un.SY.2ADM"]
MONTHS = ["Jan 26", "Feb 26", "Mar 26", "Apr 26", "May 26", "Jun 26", "Jul 26", "Aug 26", "Sep 26"]
ACT = {"All Scenarios": "Actual", "All Currencies": "NOK"}
CASES = {
    "FinMax 2x2 periodic": ("FinMax", {**ACT, "All Entities": "un.SY.2WHS", "All Balance Types": "Periodic"},
                            [Axis("All Accounts", ["ac.C30", "ac.EBITDA"])], [Axis("All Time Periods", MONTHS[-2:])]),
    "FinMax BA x P&L x 9 months": ("FinMax", {**ACT, "All Balance Types": "Periodic"},
                                   [Axis("All Entities", BA), Axis("All Accounts", PNL)],
                                   [Axis("All Time Periods", MONTHS)]),
    "FinMax YTD": ("FinMax", {**ACT, "All Entities": "un.SY.2WHS", "All Balance Types": "Accumulated"},
                   [Axis("All Accounts", PNL)], [Axis("All Time Periods", ["Sep 26"])]),
    "FinLite cost base": ("FinLite", {**ACT, "All Balance Types": "Periodic"},
                          [Axis("All Entities", BA)], [Axis("All Time Periods", MONTHS[-3:])]),
    "FinPlan actual vs target": ("FinPlan", {"Entity": "2WHS Wholesale", "Years": "FY26", "Period": "Sep",
                                             "All Balance Types": "Accumulated", "Version": "Final"},
                                 [Axis("Account", ["Revenue", "COGS", "OPEX", "EBITDA"])],
                                 [Axis("Scenario", ["Target"])]),
}


def same(a, b) -> bool:
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) <= 0.005
    return a == b


def main() -> int:
    http, host, bad = svhttp.shared(), ExcelHost().ensure(), 0
    try:
        for name, (conn, pov, rows, cols) in CASES.items():
            t0 = time.monotonic()
            _, ph, _ = http.retrieve(conn, pov, rows, cols, False, False)
            th = time.monotonic() - t0
            t0 = time.monotonic()
            _, pc, _ = host.on_host(lambda: ops.query(host, conn, pov, rows, cols, False, False), timeout=600)
            tc = time.monotonic() - t0
            cells = sum(len(r) for r in pc.rows)
            diff = [(i, j) for i, (rh, rc) in enumerate(zip(ph.rows, pc.rows)) for j, (a, b) in enumerate(zip(rh, rc))
                    if not same(a, b)]
            ok = len(ph.rows) == len(pc.rows) and not diff and ph.row_labels == pc.row_labels
            bad += not ok
            print(f"{'PASS' if ok else 'FAIL'}  {name:30} {cells:6} cells  http {th:5.1f}s  com {tc:5.1f}s"
                  + ("" if ok else f"  first differing cells: {diff[:5]}"))
    finally:
        host.shutdown()
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 2: Run it**

Run: `.venv\Scripts\python.exe tests\live\svhttp_vs_com.py`
Expected: five `PASS` lines; HTTP in seconds, COM in tens of seconds.
- **If a case fails:** fix `svxml`/`svhttp` (add a unit test that reproduces the difference first) and rerun.
- **If FinPlan cannot be fixed:** set the default to `source="com"` for FinPlan in `_retrieve` (`if connection == "FinPlan" and source == "auto": source = "com"`) and say so in the stage 0 note.

- [ ] **Step 3: Record and document**
- Add a section "Live comparison (Task 7)" to `docs\specs\2026-10-25-stage0-epm-access-probe.md` with the five
  PASS/FAIL lines: names, cell counts and timings, **no figures**.
- In `skills\smartass\reference\operations.md`, under "Working in the user's Excel", add one bullet:
  ```
  - sv_query reads over Smart View's HTTP protocol by default (`source='auto'`: ~1 s, no Excel; signs in through the
    drill's hidden Edge) and falls back to Smart View in Excel. The answer's `route` says which was used. Workbooks
    (sv_ws_*) still build live Smart View sheets in Excel.
  ```

- [ ] **Step 4: Commit and push**

```bash
git add engine/tests/live/svhttp_vs_com.py docs/specs/2026-10-25-stage0-epm-access-probe.md skills/smartass/reference/operations.md
python scripts/check_staged.py
git commit -m "live: Smart View HTTP route equals COM cell for cell; docs"
git push
```

---

## Part B: Semantic catalogue

### Task 8: `knowledge\catalog.toml` + `catalog.load`

**Files:**
- Create: `knowledge\catalog.toml`
- Create: `engine\smartview_mcp\catalog.py`
- Test: `engine\tests\test_catalog.py`

- [ ] **Step 1: Write the failing tests**

```python
# engine/tests/test_catalog.py
import pytest

from smartview_mcp import catalog

OLD_PNL = [
    ("rev", "ac.C30", "line", None), ("cogs", "ac.C40S", "line", None), ("gm", "ac.C48", "sum", None),
    ("pers", "ac.C50S", "line", None), ("oth", "ac.C60", "line", None), ("opex", "ac.C70", "sum", None),
    ("boi", "ac.EBITDA_BOI", "sum", None), ("c73", "ac.C73", "line", None), ("ebitda", "ac.EBITDA", "grand", None),
    ("da", "ac.C75", "line", None), ("fte", "ac.STA20Y", "line", "tall1"),
]
OLD_PNL_CHECKS = [("C30 + C40S = C48", "rev+cogs-gm"), ("C50S + C60 = C70", "pers+oth-opex"),
                  ("C48 + C70 = EBITDA_BOI", "gm+opex-boi"), ("EBITDA_BOI + C73 = EBITDA", "boi+c73-ebitda")]
OLD_FP_PNL = [("rev", "Revenue", "line"), ("cogs", "COGS", "line"), ("gp", "Gross_Profit", "sum"),
              ("opex", "OPEX", "line"), ("boi", "EBITDA before Other Items", "sum"), ("oth", "Total Other Items", "line"),
              ("ebitda", "EBITDA", "grand")]
OLD_FP_CHECKS = [("Revenue + COGS = Gross profit", "rev+cogs-gp"),
                 ("Gross profit + OPEX = EBITDA before other items", "gp+opex-boi"),
                 ("EBITDA before other items + Other items = EBITDA", "boi+oth-ebitda")]


def test_repo_catalogue_carries_the_template_lists_unchanged():
    cat = catalog.load()
    assert [(x.key, x.account, x.kind, x.num) for x in cat.pnl["FinMax"]] == OLD_PNL
    assert [(c.label, c.expr) for c in cat.checks["FinMax"]] == OLD_PNL_CHECKS
    assert [(x.key, x.account, x.kind) for x in cat.pnl["FinPlan"]] == OLD_FP_PNL
    assert [(c.label, c.expr) for c in cat.checks["FinPlan"]] == OLD_FP_CHECKS


def test_repo_catalogue_has_areas_scenarios_kpis_materiality():
    cat = catalog.load()
    keys = {e.key for e in cat.entities}
    assert {"company", "wholesale", "business", "ssm", "ei", "technology", "it", "adm", "consumer"} <= keys
    assert cat.scenarios["budget"].scenario == "Target" and cat.scenarios["forecast"].rule == "latest_forecast"
    assert {k.key for k in cat.kpis} >= {"ebitda_margin", "opex_per_fte"}
    assert cat.materiality.pct == 0.10 and "ebitda" in cat.materiality.always


BASE = {
    "pnl": {"FinMax": [{"key": "rev", "account": "ac.C30", "kind": "line", "label": "Revenue"},
                       {"key": "ebitda", "account": "ac.EBITDA", "kind": "grand", "label": "EBITDA"}]},
    "checks": {"FinMax": [{"label": "x", "expr": "rev-ebitda"}]},
    "scenarios": {"actual": {"connection": "FinMax", "scenario": "Actual", "label": "Actual"}},
    "materiality": {"pct": 0.1, "abs_nok": 1, "always": ["rev"]},
    "entity": [{"key": "w", "label": "W", "finmax": "un.SY.2WHS", "aliases": ["wholesale"]}],
    "kpi": [{"key": "m", "label": "M", "connection": "FinMax", "numerator": "ebitda", "denominator": "rev",
             "unit": "pct"}],
}


def _with(path, value):
    import copy
    d = copy.deepcopy(BASE)
    node = d
    for k in path[:-1]:
        node = node[k]
    node[path[-1]] = value
    return d


@pytest.mark.parametrize("data, msg", [
    (_with(["pnl", "FinMax", 1, "kind"], "total"), "kind"),
    (_with(["pnl", "FinMax", 1, "key"], "rev"), "duplicate"),
    (_with(["checks", "FinMax", 0, "expr"], "rev-nope"), "nope"),
    (_with(["kpi", 0, "denominator"], "nope"), "nope"),
    (_with(["materiality", "always"], ["nope"]), "nope"),
    (_with(["entity"], BASE["entity"] * 2), "duplicate"),
    (_with(["scenarios", "actual"], {"connection": "FinMax", "label": "A"}), "scenario or rule"),
])
def test_parse_rejects_broken_catalogues(data, msg):
    with pytest.raises(catalog.CatalogError, match=msg):
        catalog.parse(data)
```

- [ ] **Step 2: Run them and check they fail**

Run: `.venv\Scripts\python.exe -m pytest tests/test_catalog.py -q`
Expected: FAIL, `ImportError: cannot import name 'catalog'`

- [ ] **Step 3: Write `knowledge\catalog.toml`**

```toml
# Semantic catalogue: what finance words mean in the cubes (Stage 1 of docs/specs/2026-10-25-finance-expert-design.md).
# Hand-maintained. catalog.parse() checks the structure; catalog.validate() checks every member against the atlas
# (`smartass doctor`). Sign convention as in FinMax: revenue positive, costs negative - so for every P&L line a
# higher value is better (better = "none" where that does not hold).

[pnl]
# Report order. kind: line | sum | grand. num: number-format override (tall1 = one decimal).
FinMax = [
  { key = "rev",    account = "ac.C30",        kind = "line",  label = "Revenue" },
  { key = "cogs",   account = "ac.C40S",       kind = "line",  label = "COGS" },
  { key = "gm",     account = "ac.C48",        kind = "sum",   label = "Gross margin" },
  { key = "pers",   account = "ac.C50S",       kind = "line",  label = "Personnel cost" },
  { key = "oth",    account = "ac.C60",        kind = "line",  label = "Other opex" },
  { key = "opex",   account = "ac.C70",        kind = "sum",   label = "Opex" },
  { key = "boi",    account = "ac.EBITDA_BOI", kind = "sum",   label = "EBITDA before other items" },
  { key = "c73",    account = "ac.C73",        kind = "line",  label = "Other items" },
  { key = "ebitda", account = "ac.EBITDA",     kind = "grand", label = "EBITDA" },
  { key = "da",     account = "ac.C75",        kind = "line",  label = "Depreciation and amortisation" },
  { key = "fte",    account = "ac.STA20Y",     kind = "line",  label = "FTE", better = "none", num = "tall1" },
]
# FinPlan stops at EBITDA: Target carries no depreciation/write-downs/other items, so EBIT variances would mislead.
FinPlan = [
  { key = "rev",    account = "Revenue",                   kind = "line",  label = "Revenue" },
  { key = "cogs",   account = "COGS",                      kind = "line",  label = "COGS" },
  { key = "gp",     account = "Gross_Profit",              kind = "sum",   label = "Gross profit" },
  { key = "opex",   account = "OPEX",                      kind = "line",  label = "Opex" },
  { key = "boi",    account = "EBITDA before Other Items", kind = "sum",   label = "EBITDA before other items" },
  { key = "oth",    account = "Total Other Items",         kind = "line",  label = "Other items" },
  { key = "ebitda", account = "EBITDA",                    kind = "grand", label = "EBITDA" },
]

[checks]
# Accounting identities, as expressions over line keys that must come to 0.
FinMax = [
  { label = "C30 + C40S = C48",          expr = "rev+cogs-gm" },
  { label = "C50S + C60 = C70",          expr = "pers+oth-opex" },
  { label = "C48 + C70 = EBITDA_BOI",    expr = "gm+opex-boi" },
  { label = "EBITDA_BOI + C73 = EBITDA", expr = "boi+c73-ebitda" },
]
FinPlan = [
  { label = "Revenue + COGS = Gross profit",                    expr = "rev+cogs-gp" },
  { label = "Gross profit + OPEX = EBITDA before other items",  expr = "gp+opex-boi" },
  { label = "EBITDA before other items + Other items = EBITDA", expr = "boi+oth-ebitda" },
]

[scenarios]
actual         = { connection = "FinMax",  scenario = "Actual", label = "Actual (ledger)" }
actual_finplan = { connection = "FinPlan", scenario = "Actual", version = "Working", label = "Actual (FinPlan)" }
budget         = { connection = "FinPlan", scenario = "Target", version = "Final", label = "Target (budget)" }
forecast       = { connection = "FinPlan", rule = "latest_forecast", version = "Final", label = "Latest forecast" }

[materiality]
pct = 0.10                            # a variance is material from 10 % of the comparison value ...
abs_nok = 5_000_000                   # ... and from NOK 5 million (both must hold)
always = ["rev", "pers", "ebitda"]    # always commented, material or not

# Business areas. Aliases are matched as whole words in a question; aliases shorter than 3 characters only when
# they are the whole question. abolished = the date the area was emptied (history is restated to today's org).
[[entity]]
key = "company"
label = "Telenor Norge (total)"
finmax = "un.SY"
finplan = "BANORWAY"
aliases = ["telenor norge", "total", "company", "selskapet", "tn"]

[[entity]]
key = "wholesale"
label = "Wholesale"
finmax = "un.SY.2WHS"
finplan = "2WHS Wholesale"
aliases = ["wholesale", "grossist", "whs", "2whs"]

[[entity]]
key = "business"
label = "Business"
finmax = "un.SY.2BUS"
finplan = "2BUS Business"
aliases = ["business", "bedrift", "bedriftsmarkedet", "b2b", "2bus"]

[[entity]]
key = "ssm"
label = "Security Solutions & Mobile"
finmax = "un.SY.2SSM"
finplan = "2SSM Security Solutions & Mobile"
aliases = ["security solutions & mobile", "security solutions and mobile", "ssm", "2ssm", "mobil"]

[[entity]]
key = "ei"
label = "Entertainment & Internet"
finmax = "un.SY.2EI"
finplan = "2EI Entertainment & Internet"
aliases = ["entertainment & internet", "entertainment and internet", "tv og internett", "e&i", "2ei", "ei"]

[[entity]]
key = "technology"
label = "Technology"
finmax = "un.SY.2TEC"
finplan = "2TEC Technology"
aliases = ["technology", "teknologi", "tech", "2tec"]

[[entity]]
key = "it"
label = "IT"
finmax = "un.SY.2IT"
finplan = "2IT IT"
aliases = ["it division", "it-divisjonen", "2it", "it"]

[[entity]]
key = "adm"
label = "Norway Adm"
finmax = "un.SY.2ADM"
finplan = "2ADM Norway Adm"
aliases = ["norway adm", "administration", "administrasjon", "staff units", "stab", "2adm", "adm"]

[[entity]]
key = "nordics"
label = "Nordics"
finmax = "un.SY.2NRD_N"
finplan = "2NRD_N Nordics"
aliases = ["nordics", "norden", "2nrd"]

[[entity]]
key = "ig"
label = "Infrastructure Growth"
finmax = "un.SY.2IG"
finplan = "2IG Infrastructure Growth"
aliases = ["infrastructure growth", "2ig", "ig"]
abolished = "2026-06-01"
note = "Abolished in the 1 June 2026 reorg; history is restated to today's organisation."

[[entity]]
key = "consumer"
label = "Consumer Norway"
finmax = "un.SY.2CON"
aliases = ["consumer", "privat", "privatmarkedet", "2con"]
abolished = "2026-06-01"
note = "Emptied into 2SSM and 2EI on 1 June 2026; history is restated, so 2CON shows no data."

# KPIs: ratios of two P&L line keys of one cube (computed in the workbook, never by the model).
[[kpi]]
key = "ebitda_margin"
label = "EBITDA margin"
connection = "FinMax"
numerator = "ebitda"
denominator = "rev"
unit = "pct"

[[kpi]]
key = "gross_margin_pct"
label = "Gross margin %"
connection = "FinMax"
numerator = "gm"
denominator = "rev"
unit = "pct"

[[kpi]]
key = "opex_per_fte"
label = "Opex per FTE"
connection = "FinMax"
numerator = "opex"
denominator = "fte"
unit = "nok"

[[kpi]]
key = "personnel_per_fte"
label = "Personnel cost per FTE"
connection = "FinMax"
numerator = "pers"
denominator = "fte"
unit = "nok"
```

- [ ] **Step 4: Write `catalog.py` (load and parse)**

```python
# engine/smartview_mcp/catalog.py
"""Semantic catalogue (knowledge/catalog.toml): P&L lines and identity checks per cube, business areas with aliases,
scenarios, KPIs and materiality - the one place that says what finance words mean in the cubes. templates.py and the
coming fin_* tools read it. parse() checks the structure; validate() checks every member against the atlas."""
from __future__ import annotations

import functools
import re
import tomllib
from dataclasses import dataclass
from pathlib import Path

from . import config

KINDS = {"line", "sum", "grand"}
RULES = {"latest_forecast"}


class CatalogError(ValueError):
    """knowledge/catalog.toml is broken (the message names the entry)."""


@dataclass(frozen=True)
class Line:
    key: str
    account: str
    kind: str
    label: str
    better: str = "higher"
    num: str | None = None


@dataclass(frozen=True)
class Check:
    label: str
    expr: str


@dataclass(frozen=True)
class Scenario:
    key: str
    connection: str
    label: str
    scenario: str | None = None
    version: str | None = None
    rule: str | None = None


@dataclass(frozen=True)
class Entity:
    key: str
    label: str
    aliases: tuple[str, ...]
    finmax: str | None = None
    finplan: str | None = None
    abolished: str | None = None
    note: str = ""


@dataclass(frozen=True)
class Kpi:
    key: str
    label: str
    connection: str
    numerator: str
    denominator: str
    unit: str


@dataclass(frozen=True)
class Materiality:
    pct: float
    abs_nok: float
    always: tuple[str, ...]


@dataclass(frozen=True)
class Catalog:
    pnl: dict[str, tuple[Line, ...]]
    checks: dict[str, tuple[Check, ...]]
    scenarios: dict[str, Scenario]
    materiality: Materiality
    entities: tuple[Entity, ...]
    kpis: tuple[Kpi, ...]


def _norm(s: str) -> str:
    return re.sub(r"[^\w&]+", " ", s.casefold()).strip()


def _unique(keys, what: str) -> None:
    seen = set()
    for k in keys:
        if k in seen:
            raise CatalogError(f"duplicate {what} '{k}'")
        seen.add(k)


def parse(data: dict) -> Catalog:
    try:
        pnl = {cube: tuple(Line(**x) for x in lines) for cube, lines in data["pnl"].items()}
        checks = {cube: tuple(Check(**x) for x in cs) for cube, cs in data.get("checks", {}).items()}
        scenarios = {k: Scenario(key=k, **v) for k, v in data["scenarios"].items()}
        materiality = Materiality(float(data["materiality"]["pct"]), float(data["materiality"]["abs_nok"]),
                                  tuple(data["materiality"]["always"]))
        entities = tuple(Entity(**{**e, "aliases": tuple(e.get("aliases", ()))}) for e in data["entity"])
        kpis = tuple(Kpi(**k) for k in data.get("kpi", ()))
    except (KeyError, TypeError) as e:
        raise CatalogError(f"missing or unknown field: {e}") from None
    keys = {cube: {x.key for x in lines} for cube, lines in pnl.items()}
    for cube, lines in pnl.items():
        _unique([x.key for x in lines], f"{cube} line key")
        for x in lines:
            if x.kind not in KINDS:
                raise CatalogError(f"{cube} line '{x.key}': kind '{x.kind}' is not one of {sorted(KINDS)}")
    for cube, cs in checks.items():
        for c in cs:
            for tok in re.findall(r"[a-z_][a-z0-9_]*", c.expr):
                if tok not in keys.get(cube, set()):
                    raise CatalogError(f"{cube} check '{c.label}': '{tok}' is not a {cube} line key")
    for s in scenarios.values():
        if not s.scenario and not s.rule:
            raise CatalogError(f"scenario '{s.key}' needs a scenario or rule")
        if s.rule and s.rule not in RULES:
            raise CatalogError(f"scenario '{s.key}': rule '{s.rule}' is not one of {sorted(RULES)}")
    for k in kpis:
        for part in (k.numerator, k.denominator):
            if part not in keys.get(k.connection, set()):
                raise CatalogError(f"kpi '{k.key}': '{part}' is not a {k.connection} line key")
    every = set().union(*keys.values())
    for a in materiality.always:
        if a not in every:
            raise CatalogError(f"materiality.always: '{a}' is not a line key")
    _unique([e.key for e in entities], "entity key")
    _unique([_norm(a) for e in entities for a in e.aliases], "entity alias")
    return Catalog(pnl, checks, scenarios, materiality, entities, kpis)


@functools.lru_cache(maxsize=4)
def _load(path: str) -> Catalog:
    try:
        data = tomllib.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, tomllib.TOMLDecodeError) as e:
        raise CatalogError(f"cannot read {path}: {e}") from None
    return parse(data)


def load(path: Path | None = None) -> Catalog:
    return _load(str(path or config.knowledge_file("catalog.toml")))
```

- [ ] **Step 5: Run the tests again**

Run: `.venv\Scripts\python.exe -m pytest tests/test_catalog.py -q`. Expected: all passed.

- [ ] **Step 6: Commit**

```bash
git add knowledge/catalog.toml engine/smartview_mcp/catalog.py engine/tests/test_catalog.py
git commit -m "catalog: semantic catalogue (P&L lines, checks, areas, scenarios, KPIs, materiality)"
```

### Task 9: `catalog.resolve_entity`

**Files:**
- Modify: `engine\smartview_mcp\catalog.py` (append)
- Test: `engine\tests\test_catalog.py` (append)

- [ ] **Step 1: Write the failing test**

```python
@pytest.mark.parametrize("text, key", [
    ("P&L for Wholesale Sep 26", "wholesale"),
    ("hvordan går det med bedriftsmarkedet?", "business"),
    ("EBITDA i TV og internett hittil i år", "ei"),
    ("un.SY.2TEC", "technology"),
    ("2SSM Security Solutions & Mobile", "ssm"),
    ("it", "it"),
    ("Telenor Norge total", "company"),
])
def test_resolve_entity(text, key):
    assert catalog.resolve_entity(catalog.load(), text).key == key


@pytest.mark.parametrize("text", ["is it good?", "what about ei?", "something else entirely"])
def test_resolve_entity_needs_a_real_match(text):
    assert catalog.resolve_entity(catalog.load(), text) is None
```

- [ ] **Step 2: Run it and check it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_catalog.py -q -k resolve`
Expected: FAIL, `AttributeError: ... has no attribute 'resolve_entity'`

- [ ] **Step 3: Implement** (append to `catalog.py`)

```python
def resolve_entity(cat: Catalog, text: str) -> Entity | None:
    """The business area a question or member name refers to: a member name (finmax/finplan) or the whole text as
    an alias first, else the longest alias found as whole words (aliases under 3 characters only as the whole text)."""
    n = _norm(text)
    for e in cat.entities:
        names = [x for x in (e.finmax, e.finplan) if x]
        if n in {_norm(x) for x in names} or n in {_norm(a) for a in e.aliases} or n == _norm(e.label):
            return e
    best: tuple[int, Entity] | None = None
    for e in cat.entities:
        for a in (*e.aliases, e.label):
            a = _norm(a)
            if len(a) < 3:
                continue
            if re.search(rf"(?<![\w&]){re.escape(a)}(?![\w&])", n) and (best is None or len(a) > best[0]):
                best = (len(a), e)
    return best[1] if best else None
```

- [ ] **Step 4: Run the tests again**

Run: `.venv\Scripts\python.exe -m pytest tests/test_catalog.py -q`. Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/catalog.py engine/tests/test_catalog.py
git commit -m "catalog: resolve a business area from a question or member name"
```

### Task 10: `catalog.validate` against the atlas + doctor check

**Files:**
- Modify: `engine\smartview_mcp\catalog.py` (append), `engine\smartview_mcp\doctor.py` (new `_catalog`, `CHECKS`),
  `engine\tests\live\scenarios.py:737`
- Test: `engine\tests\test_catalog.py` (append)

- [ ] **Step 1: Write the failing tests**

```python
def _atlas_with(tmp_path, members):
    from smartview_mcp import atlas
    db = atlas.connect(tmp_path / "a.db")
    for (conn, dim), names in members.items():
        atlas.store(db, conn, dim, [{"name": n, "label": n, "parent": None, "depth": 1, "order": i}
                                    for i, n in enumerate(names)], 0.1)
    return db


def test_validate_names_missing_members_and_unchecked_dimensions(tmp_path):
    cat = catalog.parse(BASE)
    db = _atlas_with(tmp_path, {("FinMax", "All Accounts"): ["ac.C30"],
                                ("FinMax", "All Entities"): ["un.SY.2WHS"]})
    problems, unchecked = catalog.validate(cat, db)
    assert problems == ["FinMax line 'ebitda': 'ac.EBITDA' is not in FinMax / All Accounts"]
    assert unchecked == ["FinMax / All Scenarios"]


def test_validate_passes_a_complete_atlas(tmp_path):
    cat = catalog.parse(BASE)
    db = _atlas_with(tmp_path, {("FinMax", "All Accounts"): ["ac.C30", "ac.EBITDA"],
                                ("FinMax", "All Entities"): ["un.SY.2WHS"],
                                ("FinMax", "All Scenarios"): ["Actual"]})
    assert catalog.validate(cat, db) == ([], [])


def test_doctor_reports_catalogue_problems(monkeypatch, tmp_path):
    from smartview_mcp import doctor
    monkeypatch.setattr(doctor.config, "ATLAS_DB", tmp_path / "missing.db")
    c = doctor._catalog()
    assert c.name == "catalog" and c.status == "WARN" and "no atlas" in c.detail
```

- [ ] **Step 2: Run them and check they fail**

Run: `.venv\Scripts\python.exe -m pytest tests/test_catalog.py -q -k "validate or doctor"`
Expected: FAIL (`no attribute 'validate'`, `no attribute '_catalog'`)

- [ ] **Step 3: Implement**

Append to `catalog.py`:

```python
ACCOUNT_DIM = {"FinMax": "All Accounts", "FinLite": "All Accounts", "FinPlan": "Account"}
ENTITY_DIM = {"FinMax": "All Entities", "FinPlan": "Entity"}
SCENARIO_DIM = {"FinMax": "All Scenarios", "FinPlan": "Scenario"}
VERSION_DIM = {"FinPlan": "Version"}


def validate(cat: Catalog, db) -> tuple[list[str], list[str]]:
    """(problems, unchecked): catalogue members the atlas does not have, and dimensions not checked because the
    atlas has not built them (or they are stale)."""
    from . import atlas
    problems: list[str] = []
    unchecked: list[str] = []

    def check(conn: str, dim: str | None, name: str, what: str) -> None:
        if dim is None:
            return
        if not atlas.usable(db, conn, dim):
            if f"{conn} / {dim}" not in unchecked:
                unchecked.append(f"{conn} / {dim}")
            return
        if not atlas.has_member(db, conn, dim, name):
            problems.append(f"{what}: '{name}' is not in {conn} / {dim}")

    for cube, lines in cat.pnl.items():
        for x in lines:
            check(cube, ACCOUNT_DIM.get(cube), x.account, f"{cube} line '{x.key}'")
    for e in cat.entities:
        if e.finmax:
            check("FinMax", ENTITY_DIM["FinMax"], e.finmax, f"entity '{e.key}'")
        if e.finplan:
            check("FinPlan", ENTITY_DIM["FinPlan"], e.finplan, f"entity '{e.key}'")
    for s in cat.scenarios.values():
        if s.scenario:
            check(s.connection, SCENARIO_DIM.get(s.connection), s.scenario, f"scenario '{s.key}'")
        if s.version:
            check(s.connection, VERSION_DIM.get(s.connection), s.version, f"scenario '{s.key}' version")
    return problems, unchecked
```

Add to `doctor.py` after `_atlas`:

```python
def _catalog() -> Check:
    from . import atlas, catalog
    try:
        cat = catalog.load()
    except catalog.CatalogError as e:
        return Check("catalog", "FAIL", str(e), "Fix knowledge\\catalog.toml (see the error), or update smartass.")
    if not config.ATLAS_DB.exists():
        return Check("catalog", "WARN", "loaded; members not checked (no atlas)", "Build the atlas with sv_atlas_build.")
    problems, unchecked = catalog.validate(cat, atlas.connect(config.ATLAS_DB))
    if problems:
        return Check("catalog", "WARN", f"{len(problems)} member(s) not in the cubes: " + "; ".join(problems[:5]),
                     "Update knowledge\\catalog.toml to the current members (a reorg or rename?).")
    detail = (f"{sum(len(v) for v in cat.pnl.values())} P&L lines and {len(cat.entities)} business areas match the "
              "atlas" + (f"; not checked: {', '.join(unchecked)}" if unchecked else ""))
    return Check("catalog", "OK", detail)
```

Change `CHECKS` to:

```python
CHECKS = (_config, _excel, _smartview, _epm, _drill_browser, _atlas, _catalog, _output, _http, _update)
```

In `engine\tests\live\scenarios.py:737`, add `"catalog"` to `DOCTOR_CHECKS`.

- [ ] **Step 4: Run the tests and the real doctor**

Run: `.venv\Scripts\python.exe -m pytest -q`. Expected: all passed.
Run: `.venv\Scripts\python.exe -m smartview_mcp.cli doctor`. Expected: `catalog OK ...` against the real atlas.
- **If it reports missing members:** correct `knowledge\catalog.toml` to the member the atlas has (look it up with
  `atlas.search`) and rerun until OK.
- **Never change `pnl` entries:** that would change the templates. If a template line is missing in the atlas, report
  it to Tom instead.

- [ ] **Step 5: Commit**

```bash
git add engine/smartview_mcp/catalog.py engine/smartview_mcp/doctor.py engine/tests/test_catalog.py engine/tests/live/scenarios.py knowledge/catalog.toml
git commit -m "catalog: validate against the atlas; doctor check"
```

### Task 11: templates take their lists from the catalogue

**Files:**
- Modify: `engine\smartview_mcp\templates.py:195-202` (PNL, PNL_CHECKS) and `:405-409` (FP_PNL, FP_CHECKS)
- Test: `engine\tests\test_catalog.py` (append)

- [ ] **Step 1: Write the failing test**

```python
def test_templates_read_the_catalogue():
    import inspect

    from smartview_mcp import templates
    assert templates.PNL == OLD_PNL and templates.PNL_CHECKS == OLD_PNL_CHECKS
    assert templates.FP_PNL == OLD_FP_PNL and templates.FP_CHECKS == OLD_FP_CHECKS
    assert '("rev", "ac.C30"' not in inspect.getsource(templates)   # no second copy of the list
```

- [ ] **Step 2: Run it and check it fails**

Run: `.venv\Scripts\python.exe -m pytest tests/test_catalog.py -q -k templates`
Expected: FAIL on the `inspect.getsource` assertion.

- [ ] **Step 3: Implement**

In `templates.py`, add `catalog` to the `from . import ...` line. Replace lines 195-202 with:

```python
_CAT = catalog.load()
PNL = [(x.key, x.account, x.kind, x.num) for x in _CAT.pnl["FinMax"]]       # key, FinMax account, kind, num
PNL_CHECKS = [(c.label, c.expr) for c in _CAT.checks["FinMax"]]
```

Replace the `FP_PNL = [...]` and `FP_CHECKS = [...]` literals (lines 405-409) with the lines below. Keep the comment
above them, which explains why FinPlan stops at EBITDA.

```python
FP_PNL = [(x.key, x.account, x.kind) for x in _CAT.pnl["FinPlan"]]
FP_CHECKS = [(c.label, c.expr) for c in _CAT.checks["FinPlan"]]
```

- [ ] **Step 4: Run everything**

Run: `.venv\Scripts\python.exe -m pytest -q` and `.venv\Scripts\python.exe -m ruff check .`.
Expected: all passed; `All checks passed!`

- [ ] **Step 5: Commit and push**

```bash
git add engine/smartview_mcp/templates.py engine/tests/test_catalog.py
python scripts/check_staged.py
git commit -m "templates: P&L lines and checks come from the catalogue"
git push
```

---

## Done when

- `pytest -q` is green (295 existing + new), `ruff` is clean, and `check_staged.py` passes.
- `tests\live\svhttp_vs_com.py` prints five PASS lines (or FinPlan is explicitly routed to COM and documented).
- `smartass doctor` shows `catalog OK`.
- `sv_query` returns `"route": "http"` in a normal session. Restart the MCP server first (`/mcp` reconnect).
