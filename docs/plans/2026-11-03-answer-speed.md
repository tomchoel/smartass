# Answer Speed Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Every answer waits less: grids fetched in parallel (and while Excel opens), the pnl live read in parallel, no fixed 12 s login wait if the spike allows, and a COM-call counter that shows which Excel writes to trim next.

**Architecture:** `SvHttp.fetch_many` runs `fetch` on a thread pool (the session pool from step 1b keeps it safe). `source.Prefetch` starts it in the background before `open target`; `source.pull` consumes the result. `extract._fill` runs its `retrieve` calls on a pool and stores the results in grid order. A thread-local counter in `excel/jobs.com` is reported per stage.

**Tech Stack:** Python 3.12, concurrent.futures, pywin32 (COM), pytest. Run tests from `C:\dev\smartass\engine` with `.venv\Scripts\python.exe -m pytest -q tests` (Bash: `.venv/Scripts/python.exe`). Lint: `.venv/Scripts/python.exe -m ruff check smartview_mcp tests`.

Spec: `docs/specs/2026-11-03-answer-speed-design.md`. Commits end with:
```
Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01MW9ANsye1x7As56S36PZbu
```

**Excel-write batching (spec section 3, second half) gets its own plan after Task 6 has measured where the calls are**: the writers are already block writes inside `jobs.fast_excel`, so which calls to cut is a measurement result, not a guess.

## File map
- `smartview_mcp/svhttp.py`: + `SvHttp.fetch_many`.
- `smartview_mcp/config.py`: + `ANSWER_PREFETCH`.
- `smartview_mcp/core/source.py`: + `too_big(pin)`, `Prefetch`; `pull(..., pf=None)`.
- `smartview_mcp/core/pipeline.py`: `run` starts a `Prefetch` before `open target`; `_new`/`_follow` pass it to `pull`.
- `smartview_mcp/hostproc.py`: + `on_host_thread()`.
- `smartview_mcp/extract.py`: `_fill` retrieves in parallel (not on the host thread), stores in order.
- `smartview_mcp/excel/target.py`: `appear_grace` (only if the Task 5 spike says so).
- `smartview_mcp/excel/jobs.py`, `smartview_mcp/timing.py`: COM call counter per stage.
- Tests: `tests/test_svhttp.py`, `tests/test_core_skeleton.py`, `tests/test_extract.py`, `tests/test_timing.py` (new if absent), `tests/conftest.py`.

---

### Task 1: `SvHttp.fetch_many`

**Files:** Modify `smartview_mcp/svhttp.py` (in class `SvHttp`, after `fetch`). Test `tests/test_svhttp.py` (append).

- [ ] **Step 1: Write the failing tests** (append to `tests/test_svhttp.py`)

```python
def test_fetch_many_keeps_order_returns_errors_in_place_and_caps_concurrency(db_path, monkeypatch):
    import random
    import threading
    import time as _t
    monkeypatch.setattr(config, "SVHTTP_SESSIONS", 3)
    sv = _sv(FakeTransport(), db_path)
    live, peak, lock = [0], [0], threading.Lock()

    def fake_fetch(conn, pov, rows, cols):
        with lock:
            live[0] += 1
            peak[0] = max(peak[0], live[0])
        _t.sleep(random.uniform(0.0, 0.03))
        with lock:
            live[0] -= 1
        if pov == "bad":
            raise svhttp.SvHttpError("no")
        return f"grid {pov}"
    monkeypatch.setattr(sv, "fetch", fake_fetch)
    items = [("FinMax", str(i) if i != 4 else "bad", [], []) for i in range(10)]
    out = sv.fetch_many(items)
    assert [x if isinstance(x, str) else "ERR" for x in out] == [f"grid {i}" if i != 4 else "ERR" for i in range(10)]
    assert isinstance(out[4], svhttp.SvHttpError)
    assert 1 < peak[0] <= 3                                 # parallel, never above the per-cube session cap
    assert sv.fetch_many([]) == []
```

- [ ] **Step 2: Run, expect FAIL** — `.venv/Scripts/python.exe -m pytest -q tests/test_svhttp.py -k fetch_many` → `AttributeError: 'SvHttp' object has no attribute 'fetch_many'`.

- [ ] **Step 3: Implement** (in `SvHttp`, right after `fetch`)

```python
    def fetch_many(self, items: list[tuple]) -> list:
        """fetch(conn, pov, rows, cols) for every item at once (at most SVHTTP_SESSIONS threads; each request has a
        Smart View session of its own). Results in input order; an item that failed holds its exception."""
        if not items:
            return []
        with cf.ThreadPoolExecutor(max_workers=min(len(items), config.SVHTTP_SESSIONS),
                                   thread_name_prefix="svhttp-many") as ex:
            futs = [ex.submit(self.fetch, *it) for it in items]
        return [f.exception() if f.exception() is not None else f.result() for f in futs]
```

- [ ] **Step 4: Run, expect PASS** — same command.
- [ ] **Step 5: Commit** — `git add smartview_mcp/svhttp.py tests/test_svhttp.py` and commit `feat(svhttp): fetch_many - grids at once, results in order, errors in place`.

---

### Task 2: `source.Prefetch` and `pull(..., pf=...)`

**Files:** Modify `smartview_mcp/config.py`, `smartview_mcp/core/source.py`, `tests/conftest.py`. Test `tests/test_core_skeleton.py` (append).

- [ ] **Step 1: Config + test guard.** In `config.py` next to `DRILL_DERIVE` add:

```python
ANSWER_PREFETCH = os.environ.get("SV_ANSWER_PREFETCH", "1") != "0"   # fetch an answer's pulls at once, while Excel opens
```

In `tests/conftest.py`, inside the autouse fixture that sets `DRILL_DIRECT` False, add:

```python
    monkeypatch.setattr(config, "ANSWER_PREFETCH", False)        # tests never reach the live HTTP route
```

- [ ] **Step 2: Write the failing tests** (append to `tests/test_core_skeleton.py`; `PIN`, `source`, `adhoc`, `config`, `pytest` are already imported there — add any missing import at the top)

```python
def _pins(*specs):
    return [PIN.model_copy(update={"id": i, "rows": PIN.rows, "columns": PIN.columns}) for i in specs]


def test_prefetch_fetches_every_http_pull_at_once_and_pull_uses_it(monkeypatch):
    from smartview_mcp import svhttp
    monkeypatch.setattr(config, "ANSWER_PREFETCH", True)
    seen = []

    class Http:
        def fetch_many(self, items):
            seen.append(len(items))
            return ["F0", svhttp.SvHttpError("down"), "F2"]
    monkeypatch.setattr(svhttp, "shared", lambda: Http())
    written = []
    monkeypatch.setattr(source, "write", lambda t, q, title, pin, f: written.append(f) or ("W", f))
    monkeypatch.setattr(adhoc, "pull", lambda h, t, q, conn, pov, rows, cols, title: "ADHOC")
    pins = _pins("a", "b", "c")
    pf = source.Prefetch(pins)
    assert source.pull(None, None, 1, 0, pins[0], pf=pf) == ("W", "F0")
    s = source.pull(None, None, 1, 1, pins[1], pf=pf)                   # its HTTP fetch failed: Excel, with the note
    assert s.route == "adhoc" and "HTTP route failed (down)" in s.note
    assert source.pull(None, None, 1, 2, pins[2], pf=pf) == ("W", "F2")
    assert seen == [3]                                                  # one batch for all three


def test_prefetch_off_or_too_big_pulls_fetch_as_before(monkeypatch):
    from smartview_mcp import svhttp
    calls = []

    class Http:
        def fetch_many(self, items):
            calls.append("many")
            return []

        def fetch(self, *a):
            calls.append("one")
            return "F"
    monkeypatch.setattr(svhttp, "shared", lambda: Http())
    monkeypatch.setattr(source, "write", lambda t, q, title, pin, f: ("W", f))
    pf = source.Prefetch(_pins("a"))                                    # ANSWER_PREFETCH is False in tests
    assert source.pull(None, None, 1, 0, _pins("a")[0], pf=pf) == ("W", "F") and calls == ["one"]
    monkeypatch.setattr(config, "ANSWER_PREFETCH", True)
    monkeypatch.setattr(config, "SVHTTP_TIMEOUT_S", 0.0)               # every pull is 'too big' for HTTP
    assert source.Prefetch(_pins("a")).has(0) is False


def test_a_prefetched_input_error_is_raised_not_sent_to_excel(monkeypatch):
    from smartview_mcp import svhttp
    from smartview_mcp.host import InputError
    monkeypatch.setattr(config, "ANSWER_PREFETCH", True)

    class Http:
        def fetch_many(self, items):
            return [InputError("Invalid query: x")]
    monkeypatch.setattr(svhttp, "shared", lambda: Http())
    with pytest.raises(InputError, match="Invalid query"):
        source.pull(None, None, 1, 0, _pins("a")[0], pf=source.Prefetch(_pins("a")))
```

- [ ] **Step 3: Run, expect FAIL** — `.venv/Scripts/python.exe -m pytest -q tests/test_core_skeleton.py -k prefetch` → `AttributeError: module ... has no attribute 'Prefetch'`.

- [ ] **Step 4: Implement** in `smartview_mcp/core/source.py`. Add `import concurrent.futures as cf` and `import math` (if absent) at the top, and from `..host import InputError` if absent. Replace `pull` and add the helpers:

```python
_BG = cf.ThreadPoolExecutor(max_workers=2, thread_name_prefix="prefetch")


def too_big(pin: PullSpec) -> bool:
    """More cells than one HTTP request should carry: fetched through Smart View in Excel."""
    n = math.prod(len(a.members) for a in [*pin.rows, *pin.columns])
    return n / 200 > config.SVHTTP_TIMEOUT_S


class Prefetch:
    """An answer's HTTP pulls fetched at once in the background (svhttp.fetch_many), started before the workbook is
    opened so the grids arrive while Excel starts. pull() takes pull i's result from here (has(i)), else fetches it
    itself as before. Off when config.ANSWER_PREFETCH is False."""

    def __init__(self, pins: list[PullSpec], only: set[int] | None = None):
        self.index = [i for i, p in enumerate(pins)
                      if config.ANSWER_PREFETCH and (only is None or i in only) and not too_big(p)]
        items = [(pins[i].connection, dict(pins[i].pov), [Axis(a.dimension, a.members) for a in pins[i].rows],
                  [Axis(a.dimension, a.members) for a in pins[i].columns]) for i in self.index]
        self._fut = _BG.submit(svhttp.shared().fetch_many, items) if items else None

    def has(self, i: int) -> bool:
        return i in self.index

    def result(self, i: int):
        """pull i's Fetched, or the exception its fetch raised."""
        try:
            got = self._fut.result(timeout=config.SVHTTP_TIMEOUT_S * 2)
        except Exception as e:   # noqa: BLE001 - the whole batch failed: every pull falls back like a failed fetch
            return svhttp.SvHttpError(f"prefetch failed: {e}")
        return got[self.index.index(i)]


def pull(h, t, q: int, i: int, pin: PullSpec, pf: Prefetch | None = None) -> Source:
    """Pull i of question q: HTTP first (prefetched in `pf` when given), the ad hoc grid if the HTTP route fails or
    the grid is too big for one request. A wrong query (InputError) is never retried through Excel."""
    title = sheet_title(q, i, pin)
    rows = [Axis(a.dimension, a.members) for a in pin.rows]
    cols = [Axis(a.dimension, a.members) for a in pin.columns]
    if too_big(pin):
        n = math.prod(len(a.members) for a in rows + cols)
        note = f"{n} cells is too big for one HTTP request; fetched through Smart View in Excel."
    else:
        try:
            with stage("fetch (HTTP)"):
                if pf is not None and pf.has(i):
                    f = pf.result(i)
                    if isinstance(f, BaseException):
                        raise f
                else:
                    f = svhttp.shared().fetch(pin.connection, dict(pin.pov), rows, cols)
        except svhttp.SvHttpError as e:
            note = f"HTTP route failed ({e}); fetched through Smart View in Excel instead."
        else:
            with stage("write source sheet"):
                return write(t, q, title, pin, f)
    with stage("pull (Smart View ad hoc)"):
        p = adhoc.pull(h, t, q, pin.connection, dict(pin.pov), rows, cols, title=title)
    return Source(p, "adhoc", [], note)
```

(An `InputError` or any other exception from `pf.result` is re-raised by `raise f`, exactly as `fetch` raising it did before.)

- [ ] **Step 5: Run, expect PASS** — `.venv/Scripts/python.exe -m pytest -q tests/test_core_skeleton.py` (all, including the existing `source.pull` tests).
- [ ] **Step 6: Commit** — `feat(answer): Prefetch - an answer's HTTP pulls fetched at once; pull takes them (errors fall back as before)`.

---

### Task 3: pipeline starts the prefetch before `open target`

**Files:** Modify `smartview_mcp/core/pipeline.py` (`run`, `_new`, `_follow`). Test `tests/test_core_skeleton.py`.

- [ ] **Step 1: Write the failing test** (append; reuses `_fake_run` and `SPEC` from the same file). First change the fake in `_fake_run` to accept the new keyword: `def pull_(h, t, q, i, pin, pf=None):` and record it: `calls.append(("pull", pin.id))` stays; add `seen_pf.append(pf)` via a list stored on the function — simplest: replace the body's first line with `calls.append(("pull", pin.id, pf is not None))`? No - keep the existing assertion on `[c[0] for c in calls]` intact; so add a separate test:

```python
def test_run_starts_the_prefetch_before_opening_the_workbook(monkeypatch):
    calls = []
    _fake_run(monkeypatch, calls)
    order = []
    real_resolve = pipeline.xltarget.resolve_target
    monkeypatch.setattr(pipeline.xltarget, "resolve_target",
                        lambda *a, **k: order.append("open") or real_resolve(*a, **k))
    monkeypatch.setattr(pipeline.source, "Prefetch", lambda pins, only=None: order.append("prefetch") or "PF")
    got = []
    monkeypatch.setattr(pipeline.source, "pull", lambda h, t, q, i, pin, pf=None: got.append(pf) or
                        pipeline.source.Source(SimpleNamespace(sheet="Q3 FinMax"), "http", []))
    pipeline.run(None, SPEC)
    assert order[:2] == ["prefetch", "open"] and got == ["PF"]
```

Note for the implementer: check how `_fake_run` stubs `xltarget.resolve_target` (it must already, since `run` calls it); if it stubs it with `monkeypatch.setattr(pipeline.xltarget, "resolve_target", ...)`, the wrapper above composes with it. If `SimpleNamespace` is not imported in the file, add `from types import SimpleNamespace`.

- [ ] **Step 2: Run, expect FAIL** — `-k prefetch_before` → `order` lacks `"prefetch"`.

- [ ] **Step 3: Implement** in `core/pipeline.py`:

In `run`, after the ambiguous-gate check and before `with stage("open target"):`

```python
    pf = source.Prefetch(spec.pulls) if old is None else None   # grids arrive while Excel opens the workbook
```

Pass it on: `return _new(h, t, q, raw, spec, frame_done, drills, value_findings, pf=pf)`; add `pf: source.Prefetch | None = None` as the last parameter of `_new`, and in its loop call `sources.append(source.pull(h, t, q, i, pin, pf=pf))`.

In `_follow`, after `fresh = [p for p in spec.pulls if p.id not in plan.reuse]` add:

```python
        pf = source.Prefetch(spec.pulls, only={i for i, p in enumerate(spec.pulls) if p.id not in plan.reuse})
```

and change `src = source.pull(h, t, q, i, pin)` to `src = source.pull(h, t, q, i, pin, pf=pf)`.

- [ ] **Step 4: Run, expect PASS** — `.venv/Scripts/python.exe -m pytest -q tests/test_core_skeleton.py tests/` (full suite; the follow-up tests must keep passing).
- [ ] **Step 5: Commit** — `perf(answer): start fetching the pulls before the workbook opens`.

---

### Task 4: the pnl live read in parallel (`extract._fill`)

**Files:** Modify `smartview_mcp/hostproc.py`, `smartview_mcp/extract.py`. Test `tests/test_extract.py` (append).

- [ ] **Step 1: Write the failing tests** (append to `tests/test_extract.py`; `extract`, `FactStore`, `CAT` are imported there — check the file's imports and reuse its grid helpers; the test below builds its own grids with `extract.finmax_grids`, used elsewhere in the file)

```python
def test_fill_retrieves_at_once_and_stores_in_grid_order(monkeypatch):
    import random
    import threading
    import time as _t
    from smartview_mcp import config
    grids = extract.finmax_grids(CAT, ["Jan 26", "Feb 26", "Mar 26"], ["un.SY.2WHS", "un.SY.2BUS"])
    assert len(grids) >= 2
    plain = FactStore()

    def answer(conn, pov, rows, cols):                     # the http fake used by the other tests in this file
        return _fake_http_retrieve(conn, pov, rows, cols)
    seq = extract.fill(plain, answer, CAT, grids, "t", "Mar 26")
    live, peak, lock = [0], [0], threading.Lock()

    def slow(conn, pov, rows, cols):
        with lock:
            live[0] += 1
            peak[0] = max(peak[0], live[0])
        _t.sleep(random.uniform(0.0, 0.02))
        with lock:
            live[0] -= 1
        return answer(conn, pov, rows, cols)
    par = extract.fill(FactStore(), slow, CAT, grids, "t", "Mar 26")
    assert [r.key for r in par.rows] == [r.key for r in seq.rows]          # same rows, same order
    assert par.fill.result == seq.fill.result and peak[0] > 1


def test_fill_on_the_host_thread_stays_sequential(monkeypatch):
    from smartview_mcp import hostproc
    monkeypatch.setattr(hostproc, "on_host_thread", lambda: True)
    grids = extract.finmax_grids(CAT, ["Jan 26", "Feb 26"], ["un.SY.2WHS"])
    order = []

    def rec(conn, pov, rows, cols):
        import threading
        order.append(threading.current_thread().name)
        return _fake_http_retrieve(conn, pov, rows, cols)
    extract.fill(FactStore(), rec, CAT, grids, "t", "Feb 26")
    assert len(set(order)) == 1                                  # no pool: a COM fallback could not deadlock
```

`_fake_http_retrieve` does not exist yet: the existing tests build `extract.route(http=http)` with a fake `http`. Before writing the tests, read `tests/test_extract.py` lines 1-60, then define at the top of the new tests:

```python
def _fake_http_retrieve(conn, pov, rows, cols):
    return extract.route(http=_HTTP)(conn, pov, rows, cols)
```

where `_HTTP` is the module's existing fake http object (named in those first 60 lines; if it is built inside a test, lift it to module level unchanged).

- [ ] **Step 2: Run, expect FAIL** — `-k "at_once or host_thread"` → peak stays 1 / `on_host_thread` missing.

- [ ] **Step 3: Implement.** In `smartview_mcp/hostproc.py` (module level, after `class ComThread`):

```python
def on_host_thread() -> bool:
    """True on an ExcelHost's own COM thread (named 'excel-com'): work there must not wait on threads that call
    on_host, or it deadlocks."""
    t = threading.current_thread()
    return isinstance(t, ComThread) and t.name == "excel-com"
```

In `smartview_mcp/extract.py` `_fill`, replace the start of the loop:

```python
def _fill(store, retrieve, cat, grids, job, closed, closed_label, now, atlas_db, t0, started):
    rows, errors, warnings, routes = [], [], [], set()
    for g, got_or_err in zip(grids, _retrieve_all(retrieve, grids)):
        if isinstance(got_or_err, Exception):
            e = got_or_err
            errors.append(f"{g.what}: {type(e).__name__}: {str(e).splitlines()[0][:200] if str(e) else ''}")
            continue
        lay, parsed, _secs, how = got_or_err
        routes.add(how)
        # ... the rest of the loop body is unchanged ...
```

and add above `_fill`:

```python
def _retrieve_all(retrieve, grids) -> list:
    """retrieve() for every grid at once (HTTP route: a Smart View session each, at most SVHTTP_SESSIONS); results
    in grid order, an exception in place of a failed grid. Sequential on the host COM thread (a COM fallback there
    would wait on itself)."""
    import concurrent.futures as cf

    from . import hostproc

    def one(g):
        try:
            return retrieve(g.conn, g.pov, g.rows, g.cols)
        except Exception as e:   # noqa: BLE001 - one failed grid must not lose the others
            return e
    if len(grids) < 2 or hostproc.on_host_thread():
        return [one(g) for g in grids]
    with cf.ThreadPoolExecutor(max_workers=min(len(grids), config.SVHTTP_SESSIONS),
                               thread_name_prefix="fill") as ex:
        return list(ex.map(one, grids))
```

(`config` is already imported in `extract.py`; check, add `from . import config` if not.)

- [ ] **Step 4: Run, expect PASS** — `.venv/Scripts/python.exe -m pytest -q tests/test_extract.py tests/` (full suite: the fin tests use `extract.fill`).
- [ ] **Step 5: Commit** — `perf(pnl): the live read's grids at once, stored in grid order`.

---

### Task 5: login-wait spike, then `appear_grace`

**Files:** scratch script (session scratchpad), maybe `smartview_mcp/excel/target.py:236,316`.

- [ ] **Step 1: Spike (live, read-only, no Smart View call).** Write `<scratchpad>\spike_login.py`:

```python
"""Does a fresh Excel show the EPM login when no Smart View call is made? Watches 40 s; prints seconds only."""
import subprocess
import sys
import time

sys.path.insert(0, r"C:\dev\smartass\engine")
from smartview_mcp.excel import instances, login  # noqa: E402

before = {pid for pid, _, _ in instances.interactive_excels()}
p = subprocess.Popen([r"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE", "/x", "/e"])
t0, pid, seen = time.monotonic(), None, None
while time.monotonic() - t0 < 40:
    if pid is None:
        new = [x for x, _, _ in instances.interactive_excels() if x not in before]
        pid = new[0] if new else None
    elif seen is None and login.login_open(pid):
        seen = round(time.monotonic() - t0, 1)
    time.sleep(0.5)
print(f"excel pid {pid}; login window seen at {seen} s" if seen else f"excel pid {pid}; no login window in 40 s")
```

Run it three times with no other Excel open (ask the owner to close Excel first, or skip if Excel is open: the result is only valid for a fresh Excel). Close the spawned Excel after each run (`taskkill /PID <pid>` for that pid only).

- [ ] **Step 2: Record** the result under a "Spike result" heading in `docs/specs/2026-11-03-answer-speed-design.md` (section 2).

- [ ] **Step 3 (only if no login window appeared in all runs):** in `excel/target.py`, change both `login.wait_login_cleared(pid, appear_grace=12)` to `login.wait_login_cleared(pid, appear_grace=2)` and update the two comments to `# a login window still holds us for as long as it shows`. Run the full suite; commit `perf(excel): a fresh Excel waits 2 s for a login window, not 12 (spike: none appears without a Smart View call)`. If a login window appeared: commit only the spike note (`docs: login spike - a fresh Excel does show the EPM login; the 12 s grace stays`).

---

### Task 6: COM call counter per stage

**Files:** Modify `smartview_mcp/excel/jobs.py` (`com`), `smartview_mcp/timing.py` (`stage`, `Timer.summary`). Test `tests/test_timing.py` (create if absent; else append).

- [ ] **Step 1: Read** `smartview_mcp/timing.py` fully (Timer, stage, summary) before writing the test, to use its real API.

- [ ] **Step 2: Write the failing test**

```python
def test_com_calls_are_counted_per_stage(monkeypatch):
    from smartview_mcp import timing
    from smartview_mcp.excel import jobs
    monkeypatch.setattr(jobs, "_check_alive", lambda: None)
    tm = timing.Timer.start()
    with timing.stage("write"):
        for _ in range(3):
            jobs.com(lambda: 1)
        with timing.stage("inner"):
            jobs.com(lambda: 2)
    s = tm.summary()
    assert s["com calls.write"] == 4 and s["com calls.write.inner"] == 1
```

(If `Timer.start()`/`summary()` differ in the real API, adapt the calls, not the assertion's meaning: per stage, including nested stages, the number of `jobs.com` calls, reported in the summary under `com calls.<stage path>`.)

- [ ] **Step 3: Run, expect FAIL.**

- [ ] **Step 4: Implement.** In `timing.py`: `Timer` gets `self.com: dict[str, int] = {}`; add

```python
def count_com() -> None:
    """One COM call in the current stage (and its parents), for the stage summary."""
    t = Timer.current()
    if t is None or not t.stack:
        return
    for k in range(1, len(t.stack) + 1):
        key = ".".join(t.stack[:k])
        t.com[key] = t.com.get(key, 0) + 1
```

and in `summary()` add `**{f"com calls.{k}": v for k, v in self.com.items()}` to the returned dict. In `excel/jobs.py` `com()`, first line: `timing.count_com()` (import `from .. import timing` at the top if absent; check for an import cycle — if `timing` imports `jobs`, import inside the function).

- [ ] **Step 5: Run, expect PASS**, then the full suite.
- [ ] **Step 6: Commit** — `feat(timing): COM calls per stage in the run summary`.

---

### Task 7: live bench and the next plan

- [ ] **Step 1:** The server reloads by itself on the next tool call (run_server relay). Ask the owner, or run through smartass, three answers on a fresh question each: the consultant question (`lines_ranking`), a P&L (`pnl`, e.g. Wholesale YTD), a trend (`monthly_trend`). Read their `timings` from `%LOCALAPPDATA%\smartass\logs\runs.jsonl` (times and COM counts only).
- [ ] **Step 2:** Append a "Answer speed (2026-11-03)" section to `docs/audit/2026-11-02-speed-bench.md`: before (means from the spec table) vs after for `live read`, `fetch (HTTP)`, `open target`, and each write stage with its COM-call count.
- [ ] **Step 3:** From the counts, list the three write stages with the most COM calls; that list is the input for the follow-up plan (Excel-write batching). Commit `docs: answer speed bench`.
