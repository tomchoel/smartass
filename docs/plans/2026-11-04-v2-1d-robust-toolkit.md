# smartass v2 · Plan 1D: robustness and the Excel toolkit

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or
> superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make v2 safe and steady in daily use and give the agent v2's own Excel toolkit:
- AutoSave can never store a half-written answer.
- The session workbook survives a `/mcp` reconnect.
- A cold browser no longer fails the first request.
- `xl_snapshot` / `xl_revert`: a safety net before any Excel MCP edit.
- `xl_tidy`: TN style and the tn-formelstil check.

**Architecture:** All changes are inside `engine/smartass/`. v1 functions are reached through the legacy seam and
are never edited:
- `workbench/history`, `workbench/books`, `workbench/fmt`
- `excel/target._wait_ready`
- `session`

**Roadmap after this plan:**
- **1E:** `fin_drill` and the `lines` block, the ad hoc Excel fallback, and the test-set runner with golden numbers.
- **1F:** lift the borrowed v1 modules into v2 with the runtime fixes (atomic `pids.json`, reload on a version
  file only, DPAPI secret, Smart View options restored at startup), then delete v1.

**Conventions:** as Plan 1C.
- Branch `v2/1d` from `main`. Git Bash from `/c/dev/smartass/engine`.
- Tests go in `engine/tests/v2/test_v2_*.py`.
- Never run python through a stdin heredoc. Bare python only with `< /dev/null`.
- Stage files explicitly.
- No live Excel, EPM or browser. Task 8 waits for Tom.
- Never edit `engine/smartview_mcp` or `knowledge/`.
- Commit messages end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- Read the current v2 code before each task. Where a snippet differs from the code, follow the code and say so.

---

### Task 0: Branch

- [ ] `cd /c/dev/smartass && git switch main && git switch -c v2/1d`

---

### Task 1: AutoSave never stores a half-written answer

**Files:** `engine/smartass/pipeline.py`. Test: `engine/tests/v2/test_v2_autosave.py`.

**Rule:**
- **v2's own file:** AutoSave is off while v2 writes. v2 saves the file itself, then puts AutoSave back as it was.
- **The user's workbook** (`target='active'` or a path outside the analyser folder): AutoSave is turned off and
  left off, as in v1, because AutoSave would otherwise save our edits into the user's file. The reply then carries
  the note: "AutoSave was switched off in <name> so the answer is not saved for you; save yourself or switch it
  back on."

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_autosave.py
from smartass import pipeline


class WB:
    def __init__(self, on):
        self.AutoSaveOn = on
        self.Name = "x.xlsx"
        self.log = []

    def Save(self):
        self.log.append(("save", self.AutoSaveOn))


class T:
    def __init__(self, on, ours):
        self.wb = WB(on)
        self.created = ours
        self.path = "C:/analyser/v2/p/x.xlsx" if ours else "C:/Users/me/own.xlsx"


def test_our_file_autosave_off_while_writing_then_restored(monkeypatch):
    monkeypatch.setattr(pipeline, "_is_ours", lambda t: t.created)
    t = T(True, ours=True)
    with pipeline.autosave_paused(t) as notes:
        assert t.wb.AutoSaveOn is False
        pipeline._save(t)
    assert t.wb.log == [("save", False)]
    assert t.wb.AutoSaveOn is True and notes == []


def test_user_workbook_autosave_left_off_with_a_note(monkeypatch):
    monkeypatch.setattr(pipeline, "_is_ours", lambda t: t.created)
    t = T(True, ours=False)
    with pipeline.autosave_paused(t) as notes:
        assert t.wb.AutoSaveOn is False
    assert t.wb.AutoSaveOn is False
    assert notes and "AutoSave was switched off" in notes[0]


def test_autosave_already_off_or_unsupported_is_quiet(monkeypatch):
    monkeypatch.setattr(pipeline, "_is_ours", lambda t: t.created)
    t = T(False, ours=False)
    with pipeline.autosave_paused(t) as notes:
        pass
    assert notes == []

    class Bad(T):
        pass
    b = Bad(True, ours=True)
    type(b.wb).AutoSaveOn = property(lambda s: (_ for _ in ()).throw(RuntimeError("no autosave")))
    with pipeline.autosave_paused(b) as notes:
        pass
    assert notes == []
```

- [ ] **Step 2:** Run it and see it fail with `AttributeError: autosave_paused`.

- [ ] **Step 3: Implement** in `pipeline.py`:

```python
import contextlib


def _is_ours(t) -> bool:
    from smartview_mcp import server as v1
    return bool(getattr(t, "created", False)) or v1._is_our_analysis_file(str(t.path))


@contextlib.contextmanager
def autosave_paused(t):
    """AutoSave (OneDrive/SharePoint files) off while v2 writes, so a half-written answer is never stored.
    Our own analysis file: restored afterwards (we save it ourselves). The user's workbook: left off, with a note
    (AutoSave would otherwise save our edits into their file; v1 does the same)."""
    from smartview_mcp.excel.jobs import com
    notes: list[str] = []
    try:
        prev = bool(com(lambda: t.wb.AutoSaveOn))
    except Exception:   # noqa: BLE001 - a local file or an Excel without AutoSave
        prev = False
    if prev:
        com(lambda: setattr(t.wb, "AutoSaveOn", False))
    try:
        yield notes
    finally:
        if prev and _is_ours(t):
            with contextlib.suppress(Exception):
                com(lambda: setattr(t.wb, "AutoSaveOn", True))
        elif prev:
            notes.append(f"AutoSave was switched off in {t.wb.Name} so the answer is not saved for you; "
                         "save yourself or switch it back on.")
```

  Wrap the body of `excel_job` and `redraw_job`, from target resolution up to and including `_save`, in
  `with autosave_paused(t) as notes:`, and append `notes` to the reply's `warnings`. Reuse `_is_ours` wherever
  `_save` decides whether to save; replace any duplicate of that logic.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes, including the existing job and redraw order tests.

- [ ] **Step 5: Commit.** Message: `fix(v2): AutoSave paused while v2 writes; restored for our own files`.

---

### Task 2: The session workbook survives a reconnect

**Files:** `engine/smartass/legacy.py`. Test: `engine/tests/v2/test_v2_session_reconnect.py`.

**Rule:**
- `set_workbook` also records `_last:<cwd>` = `{workbook, ts}`. That entry is never pruned.
- `workbook()` first tries the current relay id. Otherwise it takes `_last:<cwd>` when the file still exists, and
  re-binds it under the new id.

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_session_reconnect.py
from smartview_mcp import session as v1session

from smartass import legacy


def configure(monkeypatch, tmp_path, rid, cwd):
    monkeypatch.setattr(legacy, "_configured", False)
    monkeypatch.setattr(v1session, "_STATE", tmp_path / "sessions.json")
    monkeypatch.setattr(legacy, "_session_info", lambda: {"id": rid, "cwd": cwd, "project": "p"})
    legacy.configure_v2()
    monkeypatch.setattr(v1session, "_STATE", tmp_path / "sessions_v2.json")


def test_reconnect_finds_the_last_workbook_of_the_same_folder(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"x")
    configure(monkeypatch, tmp_path, "v2-1-1", "C:/proj")
    v1session.set_workbook(str(book))
    configure(monkeypatch, tmp_path, "v2-2-2", "C:/proj")            # new relay after /mcp
    assert v1session.workbook() == str(book)


def test_other_folder_and_deleted_file_do_not_match(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"x")
    configure(monkeypatch, tmp_path, "v2-1-1", "C:/proj")
    v1session.set_workbook(str(book))
    configure(monkeypatch, tmp_path, "v2-3-3", "C:/other")
    assert v1session.workbook() is None
    configure(monkeypatch, tmp_path, "v2-4-4", "C:/proj")
    book.unlink()
    assert v1session.workbook() is None
```

  Read `legacy.py` first. The test has to drive whatever `configure_v2` patches (`_info`, `_STATE`,
  `set_workbook`, `_alive_relay`). Adapt the `configure` helper so each call simulates a new relay with the
  given id and cwd, and keep the assertions.

- [ ] **Step 2:** Run it and see it fail on the first assertion (`None`).

- [ ] **Step 3: Implement** in `legacy.py`:
  - In `_set_workbook_v2`, write the entry under the id as now, plus `data[f"_last:{cwd}"] = {"workbook": path,
    "ts": time.time()}`. Pruning must skip keys starting with `_last:`.
  - Add `_workbook_v2()`:
    1. The current id's workbook, if the file exists.
    2. Otherwise `_last:<info()["cwd"]>`, if the file exists. Re-bind it under the current id by calling
       `_set_workbook_v2` and return it.
    3. Otherwise `None`.
  - In `configure_v2`, set `v1session.workbook = _workbook_v2`.
  - Two sessions in the same folder now share the last workbook. Excel's per-workbook turn lock
    (`hold_turn`) keeps them from writing at the same time. Document this in the docstring.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.

- [ ] **Step 5: Commit.** Message: `fix(v2): the session workbook survives an /mcp reconnect (same folder)`.

---

### Task 3: A cold browser no longer fails the first request

**Files:** `engine/smartass/data/provider.py`. Test: `engine/tests/v2/test_v2_provider_retry.py`.

**Rule:** A grid that fails with a lost request (`SvHttpError` whose text contains "lost the request") is tried
once more after `RETRY_WAIT_S` (5 s; the sign-in has normally finished by then). A second failure is raised as
before. Any other `SvHttpError` is raised at once.

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_provider_retry.py
import pytest

from smartass.data import provider as prov
from smartass.data.provider import SvHttpProvider
from smartass.model.requests import GridRequest
from smartview_mcp.svhttp import SvHttpError

REQ = GridRequest("fm", "FinMax", ("ac.C30",), ("E1",), ("Sep 26",), "Actual", "", "Periodic")


class Parsed:
    columns = ["All Accounts", "All Entities", "Sep 26"]
    rows = [["ac.C30", "E1", 5.0]]
    unknown_members = []


class SV:
    def __init__(self, fails, text="lost the request (TimeoutError: signal timed out)"):
        self.fails, self.text, self.calls = fails, text, 0

    def retrieve(self, *a):
        self.calls += 1
        if self.calls <= self.fails:
            raise SvHttpError(f"Smart View provider at x: {self.text}")
        return None, Parsed(), 0.1


def test_one_lost_request_is_retried(monkeypatch):
    monkeypatch.setattr(prov, "RETRY_WAIT_S", 0)
    sv = SV(fails=1)
    got = SvHttpProvider(sv).fetch(REQ)
    assert sv.calls == 2 and [f.value for f in got.facts] == [5.0]
    assert any("retried" in w for w in got.warnings)


def test_two_lost_requests_fail_and_other_errors_fail_at_once(monkeypatch):
    monkeypatch.setattr(prov, "RETRY_WAIT_S", 0)
    with pytest.raises(SvHttpError):
        SvHttpProvider(SV(fails=2)).fetch(REQ)
    sv = SV(fails=1, text="rejected members: ac.X")
    with pytest.raises(SvHttpError):
        SvHttpProvider(sv).fetch(REQ)
    assert sv.calls == 1
```

- [ ] **Step 2:** Run it and see it fail: the first test raises.

- [ ] **Step 3: Implement** in `provider.py`:
  - Add `RETRY_WAIT_S = 5.0`.
  - Wrap the `self.sv.retrieve(...)` call per grid in `_retrieve(self, out, *args)`:
    1. Call `retrieve`.
    2. On a `SvHttpError` whose text contains "lost the request", `time.sleep(RETRY_WAIT_S)`, append the warning
       `"<cube>: the first request was lost (cold sign-in) and retried"`, and call once more.
    3. Re-raise anything else.
  - Import `SvHttpError` from `smartview_mcp.svhttp`.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.

- [ ] **Step 5: Commit.** Message: `fix(v2): a lost first request after a cold sign-in is retried once`.

---

### Task 4: `xl_snapshot` and `xl_revert`

**Files:** Create `engine/smartass/workbench.py`; modify `engine/smartass/server.py`. Test:
`engine/tests/v2/test_v2_workbench.py`.

**Rule:**
- `xl_snapshot(workbook, sheet, range)` saves the current state of a range (or the used range of the sheet when
  `range` is omitted) in v1's snapshot store, without changing anything. It returns `{change_id, workbook, sheet,
  address}`.
- `xl_revert(change_id, workbook, force=False)` is v1's `history.revert`.
- Together they are the safety net around any Excel MCP edit to an open workbook.

- [ ] **Step 1: Write the failing test**

```python
# engine/tests/v2/test_v2_workbench.py
from smartass import workbench


def test_snapshot_captures_without_mutating(monkeypatch):
    calls = []

    class B:
        name = "Book1.xlsx"
        app = object()

        def sheet(self, name):
            calls.append(("sheet", name))
            return "WS"

    monkeypatch.setattr(workbench, "_find_book", lambda wb: B())
    monkeypatch.setattr(workbench, "_wait_ready", lambda app: calls.append("ready"))
    monkeypatch.setattr(workbench, "_range", lambda ws, address: ("RNG", address or "A1:F30"))
    monkeypatch.setattr(workbench, "_capture", lambda ws, rng: {"snap": 1})
    monkeypatch.setattr(workbench, "_record", lambda b, ws, address, snap, label: {"id": "c1", "address": address})
    out = workbench.snapshot("Book1", "Sheet1", None, label="before Excel MCP edit")
    assert out == {"change_id": "c1", "workbook": "Book1.xlsx", "sheet": "Sheet1", "address": "A1:F30"}
    assert ("sheet", "Sheet1") in calls and "ready" in calls


def test_revert_delegates(monkeypatch):
    monkeypatch.setattr(workbench, "_history_revert", lambda cid, wb, force: {"reverted": cid, "force": force})
    assert workbench.revert("c1", "Book1", force=True) == {"reverted": "c1", "force": True}
```

- [ ] **Step 2:** Run it and see it fail with `ModuleNotFoundError`.

- [ ] **Step 3: Implement**

```python
# engine/smartass/workbench.py
"""v2's safety net for edits to an open workbook: xl_snapshot records a range's state in v1's snapshot store
(workbench/history, kept 7 days) without changing anything; xl_revert puts it back. Use before any Excel MCP edit.
Runs inside a user-Excel job (server wraps it in _call_user)."""
from __future__ import annotations


def _find_book(workbook):
    from smartview_mcp.workbench import books
    return books.find_book(workbook, write=True)            # also pauses AutoSave, as v1 does for edits


def _wait_ready(app):
    from smartview_mcp.excel.target import _wait_ready as wait
    wait(app)


def _range(ws, address):
    from smartview_mcp.excel.jobs import com
    from smartview_mcp.workbench import books
    rng = books._rng(ws, address) if address else com(lambda: ws.UsedRange)
    return rng, books._addr(rng)


def _capture(ws, rng):
    from smartview_mcp.workbench import history
    return history._capture(ws, rng)


def _record(b, ws, address, snap, label):
    from smartview_mcp.workbench import history
    ch = history._new_change(b, ws, "write", address, snap, label=label)
    history._save_change(ch)
    return ch


def _history_revert(change_id, workbook, force):
    from smartview_mcp.workbench import history
    return history.revert(change_id, workbook, force=force)


def snapshot(workbook: str | None, sheet: str | None, address: str | None, label: str = "xl_snapshot") -> dict:
    b = _find_book(workbook)
    ws = b.sheet(sheet)
    _wait_ready(b.app)
    rng, addr = _range(ws, address)
    ch = _record(b, ws, addr, _capture(ws, rng), label)
    return {"change_id": ch["id"], "workbook": b.name, "sheet": sheet, "address": addr}


def revert(change_id: str | None, workbook: str | None, force: bool = False) -> dict:
    return _history_revert(change_id, workbook, force)
```

  Check against v1:
  - `history._new_change`'s return key for the id (`"id"` or `"change_id"`).
  - `b.sheet(None)` returning the active sheet.
  - The sheet name to return when `sheet` is None: read it from `ws.Name`.
  - Fix the test only where v1's real shapes demand it, and say so.

  In `server.py` add:

```python
@_register(v1.WRITE)
async def xl_snapshot(workbook: str | None = None, sheet: str | None = None, range: str | None = None) -> dict:
    """Save the current state of a range (default: the sheet's used range) of an open workbook, changing nothing.
    Call before any Excel MCP edit to a workbook the user has open; xl_revert(change_id) undoes the edit."""
    from .workbench import snapshot
    return await _user(lambda h: snapshot(workbook, sheet, range), "xl_snapshot")


@_register(v1.WRITE)
async def xl_revert(change_id: str | None = None, workbook: str | None = None, force: bool = False) -> dict:
    """Put a range back as xl_snapshot (or a v1 xl_* edit) recorded it. Default: the latest change. Not saved."""
    from .workbench import revert
    return await _user(lambda h: revert(change_id, workbook, force), "xl_revert")
```

  Use the existing `_user` helper's real signature.

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): xl_snapshot and xl_revert - a safety net for Excel MCP edits`.

---

### Task 5: `xl_tidy`

**Files:** `engine/smartass/workbench.py`, `engine/smartass/server.py`. Test: extend `test_v2_workbench.py`.

**Rule:**
- `xl_tidy(workbook, sheet, rows=None, last_col=None)` always runs the tn-formelstil check (v1 `fmt.check`) and
  returns `{ok, block, warn, counts}`.
- When `rows` (`{"1": "title", "4-9": "line", ...}`) and `last_col` are given, it first TN-styles the sheet with v1
  `fmt.format_sheet`. That call snapshots first, so it can be reverted.

- [ ] **Step 1: Write the failing test**

```python
def test_tidy_checks_and_optionally_styles(monkeypatch):
    log = []
    monkeypatch.setattr(workbench, "_format_sheet", lambda wb, sh, last_col, rows: log.append("style") or
                        {"change_id": "c9"})
    monkeypatch.setattr(workbench, "_check", lambda wb, sh: {"ok": True, "block": [], "warn": [], "counts": {}})
    assert workbench.tidy("B", "S")["check"]["ok"] is True and log == []
    out = workbench.tidy("B", "S", rows={"1": "title"}, last_col="F")
    assert log == ["style"] and out["styled"]["change_id"] == "c9"
```

- [ ] **Step 2:** Run it and see it fail.

- [ ] **Step 3: Implement** in `workbench.py`:

```python
def _format_sheet(workbook, sheet, last_col, rows):
    from smartview_mcp.workbench import fmt
    return fmt.format_sheet(workbook, sheet, last_col, rows)


def _check(workbook, sheet):
    from smartview_mcp.workbench import fmt
    return fmt.check(workbook, sheet)


def tidy(workbook, sheet, rows: dict | None = None, last_col: str | None = None) -> dict:
    out: dict = {}
    if rows and last_col:
        out["styled"] = _format_sheet(workbook, sheet, last_col, rows)
    out["check"] = _check(workbook, sheet)
    return out
```

  Add the `xl_tidy` tool in `server.py` (`v1.WRITE`, wrapped in `_user`), with the docstring: "TN report style
  (when rows + last_col are given) and the tn-formelstil formula check; zero BLOCK findings before work is done."

- [ ] **Step 4:** `uv run pytest tests/v2 -q` passes.
- [ ] **Step 5: Commit.** Message: `feat(v2): xl_tidy - TN style and the formula check`.

---

### Task 6: Instructions and tool list

**Files:** `engine/smartass/server.py`. Test: `test_v2_server.py`.

- [ ] **Step 1:** Update the tests:
  - `TOOL_NAMES == ("fin_analyze", "fin_lookup", "fin_status", "fin_refresh", "xl_snapshot", "xl_revert",
    "xl_tidy")`.
  - `INSTRUCTIONS` mentions `xl_snapshot`, `Excel MCP` and `/mcp`.
  - `INSTRUCTIONS` stays under 2500 characters.
- [ ] **Step 2:** Run them and see them fail.
- [ ] **Step 3:** Update `TOOL_NAMES`. Append two sentences to `INSTRUCTIONS`, shortening elsewhere to stay under
  the limit:

```text
6) Editing a workbook the user has open with Excel MCP: xl_snapshot first, then edit, xl_revert to undo; finish
with xl_tidy (zero BLOCK findings). 7) After the server reloads, Claude Code may keep an old tool schema: if a
parameter seems missing, ask the user to run /mcp and reconnect smartass2.
```

- [ ] **Step 4:** `uv run pytest -q && uv run ruff check .` are both green.
- [ ] **Step 5:** Add a CHANGELOG line under `## Unreleased` / `### Added`:
  `v2 robustness and toolkit: AutoSave paused while writing, session workbook kept across /mcp reconnects, a lost
  first request retried, xl_snapshot / xl_revert / xl_tidy.`
- [ ] **Step 6: Commit.** Message: `feat(v2): instructions and tool list for the Excel toolkit`.

---

### Task 7: Live check (only on Tom's go)

After `/mcp` reconnects `smartass2`, the controller:
1. Checks that `fin_status` responds. With no reconnect step in between, a follow-up on Q3 without `target` works,
   which proves the session survived the reconnect.
2. Runs `fin_analyze` on a new quick question. The reply has no AutoSave note (our own file), and the file shows
   AutoSave as it was before.
3. Runs `xl_snapshot` on `Q3 Answer!B1`, changes B1 with Excel MCP, runs `xl_revert` and checks that the old
   title is back.
4. Runs `xl_tidy` on `Q3 Answer` (check only). Expect zero BLOCK findings, or report them.

Whatever fails becomes a task in 1E with the observed error.
