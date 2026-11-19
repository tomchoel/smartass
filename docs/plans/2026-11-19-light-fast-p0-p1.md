# Light and fast, phases 0-1: baseline timings and drill lines over the token - Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Every v2 answer logs its stage timings again, and drill lines are fetched with the user's OAuth token instead of the hidden browser, after which the whole browser stack (Playwright) is deleted.

**Architecture:** One small token session per EPM pod (`epm/epmhttp.py`: stdlib HTTP, cookie jar, bearer from `epm/oauth.py`, a path allowlist) replaces the browser as the carrier of the drill requests. `epm/drill_token.py` offers the same three calls `drill_direct.DirectDrill` offers today (`query`, `lines`, `fetch`), so `drillstep.py` only swaps its source. The browser route stays behind a switch until a live comparison has matched it; then it is deleted together with `BrowserTransport`.

**Tech Stack:** Python 3.12, stdlib `urllib` / `http.cookiejar` / `asyncio`, pytest. No new dependency; `playwright` is removed.

**Spec:** `docs/specs/2026-11-19-light-fast-design.md` (phases 0 and 1). Phases 2-4 get their own plans.

## Global Constraints

- Read-only towards EPM: the token session sends nothing outside its allowlist; the only POST targets are `/HyperionPlanning/SmartView` (bodies admitted by `svxml.check_read_only`) and `/aif/ui/model/drilldown/getData`.
- A token, the `x-appbuilder-client-id` key, a cookie or a URL query string never appears in a log line, an exception text or a tool reply (`drillcore.redact` on every URL that is logged).
- The bearer is only ever sent to the origin of a pod that has a scope in `config.OAUTH_SCOPES`.
- Nothing is cached: every drill is live (owner decision 2026-11-03).
- No change to tool names, tool parameters or reply fields.
- The old route is deleted only after Task 7's live comparison is recorded as passed.
- Tests never reach EPM, Excel or a browser (`engine/tests/conftest.py` keeps the suite offline); live checks are scripts under `scripts/` run on the owner's PC.
- Line length 120; `uv run --frozen ruff check .` and `uv run --frozen pytest -q` green before every commit; commits end with the session's `Co-Authored-By` trailer.
- Work on branch `feat/drill-token` cut from `main`; one PR for phase 0 (Tasks 1-2) and one for phase 1 (Tasks 3-9).

## Review Focus

1. **The hour ends mid-drill:** a request answered 401 gets one new access token (`oauth.access_token(base, reject=old)`) and is sent again; the drill does not fail and does not start over. (Task 4)
2. **The server session or its key is gone mid-paging:** a `getData` answered with a redirect to `/interop/epmexit.jsp`, or with 200 that is not JSON, takes a new key once and resumes at the same offset; a second such answer is a `DrillError` that names the offset. (Task 5)
3. **Many pieces at once:** twenty pieces of a split drill started together take the key once, not twenty times, and a renewal needed by several pieces at once happens once. (Task 5)
4. **Not signed in:** a tool call without a valid sign-in answers with the sign-in page and code for every pod that needs one - not a stack trace, not the Excel fallback, and not a second code when one is already waiting. (Task 8)
5. **A drill link that points elsewhere:** a link whose host is not the pod's, or a redirect to another host, is refused before the bearer is sent. (Task 4)

---

## File structure

| File | Responsibility |
|---|---|
| `engine/smartass/core/runlog.py` (new) | One JSON line per tool run in `config.RUN_LOG`; spans collected per tool call. |
| `scripts/bench_answers.py` (new) | Live timing of the test-set questions on the owner's PC; writes times only. |
| `engine/smartass/epm/drillcore.py` (new) | The drill's pure parts, moved unchanged out of `drill_browser.py` and `drill_direct.py`: errors, `redact`, `Pager`, `SuiQuery`, `sui_url`, `parse_sui`. |
| `engine/smartass/epm/epmhttp.py` (new) | `EpmSession`: allowlisted requests to one pod with the bearer and a cookie jar. |
| `engine/smartass/epm/drill_token.py` (new) | `TokenDrill`: the server key, the drill query, the view, the lines. |
| `engine/smartass/epm/drillstep.py` (modify) | Uses `drill_token`; the page route and its browser helpers go. |
| `engine/smartass/epm/oauth.py` (modify) | `pending_logins()`: start or reuse a device login per pod, finished in the background. |
| `engine/smartass/core/serverkit.py` (modify) | `guard` turns `oauth.NotSignedIn` into the sign-in message; warm-up and browser shutdown go. |
| `engine/smartass/epm/svhttp.py` (modify) | `BrowserTransport` goes; `BearerTransport` stands alone. |
| deleted | `epm/drill_browser.py`, `epm/drill_tabs.py`, `epm/drill_direct.py` and their tests; the `playwright` dependency. |

---

### Task 1: Run log and spans for v2 answers (phase 0)

**Files:**
- Create: `engine/smartass/core/runlog.py`
- Modify: `engine/smartass/core/serverkit.py` (`guard`), `engine/smartass/server.py` (`_prepare`, `_new_question`, the `fin_analyze` follow-up path, `fin_refresh`), `engine/smartass/pipeline.py` (`excel_job`, `redraw_job`, `_draw`)
- Test: `engine/tests/test_runlog.py`

**Interfaces:**
- Produces:
  - `runlog.begin(tool: str) -> None` - start collecting for this tool call (a `contextvars` value, so it follows the call across `await` and `asyncio.to_thread`).
  - `runlog.span(name: str)` - context manager; adds wall seconds to `name` (summed when entered more than once). A no-op outside `begin`.
  - `runlog.add(timings: dict[str, float]) -> None` - merge seconds measured elsewhere (a `Timer.summary()`, `drilltime.take()`).
  - `runlog.note(**fields) -> None` - set `q`, `workbook`, `question`.
  - `runlog.finish(ok: bool, error: str | None = None) -> dict` - append one line to `config.RUN_LOG` and return it: `{"ts", "tool", "ok", "error", "q", "workbook", "question", "timings"}` with `timings["_total"]`.

- [ ] **Step 1: Write the failing tests**

```python
import json
from smartass.core import config, runlog

def test_a_run_is_one_json_line_with_its_spans(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "logs" / "runs.jsonl")
    runlog.begin("fin_analyze")
    with runlog.span("prepare"):
        with runlog.span("prepare.fetch"):
            pass
    with runlog.span("prepare"):
        pass
    runlog.add({"excel job.answer": 1.25, "com calls.excel job": 40})
    runlog.note(q=3, workbook="C:\\a\\b.xlsx", question="P&L Wholesale")
    line = runlog.finish(ok=True)
    stored = [json.loads(x) for x in (tmp_path / "logs" / "runs.jsonl").read_text(encoding="utf-8").splitlines()]
    assert stored == [line] and line["tool"] == "fin_analyze" and line["ok"] is True and line["q"] == 3
    assert set(line["timings"]) == {"prepare", "prepare.fetch", "excel job.answer", "com calls.excel job", "_total"}
    assert line["timings"]["excel job.answer"] == 1.25 and line["timings"]["_total"] >= line["timings"]["prepare"]

def test_outside_a_run_nothing_is_recorded_and_nothing_raises(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")
    with runlog.span("x"):
        pass
    runlog.add({"y": 1.0}); runlog.note(q=1)
    assert not (tmp_path / "runs.jsonl").exists()

def test_a_failed_run_keeps_its_timings_and_a_redacted_error(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")
    runlog.begin("fin_drill")
    line = runlog.finish(ok=False, error="boom https://x/y?sso_token=abcdefghijklmnopqrstuvwxyz")  # allow-secret
    assert line["ok"] is False and "abcdefghij" not in line["error"] and "_total" in line["timings"]

def test_a_log_that_cannot_be_written_never_breaks_the_tool(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "RUN_LOG", tmp_path)          # a directory: open() fails
    runlog.begin("fin_analyze")
    assert runlog.finish(ok=True)["ok"] is True

def test_spans_follow_the_call_into_a_worker_thread(monkeypatch, tmp_path):
    import asyncio
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")
    async def tool():
        runlog.begin("fin_analyze")
        def work():
            with runlog.span("prepare.fetch"):
                pass
        await asyncio.to_thread(work)
        return runlog.finish(ok=True)
    assert "prepare.fetch" in asyncio.run(tool())["timings"]
```

- [ ] **Step 2: Run them to see them fail**

Run: `uv run --frozen pytest -q tests/test_runlog.py` (from `engine/`). Expected: FAIL, `No module named 'smartass.core.runlog'`.

- [ ] **Step 3: Implement `core/runlog.py`** with the five functions above. The collector is one mutable dict held in a `contextvars.ContextVar` (a thread started by `asyncio.to_thread` copies the context, so it sees the same dict; guard the dict with a lock). `error` goes through `logsetup.redact`. The write is append-only, UTF-8, one `json.dumps(..., ensure_ascii=False, default=str)` line; any `OSError` is logged at info and swallowed.

- [ ] **Step 4: Wire it in.**
  - `serverkit.guard`: `runlog.begin(fn.__name__)` before the call for the tools `fin_analyze`, `fin_drill` and `fin_refresh` only; `runlog.finish(ok=True)` after it, `runlog.finish(ok=False, error=str(e))` in both `except` branches (before the `raise`).
  - `server._prepare`: the whole function in `runlog.span("prepare")`, the provider fetch in `runlog.span("prepare.fetch")`, the drills (`_line_sets`) in `runlog.span("prepare.drill")`, and after the drills `runlog.add({f"drills.{k}": v for k, v in drilltime.take().items()})`.
  - The Excel job runs on the host's COM thread, which does **not** inherit the tool call's context, so its timings travel back with its result. Add `pipeline.with_timings(job)`: it returns a job that calls `Timer.start()` (today nothing starts a Timer in v2, so `timing.stage` and `count_com` are silent), runs `job(h)` and returns `(result, Timer.current().summary())`. In `server.py` every `_user(lambda h: excel_job(...))` and `_user(lambda h: redraw_job(...))` call goes through it; on the event-loop side the summary is logged with `runlog.add({f"excel job.{k}": v for k, v in summary.items()})` and `runlog.note(q=out["q"], workbook=out["workbook"], question=pre.spec.question)`.
  - Stages inside the job (fixed names; Task 2 reads them): `open target` around `_target`; `draw` around `_draw`, and inside `_draw` one stage around each thing it calls - `data sheet`, `answer`, `charts`, `lines` (`_write_lines`), `sources` (`_write_sources`), `read back` (`_check`); `save` around `_save`.

- [ ] **Step 5: Add the wiring test** to `engine/tests/test_runlog.py`:

```python
def test_guard_logs_a_run_for_fin_analyze_and_not_for_fin_lookup(monkeypatch, tmp_path):
    import asyncio, json
    from smartass.core import serverkit
    monkeypatch.setattr(config, "RUN_LOG", tmp_path / "runs.jsonl")
    async def fin_analyze():
        with runlog.span("prepare"):
            return {"q": 1}
    async def fin_lookup():
        return {}
    async def fin_drill():
        raise RuntimeError("no")
    asyncio.run(serverkit.guard(fin_analyze)()); asyncio.run(serverkit.guard(fin_lookup)())
    try:
        asyncio.run(serverkit.guard(fin_drill)())
    except Exception:
        pass
    lines = [json.loads(x) for x in (tmp_path / "runs.jsonl").read_text(encoding="utf-8").splitlines()]
    assert [(x["tool"], x["ok"]) for x in lines] == [("fin_analyze", True), ("fin_drill", False)]
    assert "prepare" in lines[0]["timings"]
```

- [ ] **Step 6: Run** `uv run --frozen ruff check . && uv run --frozen pytest -q`. Expected: all green (1424 existing + 6 new).

- [ ] **Step 7: Commit** `feat(core): run log with stage timings for v2 answers`.

### Task 2: Baseline of today's engine (phase 0, live)

**Files:**
- Create: `scripts/bench_answers.py`
- Create: `docs/audit/2026-11-19-baseline-v2.md` (written by the script's output, times only)

**Interfaces:**
- Consumes: the `runs.jsonl` lines of Task 1 (`timings` keys `prepare`, `prepare.fetch`, `prepare.drill`, `excel job.<stage>`, `com calls.excel job`, `_total`).
- Produces: `python scripts/bench_answers.py --rounds 3 [--case b01 ...]` prints a markdown table: per test-set case the median `_total`, `prepare.fetch`, `prepare.drill`, `excel job` and COM-call count; one line with the medians over all cases. No figure from EPM is printed.

- [ ] **Step 1: Write `scripts/bench_answers.py`.** It loads the cases with `smartass.testset.runner.load_cases()`, and for each case and round calls the server's own `_new_question(spec, target)` inside `serverkit.guard` under the tool name `fin_analyze` (so Task 1 logs it), with `target="session"` after one initial `target="new"` call (all rounds land in ONE new workbook in the analyser folder). It then reads back the lines it caused from `config.RUN_LOG` (by `ts` >= its own start) and prints the table. A failed run is listed with its error, not hidden.

- [ ] **Step 2: Run it live** (owner's PC, Excel open, signed in): `engine\.venv\Scripts\python.exe scripts\bench_answers.py --rounds 3`. Expected: a table with 9 cases; no case missing.

- [ ] **Step 3: Write `docs/audit/2026-11-19-baseline-v2.md`:** the table, the date, the engine version, the Excel build (`Application.Version` + `Build`), and one paragraph naming the three largest stages. This is the baseline every later phase reports against.

- [ ] **Step 4: Commit** `docs(audit): v2 baseline timings; scripts/bench_answers.py`, push the branch, open the phase-0 PR (`Tasks 1-2`), merge when CI is green (the owner's standing flow), then bump the version with `scripts\bump.ps1` so running servers reload.

### Task 3: `drillcore.py` - the drill's pure parts in one module

**Files:**
- Create: `engine/smartass/epm/drillcore.py`
- Modify: `engine/smartass/epm/drill_browser.py`, `drill_direct.py`, `drill_tabs.py`, `drillstep.py`, `svhttp.py`, `engine/smartass/core/serverkit.py` (imports only)
- Test: `engine/tests/test_drillcore.py` (the existing tests of these names, moved from `test_drill.py` and `test_drill_direct.py`)

**Interfaces:**
- Produces, all moved **unchanged** (same names, same behaviour):
  - from `drill_browser.py`: `DrillError`, `DrillBusy`, `TooManyLines`, `CellTooWide`, `redact`, `_payload` (now public `payload`), `_expired` (now `expired`), `Pager`.
  - from `drill_direct.py`: `NoDrillData`, `SignedOut`, `SuiQuery`, `sui_url`, `parse_sui`.
- `drill_browser.py` and `drill_direct.py` re-import these names from `drillcore` so every existing import keeps working until Task 9.

- [ ] **Step 1: Move the tests first.** Create `tests/test_drillcore.py` from the tests that exercise only the names above (`test_sui_url_*`, `test_parse_sui_*` from `test_drill_direct.py`; the `redact` and `Pager`/`page_through` bookkeeping tests in `test_drill.py` that do not start a browser), importing from `smartass.epm.drillcore`.
- [ ] **Step 2: Run** `uv run --frozen pytest -q tests/test_drillcore.py`. Expected: FAIL, `No module named 'smartass.epm.drillcore'`.
- [ ] **Step 3: Create `drillcore.py`** by moving the code; leave `from .drillcore import ...` lines in the two old modules.
- [ ] **Step 4: Run the whole suite.** Expected: green, same count as before.
- [ ] **Step 5: Commit** `refactor(drill): pure drill parts in epm/drillcore.py`.

### Task 4: `epmhttp.EpmSession` - allowlisted requests with the token

**Files:**
- Create: `engine/smartass/epm/epmhttp.py`
- Test: `engine/tests/test_epmhttp.py`

**Interfaces:**
- Consumes: `oauth.scope_for(base)`, `oauth.access_token(base, reject=None)`, `oauth.origin(base)`, `oauth.OAuthError`, `oauth.NotSignedIn`; `svxml.check_read_only(body)`; `drillcore.redact`.
- Produces:
  - `class Response(NamedTuple): status: int; text: str; headers: dict[str, str]` (header names lower-cased).
  - `class Refused(ValueError)` - the request is not on the allowlist or leaves the pod.
  - `class EpmSession` with `__init__(self, base: str, tokens=oauth, opener=None)` and `request(self, method: str, url: str, headers: dict | None = None, body: str | None = None, timeout: float = 120.0) -> Response`. `url` is a path (`/aif/...`) or an absolute URL on the pod.
  - `session(base: str) -> EpmSession` - one per pod origin per process.
- The allowlist, exactly:

| Method | Path | Extra rule |
|---|---|---|
| POST | `/HyperionPlanning/SmartView` | `svxml.check_read_only(body)` |
| GET | `/aif/ui/model/applications/servicetype` | - |
| GET | `/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp` | the query has `FORMAT=SUI` |
| GET | `/aif/ui/model/view` | the query has `page=DRILL_THROUGH` |
| POST | `/aif/ui/model/drilldown/getData` | - |

- Behaviour: redirects are never followed (a 3xx is returned as the answer); one cookie jar per session; `Authorization: Bearer` on every request; a 401 answer takes one new token with `reject=<the refused one>` and sends the request once more; `OSError` from the network is raised as it is.

- [ ] **Step 1: Write the failing tests** (a fake `opener` with `open(req, timeout)` records `req.full_url`, `req.get_method()`, `req.headers`; a fake `tokens` as `_Tokens` in `tests/test_svhttp.py`):

```python
POD = "https://plan1.example"

def test_only_the_allowlist_is_sent(make):                      # make(...) -> (session, opener, tokens)
    s, op, tok = make()
    for method, url, body in [("POST", "/aif/ui/model/integration/run", "{}"),
                              ("DELETE", "/aif/ui/model/view?page=DRILL_THROUGH", None),
                              ("GET", "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=x", None),
                              ("GET", "/aif/ui/model/view?page=SETUP", None),
                              ("POST", "/HyperionPlanning/SmartView", "<req_SubmitData></req_SubmitData>"),
                              ("GET", "/HyperionPlanning/rest/v3/applications", None)]:
        with pytest.raises(ValueError, match="refused"):
            s.request(method, url, body=body)
    assert op.sent == [] and tok.asked == []

def test_the_bearer_never_leaves_the_pod(make):
    s, op, tok = make()
    for url in ("https://evil.example/aif/ui/model/drilldown/getData", "http://plan1.example/aif/ui/model/view?page=DRILL_THROUGH",
                "//evil.example/aif/ui/model/view?page=DRILL_THROUGH"):
        with pytest.raises(epmhttp.Refused):
            s.request("GET" if "view" in url else "POST", url, body=None if "view" in url else "{}")
    assert op.sent == []

def test_a_pod_without_a_scope_is_refused(make):
    with pytest.raises(epmhttp.Refused, match="no OAuth scope"):
        epmhttp.EpmSession("https://other.example", tokens=make()[2]).request("GET", "/aif/ui/model/applications/servicetype")

def test_a_request_carries_the_bearer_and_the_sessions_cookies(make):
    s, op, tok = make(answers=[(200, "{}", {"Set-Cookie": "JSESSIONID=abc; Path=/"}), (200, "{}", {})])
    s.request("GET", "/aif/ui/model/applications/servicetype", headers={"page_id": "smartass"})
    s.request("GET", "/aif/ui/model/view?page=DRILL_THROUGH&appId=1&partitionkey=2")
    assert op.sent[0].get_header("Authorization") == "Bearer at-1" and op.sent[0].get_header("Page_id") == "smartass"
    assert "JSESSIONID=abc" in (op.sent[1].get_header("Cookie") or "")

def test_a_401_gets_one_new_token_and_one_more_try(make):        # Review Focus 1
    s, op, tok = make(tokens=["at-1", "at-2"], answers=[(401, "", {}), (200, '{"ok": 1}', {})])
    r = s.request("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0", body="{}")
    assert r.status == 200 and [x.get_header("Authorization") for x in op.sent] == ["Bearer at-1", "Bearer at-2"]
    assert tok.asked == [(POD, None), (POD, "at-1")]
    s2, op2, _ = make(tokens=["at-1", "at-2"], answers=[(401, "", {})])
    assert s2.request("GET", "/aif/ui/model/applications/servicetype").status == 401 and len(op2.sent) == 2

def test_a_redirect_is_an_answer_not_followed(make):             # Review Focus 5
    s, op, _ = make(answers=[(302, "", {"Location": "https://evil.example/x"})])
    r = s.request("POST", "/aif/ui/model/drilldown/getData?limit=5000&offset=0", body="{}")
    assert r.status == 302 and r.headers["location"] == "https://evil.example/x" and len(op.sent) == 1

def test_not_signed_in_is_raised_as_it_is(make):
    s, op, _ = make(tokens=[oauth.NotSignedIn("not signed in")])
    with pytest.raises(oauth.NotSignedIn):
        s.request("GET", "/aif/ui/model/applications/servicetype")

def test_one_session_per_pod(monkeypatch):
    monkeypatch.setattr(epmhttp, "_sessions", {})
    assert epmhttp.session("https://PLAN1.example:443/x") is epmhttp.session("https://plan1.example")
```

- [ ] **Step 2: Run** `uv run --frozen pytest -q tests/test_epmhttp.py`. Expected: FAIL, module missing.
- [ ] **Step 3: Implement `epmhttp.py`** (`urllib.request.build_opener` with a no-redirect handler and `HTTPCookieProcessor`, as `svhttp.BearerTransport._opener` does today; `urllib.error.HTTPError` is an answer). The allowlist check runs before a token is asked for.
- [ ] **Step 4: Run the tests.** Expected: PASS.
- [ ] **Step 5: Commit** `feat(epm): EpmSession - allowlisted requests with the user's token`.

### Task 5: `drill_token.TokenDrill` - the key, the query, the view, the lines

**Files:**
- Create: `engine/smartass/epm/drill_token.py`
- Modify: `engine/smartass/core/config.py` (add `DRILL_PARALLEL = int(os.environ.get("SV_DRILL_PARALLEL", "4"))` - `getData` calls in flight at once)
- Test: `engine/tests/test_drill_token.py`

**Interfaces:**
- Consumes: `epmhttp.session(base).request(...)` and `epmhttp.Response` (Task 4); `drillcore.Pager`, `SuiQuery`, `sui_url`, `parse_sui`, `DrillError`, `SignedOut`, `NoDrillData`, `CellTooWide`, `TooManyLines` (Task 3); `config.EPM_BASE`, `config.DRILL_PAGE`.
- Produces (the same calls and result shape as `drill_direct.DirectDrill`, so `drillstep` needs no other change):
  - `class TokenDrill` with `__init__(self, base: str | None = None, http=None)` (`http`: an object with `request`, default `epmhttp.session(base)`);
  - `async query(self, url: str) -> SuiQuery`;
  - `async lines(self, q: SuiQuery, warm_url: str, max_lines: int, progress=None) -> dict` returning `{"items": list[dict], "total": int | None, "view": dict, "secs": float, "query": SuiQuery}` (`warm_url` is accepted and ignored: there is no page to warm);
  - `async fetch(self, url: str, max_lines: int, progress=None) -> dict`;
  - `shared() -> TokenDrill` (one per process).
- The protocol, exactly (`PAGE_ID = "smartass"`):
  1. Key: `GET /aif/ui/model/applications/servicetype` with header `page_id: PAGE_ID`; the key is the response header `x-appbuilder-client-id`. No such header: `DrillError("EPM gave no session key for the drill")`.
  2. Every later request carries `{"Accept": "application/json, text/javascript, */*; q=0.01", "Content-Type": "application/json", "X-Requested-With": "XMLHttpRequest", "page_id": PAGE_ID, "x-appbuilder-client-id": <key>}`.
  3. Query: `GET sui_url(base, url, now_ms)`, read with `parse_sui`.
  4. View: `GET /aif/ui/model/view?page=DRILL_THROUGH&appId={q.app_id}&partitionkey={q.partition}&_={now_ms}`, cached per `(app_id, partition)` for the life of the key.
  5. Lines: `POST /aif/ui/model/drilldown/getData?limit={config.DRILL_PAGE}&offset={off}` with body `q.body`, paged with `Pager(max_lines, progress, True)`.
- "Gone" means: status 301, 302, 303, 307, 401 or 403, or status 200 whose body is not the expected JSON. On "gone" the key is dropped, taken again and the same request sent once more; `Pager`'s own single `recover` covers `getData`. Blocking calls run in `asyncio.to_thread`; at most `config.DRILL_PARALLEL` `getData` calls are in flight.

- [ ] **Step 1: Write the failing tests** (a fake `http` whose `request(method, url, headers=None, body=None, timeout=...)` records calls and answers by path; `KEY = {"x-appbuilder-client-id": "k1"}`; `SUI` and `PAGE` helpers build the JSON the page gets, as `tests/test_drill_direct.py` does today):

```python
URL = "aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp?module=aif.drilldown&DRILLURL=x&sso_token=t&attribute=Account.id.C60"

def test_fetch_takes_the_key_once_then_query_view_lines(fake):
    h = fake(lines=[PAGE(items=3, total=3)])
    r = asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))
    assert [c.path for c in h.calls] == ["/aif/ui/model/applications/servicetype", "/aif/modules/com/hyperion/aif/gl/main/FdmeeMain.jsp",
                                         "/aif/ui/model/view", "/aif/ui/model/drilldown/getData"]
    assert h.calls[0].headers == {"page_id": "smartass"}
    assert all(c.headers["x-appbuilder-client-id"] == "k1" and c.headers["page_id"] == "smartass" for c in h.calls[1:])
    assert "FORMAT=SUI" in h.calls[1].url and "DRILLURL" not in h.calls[1].url
    assert h.calls[3].method == "POST" and "limit=5000&offset=0" in h.calls[3].url and h.calls[3].body == r["query"].body
    assert len(r["items"]) == 3 and r["total"] == 3 and r["view"] == {"items": [{"name": "AMOUNTX"}]}

def test_lines_are_paged_until_has_more_ends(fake):
    h = fake(lines=[PAGE(items=5000, total=7000, more=True), PAGE(items=2000, total=7000)])
    r = asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 10_000))
    assert len(r["items"]) == 7000 and [c.url.split("offset=")[1] for c in h.calls if c.method == "POST"] == ["0", "5000"]

def test_a_lost_session_mid_paging_takes_a_new_key_and_resumes_at_the_same_offset(fake):   # Review Focus 2
    h = fake(keys=["k1", "k2"], lines=[PAGE(items=5000, total=7000, more=True), EXIT_302, PAGE(items=2000, total=7000)])
    r = asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 10_000))
    posts = [c for c in h.calls if c.method == "POST"]
    assert len(r["items"]) == 7000 and [p.url.split("offset=")[1] for p in posts] == ["0", "5000", "5000"]
    assert [p.headers["x-appbuilder-client-id"] for p in posts] == ["k1", "k1", "k2"]

def test_a_second_lost_session_is_an_error_that_names_the_offset(fake):
    h = fake(keys=["k1", "k2"], lines=[EXIT_302, EXIT_302])
    with pytest.raises(drillcore.DrillError, match="offset 0"):
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))

def test_a_page_instead_of_json_counts_as_a_lost_session(fake):
    h = fake(keys=["k1", "k2"], sui=[(200, "<html>sign in</html>"), SUI()], lines=[PAGE(items=1, total=1)])
    assert len(asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))["items"]) == 1
    assert [c.path for c in h.calls].count("/aif/ui/model/applications/servicetype") == 2

def test_twenty_pieces_at_once_take_the_key_once_and_at_most_four_are_in_flight(fake, monkeypatch):   # Review Focus 3
    monkeypatch.setattr(config, "DRILL_PARALLEL", 4)
    h = fake(lines=[PAGE(items=2, total=2)], slow_post=0.05)
    d = drill_token.TokenDrill("https://plan1.example", http=h)
    async def go():
        return await asyncio.gather(*(d.fetch(URL, 100) for _ in range(20)))
    assert all(len(r["items"]) == 2 for r in asyncio.run(go()))
    assert [c.path for c in h.calls].count("/aif/ui/model/applications/servicetype") == 1 and h.max_posts_in_flight <= 4

def test_a_renewal_needed_by_many_pieces_happens_once(fake):
    h = fake(keys=["k1", "k2"], lines=[EXIT_302] * 6 + [PAGE(items=1, total=1)], slow_post=0.02)
    d = drill_token.TokenDrill("https://plan1.example", http=h)
    async def go():
        return await asyncio.gather(*(d.fetch(URL, 100) for _ in range(6)))
    asyncio.run(go())
    assert [c.path for c in h.calls].count("/aif/ui/model/applications/servicetype") == 2

@pytest.mark.parametrize("sui, exc", [({"details": "EPMFDM-140373 more than 3 dimensions"}, drillcore.CellTooWide),
                                      ({"details": "The drill query does not return any data"}, drillcore.NoDrillData),
                                      ({"details": "something else"}, drillcore.DrillError)])
def test_epms_refusals_are_raised_as_before_and_no_lines_are_asked(fake, sui, exc):
    h = fake(sui=[(200, json.dumps(sui))])
    with pytest.raises(exc):
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))
    assert not [c for c in h.calls if c.method == "POST"]

def test_more_lines_than_allowed_is_too_many_lines_with_the_total(fake):
    h = fake(lines=[PAGE(items=5000, total=9000, more=True)])
    with pytest.raises(drillcore.TooManyLines) as e:
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 6000))
    assert e.value.total == 9000

def test_no_key_from_epm_is_a_clear_error_without_the_url(fake):
    h = fake(keys=[None])
    with pytest.raises(drillcore.DrillError, match="no session key") as e:
        asyncio.run(drill_token.TokenDrill("https://plan1.example", http=h).fetch(URL, 100))
    assert "sso_token" not in str(e.value)

def test_lines_for_a_built_query_reuse_the_view(fake):                 # drill_derive's siblings
    h = fake(lines=[PAGE(items=1, total=1), PAGE(items=2, total=2)])
    d = drill_token.TokenDrill("https://plan1.example", http=h)
    first = asyncio.run(d.fetch(URL, 100))
    q2 = dataclasses.replace(first["query"], body='{"QUERY": "other"}')
    second = asyncio.run(d.lines(q2, URL, 100))
    assert len(second["items"]) == 2 and [c.path for c in h.calls].count("/aif/ui/model/view") == 1
    assert [c for c in h.calls if c.method == "POST"][-1].body == '{"QUERY": "other"}'
```

- [ ] **Step 2: Run** `uv run --frozen pytest -q tests/test_drill_token.py`. Expected: FAIL, module missing.
- [ ] **Step 3: Implement `drill_token.py`.** The key and the view cache are guarded by one `asyncio.Lock` created per event loop; a renewal compares the key the failed request used with the current one and only renews when they are the same (the same idea as `DirectDrill.gen` today).
- [ ] **Step 4: Run the tests.** Expected: PASS.
- [ ] **Step 5: Commit** `feat(drill): drill lines over the token (TokenDrill)`.

### Task 6: `drillstep` can use either route

**Files:**
- Modify: `engine/smartass/core/config.py` (add `DRILL_ROUTE = os.environ.get("SV_DRILL_ROUTE", "browser")`, values `"browser"` | `"token"`)
- Modify: `engine/smartass/epm/drillstep.py` (`_drill_fetch`, `_derived_piece`)
- Test: `engine/tests/test_drill.py`

**Interfaces:**
- Consumes: `drill_token.shared()` (Task 5).
- Produces: `drillstep._lines_source()` returning `drill_token.shared()` when `config.DRILL_ROUTE == "token"`, else `drill_direct.shared_direct()`. With `"token"`, `config.DRILL_DIRECT` is ignored and `_drill_fetch` has no page route: `NoDrillData` gives an empty piece as today, and every other exception (`CellTooWide`, `TooManyLines`, any `DrillError`) is raised to the caller as it is.

- [ ] **Step 1: Write the failing tests** in `tests/test_drill.py`:

```python
def test_the_token_route_is_used_when_configured(monkeypatch):
    from smartass.epm import drill_token, drillstep
    class Fake:
        async def fetch(self, url, max_lines, progress=None):
            return {"items": [{"AMOUNTX": "1"}], "total": 1, "view": {"items": []}, "secs": 0.1, "query": None}
    monkeypatch.setattr(config, "DRILL_ROUTE", "token"); monkeypatch.setattr(config, "DRILL_DIRECT", True)
    monkeypatch.setattr(drill_token, "shared", lambda: Fake())
    assert asyncio.run(drillstep._drill_fetch("u", 10))["total"] == 1

def test_the_token_route_has_no_page_route_behind_it(monkeypatch):
    from smartass.epm import drill_token, drillcore, drillstep
    class Fails:
        async def fetch(self, url, max_lines, progress=None):
            raise drillcore.DrillError("getData returned HTTP 500 with an error at offset 0.")
    monkeypatch.setattr(config, "DRILL_ROUTE", "token"); monkeypatch.setattr(config, "DRILL_DIRECT", True)
    monkeypatch.setattr(drill_token, "shared", lambda: Fails())
    monkeypatch.setattr(drillstep, "_page_fetch", lambda *a: (_ for _ in ()).throw(AssertionError("page route used")))
    with pytest.raises(drillcore.DrillError, match="offset 0"):
        asyncio.run(drillstep._drill_fetch("u", 10))
```

- [ ] **Step 2: Run** them. Expected: FAIL (`config` has no `DRILL_ROUTE`).
- [ ] **Step 3: Implement** `_lines_source()` and use it in `_drill_fetch` and `_derived_piece`.
- [ ] **Step 4: Run the whole suite.** Expected: green.
- [ ] **Step 5: Commit** `feat(drill): SV_DRILL_ROUTE=token switches the drill to the token route`.

### Task 7: Live comparison of the two routes (the gate for deleting the browser)

**Files:**
- Create: `scripts/compare_drill_routes.py` (deleted again in Task 9)
- Create: `docs/audit/2026-11-19-drill-token-vs-browser.md`

**Interfaces:**
- Consumes: `drillstep.drill(cell, say=...)` under `SV_DRILL_ROUTE=browser` and `=token` (one fresh Python process per route, so the route is read at import).
- Produces: per cell a line `<cell>: lines N/N, same lines: True, tie browser/token: exact/exact, cube value equal: True, seconds B/T`; exit code 1 if any `same lines`, tie or cube value differs.

The three cells (FinLite, defaults for the other dimensions):
1. `{"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"}` - one piece, 2 839 lines on 19.11.
2. `{"All Accounts": "ac.C60", "All Entities": "un.SY", "All Time Periods": "Sep 26"}` - expected over 5 000 lines: split into child cells.
3. `{"All Accounts": "ac.C60080", "All Entities": "un.SY", "All Time Periods": "Yr 26"}` - EPM refuses it as too wide: split by entity, siblings built with `drill_derive`.

- [ ] **Step 1: Explain the view difference found in the spike.** Fetch the view definition of cell 1 on both routes and diff the two JSON documents key by key (print key paths and value types, not business text). Write the finding into the audit file. If the differences are confined to volatile fields (links, timestamps, ordering of equal items), continue. If column names, order or count differ, STOP and report to the owner: the token route is not equivalent yet.
- [ ] **Step 2: Write `scripts/compare_drill_routes.py`.** "Same lines" compares `DrillResult.rows` as sorted lists of `json.dumps(row, sort_keys=True)`; it also requires equal `columns`, equal `reconciliation["status"]` and `cell_value` within 0.005. Only counts, booleans and seconds are printed.
- [ ] **Step 3: Run it live** (signed in on plan1): `engine\.venv\Scripts\python.exe scripts\compare_drill_routes.py`. Expected: three lines, all `True`, exit code 0. If a cell no longer behaves as described (fewer than 5 000 lines, or not refused), pick the nearest cell that does and say so in the audit file.
- [ ] **Step 4: Run the test set's drill questions on the token route:** `SV_DRILL_ROUTE=token` and `scripts\bench_answers.py --rounds 1 --case b16` plus every other case whose spec has a `lines` block; each must finish with `checks.sheet_equals_model` ok.
- [ ] **Step 5: Write the audit file** (the output of steps 1, 3 and 4; times only, no figures) and **commit** `docs(audit): drill lines over the token match the browser route`.

### Task 8: Sign-in from the tool call; `BearerTransport` without the browser behind it

**Files:**
- Modify: `engine/smartass/epm/oauth.py`, `engine/smartass/epm/svhttp.py`, `engine/smartass/core/serverkit.py`, `engine/smartass/data/provider.py` (only if it catches broader than `SvHttpError`)
- Test: `engine/tests/test_oauth.py`, `engine/tests/test_svhttp.py`, `engine/tests/test_errors.py`

**Interfaces:**
- Produces:
  - `oauth.pending_logins() -> list[dict]`: for every configured pod whose `status()` state is not `"signed_in"`, start a device login (`start_login`) unless one started earlier is still valid, and finish it in a daemon thread (`finish_login`). Returns `[{"pod": str, "verification_uri": str, "user_code": str, "minutes_left": int}]` - empty when every pod is signed in. A pod whose login could not be started is returned with `"error": str` instead of a code.
  - `oauth.sign_in_message() -> str | None`: `None` when `pending_logins()` is empty, else exactly: `"EPM sign-in needed. Open {uri} and enter the code {code} ({host}, valid {n} min), approve with Telenor SSO, then ask again."` - one such sentence per pod, joined by a space.
  - `svhttp.BearerTransport(tokens=oauth, send=None)` - no `fallback` argument. `oauth.NotSignedIn` is raised to the caller unchanged; any other `oauth.OAuthError`, and a token EPM refuses twice, is `SvHttpError`; a pod without a scope is `SvHttpError("no OAuth scope is configured for {base}: add it under [oauth.scopes] in config.toml")`.
  - `serverkit.guard`: an `oauth.NotSignedIn` escaping a tool becomes `ToolError(oauth.sign_in_message() or str(e))`, with no "Next step" appended and no traceback logged.

- [ ] **Step 1: Write the failing tests.**

```python
# tests/test_oauth.py
def test_a_pending_login_is_started_once_and_reused_until_it_ends(oa, monkeypatch):          # Review Focus 4
    idcs = oa(device=[(200, {"device_code": "dc", "user_code": "ABCD", "verification_uri": "https://idcs.example/ui",
                             "expires_in": 300, "interval": 5})], token=[(400, {"error": "authorization_pending"})])
    monkeypatch.setattr(oauth, "_finish_in_background", lambda pod, device: None)
    first, again = oauth.pending_logins(), oauth.pending_logins()
    assert [p["user_code"] for p in first] == ["ABCD"] and again[0]["user_code"] == "ABCD" and len(idcs.to("device")) == 1
    assert oauth.sign_in_message() == ("EPM sign-in needed. Open https://idcs.example/ui and enter the code ABCD "
                                       "(strat.example, valid 5 min), approve with Telenor SSO, then ask again.")

def test_no_message_when_every_pod_is_signed_in(oa):
    oa(token=[(500, {})]); _signed_in()
    assert oauth.pending_logins() == [] and oauth.sign_in_message() is None

def test_a_login_that_cannot_start_says_why_instead_of_a_code(oa):
    oa(device=[(400, {"error": "invalid_scope", "error_description": "Invalid scope."})])
    assert "invalid_scope" in oauth.pending_logins()[0]["error"] and "invalid_scope" in oauth.sign_in_message()

# tests/test_svhttp.py  (replace the fallback tests of BearerTransport)
def test_not_signed_in_reaches_the_caller_as_it_is():
    t = svhttp.BearerTransport(tokens=_Tokens(oauth.NotSignedIn("not signed in")), send=lambda *a: OK_XML)
    with pytest.raises(oauth.NotSignedIn):
        t.post("https://strat.example", svxml.connect())

@pytest.mark.parametrize("tokens, send, msg", [
    (_Tokens(oauth.OAuthError("the identity domain did not answer")), lambda *a: OK_XML, "identity domain"),
    (_Tokens("at-1", "at-2"), lambda *a: GATE_401, "refused the access token"),
    (_Tokens("at-1", scopes=()), lambda *a: OK_XML, "no OAuth scope is configured")])
def test_everything_else_is_an_svhttp_error(tokens, send, msg):
    with pytest.raises(svhttp.SvHttpError, match=msg):
        svhttp.BearerTransport(tokens=tokens, send=send).post("https://strat.example", svxml.connect())

# tests/test_errors.py
def test_a_missing_sign_in_becomes_the_sign_in_message(monkeypatch):
    from smartass.epm import oauth
    monkeypatch.setattr(oauth, "sign_in_message", lambda: "EPM sign-in needed. Open https://idcs.example/ui and enter the code ABCD "
                        "(plan1.example, valid 5 min), approve with Telenor SSO, then ask again.")
    async def fin_analyze():
        raise oauth.NotSignedIn("not signed in to https://plan1.example: run 'smartass login'")
    with pytest.raises(ToolError) as e:
        asyncio.run(serverkit.guard(fin_analyze)())
    assert str(e.value).startswith("EPM sign-in needed.") and "Next step" not in str(e.value) and "code ABCD" in str(e.value)

def test_not_signed_in_is_not_swallowed_into_the_excel_fallback(db_path):
    """run_prepare with a provider whose transport raises NotSignedIn raises NotSignedIn (no 'routes: excel')."""
```

The last test is written against `pipeline.run_prepare` with `SvHttpProvider(svhttp.SvHttp(BearerTransport(tokens=_Tokens(oauth.NotSignedIn("x"))), ...))` and the smallest valid spec from `tests/v2`; its assertion is `pytest.raises(oauth.NotSignedIn)`.

- [ ] **Step 2: Run** them. Expected: FAIL.
- [ ] **Step 3: Implement**, in this order: `oauth` (pending logins live in a module dict `pod -> (device, deadline)` under a lock; `_finish_in_background(pod, device)` starts the daemon thread), `BearerTransport`, `guard`. Remove `config.SVHTTP_BEARER_RETRY_S` and `BearerTransport._off`. `svhttp.shared()` becomes `SvHttp(BearerTransport())`. `BrowserTransport` itself is deleted in Task 9.
- [ ] **Step 4: Run the whole suite.** Expected: green.
- [ ] **Step 5: Commit** `feat(oauth): a missing sign-in is answered with the page and code; no browser behind the bearer route`.

### Task 9: Delete the browser stack

**Files:**
- Delete: `engine/smartass/epm/drill_browser.py`, `drill_tabs.py`, `drill_direct.py`; `engine/tests/test_drill_direct.py`, `test_drill_inflight.py`, `test_drill_keepwarm.py`, `test_drill_tabs.py`; `scripts/compare_drill_routes.py`; `scripts/probe_sv_http.py`
- Modify: `engine/smartass/epm/drillstep.py`, `svhttp.py`, `engine/smartass/core/serverkit.py`, `core/config.py`, `engine/smartass/ops/doctor.py`, `ops/cli.py`, `engine/smartass/knowledge/cards_fetch.py`, `engine/defaults.toml`, `engine/pyproject.toml`, `engine/uv.lock`, `engine/tests/conftest.py`, `test_drill.py`, `test_svhttp.py`, `test_doctor.py`, `test_errors.py`
- Modify (docs): `docs/DESIGN.md`, `docs/troubleshooting.md`, `skills/smartass/reference/drill.md`, `skills/smartass/reference/operations.md`, `README.md`, `CHANGELOG.md`

**Interfaces:**
- Consumes: Task 7's audit file says "passed". Without it this task does not start.
- Produces: `drillstep._drill_fetch` always uses `drill_token.shared()`; `config.DRILL_ROUTE` and `config.DRILL_DIRECT` are gone; no module imports `playwright`.

- [ ] **Step 1: Write the failing test** in `engine/tests/test_doctor.py`:

```python
def test_nothing_in_the_engine_needs_a_browser():
    import pathlib, smartass
    root = pathlib.Path(smartass.__file__).parent
    hits = [p.name for p in root.rglob("*.py") if "playwright" in p.read_text(encoding="utf-8") or "drill_browser" in p.read_text(encoding="utf-8")]
    assert hits == []
    assert "playwright" not in (root.parent / "pyproject.toml").read_text(encoding="utf-8")
    assert "drill_browser" not in [c.__name__.lstrip("_") for c in doctor.CHECKS]
```

- [ ] **Step 2: Run** it. Expected: FAIL with the list of files.
- [ ] **Step 3: Delete and unhook.** Remove the three modules; in `drillstep.py` remove `browser()`, `_page_fetch`, `_tick`/`_ticks`/`_progress` and import the errors from `drillcore`; in `svhttp.py` remove `BrowserTransport`, `_origin`, `_fetch_ms`'s browser use, `_close`, the `drill_browser` import and `_first_line`'s use of it (use `drillcore.redact`); in `serverkit.py` remove `_warm_up` and the `shutdown_shared` call in `lifespan`; in `cards_fetch.py` remove the `shutdown_shared` call; in `doctor.py` remove `_drill_browser`; in `cli.py` remove the two `os.environ.setdefault("SV_DRILL_...")` lines; in `config.py` remove `DRILL_CHANNELS`, `DRILL_PROFILE_SLOTS`, `DRILL_PROFILE`, `DRILL_LAUNCH_TRIES`, `DRILL_LOGIN_TIMEOUT_S`, `DRILL_CAPTURE_TIMEOUT_S`, `DRILL_RELOAD_WAIT_S`, `DRILL_SETUP_S`, `DRILL_IDLE_S`, `DRILL_KEEP_WARM`, `DRILL_YIELD_IDLE_S`, `DRILL_KEEPALIVE_S`, `DRILL_TABS`, `DRILL_DIRECT`, `DRILL_ROUTE`; in `defaults.toml` remove `[drill]`; in `pyproject.toml` remove `playwright` and run `uv lock`; in `conftest.py` remove the `DRILL_DIRECT`, `DRILL_KEEP_WARM` and `DRILL_PROFILE_SLOTS` lines and, in the same autouse fixture, replace `drill_token.shared` with a function that raises `AssertionError("tests never reach EPM")` (a test that drills sets its own fake). `drillstep._children` keeps `config.DRILL_DERIVE` as its only condition for derived pieces.
- [ ] **Step 4: Sort the tests.** A test of deleted behaviour (browser launch, profile lock, tabs, warm tab, sign-in clicks, `BrowserTransport`) is deleted. A test of behaviour that still exists (`drillstep`'s splitting, budgets, derived pieces, reconciliation) stays and has its fakes moved from `drill_direct.shared_direct` to `drill_token.shared`.
- [ ] **Step 5: Run** `uv run --frozen ruff check . && uv run --frozen pytest -q`. Expected: green; the count drops only by the deleted tests (state the before and after counts in the PR).
- [ ] **Step 6: Update the docs** to what is left: the drill needs a sign-in (`smartass login` or the code in the tool reply), no Edge; the error rows about the hidden browser, blank drill pages and the drill profile go; `troubleshooting.md` loses `drill_browser`; `DESIGN.md`'s Runtime section loses "Drill browser" and its bearer paragraph says the token carries grids and drill lines; `CHANGELOG.md` gets an entry with Task 7's measured seconds.
- [ ] **Step 7: Live check** on the owner's PC after `scripts\bump.ps1`: `fin_drill` on cell 1 of Task 7 through the MCP server; then `smartass login --forget` followed by the same call must answer with the sign-in message, and after approval the next call must succeed.
- [ ] **Step 8: Commit** `feat(drill)!: drill lines over the token; the hidden browser and Playwright are gone`, push, open the phase-1 PR (`Tasks 3-9`) with Task 7's table in the description, merge when CI is green.

---

## What comes after this plan

Phase 2 (REST second engine, atlas over REST, hidden Excel deleted), phase 3 (file-built answers) and phase 4 (slim down) each get a plan of their own, written when the phase before it has landed. Until phase 2, a grid the token route cannot fetch still falls back to Smart View in Excel, and `drillstep._link` / `_child_values` keep their Excel fallback.
