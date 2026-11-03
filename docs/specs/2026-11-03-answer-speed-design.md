# Answer speed: data in parallel while Excel opens, fewer Excel calls (design, 2026-11-03)

Owner goal: less waiting on **every** answer type (owner, 2026-11-03: "all of them", approach "overlap + trim").
Sheets, formulas, checks and numbers stay exactly as today. Follows `docs/specs/2026-11-02-speed-design.md`
(drills: done, step 1 + 1b) and replaces its steps 2-3.

Evidence: the last 40 answer runs on the owner's PC (`%LOCALAPPDATA%\smartass\logs\runs.jsonl`, times only):

| Stage | Mean | Worst | Cause |
|---|---:|---:|---|
| `live read` (pnl recipe) | 14 s | 76 s | `extract._fill` retrieves its grids one at a time |
| `open target` | 4 s | 39 s | a cold Excel start, plus a fixed 12 s wait for a login window |
| Excel writing (`answer`, `extra 'By project'`, `write lines d1/d2`, `write source sheet`, `log sources`) | ~8-12 s per answer | - | many single COM calls |
| `fetch (HTTP)` | 1 s | 8 s | `pipeline._new` fetches each pull before writing it, one at a time |

Since step 1b the HTTP route is safe in parallel: a pool of Smart View sessions (never two requests on one sID),
at most `SVHTTP_SESSIONS` (8) in flight per cube; 12 grids at once equal one at a time (live, 3 rounds).

## 1. Data in parallel, while Excel opens

### `svhttp.SvHttp.fetch_many(items) -> list[Fetched | Exception]`
- `items`: `(conn, pov, rows, cols)` tuples. Runs `fetch` for each on a thread pool (size `SVHTTP_SESSIONS`);
  results in input order; an item's exception is returned in its place (never raised for the batch).
- The per-cube cap stays the session pool's semaphore; `fetch_many` adds no state.

### Answer pulls prefetched (`core/pipeline.py`, `core/source.py`)
- `prefetch(spec) -> Prefetch`: starts `fetch_many` for every pull that will go over HTTP (same size rule as
  `source.pull`: not when `cells / 200 > SVHTTP_TIMEOUT_S`) in a background thread and returns at once. The
  recipes and `sv_answer` call it **before** `open target`, so the grids arrive while Excel opens/creates the
  workbook. `pipeline.run` accepts it (`prefetched=`); without it, `run` starts one itself (same behaviour, less
  overlap).
- `source.pull(..., got=...)` takes the prefetched `Fetched` (written as today) or the prefetched exception: an
  `SvHttpError` goes to Smart View in Excel with today's note; an `InputError` is raised as today.
- Follow-ups prefetch only the pulls that are not reused.
- Stage names stay (`fetch (HTTP)` becomes the wait for the prefetch, usually ~0 s), so `runs.jsonl` stays
  comparable.

### P&L reads in parallel (`extract._fill`)
- The `retrieve` calls of one `fill` run on a thread pool (size `SVHTTP_SESSIONS`); results are processed and
  stored **in grid order**, exactly as the sequential loop does (same rows, errors, warnings, checks). A COM
  fallback inside `retrieve` keeps queueing on the host thread as today.
- Base read and top-up are separate `fill` calls; top-up depends on nothing from the base read except
  `_forecast`, which stays after it.

## 2. Excel startup
- Spike first (live, read-only): start Excel on the TN template with no Smart View call; does the EPM login window
  appear? Record the answer in this file.
- If it does not: `wait_login_cleared(pid, appear_grace=12)` in `excel/target.py` (both places) becomes
  `appear_grace=2`; the wait still lasts as long as a login window is showing. If it does: unchanged.
- The user's running Excel stays the first choice (today's behaviour).

## 3. Fewer Excel calls
- Measure: a COM call counter per stage (`excel/jobs.com` increments a thread-local counter read by `stage()`),
  logged with the timings, on a consultant answer and a P&L answer.
- Cut the largest counts in: `write lines d1/d2` (`drill_xl`), `extra 'By project'`, `write source sheet`, the
  answer write/styling and `log sources`: one `Range.Value` block per contiguous range, number formats and styles
  per range, no per-cell reads back where the written value is known.
- Target: Excel writing under ~6 s per answer.

## Tests
- `fetch_many`: input order with out-of-order completion, per-item exceptions, never more than the pool cap at once.
- `prefetch` + `source.pull(got=...)`: Fetched written as before; SvHttpError -> Excel route with the note;
  InputError raised; too-big pulls not prefetched; follow-up prefetches only changed pulls.
- `_fill` parallel: the same FactStore rows / FillResult as the sequential loop for a fake retrieve with random
  delays (order kept), one failing grid named, the others stored.
- COM counter: counts per stage in the timings; existing writer tests keep passing (output unchanged).
- Live: consultant, P&L and trend answers before/after; all checks pass; times into
  `docs/audit/2026-11-02-speed-bench.md` (times only).

## Out of scope
Building the workbook without Excel; caching; changing what the sheets contain.
