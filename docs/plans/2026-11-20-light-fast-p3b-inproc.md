# Light and fast, phase 3b: the Excel work runs inside Excel

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan
> task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** the Excel part of an answer runs as 4-5 batches executed inside Excel by a small Excel-DNA add-in, instead
of 35-70 COM round trips each holding many cross-process calls. Target: a follow-up's Excel job at or under **1.5 s**
(2.3-2.8 s today), and bench round 3 (61 sheets) at or under the design gate of **3 s / 5 s** (3.29 / 5.19 s today).

**Architecture:** operations as data. Python records what Excel should do into a `Batch` (a list of late-bound
get / call / set operations with references, a few named verbs, a `finally` list). One contract, two executors: the
**in-process executor** writes the batch as JSON and calls one `Application.Run` into the add-in, which executes it
inside Excel and writes the result; the **COM executor** interprets the same operations over the COM objects in
Python (the fallback when the add-in cannot load, and what the unit tests run on the fake Excel).

**Tech stack:** Python 3.12, pywin32 (`xl/fastcom.py`), Excel-DNA 1.9.0 (vendored binaries, zlib licence), C# 5
compiled on first use with the .NET Framework compiler that ships with Windows.

**Why (measured by the controller, 2026-11-19/20; throwaway code in the session scratchpad, paths in the briefs):**
- Same built test answer in the owner's Excel (move-in, tables, formula blocks, calculate, read-back, identical
  result): 1 240 ms over COM, 520 ms as one in-process batch (about 300 ms of it is `Workbooks.Open`).
- One `Application.Run` into the add-in costs 3 ms; a plain COM property read 1.6 ms. Setting and restoring the four
  quiet settings costs about 17 ms.
- The file renderer's bench (docs/audit/2026-11-19-file-render-bench.md): the job grows 1.25 s from round 1 to
  round 3 with the COM-call count flat - reads that walk every sheet inside one call, the spec store, the move-in.
- `Application.RegisterXLL(path)` loads an add-in into a running Excel for the session, silently, also in the owner's
  Excel with Smart View, think-cell, Arixcel and his own Excel-DNA add-in loaded. The tenant blocks Office web
  add-ins; per-user XLLs are allowed.
- **Nothing repaints inside a Run** (ScreenUpdating toggle, RedrawWindow, a WM_PAINT pump: all shown by window
  snapshots not to paint). Between Runs Excel repaints once it has had at least 25 ms idle (0 and 10 ms: the old
  picture; 25 ms and more: the new one). A visible step is a batch boundary plus a pause of 40 ms.
- `BreakLink` works inside a batch as "read `LinkSources`, break those whose file name matches"; a bare
  `BreakLink(path)` without that read failed with 0x800A03EC.
- A spike that hard-coded the quiet settings left DisplayAlerts off in the owner's Excel after a failing operation:
  the settings are captured and restored in a C# `finally`, never hard-coded.
- A loaded add-in's file is locked: every build has its own folder, file name and function names.

## Global constraints

- The user's Excel is left as found (Calculation, ScreenUpdating, EnableEvents, DisplayAlerts, the active workbook)
  on every path, also when an operation fails mid-batch.
- The user's own workbook is never saved. Nothing writes to EPM. No figure from EPM is printed by a script.
- A failed job leaves the workbook as it was (today's rollback / park-restore semantics, unchanged).
- Same numbers and same sheets as today: `sheet_equals_model` OK; `scripts/check_followups.py` "checks OK".
- Loading the add-in must never show a window or dialog in the user's Excel; any failure to build, load or reach it
  falls back to the COM executor with one warning per Excel process. `SV_INPROC=0` forces the COM executor.
- Tests never start Excel. Live scripts use a hidden Excel of their own, except the two measurement scripts that are
  meant for the owner's Excel (`check_followups.py`, `bench_answers.py`).
- C# 5 only (`C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`). No new Python dependency.
- Commits end with `Co-Authored-By: Claude Fable 5.2 <noreply@anthropic.com>`.

---

### Task 1: The batch, its two executors and the add-in

**Files:**
- Create: `engine/smartass/xl/inproc/__init__.py`, `batch.py`, `runner.py`, `build.py`, `Runner.cs`,
  `vendor/ExcelDna64.xll`, `vendor/ExcelDna.Integration.dll`, `vendor/LICENSE-ExcelDna.txt` (copies of 1.9.0 are in
  `%LOCALAPPDATA%\svkit\addin\_nuget\exceldna.addin\tools\net462` and `...\exceldna.integration\lib\net462`)
- Create: `engine/tests/test_inproc_batch.py`, `test_inproc_runner.py`, `test_inproc_build.py`,
  `scripts/check_inproc.py`
- Modify: `engine/smartass/core/config.py` (`INPROC`, `INPROC_DIR`), `engine/pyproject.toml` (the new data files are
  packaged the way `config.BAS_FILES` are), `engine/tests/fakexl.py` only if an operation needs it

**Interfaces (produced):**
- `Batch(quiet=True)`: a recorder. `b.app` is the Application node; `b.workbook(name)` the open workbook of that
  name. Attribute access and calls on a node record operations and return nodes
  (`ws = wb.Worksheets("Q3 Data"); lo = ws.ListObjects.Add(1, ws.Range(ref), MISSING, 1)`); assignment records a set
  (`lo.Name = "Data_Q3"`); `b.ret(node, "name")` asks for that value back; `b.mark("name")` records the milliseconds
  reached; `b.verb("break_links", wb, file_name)` a named verb; `with b.always(): ...` records into the `finally`
  list; `b.attempt(node_or_set)` / a `try` flag marks an operation whose failure is reported in the result and does
  not stop the batch. `MISSING` is an omitted optional argument. A batch is plain data (`b.to_dict()`, version 1).
- `run(app, batch, limit=120) -> Result` with `Result.values` (name -> value), `.marks`, `.ms`, `.skipped` (failed
  `try` operations), `.where` ("inproc" or "com"). Raises `BatchError(index, text, message)` when an operation
  fails: the `finally` operations have run and the application state is restored. `available(app) -> bool`.
- Values: None, bool, int, float, str; a 2-D grid (tuple of tuples) in and out; a list (for `Worksheets([...])`);
  references. Both executors return the SAME Python values for the same batch (2-D reads as tuples of tuples, 1x1
  reads as scalars, empty cells None, cell errors the integers pywin32 gives, Value2 dates as floats) and make the
  SAME call for None (pywin32 sends VT_NULL: `DBNull.Value` in C#). A COM object asked back is an error.

**The add-in (`Runner.cs`, a template: the build puts the hash into the function names):**
- `SmartassRun_<hash8>(jobPath)`, a macro-type `[ExcelFunction]` (hidden): reads the job, captures the four settings
  when `quiet`, executes the operations with late-bound `InvokeMember` on `ExcelDnaUtil.Application`, runs the
  `finally` operations and restores the settings in a C# `finally`, writes `<jobPath>.out` (UTF-8, no BOM) with
  ok / values / marks / ms / skipped or the failing operation's index, text and message. `SmartassPing_<hash8>()`
  returns the hash.
- Verbs: `break_links` (above). Add a verb only where an operation list cannot say it; each verb exists in both
  executors and has a test.
- Excel-DNA must not show its log window on a load problem (its diagnostics switched off in the `.xll.config` or
  `.dna`): prove it by loading a deliberately broken build in the hidden Excel.

**Build and load (`build.py`), like the Smart View bridge (`xl/instances.py` `bridge_path`, `ensure_bridge_file`):**
- hash8 = sha1 over `Runner.cs`, the vendored files and the batch version. Folder
  `config.INPROC_DIR / smartass_run_<hash8>` holding `smartass_run_<hash8>.xll` (the renamed `ExcelDna64.xll`), its
  `.dna` and the compiled dll; built in a temporary folder and renamed into place (two jobs at once); older folders
  pruned when not locked.
- 64-bit Excel only (else the COM executor). Loaded once per Excel process with `RegisterXLL` inside `jobs.com` with
  a limit, then the ping. A `Run` that says the macro is not there registers once more.
- Failures: no compiler, a compile error (kept in a marker file per hash so it is not retried every job),
  RegisterXLL False or timing out, a ping mismatch -> the COM executor for that Excel process, one warning.
  A batch falls back to the COM executor only when the in-process executor failed **before any operation ran**;
  a failure mid-batch is a `BatchError`, never replayed.

- [ ] **Step 1: Failing tests for the recorder** (`test_inproc_batch.py`): the operations a chain records, sets,
  `ret`, `MISSING`, grids, lists, `always`, `attempt`, the version, a node used as an argument, a COM-object `ret`
  refused at build time.
- [ ] **Step 2: Implement `batch.py`.** Run the tests: green.
- [ ] **Step 3: Failing tests for the COM executor** on the fake Excel (`test_inproc_runner.py`): every operation
  kind; values back; `quiet` captured and restored, also when operation k fails; `finally` run on failure; a `try`
  operation skipped and reported; `BatchError` names the operation; one `jobs.com` per batch.
- [ ] **Step 4: Implement the COM executor.** Green.
- [ ] **Step 5: Failing tests for the in-process executor's Python side** with a fake `app.Run` that executes the
  job file through the COM executor and writes the `.out` file: JSON round trip of every value kind (lists back to
  tuples), error mapping, fallback before the first operation, no replay after a mid-batch failure, `SV_INPROC=0`.
- [ ] **Step 6: Implement `runner.py` and `build.py`** (tests with a fake compiler and a fake `RegisterXLL`: the
  hash and names, the atomic build, pruning, each failure -> fallback and one warning, the marker file).
- [ ] **Step 7: Write `Runner.cs` and `scripts/check_inproc.py`** (hidden Excel of its own, nothing saved): builds
  and loads the add-in; runs a battery of batches through BOTH executors on fresh workbooks and compares values and
  the resulting sheets (cell values, formulas, formats, tables, names): each operation kind, each value kind in and
  out (None, bool, int, float, text with quotes and non-ASCII, 1x1 and NxM grids, a cell error, an empty cell, a
  30 000-character string), `Worksheets([...]).Copy` between workbooks with `MISSING` and None arguments,
  `break_links` leaving another link alone, an operation failing mid-batch (settings restored, `finally` ran, the
  error names it), a `try` operation, the ping, two builds loaded side by side, a broken build (no window, fallback).
  Prints OK / DIFF per check and the milliseconds per executor; exit 0 only when all are OK.
- [ ] **Step 8: Run** the suite and ruff: green. Run `scripts/check_inproc.py`: exit 0. Commit
  `feat(xl): batches of Excel operations with an in-process executor`.

### Task 2: The file job as batches

**Files:**
- Modify: `engine/smartass/pipeline.py` (`_file_draw`, `excel_job`, `redraw_job`, `_save`, `autosave_paused`, the
  read-back, status, heights, breaks, `_take_places`, `_arrange`, `_check`), `xl/movein.py` (`read_book`, `base`,
  `quiet`, `move_in`, `finish`, `calculate`, `show`, `park`, `restore`, `drop`, `close_left`), `xl/sources.py`,
  `xl/sheets.py` (`save_spec`), `xl/sheetorder.py`, tests beside each, `scripts/check_followups.py`,
  `scripts/bench_answers.py` (a "runs" column: batches per answer)

**Interfaces:** consumes Task 1's `Batch` / `run`. The functions keep their names and meaning; each now records
into a batch it is given (or builds and runs its own where it stands alone, as rollback and restore do).

**Shape of a job (controller ruling; the implementer may merge or split a batch where a read forces it, and says so):**
1. **look**: the workbook's sheets (`movein.Book`), the base (Normal font, standard sizes), whether the workbook is
   on screen and active, AutoSave's state, the Sources log's read, the stored spec's row - every read Python needs
   before it builds the file. No per-sheet reads outside this batch.
2. Python builds the file (unchanged).
3. **move in**: (a follow-up: park the old sheets and tables) open the file, copy, close (in `finally`), break the
   links, the moves that put the sheets in their places, activate the Answer. **Boundary + 40 ms: the formatted,
   still empty answer is on screen.**
4. **finish**: tables, formula blocks, calculate, and the reads the next step needs (the read-back block, the chart
   labels' positions, the status cells). **Boundary + 40 ms: the numbers and charts are on screen.**
5. **close out**: status, Key point row heights, page breaks, label fixes, (a follow-up: drop the old sheets),
   the Sources rows, the spec, the sheet order, the Answer active, the layout check's reads; then the save of our own
   file with AutoSave back on (in this batch or one after it, whichever keeps today's error behaviour).
- `movein.show` is replaced by the boundaries; the pause is `movein.PAINT_PAUSE = 0.04`, skipped when the workbook is
  not the active workbook of a visible Excel.
- A new Sources log's furniture is part of the batch that writes its first rows (the first answer of a new workbook
  spends 1.6 s on it today).
- Failure behaviour stays exactly as the tests pin it today: names in `t.sheets_added` before the copy; a
  `BatchError` in move-in or finish -> rollback, then restore of the parked sheets; after the drop -> `_kept`; the
  temporary workbook closed in the batch's `finally` and tried again at the end if that failed.

- [ ] **Step 1:** For each batch: the failing tests first (on the fake Excel through the COM executor: the number
  of batches per job - at most 6 for a follow-up, 6 for a new answer - the unchanged content, each failure path),
  then the change, then the suite. One commit per batch or coherent group.
- [ ] **Step 2:** The whole suite and ruff: green.
- [ ] **Step 3: Live in the owner's Excel:** `scripts/check_followups.py` with `SV_INPROC=1` and with `SV_INPROC=0`:
  every step "checks OK"; per step the Excel-job seconds, the batches and the stage seconds, next to today's
  (second answer 2.06 s, follow-ups 2.30-2.47 s, the sixth 4.22 s, fin_refresh 4.97 s, new workbook 6.81 s).
- [ ] **Step 4: The gate again:** `scripts/bench_answers.py --rounds 3`, all nine cases, `SV_INPROC=1`: the median
  Excel job per round without a lines sheet at most 3 s and with one at most 5 s, in every round. Write
  `docs/audit/2026-11-20-inproc-bench.md` with the table beside docs/audit/2026-11-19-file-render-bench.md.
  A miss is reported with the stage table, not hidden.
- [ ] **Step 5:** Docs (`docs/DESIGN.md`, `skills/smartass/reference/operations.md`: the add-in, the fallback,
  `SV_INPROC`, troubleshooting), CHANGELOG, version. Commit.

### Task 3: Formulas that point at an answer survive a follow-up

Today a follow-up or `fin_refresh` parks the question's sheets and tables under "old ..." names, copies the new ones
in and deletes the old ones. A formula on the user's own sheet (or on another answer) that points at
`'Q3 Answer'!E12` or `Data_Q3[Value]` follows the rename to the parked name and shows #REF! after the delete (final
review I2). A monthly pack with a summary sheet breaks every month.

**Files:** `engine/smartass/xl/movein.py` (a step between the finish and the drop), `pipeline.py`, tests,
`scripts/check_followups.py` (a summary sheet with a cell reference, a table reference, a cross-sheet SUM and a chart
series pointing at a Q sheet; each must show the new answer's value after every follow-up and after `fin_refresh`),
docs (the known limit removed from operations.md, DESIGN.md, CHANGELOG).

- Before the drop, every formula outside the question's own sheets that mentions a parked sheet or table name is
  re-pointed to the new name. Probe first, in a hidden Excel, which way is exact and cheap in-process: reading each
  other sheet's formulas once (`UsedRange.Formula2`), replacing in Python and writing back only the changed cells
  (as a verb, so the scan runs inside Excel); or `Range.Replace` on formula text (it changes the user's Find
  settings, which must then be put back). Also: defined names and chart series that refer to the parked sheets,
  conditional formats and data validation lists (report what Excel offers; cover what is cheap, say what is not).
- Only sheets with a formula that mentions a parked name are written to; a sheet without one is not touched (the
  user's workbook can be large: one `Find` per sheet for the parked prefix decides).
- A cell the new answer no longer has (the layout changed) keeps Excel's #REF!: that is correct.
- The re-pointing is part of the job's rollback story: a failure before the drop restores the old names, and the
  references with them.

- [ ] Step 1: the probe and its findings in the report. Step 2: failing tests on the fake Excel. Step 3: the
  change. Step 4: suite, ruff. Step 5: live `check_followups.py` in the owner's Excel with the summary sheet.

### Task 4: Carry-overs from the final review of phase 3

Each with a test: (M1) the style findings the old renderer reported in `reply.warnings` ("text is cut off", "text
does not fit", a wrong font) computed from the look model and reported again; (M2) the notes of a failed job (a
temporary workbook left open, AutoSave left off) reach the error text; (M3) settings found off at the start of a job
that the previous job left off (a killed server, a restore that timed out) are not taken for the user's own: the
job records what it changed in a small state file and the next job puts it right; (M4) a spec store that fails
after the drop says so; (M5) a new answer whose save fails removes its spec row and Sources rows with its sheets;
(M9) a chart title that XlsxWriter reads as a cell reference ("Revenue beat budget!FY26") is written as text.

### Task 5 (conditional, decided by the controller on Task 2's stage table)

Only what the measurement justifies, each as its own small task: the save handed to Excel after the reply
(`ExcelAsyncUtil.QueueAsMacro`, its result reported by the next job); drawing the Answer and Data sheets by
operations with no temporary file (removes Open, Copy and Close; needs a second writer for the look and the charts,
so only with a clear gain); the 2.9 s a new workbook's creation takes.
