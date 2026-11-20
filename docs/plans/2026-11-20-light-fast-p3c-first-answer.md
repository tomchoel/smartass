# Light and fast, phase 3c: the first answer, the save, the carry-overs

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan
> task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** the first answer of a new workbook - what every new question starts with - at or under **5 s** in the
owner's Excel (10.6-11.3 s in 0.4.0, measured 20.11.2026 17:28), and the save out of the user's wait.

**Where the 11.3 s go (run log, 0.4.0, the owner's Excel with no other workbook open):** open target 5.2 s
(`Workbooks.Add` of the TN template 1.2 s, then name + label + `SaveAs` into the OneDrive folder 3.4 s); the
close-out 2.9 s (a new Sources log: print setup 1.4-1.5 s, furniture 0.75 s, restoring the settings 0.58 s); the
insert 0.9 s; the look 0.64 s (0.03 s in an existing workbook); the save 0.66 s; finish 0.26 s; paint 0.16 s.

**Controller spike (scratchpad `spike/seed_open.py`, copied into the ledger folder):** in the owner's Excel a COPY
of the labelled seed file opened with `Workbooks.Open` took 1.55-1.67 s (copy 3 ms) against 2.4-4.3 s for
add + name + label + SaveAs; the opened copy had AutoSave on, was Saved and carried a sensitivity label. Today's seed
has 47 cell styles (the TN template gives 101) and one empty sheet named Sources.

## Global constraints

- The user's Excel is left as found on every path; the user's own workbook is never saved; nothing writes to EPM;
  no figure from EPM is printed by a script.
- A failed job leaves no half-made workbook open and no file behind that the user did not get told about.
- Every new analysis file carries the sensitivity label (without it Excel's save fails silently on this tenant) and
  opens with AutoSave as today.
- Same answers: `sheet_equals_model` OK, `scripts/check_followups.py` "checks OK", `scripts/check_insert.py` and
  `scripts/check_inproc.py` exit 0. Tests never start Excel.
- One code path for the Excel work (batches through `xl/inproc`); the COM executor stays the fallback.
- Commits end with the trailer your own session's instructions give.

---

### Task 1: The first answer of a new workbook

**Files:** `engine/smartass/xl/target.py` (`_new_target`, `_save_new`, `_new_analysis_file`, the seed),
`engine/smartass/pipeline.py` (`excel_job`, `_file_draw` for a new workbook), `engine/smartass/render/xlsx.py`
(a workbook-level build), `engine/smartass/xl/sources.py` (the log's furniture as data both writers share),
`engine/smartass/xl/movein.py`, tests beside each, `scripts/check_insert.py` (a "new workbook" part),
`scripts/bench_answers.py` / `scripts/check_followups.py` (the stage columns), docs, CHANGELOG.

**Two routes; probe both in a hidden Excel of your own first, then build the better one that holds.**

- **Route A - the built file IS the new workbook.** For a question that starts a new workbook, the renderer writes
  the whole workbook to its final path in the analysis folder: the answer's sheets as today, plus the furnished
  Sources sheet (header, fonts, column formats and widths, filter, freeze, print setup - what `sources._furnish` and
  the print-setup step make today), in the sheet order `sheetorder` wants, with the sensitivity label (the label is
  custom document properties, `MSIP_Label_<id>_...`; take them from the labelled seed - see `target._seed_has_label` -
  and write them with XlsxWriter's custom properties). Excel then only opens it (`Workbooks.Open`) and the job goes
  on with the finish (tables, formula blocks, labels, calculate, reads), the close-out (status, heights, the
  Sources ROWS, the spec store) and the save. No `Workbooks.Add`, no `SaveAs`, no insert, no Sources set-up in Excel.
  Probe: Excel opens it without a repair prompt or a Protected View bar; AutoSave is on and `Workbook.Saved` True
  after the open; the label is recognised (`wb.SensitivityLabel.GetLabel()`) and a later `Save` works; the TN cell
  styles (the workbook has only XlsxWriter's formats: check what `xl/styling.apply_report` / `xl_tidy` and
  `styling.ensure_template` do when the styles are missing, and whether they install them on demand as the comment in
  `_new_target` says); the hidden `_svmeta` sheet (made in the file or by the close-out as today); what the first
  answer's Sources links, the Data header link and the print setup look like against today's (extend the parity
  comparison of `scripts/check_insert.py`: a new workbook made today's way against one made this way, item by item).
- **Route B - a furnished seed.** The seed file carries the TN styles and the furnished Sources sheet (print setup
  included); a new workbook is a copy of it on disk opened with `Workbooks.Open`; the insert and the rest stay as
  they are. The seed is rebuilt when its recipe changes (a version stamp) and when it lacks the label.
- Route A is the goal if its probes hold (expected: a first answer at 3.5-4.5 s); Route B (expected 6.5-7.5 s) if A
  fails on something that cannot be put right cheaply - say exactly what. Either way: a new Sources log costs no
  set-up calls in Excel; the look of a new workbook costs what it costs in an existing one (find what the 0.6 s is).
- Keep: the choice of a free Excel for parallel agents (`_new_target`'s turn logic), the file name rules
  (`_new_path`), the error when the folder is not writable, the clean-up when anything fails after the file exists
  (close it without saving and delete the file the job itself created - never a file that existed before).
- The "no Excel at all" path (`_new_analysis_file`: the hidden Excel makes the file, then it is launched) stays
  working; say whether Route A simplifies it.

- [ ] Step 1: the probes, with their findings in the report. Step 2: failing tests on the fake Excel (a new
  workbook: the number of batches and COM calls, no SaveAs, no Add, no furniture calls; the label properties in the
  built file; the failure clean-up). Step 3: build it. Step 4: suite, ruff, `check_inproc.py`, `check_insert.py`
  (hidden Excel). Step 5 (controller, owner's Excel): `check_followups.py` and the first-answer seconds next to
  10.6-11.3 s.

### Task 2: The save happens after the reply

The save of our own file is 0.4-1.1 s of every answer. With the add-in loaded, the last batch asks Excel to save
right after the Run returns (`ExcelAsyncUtil.QueueAsMacro`): the user has the answer and the reply while Excel saves.
- Only our own analysis file; the user's workbook is never saved. AutoSave is switched back on by the queued step
  after the save, in today's order.
- The queued step writes its outcome (ok or the error text, the seconds) to a small file under the smartass home; the
  next job's look reads it and, on a failure, adds a note to its reply and saves in the foreground. `fin_status`
  shows the last outcome.
- A job that arrives while the save runs waits for Excel as it does for any busy Excel (prove it: a follow-up sent
  at once after a reply).
- Excel must have painted the answer before the save starts (the save blocks painting): queue it so that it runs
  after Excel has been idle once, and prove with the window-capture script (`<ledger>/spike/watch_paint.py`).
- The COM executor (no add-in) saves in the foreground as today. `SV_SAVE=sync` forces that with the add-in too.
- The bench reports both: the Excel job as the user waits for it, and the save's own seconds.

### Task 3: Carry-overs

Each with a test. From the final review of 0.3.0: (M1) the style findings the old renderer reported in
`reply.warnings` ("text is cut off", "text does not fit", a wrong font) computed from the look model and reported
again; (M3) settings found off at the start of a job that the previous job left off (a killed server, a restore that
timed out) are not taken for the user's own: the job records what it changed in a small state file and the next job
puts it right; (M4) a spec store that fails after the drop says so; (M5) a new answer whose save fails removes its
spec row and Sources rows with its sheets; (M9) a chart title that XlsxWriter reads as a cell reference ("Revenue
beat budget!FY26") is written as text. From 0.4.0's reviews (the ledger of phase 3b lists them): an unknown outcome
that recurs gives up on the add-in for that Excel and falls back; a used range reaching both the last row and the
last column must not cost hundreds of thousands of calls in the formula scan; the names read over COM; a protected
sheet warned about only when it could matter; the Excel job moved out of `pipeline.py` into its own module.
