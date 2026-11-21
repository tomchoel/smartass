# Changelog

## Unreleased

The sections below are 0.5.0 (phase 3c). After the update, reconnect every open Claude Code session (`/mcp`):
a session still on 0.4.0 in the same Excel has no first-contact check, so a 0.5.0 save after the reply can run in
the middle of its job.

### Carry-overs from the reviews (0.5.0)

- The report style's text findings are back in the reply's warnings ("B5: text is cut off or spills over (~65 > 46
  chars)", "B13: text does not fit (...)"), in the words of the renderer before 0.3.0, computed from the look model
  over the block the job reads back anyway: no Excel call. The old lint's font, vertical-border and "content outside
  the table width" findings read what Excel shows anywhere on the sheet, which the look cannot know (it writes Arial
  with top and bottom edges only, inside its block).
- Excel settings a job that did not finish left changed (a killed server, a put-back Excel refused) are put back by
  the next job in that Excel and said in its warnings ("an earlier smartass job in this Excel did not finish and left
  Calculation manual: put back now"): before its first quiet batch a job records in the state file
  `%LOCALAPPDATA%\smartass\jobstate\<Excel's pid>.json` what it found and what it sets (ScreenUpdating,
  EnableEvents, DisplayAlerts, Calculation, AutoSave of our own file; no figures); its end removes it. Only a setting
  that still reads what that job set is put back, and never in another Excel process that has the same pid. What
  Excel refuses to put back stays in the file, and only that, also through the next job, until a put-back goes
  through; when only AutoSave could not be switched back on, only that is kept. The file of an Excel process that is
  known to be gone is removed.
- AutoSave of our file that could not be switched back on is said: "AutoSave could not be switched back on in <file>
  (...)", with "the next answer in this Excel switches it on" only when the state file holds that workbook, else
  "switch it on yourself"; after a save after the reply, the next answer says so and switches it on.
- A follow-up whose close-out fails after the old sheets are dropped says when its spec was not stored (the next
  follow-up would patch the spec before it) - repeat the follow-up. A new answer whose save (or close-out) fails
  removes its spec row and Sources rows with its sheets: no ghost question for a follow-up to pick.
- A chart title or series name shaped like a cell reference ("Revenue beat budget!FY26") is written as text, not as
  a broken reference.
- Two unknown batch outcomes in a row in one Excel make that Excel use the COM executor for the rest of the server's
  life, with one warning, instead of failing every later job.
- The look for formulas that point at a follow-up's sheets halves each used range by cells: a sheet whose used range
  a stray format stretched to XFD1048576 costs a few dozen steps instead of ~175 000 (in a hidden Excel 48 ms
  in-process, 288 ms over COM); a sheet that would take more than 2 000 steps is said as unread. A sheet whose
  formulas cannot be read (a protected sheet with a hidden formula) is warned about only when a cell there turned
  #REF! with the follow-up, and the warning names those cells ("... and 1 cell(s) there show #REF! since the old Q3
  was dropped (B2): a formula there pointed at it"; #REF! told by Excel's error code, not its text, with the verb
  `errors`, which runs only for such a sheet). A sheet too large to check that way as well (more than 2 000 blocks:
  about 100-200 million cells) is warned about on every follow-up.
- Not changed: the defined names are still read one by one over COM when the add-in is not there (3-5 s for 2 000
  names); Excel offers no exact lookup of the names that mention a sheet.

### After Excel froze in the live check (0.5.0)

- No COM object outlives the thread of the job that made it: every Excel job - one that answers, one that fails, one
  the server gave up on - has the Excel objects it made released on its own thread, after its answer went out (the
  frames its error carries, and every reference cycle: the garbage collector runs there), and only then are its turns
  in that Excel released. Released from another thread such an object calls that Excel from there and waits for it -
  for ever, when that Excel has stopped responding (live, 21.11.2026: Excel stopped answering between one job's thread
  ending and the next one's starting, and that next thread hung releasing an object). A release that has not returned
  10 s after the answer gives the request up as below. A job's result is checked to be plain data.
- A request the server gave up on because Excel did not answer in time no longer keeps the next request in that Excel
  waiting silently for its turn (live, 21.11.2026: the next one waited until it was given up on too): after 5 s a
  request that waits for that Excel's turn is refused with "An earlier request in that Excel (...) was given up after
  Excel did not respond in time, and it still holds that Excel: ... was not started", and what to check - also in
  another smartass server process (`user_excel_<pid>.given_up.json` next to the turn's lock says which request holds
  it). The request given up on never writes again and frees the turn when Excel answers it; the turn is never freed
  before that. A new question's request that looks for a free Excel (target 'new') still counts that Excel as busy:
  it takes another free Excel, or starts one for a parallel agent, as before.
- Excel's "call rejected" while it is busy (a save, AutoSave) no longer fails a job with "Unexpected AttributeError:
  <unknown>.FullName" (or "Property '<unknown>.Name' can not be set" in a batch over COM): the object wrapper waits
  for Excel's type information instead of going on without it, and a rejected lookup of a name is said as busy, so
  the usual wait for a busy Excel asks again. A call is never run twice for it.
- What froze Excel is not known; it was not reproduced under the conditions below, which do not rule anything out.
  Through the server's own job path (`scripts/check_followups.py`'s steps: a new answer, a second one, six follow-ups,
  fin_refresh; each Excel job on its own thread; the engine's own target resolution pointed at a hidden Excel of our
  own; the tests' fake data, no EPM): 20 runs of two add-in processes then one COM process in one Excel with an old
  add-in build loaded on the code before the first point above, 20 runs of the same on the code after it, and 7 runs
  of three COM processes in an Excel with no smartass add-in loaded on each code. Earlier (Task D): 19 full and 2
  partial sequences of a stand-in, saved. Not tried: a visible Excel, the live EPM and Smart View session, the owner's
  workbooks, and think-cell disconnected (it is installed for all users; Excel refuses to disconnect it without an
  administrator, and its per-user setting was not changed).
- `scripts/check_insert.py` part "mixed" runs the executors mixed in an Excel of its own with an old add-in build
  loaded, each phase from a new server's state, and the add-in phase's last save after the reply still queued when the
  COM jobs start; a job that does not answer is a DIFF and that Excel is killed by its pid. The two "manual page
  breaks" DIFFs of the default parts were a read that fails in a hidden Excel (the sheet's list of page breaks answers
  "invalid index"), on every route alike: the check now reads each row and column instead when that happens.

### The save happens after the reply (0.5.0)

- With the add-in, the answer's own file is no longer saved inside the job (0.4-1.1 s of every answer in the owner's
  Excel): the job's save step asks the add-in, in one `Application.Run` of a few ms, to save it, and the reply goes
  out (not for the first answer of a workbook the question created: that one is saved in the job, so the file on
  disk holds the answer before the session points at it). The add-in waits `SV_SAVE_DELAY` (0.3 s) on a timer of its
  own so Excel paints the answer first, then saves
  as a queued macro (`ExcelAsyncUtil.QueueAsMacro`) and switches AutoSave back on (also after a failed save), and
  writes the outcome - no figures: the workbook's name, ok or Excel's error, the seconds - to
  `inproc\saved\<Excel's pid>.json`. Live in a visible Excel of our own (20.11.2026): with the delay the answer's
  last picture change came 0.2-0.35 s before the save started and nothing changed during or after it; with no
  delay (the control) the save started at once and the answer's last changes were painted only after it.
- A follow-up that comes before that save ran has it run first, at the start of its first batch (never between its
  batches, never twice); one that comes while it runs waits as for a busy Excel. Excel answers calls in the middle
  of a save of a OneDrive file, so the add-in answers "busy" to a batch while its save runs and the batch is asked
  again (live: one in four batches sent 30 ms into a save got that answer; none ran during a save).
- The next answer in that Excel reads the outcome once: its seconds go to its run-log line as `previous save`; a
  failed save is a warning in that reply ("the save of <file> after the last answer failed (...); it is saved now")
  and that answer saves in the job, also that workbook when it is another one; a workbook closed before its save is
  said ("Excel asks whether to save it when it closes" - observed: Excel's own "Save your changes to this file?"
  dialog, and the waiting save does not run while it is shown). `fin_status` shows the last outcome
  (`save_after_reply`).
- Over COM (no add-in) and with `SV_SAVE=sync` the job saves before it replies, as before; a follow-up that fails
  once its old sheets are gone still saves our file in the job. The user's own workbook is never saved: the add-in
  keeps the workbook's full name from the request and does not save another file opened under that name, and a
  failed save is repeated only while that file is open. A failed save that the next answer could not repeat (it
  failed first) is left for the answer after it.
- A job's first contact with an Excel runs the waiting save of every smartass build loaded there (two servers of
  different versions each load their own add-in, with a save of its own): `Application.RegisteredFunctions` names
  the builds' `.xll` files (1-5 ms). Before, another build's save ran on its own timer in the middle of the job,
  and its outcome could be overwritten by the job's own before anyone read it (a failed save then went untold). A
  first contact that cannot be made is logged once per Excel process.
- The reads made before a job's turn (`target='active'`'s active workbook, the open workbooks that every `xl_*`
  tool scans) wait for a busy Excel, as Excel rejects calls while our save runs: a question or a tool sent within
  about 1.4 s of a reply no longer fails with "Call was rejected by callee". A job's start waits for an Excel that
  rejects calls (a cell in edit mode, a dialog) 30 s in all, then says so, as before the save after the reply.
  Only the seconds Excel rejects calls count: a save of ours is waited for through - the add-in's answer that one
  runs, a waiting save the job's first contact runs itself, and Excel's rejections while a save this server asked
  for may still run. Another smartass server's save is said as that server's ("(asked by another smartass server in
  this Excel)"). A server with `SV_INPROC=0` makes no first contact, so another server's waiting save can run between
  its batches. An abandoned job stops in every one of these waits.
- `bench_answers.py` adds "saved after s" (the save's own seconds) beside the Excel job the user waits for;
  `check_followups.py` waits for each step's saves before it reads `Workbook.Saved` or the file and prints their
  seconds; `check_insert.py --parts newbook` waits for them too (`--parts newbook,failures`: 82 checks);
  `check_inproc.py` checks the save after the reply (61 checks in all).

### The first answer of a new workbook is its built file (0.5.0)

- A question that starts a new workbook no longer has Excel create one from the TN template and save it as a new
  file in the OneDrive folder, then set up a Sources log (10.6-11.3 s in the owner's Excel, 0.4.0): the renderer
  writes the whole workbook to a reserved name in the analysis folder - the answer's sheets, the Sources log
  furnished (header, fonts, widths, filter, frozen header, print setup and its two hidden names), the sensitivity
  label as custom document properties and the TN template's theme - and Excel only opens it. The job then runs its
  usual batches: look, bring in (the finish only), close out (the log's rows), save. In a hidden Excel on the
  OneDrive folder the whole first answer's Excel work took 1.92-2.46 s against 7.05-13.72 s made today's way (four
  rounds each, 20.11.2026; the owner's Excel is measured separately).
- The file name is reserved by creating it empty and exclusively (parallel agents get different names) and is never
  the name of a workbook open in that Excel. A failure once it is reserved closes the workbook without saving and
  deletes the file (found by its full path: never another session's workbook of the same name), and the error says
  so; the workbook becomes the session's only once the answer is in it and saved - in the job, also with the add-in.
- The new workbook is built in the Normal style and sizes of a workbook added from the template, which the look of
  the first one made that way stores; until then (and without a label id, without the template, in an Excel started
  for a parallel agent, or with no Excel open) a new workbook is made as before.
- The labelled seed of the no-Excel path is now found: Excel on this tenant keeps the label in
  `docMetadata/LabelInfo.xml`, not in custom properties, so the seed was never copied before.
- A built workbook has no TN cell styles until `xl_tidy` or `apply_report` installs them (as they do in any
  workbook without them).
- `scripts/check_insert.py` has a `newbook` part (a new workbook made today's way against one built as its file,
  item by item, and a failed one); `bench_answers.py` and `check_followups.py` show the `open target` and
  `draw.open` stages.
- Live in the owner's Excel (a fresh Excel, 20.11.2026, the save still in the job): a new workbook's first answer
  5.3-6.1 s against 7.5 s with 0.4.0 in the same Excel: open target 2.1-2.6 s (of it 2.0 s not Excel's),
  `draw.open` 1.05-1.08 s, look 0.83-1.14 s, the rest 0.9 s. The workbook opened showing the formatted empty Answer
  with "Internal" in its title, no repair prompt and no Protected View; the whole answer was on screen about 1.4 s
  later; on disk it carries the label as one made the old way does. Second answers 1.2-1.6 s, follow-ups 1.5-2.1 s,
  `fin_refresh` 3.4-4.1 s. 0.4.0's bench in a fresh Excel held the 3 s / 5 s gate in every round (1.3 / 1.9 / 2.0
  s and 3.6 / 3.1 / 3.6 s). Details: docs/audit/2026-11-20-first-answer.md.
- Two of those costs are cut. The session's parent processes are read from one snapshot of the process table:
  psutil took one per ancestor, 1.5 s on the first answer of every server start (a plain script: 1.75-2.15 s before,
  0.20-0.22 s after). And a server process that meets an Excel holding its build of the add-in already (loaded by an
  earlier server) uses it as it is instead of registering it again: the look of a new server's first answer took
  0.45-0.49 s in a hidden Excel, now 0.04-0.09 s (the look's reads themselves take a few ms; that this was most of
  the live look's 0.8-1.1 s is inferred). The live figure after these cuts is still to be measured.

### The answer's Excel work runs inside Excel (0.4.0)

- An answer's Excel work is recorded as batches of operations, and each batch runs in one go inside Excel through a
  small Excel-DNA add-in the engine builds on first use (C# compiled by the .NET Framework compiler that ships with
  Windows) and loads silently with `RegisterXLL` once per Excel process: one `Application.Run` per batch instead
  of dozens of cross-process calls. Where it cannot load (no compiler, a 32-bit Excel, a refused or blocked
  `RegisterXLL`) or `SV_INPROC=0` is set, the same batches run over COM, with one warning in the server log.
- The job is four batches, 8 COM calls per answer (34-37 in 0.3.0): **look** (every read the build needs: the
  sheets, the base, which workbook is active, the Sources log's kind, the stored specs' keys; then AutoSave off while
  the job writes); **bring in** (the answer's table names looked up among the defined names, one lookup per name;
  a follow-up's old sheets parked; the built file's sheets inserted with one `Sheets.Add(Type:=<the file's path>)`,
  so no second workbook is opened, nothing is copied and no link is left; then the moves, the tables, the formulas,
  the chart labels, calculate, the read-back, and a follow-up's scan for formulas pointing at its old sheets);
  **close out** (status, Key point heights and page breaks, a follow-up's formulas re-pointed and its old sheets
  dropped, the Sources rows, the spec, the sheet order, the Answer active); **save** (our own file saved and AutoSave
  back on; the user's workbook is never saved, so it gets no save batch unless a new Sources log needs its print
  setup's name). Each batch that quiets Excel puts the user's settings back itself, also when a step fails;
  Calculation is manual only where formulas are written (the bring-in, a close-out that re-points, the point-back).
  After the bring-in the job pauses once (0.15 s, `SV_PAINT_PAUSE`; only when the workbook is the active one of a
  visible Excel) so Excel paints the answer with its numbers and charts before the close-out.
- Besides plain gets, calls and sets, a batch has five verbs, each one operation inside Excel: `each` (properties
  read over a collection; a method's name is never called), `exist` (names looked up), `nudge` (the thin waterfall
  labels' lift), `mentions` (the formulas that name a text), `swap` (a text property set only when it still reads
  what was read before).
- A failure is undone as before (rollback, the old answer's names put back, the new answer kept once the old one is
  gone), also for a job abandoned by a timeout. An error names the Excel operation that failed and Excel's own
  message, and says when AutoSave was left off in the user's workbook.
- Formulas elsewhere in the workbook that point at a question's sheets or tables (`='Q3 Answer'!E12`,
  `=SUM(Data_Q3[Value])`, cross-sheet SUMs, array formulas, chart series, defined names, also on another answer's
  sheet) survive a follow-up or `fin_refresh`: they are re-pointed at the new sheets right before the old ones are
  dropped (in 0.3.0: #REF!). Exactly those references: never text in a string literal, a sheet whose name merely
  contains the answer's, another workbook's. The formulas are read directly (only the parts of a sheet that hold a
  formula); no Range.Find is used, so Excel's Find dialog keeps its options. Each write is a compare-and-set: a
  formula the user changed while the answer was drawn is left alone and reported. A formula Excel refuses to rewrite
  (a protected sheet) and a sheet whose formulas cannot be read (a hidden formula on a protected sheet) are warnings;
  a failed follow-up points them back with the old answer. When the workbook cannot be scanned at all, the old sheets
  are dropped anyway and the warning says the formulas pointing at them were not checked. Not covered: conditional
  formats, data validation, pivot tables on a `Data_Q<n>` table, linked chart titles, sparklines, form controls'
  links, linked shapes, other workbooks. `fin_refresh` now gives each question's warnings and layout findings.
- Measured (docs/audit/2026-11-20-inproc-bench.md; the Excel job in-process unless said):
  - Quiet Excel, the four-batch job: a follow-up 1.66-2.0 s (`scripts/check_followups.py`, owner's Excel; measured
    before the re-pointing came, whose scan adds the cost given last).
  - The gate run in a quiet Excel (`scripts/bench_answers.py --rounds 3`, 06:00), measured on the five-batch code
    before the batches were merged: 1.9 / 2.4 / 2.7 s per round without a lines sheet, 4.1 / 3.9 / 4.3 s with one
    (0.3.0: 2.0 / 2.7 / 3.3 s and 4.2 / 4.1 / 5.2 s). The 3 s / 5 s gate holds in every round.
  - With another session working in the same Excel (20.11.2026 11:15): follow-ups 2.5-2.9 s in-process against
    4.9-6.2 s over COM; the bench then 4.7 / 2.7 / 5.9 s without a lines sheet, not a gate run.
  - The final code (four batches, the re-pointing), 20.11.2026 17:20-17:30, the owner's Excel with no other workbook
    open: the bench 2.3 / 3.2 / 3.6 s without a lines sheet and 5.2 / 6.0 / 5.9 s with one - over the gate in rounds 2
    and 3 (the move-in as in the morning run, 0.9 / 1.1 / 1.1 s; the save 0.5 / 0.8 / 1.1 s against 0.3 / 0.5 / 0.6 s).
    Back to back with 0.3.0 in the same Excel, same minutes (`check_followups.py`, each twice): a second answer
    2.2-2.3 s against 3.6-3.9 s, follow-ups 3.0-3.8 s against 4.1-5.4 s, `fin_refresh` of two answers 5.7-6.4 s
    against 8.5-9.2 s, a new workbook's first answer 10.6-11.3 s against 12.7 s. 0.3.0 itself took 2.2-2.8 s per
    follow-up the night before, so that afternoon everything ran slower; the gate is shown only by the morning run.
  - The scan for pointing formulas when nothing points at the answer: 0.06-1.0 s in-process, 0.3-4.5 s over COM,
    by workbook size (10 sheets and 5 000 formulas to 60 sheets and 260 000 formulas).
- Thin waterfall labels now sit 10 pt higher than 0.3.0 drew them in a hidden Excel: 0.3.0 set the lift before the
  calculation, which put the label back where Excel places it; the lift is now made right after the calculation,
  inside the same batch (the other point labels are positioned before it, as before). Whether 0.3.0 lifted them in a
  visible Excel was not measured: the script's own Excel would not become visible.
- Live checks, each in a hidden Excel of its own: `scripts/check_insert.py` compares the one-call insert with the
  open, copy and close it replaced, runs whole jobs through both executors and compares the workbooks item by item,
  makes a follow-up and a new answer fail inside their batches through both, re-points formulas and points them
  back, and measures the chart labels (90 of 91: the one difference is the script's own read of the manual page
  breaks, the same on every route and on the code before the re-pointing; after the last fixes, the parts they
  touch 47 of 47).
  `scripts/check_inproc.py` proves the five verbs, the batches' settings, the add-in's guards and the job's reads
  the same in both executors (45 of 45). `scripts/check_followups.py` also watches the file's writes and reads
  the saved file.
- Removed with the copy: the "temporary workbook could not be closed" warning and its retry at the end of a job,
  and the in-process executor's `break_links` verb.

### Answers built as a file (0.3.0)

- An answer's sheets (`Q<n> Answer` with its charts, `Q<n> Data`, `Q<n> Lines <alias>`) are built as an .xlsx
  without Excel (`render/xlsx.py`, `render/xlsx_charts.py`), moved into the workbook in one copy and finished there
  (`xl/movein.py`), with the user's Excel quiet and left as it was found. The live comparison of all nine test-set
  cases passed against the old COM writers: every difference left is on its allowed list (cell style names, a
  0.2 pt chart height, Data widths within AutoFit's band; `docs/audit/2026-11-19-render-comparison.md`).
- Measured live (`scripts/bench_answers.py --rounds 3`, all nine cases, one workbook growing to 28 answers): the
  Excel job of an answer without a lines sheet went from 15.6 / 25.8 / 46.0 s per round
  (`docs/audit/2026-11-19-baseline-v2.md`) to 2.0 / 2.7 / 3.3 s; over all cases from 33.7 s to 2.8 s, the whole
  answer from 34.9 s to 4.1 s, and from 280 COM calls to 34. The gate of 3 s (5 s with a lines sheet) per round
  held in rounds 1 and 2 and was missed in round 3: 3.29 s, and 5.19 s with a lines sheet. What grows is the save of
  the file (Excel's own) and reads that walk every sheet of a 61-sheet workbook inside one COM call; the controller
  accepted the miss, and the re-gate moves to the next task (`docs/audit/2026-11-19-file-render-bench.md`). In
  short: 2-3 s for an answer without a lines sheet in a workbook of normal size (3.3 s at 61 sheets), 4-5 s with a
  lines sheet, about 6-7 s for the first answer of a new workbook (the workbook is created).
- Follow-ups and `fin_refresh` replace the question's sheets whole (the old ones are parked, the new ones moved in
  where they were, the old ones dropped). Cells typed on a `Q<n>` sheet go with it; a failure puts the old answer
  back, and once the old sheets are gone it keeps the new one and says so (also when the final save fails).
- Known limit (fixed in 0.4.0): a formula on the user's own sheet that points at a `Q<n>` sheet or its table
  (`='Q3 Answer'!E12`, `=SUM(Data_Q3[Value])`) shows #REF! after a follow-up or refresh of that question, because
  the old sheet is deleted.
- The answer lands on the new Answer whenever its workbook is the active one, also in an Excel that is not visible.
- The Sources log's print area is `A1:N1048575` (a finite last row), so openpyxl reads our files without a warning.
- Removed: the COM writers (`render/excel.py`'s Data/Answer writers, `render/charts.py`'s drawing,
  `drill_xl.write_lines`), `xl/owned.py` (the stored drawn range is no longer used), `config.RENDER` /
  `SV_RENDER`, and `scripts/compare_renders.py`. `SV_FASTCOM` stays.

### Bearer route: grids without the hidden browser (optional, per pod)

- Smart View HTTP requests go over plain HTTPS with the user's own OAuth access token where a pod has a scope and a
  sign-in (`svhttp.BearerTransport` in front of the browser transport). Anything it cannot authenticate falls back
  to the browser route; drill-through is unchanged (untested over the token).
- Live 19.11 with scopes for plan1 and stratmodel: FinMax, FinLite and FinPlan grids, the closed month and drill
  links over the token, the same as the browser route cell for cell; from a cold start no browser is started (first
  grid 0.7 s against about 20 s). Drill lines still load in the browser.
- `smartass login` (`--status`, `--force`, `--forget`): Oracle's device-code sign-in, one approval per pod; the
  sign-in is DPAPI-protected under `oauth\` in the home folder. The refresh token is single use: refreshes run
  under a cross-process lock, and the MCP server and `smartass serve` renew sign-ins daily (an unused one ends
  after 7 days).
- Off until `config.toml` has `[oauth]` `idcs`, `client_id` and `[oauth.scopes]`; doctor check `oauth`.

### Knowledge layer phase B (0.2.23)

- Behaviour cards derived from data (`smartass atlas cards`, after every full atlas build): operator per (member,
  parent), time balance, account type; manual overrides in `knowledge/semantic/cards/`. Consolidation uses the real
  operators (`~` children out of totals, lines, rankings and drivers). Smart View member-information requests do not
  work on this provider (probe 10.11) and are not allowed.
- New metrics, verified live: EBIT, D&A incl. write-downs, profit before tax, net financial items, net income
  (FinMax), capex, total assets, fixed subscriptions (total, internet, TV, telephony), internet/TV ARPU, FinPlan FTE,
  FinPlan EBITDA margin; D&A gets its FinPlan member. FinPlan Target carries D&A, EBIT, PBT and CAPEX since FY26.
- Facts on FinMax scenarios (only Actual has data), FinPlan net income (no tax), balance sheet, free cash flow (empty),
  write-downs FY26 and volume units; an access rule (numbers depend on the reader's EPM rights).

### Knowledge layer phase A (0.2.22, spec `docs/specs/2026-11-10-knowledge-layer-design.md`)

- One source: `knowledge/semantic` gains `rules/`, `entities/`, `kpis/`, `vocab/`, `notes/` and account-bundle
  metrics, folded from catalog.toml, capabilities.toml and the three 15.10 business maps (all deleted with seed.py).
- Caveats inherit over the atlas tree (`inherit` down / up / both) and say `via` the member they came from, also on
  the Answer sheet: C40S carries the C40510 trap, a division its reorgs.
- `fin_lookup` kinds `entity`, `kpi`, `vocab`, `rule`; member hits carry their meaning note; short and generic aliases
  follow the catalogue's matching rules (longest match wins).
- Scenarios and the default materiality come from the rules (raw-member default now 10 %, Tom's catalogue decision);
  FTE materiality 10 FTE / 5 %.
- The build stops on contradictions (entity placement, FinPlan division level 1xxx vs 2xxx, aliases across kinds,
  rules defined once, required rules approved). A stale atlas warns on every answer; an atlas refresh that breaks the
  knowledge keeps the old knowledge.db.

v2 stands alone: one MCP server `smartass` with nine tools; the agent writes an AnalysisSpec v2 and the engine
validates, computes and draws it (design `docs/DESIGN.md`, `docs/specs/2026-11-04-v2-overview.md`). v1 is retired.

### Fixed

- Critic score: `fin_critic(q, scores, notes)` takes the six criterion scores and computes the score itself (mean
  rounded half up). The agent no longer computes it: a pane run stored 4 for a mean of 4.83.

- Pane live-test fixes (v2): follow-ups (`fin_analyze(patch)`), `fin_refresh` and `fin_critic` default to the
  workbook of the latest answer in answers.db, not the old session file (explicit targets win; no stored answer
  keeps the session workbook); a target that is the name of an open workbook (e.g. an unsaved `Book1`, with or
  without extension) resolves to it in any Excel (two of that name: ambiguous; the archive guard applies);
  `fin_critic(target=...)` finds and scores the answer in that workbook; answer lines, KPI tiles, bridge steps,
  trend headers, drivers/ranking/matrix titles and key_values show metric names (a raw member: its metric if it
  is that metric's only member, else the atlas label) instead of the tokens passed; `xl_tidy` accepts row kinds
  `subtotal`, `total`, `heading` (= sum, grand, section). Pane: one `/api/chat` run at a time (409 with a message,
  send disabled while running); the pane prompt pins `target="active"` on every fin_analyze / fin_critic /
  fin_refresh / fin_drill call, stops on a fin_* error instead of rebuilding the analysis by hand with Excel MCP,
  and asks for metric names from fin_lookup. `npm test` in `addin/tn-claude` tests the server.

### Added

- v2 presentation (sub-spec 5, Plan 5): the answer sheet has fixed zones (title from `plan.title`, Key points,
  Charts, the tables, Checks, Notes, Chart data, Source); the plan is no longer drawn. Scenario headers (AC, PY, BU,
  FC, Δ) and a code column. `TableBlock.periods` adds column groups (default P&L/cost table:
  `{"type":"table","query":"q","periods":["closed","closed-ytd"]}`, month and YTD); `chain` completes the P&L chain
  of a table with EBITDA (new questions and patches); `material` adds a live ● Material column on labelled
  thresholds. New `chart` block (`kind` waterfall, line, column, bar, stacked, stacked100, combo, variance; `title`
  the message; `source` a block or own data laid out as Chart data; `series`, `highlight`, `sort`), drawn in IBCS
  style through COM. A new standard/deep question without a chart is refused (`guard_chart`); other paths get an
  implicit chart. Key points (at most 5) are live number sentences for the material lines, with the commentary
  driver as the cause. `Q<n> Data` gets a provenance header and a control total,
  Lines sheets a tie-out; a Sources sheet logs every pull with known
  limits. The reply has `layout` (the layout check's findings, each with a fix) and `charts`. Skill: answer-sheet
  rules, `reference/charts.md` (message → kind, one example per kind, tested against the schema), spec reference
  for `periods`, `chain`, `material`, `plan.title` and the chart block; plan guide and pane prompt require a chart
  chosen by the message.

- Excel task pane on v2 (Plan 3A): the TN Claude pane source now lives in `addin/tn-claude` (served from the repo
  on `https://localhost:3000`); its chat may use only `mcp__smartass`, `mcp__excel` and read-only tools, with a v2
  pane prompt (fin_analyze `target="active"`, xl_snapshot before Excel MCP edits, xl_tidy, never save). New CLI
  `smartass xl changes|revert|check` (one JSON object on stdout) backs the pane's Endringer and Formelsjekk tabs.
- v2 products and product concepts (Plan 2C): `Query.products` adds the `All Products` dimension to the grid (a
  `Product` column in the Data table only then); `table`/`ranking` take `rows: "products"`; 24 approved concepts
  (`splitt`, `streamix`, 22 `fam_*` families) as `knowledge/semantic/concepts/*.toml` seeded from `concepts.json`,
  found by `fin_lookup(kind="concept")`; the `concept` block shows booked, attributed, estimated (base x key share
  in helper columns) and missing items, the total and its checks as findings; the concept notes become caveats.
  knowledge.db is rebuilt when the semantic files changed (hash differs), not only when it is missing. Test-set
  case `b09` (mobile subscriptions family).
- v2 critic, playbooks and answer memory (Plan 2B): `critic_brief` in standard/deep replies and per question in
  `fin_refresh` (material lines backed by data, gaps); `fin_critic(q, score, notes)` returns the brief and the rubric
  (`knowledge/agent/critic.md`) and stores the score; the plan guide `knowledge/agent/plan.md`; four playbooks
  (monthly-pack, variance-commentary, gl-extract, prodreg-input) as MCP prompts with plan and critic, and as
  resources `smartass://playbooks/<name>` + `smartass://conventions`; answer memory `answers.db` (FTS5, 24 months)
  searched by `fin_lookup(kind="answer")`.
- v2 knowledge layer (Plan 2A): metrics and facts as TOML files in `knowledge/semantic`, compiled with the atlas
  into knowledge.db (build checks); `KnowledgeResolver` replaces the interim catalogue; every reply and Answer sheet
  carries the caveats it touched; `fin_lookup` finds metrics by alias; `smartass knowledge refresh | approve |
  status` for the monthly atlas diff, drafts and impact report; doctor check `knowledge` replaces `catalog`. The
  v1-only advisor texts `gate.md` / `findings.md` and the `products_map.json` side files are removed.
- v2: ad hoc fallback for failed HTTP grids, record/replay, test set (record / approve / run offline),
  text-fix candidates in the reply, wrapped notes, parallel drills.
- v2 ledger lines: lines block and fin_drill (supplier / project / account ranking tied to the cube), clean
  refusals, shorter SUMIFS.
- v2 robustness and toolkit: AutoSave paused while writing, session workbook kept across /mcp reconnects, a lost
  first request retried, xl_snapshot / xl_revert / xl_tidy.
- v2 complete answers: charts (trend, bridge), follow-ups via fin_analyze(patch, q), fin_refresh for monthly
  packs, drivers skip pass-through nodes, Q<n> Data sorted with data sheets.
- v2 model core (`engine/smartass/model`): typed analysis spec, period algebra, fact frame, expression trees
  (value and Excel formula from one tree), seven computed blocks (table, kpi_tiles, trend, matrix, bridge, ranking,
  drivers), data guards, material lines and key values. Not wired to tools yet (Plan 1B).
- v2 first answer (`smartass2` server, beside v1): `fin_analyze`, `fin_lookup`, `fin_status`; Smart View HTTP
  provider, interim atlas resolver, Q<n> Answer as live formulas on table Data_Q<n>, "Sheet = model" readback;
  answer files in `<analyser>\v2`. Not yet: follow-ups, packs, drill, xl_* tools, ad hoc fallback (Plan 1C).
- `sv_gate`: the advisor gate. A fixed checklist (period, scenario, entity, measure, currency/sign, granularity)
  against the semantic catalogue, the capability map and the user's learned defaults; `clear`, `ambiguous` (1-3
  questions) or `infeasible` (reason + proxy). It never pulls. `sv_remember` saves a gate answer as a default.
- `sv_answer`: one question from an answer spec v1 (`core/contract.py`): pulls over Smart View's HTTP provider
  (Smart View in Excel as the fallback), `Q<n> Answer` shown first, live formulas, checks, a findings table
  (Critical / Warning / Info), material lines and commentary drivers (data with evidence, or a hypothesis).
  - Follow-ups patch the stored spec (`followup=`): unchanged pulls are reused, the answer is redrawn inside the
    tool-owned range, a hidden backup restores it if the redraw fails.
  - Recipes draft the spec server-side: `pnl` (area or unit P&L, month and YTD vs last year, budget, forecast,
    bridges, KPIs, landing), `actual_vs_plan`, `monthly_trend`, `cost_base`, `concept`.
  - Run log: one JSON line per answer in `runs.jsonl` and a row in the workbook's hidden `_runlog` sheet.
- Recipe `pnl` reads its facts live on every call into an in-memory store; its check "Live read = source sheet"
  compares them with the source sheets. Materiality: 10 % of |comparison| (counted from NOK 500k) or NOK 5 million;
  a month swing of NOK 5 million is flagged; revenue, personnel cost and EBITDA are always commented. Consumer and
  Infrastructure Growth (abolished 1 June 2026) are refused with their successors.
- Login hardening: a Smart View Login that Windows reports as not responding counts as stuck at once (hidden Excel:
  restarted once; your Excel: "not responding; nothing was closed" after 10 s, never closed); a note after 5 s of
  an open Login; "Empty response from provider" is retried with a progress note, and in the hidden Excel a second
  one restarts only that Excel and retries once. Progress notes during every long stage.
- HTTP API: `POST /gate` and `POST /answer` beside `GET /health` (same calls as the MCP tools; 400 / 401 / 409 busy
  / 413 / 503).
- `[svhttp]` in `engine\defaults.toml`: the data source names and `allow_test`; environment variables
  `SV_HTTP_FINMAX_SERVER`, `SV_HTTP_FINPLAN_SERVER`, `SV_HTTP_ALLOW_TEST`, `SV_HTTP_TIMEOUT`.
- Semantic catalogue `knowledge\catalog.toml` (P&L lines and checks per cube, business areas, scenarios, KPIs,
  materiality) with a `smartass doctor` check.
- Faster answer sheets: answer cells written as one block, check colours set once per run of cells.

### Changed

- **One server, `smartass`, runs v2.** `run_server.py` -> `child_server.py` -> `smartass.server`; the side-by-side
  `smartass2` entry, `run_server_v2.py` and `child_server_v2.py` are gone (`smartass setup` removes a leftover
  `smartass2` MCP entry). Engine code lives in `engine/smartass/` (`core`, `model`, `data`, `render`, `epm`,
  `knowledge`, `xl`, `ops`); answer files go to the plain analyser folder again (no `v2` subfolder).
- **Reload on request only.** The relay restarts its server when `engine/VERSION` changes (`scripts/bump.ps1`
  bumps it with `pyproject.toml`; `smartass.__version__` reads it) or on `fin_status(restart=True)`, never on a
  `.py` edit.
- **Runtime fixes.** `pids.json` is written atomically under a lock file; the HTTP API secret is DPAPI-protected in
  `secret.bin` (a plain `http.secret` in `config.toml` is moved there once by `smartass setup`); Smart View options
  left by a crashed run are restored at the next start; `smartass doctor` checks Excel's VBA project access
  (Trust Center).
- **HTTP API:** `POST /analyze` (`{"spec", "target"}`, the same run as `fin_analyze`) replaces `/gate` and
  `/answer`; `GET /health` stays.
- **CLI:** `smartass atlas build | status` (the atlas build is no longer a tool); `doctor`, `setup`, `update` and
  `serve` run on the v2 modules (`python -m smartass.ops.cli`).
- `fin_status` returns the doctor's problems and takes `restart`.
- The SSO grace for the hidden Excel's sign-in (`LOGIN_GRACE_S`, env `SV_LOGIN_GRACE`) defaults to 30 s.
- Engine modules reorganised: packages `excel/` (the user's Excel) and `workbench/` (xl_* tools), `answer` +
  `answer_write`, `host` + `hostproc`.
- Tool count 35 -> 25 (v1 rework), then 8 (v2).

### Removed (breaking)

- **v1 retired:** the package `engine/smartview_mcp` (gate, recipes, answer spec v1, workbench tools), its tests,
  the live bench (`tests/live`, replaced by `python -m smartass.testset`) and the skill's `recipes.md`. Tool
  renames:

  | v1 | v2 |
  |---|---|
  | `sv_answer` (and `sv_gate`, `sv_remember`, recipes) | `fin_analyze` (AnalysisSpec v2; follow-ups by `patch`) |
  | `sv_drill` | `fin_drill` (or a `lines` block) |
  | `sv_find_member`, `sv_members`, `sv_member_info`, `sv_dimensions` | `fin_lookup` |
  | `sv_status`, `sv_restart`, `sv_atlas_status` | `fin_status` (`restart=True`) |
  | `sv_atlas_build`, `sv_atlas_changes` | `smartass atlas build` / `smartass atlas status` (CLI) |
  | `xl_workbooks`, `xl_overview`, `xl_read`, `xl_write`, `xl_rows`, `xl_add_sheet`, `xl_format`, `xl_format_sheet`, `xl_trace`, `xl_check`, `xl_changes`, `xl_revert` | Excel MCP + `xl_snapshot` / `xl_revert` / `xl_tidy` |
  | (new) | `fin_refresh` (monthly packs) |

- HTTP endpoints `POST /gate` and `POST /answer` (use `POST /analyze`).
- Tools `sv_query`, `sv_ws_pull`, `sv_ws_build`, `sv_ws_answer`, `sv_analyse`, `sv_concept`, `fin_pnl`,
  `sv_scope`, `sv_warehouse`, `sv_refresh_workbook`: use `sv_gate` + `sv_answer` (recipes for P&L, plan, trend,
  cost base and concepts); refresh an old workbook in Excel with Smart View's Refresh.
- The DuckDB warehouse (`warehouse.duckdb`, `paths.warehouse`, the doctor's warehouse check) and the `duckdb`
  dependency.
- Norwegian sheet names: sheets are `Q<n> Answer`, `Q<n> <cube>`, `Q<n> Lines` and `Sources`; old `Svar` /
  `Kilder` / `Linjer` sheets are no longer recognised as answers.

## 0.2.0

- Repo layout: the smartview MCP moves to the git repo `smartass` (package `smartview_mcp` under `engine\`),
  with `knowledge\`, `packs\`, `skills\`, `installer\`, `scripts\` and `docs\`.
- `uv` with a lockfile; Python 3.12; ruff.
- Config layering: `engine\defaults.toml` -> `%LOCALAPPDATA%\smartass\config.toml` -> environment variables;
  no user-specific paths in code; Excel, Smart View and OneDrive are auto-detected.
- `smartass doctor`: cheap health checks with a fix per check; `--bundle` zips results, versions, redacted config
  and log tails for the maintainer.
- Rotating logs with redaction of tokens and secrets.
- Error messages that name the next step.
- Localhost HTTP API (`127.0.0.1`, per-user secret) with `/health`.
- Updater: checks the release folder's `latest.json` at login, supports pinning (`release.version`), never
  downgrades automatically; `smartass update`.
- Per-user installer `installer\install.ps1` (no admin, idempotent): uv, `smartass setup` (config, MCP
  registration, skill junctions, Startup entries, atlas copy), `smartass doctor`.
- Pre-commit guard `scripts\check_staged.py`: no databases, stray workbooks or tokens in git.
- CI on `windows-latest`: ruff, pytest, data/secret guard.
