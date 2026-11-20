# In-process bench: the re-gate of phase 3b, Task 2b (2026-11-20)

Light and fast phase 3b, Task 2b, Step 4: `scripts/bench_answers.py --rounds 3` with `SV_INPROC=1`, live in the
owner's Excel (16.0, build 20527), against the file renderer's bench of the day before
(docs/audit/2026-11-19-file-render-bench.md, where round 3 missed the gate). Times, COM-call and batch counts, stage
names and OK/DIFF only, no figures from EPM.

- Engine: smartass 0.3.0 + branch feat/inproc at 7217944 (Task 1: batches and the add-in; Task 2a: the insert by
  `Sheets.Add(Type:=path)`; Task 2b: the file job as four batches and a save), `SV_INPROC=1`, `SV_FASTCOM` on.
- The add-in loaded silently in the owner's Excel (pid 140372) on the first batch of the warm-up (no window, no
  dialog; `check_followups.py` had loaded it the same way earlier that morning); every batch of the bench ran
  in-process.
- Run as before: one warm-up answer (b01, target `new`) made one new workbook in the analyser folder
  (`analyser\inproc`); then 3 rounds of the 9 test-set cases (rounds outer, cases inner) into that workbook (target
  `session`): 28 answers, the last into a workbook of 27. Medians over the ok rounds; the warm-up is not in them.
  The workbook was closed after the run (saved by the engine after every answer). The Excel held no other workbook.
- b16 is the only case with a lines block (its drill in `prepare`, outside the Excel job; the browser's start and
  sign-in fell in round 1).
- Every answer: sheet = model OK (27/27 and the warm-up). Every answer: 5 batches ("runs"), 9 COM calls (the
  batches, the workbook's name, the target's own calls in `open target`).

## The gate

Spec: the median `excel job._total` at most **3 s** without a lines sheet and **5 s** with one, **in every round**.

| round | answers without a lines sheet | their median excel job s (gate 3 s) | 2026-11-19 (file, over COM) | b16 (lines sheet) excel job s (gate 5 s) | 2026-11-19 | gate |
|---|---|---|---|---|---|---|
| 1 | 8 | 1.9 | 2.04 | 4.1 | 4.19 | within |
| 2 | 8 | 2.4 | 2.74 | 3.9 | 4.11 | within |
| 3 | 8 | **2.7** | 3.29 | **4.3** | 5.19 | **within** |

**The gate holds in every round**, also in round 3 (61 sheets), which missed it the day before.

All cases (median of all 27 ok runs): excel job 2.8 s -> 2.4 s, _total 4.1 s -> 3.6 s, COM calls 34 -> 9, 5 batches.

## Per round, every stage of the Excel job (answers without a lines block, medians, from the run log)

| round | excel job s | draw s | look s | build s | move in s | finish s | close out s | save s | COM calls | runs |
|---|---|---|---|---|---|---|---|---|---|---|
| 1 | 1.9 | 1.4 | 0.0 | 0.0 | 0.9 | 0.2 | 0.2 | 0.3 | 9 | 5 |
| 2 | 2.4 | 1.6 | 0.0 | 0.0 | 1.1 | 0.2 | 0.2 | 0.5 | 9 | 5 |
| 3 | 2.7 | 1.9 | 0.0 | 0.0 | 1.3 | 0.2 | 0.3 | 0.6 | 9 | 5 |

With a lines block (b16, one answer per round): excel job 4.1 / 3.9 / 4.3 s, of it move in 2.2 / 1.9 / 2.1 s (the
lines sheet's rows come in with the file) and build 0.6 / 0.6 / 0.5 s.

Where the time inside Excel goes (every answer, medians over the 27; `excel job.<stage>.excel.<step>`, from the
batches' marks):

| step | median s | sum s over 27 |
|---|---|---|
| move in: the insert (`Sheets.Add(Type:=path)`, Excel's own) | 0.56 | 17.8 |
| move in: AutoSave off (our file is on OneDrive) | 0.37 | 10.1 |
| move in: the settings put back, the result | 0.07 | 2.0 |
| finish: tables and formulas | 0.04 | 1.2 |
| finish: the settings put back, the result | 0.04 | 1.2 |
| close out: the settings put back, the result | 0.05 | 1.2 |
| close out: sheet order | 0.02 | 1.3 |
| every other step | 0.00-0.02 | under 0.5 each |
| the save batch (`Workbook.Save` + AutoSave back on, Excel's own) | 0.55 | 12.8 |

What grows from round 1 to round 3 (+0.8 s) is Excel's own work on a growing OneDrive file: the move-in
(+0.4 s: the insert and AutoSave off) and the save (+0.3 s). The look, which walked every sheet over COM before,
takes 0.02 s at 61 sheets; the reads that grew with the sheet count the day before are gone from the bill.

## Follow-ups (`scripts/check_followups.py`, the same morning)

| step | 2026-11-20 before Task 2b (over COM) | in-process (SV_INPROC=1) | over COM (SV_INPROC=0) |
|---|---|---|---|
| new answer (workbook created) | 6.83 s, 65 calls | 7.08 s, 5 batches (9.72 s in the morning's first run: the add-in's build and load, 1.3 s, fell in its look) | 7.55 s, 5 batches |
| second answer | 1.86 s, 33 calls | 1.69 s, 5 batches | 2.59 s, 5 batches |
| follow-ups 1-6 | 2.08-2.41 s, 34 calls | 1.64-1.91 s, 5 batches | 2.53-2.88 s, 5 batches |
| fin_refresh (2 answers) | 4.25 s, 68 calls | 3.38 s, 10 batches | 5.17 s, 10 batches |

Every step "checks OK" on both executors. The follow-up goal of 1.5 s is not reached: a follow-up's Excel job is
1.64-1.91 s, of which the move-in batch is 0.84-0.91 s (the insert 0.45-0.48 s, AutoSave off 0.15-0.24 s), the save
0.22-0.33 s, finish 0.20-0.23 s, close out 0.17-0.20 s, `open target` about 0.1 s and the two paint pauses 0.08 s.
What is left above 1.5 s is Excel's own work on our OneDrive file: AutoSave off and back on around the job, and the
save (the plan's Task 5 moves the save after the reply). Over COM (the fallback) a follow-up is 2.5-2.9 s: every
operation is a call of its own there, and each quiet batch sets and puts back the four settings over COM.


## After the review fixes (08:40-09:03, the same day)

The job became four batches (look, bring in = move in + finish, close out, save), one paint pause of 0.15 s after
the bring-in, Calculation manual only in the bring-in, the label lift inside it (verb `nudge`), AutoSave switched
off in the look. Re-run in the owner's Excel, **while another session was working in the same Excel**: from 08:48 its
own answer workbook ("... Pull unified FinPlan volume KPIs ...") was the active one there, and the paint pause (only
for the active workbook) ran only in round 1 of the bench. The inserts alone took about twice as long as in the
morning's runs above (median 1.08 s against 0.56 s), so these figures do not compare with the morning's; the gate is
reported as measured.

`scripts/bench_answers.py --rounds 3`, SV_INPROC=1, all nine cases: 27/27 ok, sheet = model OK, 4 batches and 8 COM
calls per answer.

| round | answers without a lines sheet | median excel job s (gate 3 s) | b16 (lines sheet) excel job s (gate 5 s) | gate |
|---|---|---|---|---|
| 1 | 8 | **4.8** | **5.3** | missed |
| 2 | 8 | 3.0 | 4.8 | within (at the edge) |
| 3 | 8 | **3.3** | **8.5** | missed |

| round | excel job s | draw s | look s | build s | move in s | finish s | paint s | close out s | save s | runs |
|---|---|---|---|---|---|---|---|---|---|---|
| 1 | 4.8 | 3.1 | 0.0 | 0.1 | 1.5 | 0.4 | 0.2 | 0.6 | 0.8 | 4 |
| 2 | 3.0 | 1.9 | 0.0 | 0.0 | 1.0 | 0.2 | - | 0.4 | 0.7 | 4 |
| 3 | 3.3 | 2.0 | 0.0 | 0.0 | 1.0 | 0.2 | - | 0.5 | 1.0 | 4 |

Inside Excel (medians over the 27): the insert 1.08 s (0.56 s in the morning), the bring-in's end (the settings and
Calculation put back) 0.12 s, the close-out's end 0.10 s, tables, formulas, labels and calculate 0.06 s; the save
0.78 s (0.55 s). `open target` 0.34 s (0.12 s).

`scripts/check_followups.py` in the same hour, every step "checks OK" and the saved file read back OK:

| step | in-process (SV_INPROC=1) | over COM (SV_INPROC=0) |
|---|---|---|
| new answer | 14.39 s | 19.58 s |
| second answer | 3.88 s | 7.45 s |
| follow-ups 1-6 | 2.62-4.09 s | 7.34-8.55 s |
| fin_refresh (2 answers) | 7.23 s, 8 batches | 23.83 s, 8 batches |

Earlier in that hour, before the other session's workbook came (08:28-08:42, in-process, with --settle 10): follow-ups
1.66-2.0 s at 4 batches (move in 0.60-0.68 s, finish 0.13-0.25 s, paint 0.14-0.16 s, close out 0.16-0.28 s, save
0.25-0.38 s), the second answer 1.51-1.59 s.

AutoSave (controller ruling R3, proven or not): with AutoSave left alone, the in-process job wrote our file only in
its save batch (--settle 10, every step), but over COM the file was written once in the middle of a job's batches
(+3.56 s, its save at +5.31 s). So the job keeps switching AutoSave off while it writes, now as the look's last
operations. The file is also stamped 2-11 s after a save at times of OneDrive's own, with AutoSave on or off.

The look with 2 000 defined names and 60 sheets (a hidden Excel; review M5): in-process 35 ms, over COM 18.3 s (the
`each` walk over every name is one or more cross-process calls per name). In-process it stays under the 100 ms bar;
over COM (the fallback) such a workbook pays seconds per answer.

Task 2c (re-review 2 of Task 2b) took the names out of the look: the bring-in batch looks up the answer's own table
names (the verb `exist`, one `Item` per name) instead of walking every name, and the property check of `each`'s
path steps is cached per interface IID, so a cached step costs no call to Excel. Measured again with 2 000 defined
names and 60 sheets (`scripts/check_inproc.py --look-bench`, a hidden Excel, 20.11.2026; two runs; in-process the
median of 3, over COM one run each):

| | in-process | over COM |
|---|---|---|
| the look (without the names now) | 11.5-16.9 ms | 531-675 ms |
| the walk over every name the look made before | 33.9-37.0 ms | 3.3-5.1 s |
| the lookup of an answer's two table names (now) | 7.4-12.2 ms | 7.2-8.6 ms |

The walk over COM is 3.3-5.1 s now against 18.3 s before (the property check no longer asks the type of every
item); the job no longer makes it. Over COM the look's 0.5-0.7 s is mostly, by its make-up, the walk over the 60
sheets (four paths each; not timed apart).

## Under load, on head 534fe50 (11:15-11:23, the same day)

Run from a frozen checkout of 534fe50 (four batches, after Task 2c; before the re-pointing of Task 3) in the owner's
Excel, while two workbooks of another session were open and unsaved in the same Excel and that session was possibly
working there. Not a gate run: Excel's own steps (the insert, the save) take about twice as long when the instance is
busy with other workbooks.

`scripts/check_followups.py`, every step "checks OK", 4 batches and 8 COM calls per answer:

| step | in-process (SV_INPROC=1) | over COM (SV_INPROC=0, minutes later) |
|---|---|---|
| new answer (workbook created) | 17.3 s (look 1.88 s: the add-in's build and load for the new hash; close out 3.24 s, of it a new Sources log's print setup 1.44 s and its furniture 0.81 s; the workbook's creation about 9 s) | - |
| second answer | 2.75 s | 6.05 s |
| follow-ups | 2.50-2.89 s (move in 0.92-1.12 s, of it the insert 0.81-0.99 s; finish 0.27-0.40 s; paint 0.14-0.16 s; close out 0.33-0.38 s; save 0.31-0.36 s) | 4.92-6.22 s |
| fin_refresh (2 answers) | 5.48 s | 10.7 s |

Under load the add-in route is about twice as fast as the fallback.

`scripts/bench_answers.py --rounds 3`, SV_INPROC=1: 27/27 ok. Median Excel job per round 4.7 / 2.7 / 5.9 s without a
lines sheet, 5.0 / 5.4 / 8.9 s with one; the insert alone median 1.54 s (0.56 s in a quiet Excel), the save 0.92 s.
Round 2 shows the quiet level. The gate evidence for the branch stays the quiet run of 06:00 above (five batches; the
merge into four only removed work) and the quiet follow-ups of 08:28-08:42 (1.66-2.0 s).

## The re-pointing's costs (Task 3, a hidden Excel)

The scan for formulas pointing at the answer when none do (`scripts/check_inproc.py --repoint-bench`; in-process the
median of 3, over COM one run):

| workbook | in-process | over COM |
|---|---|---|
| 10 sheets, 5 000 formulas | 64-79 ms | 298-437 ms |
| 60 sheets, 260 000 formulas | 0.90-1.00 s | 4.4-4.5 s |
| one sheet, used range to row 1 048 576, 300 formulas | 53-65 ms | 638-690 ms |

These workbooks have no defined names: the scan reads every name too, about 3-5 s per 2 000 names over COM (the walk
above), tens of milliseconds in-process. A sheet whose used range reaches both the last row and the last column was
not benched.

A follow-up of Q1 with 500 cells pointing at it, against none (`scripts/check_insert.py --parts repoint-bench`; two
runs, the second on a busier machine; all 500 re-pointed in both executors):

| executor | pointing cells | draw | close out |
|---|---|---|---|
| in-process | 0 | 0.52 / 0.83 s | 0.03 / 0.09 s |
| in-process | 500 | 0.62 / 1.14 s | 0.17 / 0.28 s |
| over COM | 0 | 0.80 / 1.64 s | 0.08 / 0.19 s |
| over COM | 500 | 2.70 / 7.08 s | 1.14 / 2.98 s |

## The final code, 17:20, Excel otherwise empty (controller run)

`scripts/bench_answers.py --rounds 3`, `SV_INPROC=1`, head 652ac47 (four batches, one paint, the re-pointing), from a
frozen checkout; the owner's Excel had no other workbook open (the nine another session had opened were closed first
on the owner's word), but it is the same Excel process that has run since 22:40 the evening before, through several
hundred test workbooks. 27 of 27 answers ok.

### Per round (ok runs without a lines block)

| round | answers | _total s | excel job s | draw s | draw.look s | draw.build s | draw.move in s | draw.finish s | draw.paint s | draw.close out s | save s | prepare s | COM calls | runs |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | 8 | 3.6 | 2.3 | 1.6 | 0.0 | 0.0 | 0.9 | 0.2 | 0.2 | 0.1 | 0.5 | 1.3 | 8 | 4 |
| 2 | 8 | 5.0 | 3.2 | 1.9 | 0.0 | 0.0 | 1.1 | 0.2 | 0.1 | 0.3 | 0.8 | 1.6 | 8 | 4 |
| 3 | 8 | 5.4 | 3.6 | 2.0 | 0.0 | 0.1 | 1.1 | 0.3 | 0.2 | 0.3 | 1.1 | 1.6 | 8 | 4 |

### Per round (ok runs with a lines block)

| round | answers | _total s | excel job s | draw s | draw.look s | draw.build s | draw.move in s | draw.finish s | draw.paint s | draw.close out s | save s | prepare s | COM calls | runs |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | 1 | 92.3 | 5.2 | 4.2 | 0.0 | 0.9 | 2.7 | 0.3 | 0.1 | 0.1 | 0.6 | 87.1 | 8 | 4 |
| 2 | 1 | 13.7 | 6.0 | 4.5 | 0.0 | 1.0 | 2.6 | 0.3 | 0.2 | 0.3 | 1.0 | 7.7 | 8 | 4 |
| 3 | 1 | 13.6 | 5.9 | 4.2 | 0.0 | 1.0 | 2.5 | 0.2 | 0.2 | 0.2 | 1.1 | 7.7 | 8 | 4 |

Reading: over the 3 s / 5 s gate in rounds 2 and 3 (and in every round with a lines sheet). Against the 06:00 run
(1.9 / 2.4 / 2.7 s): the move-in is the same (0.9 / 1.1 / 1.1 s now, 0.9 / 1.1 / 1.3 s then); the save took longer
(0.5 / 0.8 / 1.1 s against 0.3 / 0.5 / 0.6 s), so did opening the target (about +0.3 s in rounds 2 and 3), and the
one paint pause is 0.15 s where the morning code had two of 0.04 s. So the gate rests on the 06:00 run (five-batch
code) only. Owed: one run in a freshly started Excel.

## 0.3.0 against 0.4.0, back to back (17:25-17:30, same Excel, no other workbook open)

`scripts/check_followups.py` from the main checkout (0.3.0, f32352f) and from the frozen checkout of this branch
(652ac47, `SV_INPROC=1`), alternating, each twice; Excel-job seconds per step, every step "checks OK" in all four runs.

| step | 0.3.0 run 1 | 0.4.0 run 1 | 0.3.0 run 2 | 0.4.0 run 2 |
|---|---|---|---|---|
| new answer | 12.69 | 10.61 | 12.66 | 11.28 |
| second answer | 3.56 | 2.30 | 3.89 | 2.19 |
| follow-up 1 | 4.12 | 3.73 | 4.45 | 3.42 |
| follow-up 2 | 4.34 | 3.06 | 4.38 | 3.78 |
| follow-up 3 | 5.36 | 3.16 | 5.22 | 3.28 |
| follow-up 4 | 4.30 | 3.59 | 4.86 | 3.47 |
| follow-up 5 | 4.39 | 3.58 | 5.25 | 3.26 |
| follow-up 6 | 4.59 | 2.97 | 4.78 | 3.22 |
| fin_refresh | 9.24 | 5.74 | 8.53 | 6.40 |

0.4.0 is faster at every step under the same conditions (a follow-up about 3.3 s against about 4.6 s). 0.3.0 itself
took 2.2-2.8 s per follow-up the night before (docs/audit/2026-11-19-file-render-bench.md and the Task 8B / 9 runs),
so on this afternoon the same code ran about 1.8 times slower than at night: the absolute figures of this section and
of the 17:20 bench are not comparable with the night's and the morning's.

The follow-up check on the same head (17:14, nine workbooks of another session open): every step OK through both
executors, the summary sheet's formulas (cell, table, cross-sheet SUM, defined name, chart series, a reference from
another answer) correct after every follow-up and after fin_refresh; follow-ups 4.2-5.0 s in-process, 7.2-7.7 s over
COM.
