# File renderer bench: the Task 9 speed gate (2026-11-20)

Light and fast phase 3, Task 9, Step 1: `scripts/bench_answers.py --rounds 3` with `SV_RENDER=file`, live in the
owner's Excel, against the baseline of today's COM renderer (docs/audit/2026-11-19-baseline-v2.md). Times, COM-call
counts, stage names and OK/DIFF only, no figures from EPM.

- Engine: smartass 0.2.27 at 84f1d6f (branch feat/file-render, Tasks 1-8B), `SV_RENDER=file`, `SV_FASTCOM` on (the
  default).
- Run as the baseline was: one warm-up answer (b01, target `new`) made one new workbook in the analyser folder
  (`analyser\file-render`); then 3 rounds of the 9 test-set cases (rounds outer, cases inner) into that workbook
  (target `session`): 28 answers, the last one into a workbook of 27. Medians over the ok rounds; the warm-up is
  not in them. The workbook was closed after the run (it was saved by the engine after every answer).
- The Excel held no other workbook during the run (the baseline's held an earlier bench workbook of 19 answers).
- b16 is the only case with a lines block. Its drill ran in every round (browser route, direct drill); the
  browser's start (2.2 s) and sign-in (7.4 s) fell in round 1's b16, in `prepare`, outside the Excel job.
- Every answer: sheet = model OK (27/27 and the warm-up).

## The gate

Spec: the median `excel job._total` at most **3 s** without a lines sheet and **5 s** with one, **in every round**.

| round | answers without a lines sheet | their median excel job s (gate 3 s) | of them over 3 s | b16 (lines sheet) excel job s (gate 5 s) | gate |
|---|---|---|---|---|---|
| 1 | 8 (Q2-Q9) | 2.04 | 0 | 4.19 | within |
| 2 | 8 (Q11-Q18) | 2.74 | 2 | 4.11 | within |
| 3 | 8 (Q20-Q27) | **3.29** | 7 | **5.19** | **missed** |

**The gate is missed in round 3, both without a lines sheet (3.29 s) and with one (5.19 s).** Rounds 1 and 2 are
within it. Per Step 2 the work stopped there, before anything was switched over or deleted, until the controller's
ruling below.

## Controller ruling

> The gate miss in round 3 (3.29 s / 5.19 s at 61 sheets, against 46 s before) is ACCEPTED for the switch-over.
> Reason: the gate decides whether the file renderer may replace the old one, and it clearly may; the part that
> grows with the sheet count (per-sheet reads inside one COM call, the spec store) is removed by the next task (an
> in-process executor), so do NOT build a COM-side workaround for it now; the save's growth is Excel's own.

The re-gate moves to that next task.

## Before and after, per round (answers without a lines block)

| round | baseline (COM) excel job s | file excel job s | baseline _total s | file _total s | baseline COM calls | file COM calls |
|---|---|---|---|---|---|---|
| 1 | 15.6 | 2.0 | 17.2 | 3.4 | 275.5 | 34 |
| 2 | 25.8 | 2.7 | 28.2 | 3.7 | 275.5 | 34 |
| 3 | 46.0 | 3.3 | 47.0 | 4.2 | 275.5 | 34 |

All cases (median of all 27 ok runs): excel job 33.7 s -> 2.8 s, _total 34.9 s -> 4.1 s, COM calls 280 -> 34. b16:
excel job 45.3 s -> 4.2 s, _total 60.9 s -> 12.5 s (its drill, 7.8 s in `prepare.drill`, is the same route as
before).

## What grows with the workbook

The COM calls are flat (34-36 per answer). Round medians of the answers without a lines sheet, every stage of the
Excel job (from the run log):

| round | excel job s | open target s | build s | move in s | finish s | read back s | sources s | arrange s | save s | draw outside its named stages s | job outside open target, draw and save s |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | 2.04 | 0.09 | 0.08 | 0.70 | 0.18 | 0.03 | 0.13 | 0.02 | 0.38 | 0.28 | 0.04 |
| 2 | 2.74 | 0.11 | 0.08 | 0.94 | 0.16 | 0.03 | 0.17 | 0.03 | 0.61 | 0.47 | 0.19 |
| 3 | 3.29 | 0.11 | 0.09 | 0.94 | 0.16 | 0.03 | 0.20 | 0.04 | 0.83 | 0.71 | 0.21 |

From round 1 to round 3 the Excel job grows by 1.25 s:

- `save` +0.45 s: Workbook.Save of our own file (0.30 s at Q2, 1.06 s at Q28).
- Draw outside its named stages +0.43 s: `excel job.draw` less build, move in, finish, read back, sources and
  arrange. That is the job's one Book read (`movein.read_book`), quiet's way in and out, and the spec store
  (`sheets.save_spec`).
- `move in` +0.24 s (0.70 -> 0.94 s, flat from round 2).
- Outside open target, draw and save +0.17 s: `autosave_paused` (AutoSaveOn read and set around the job; our file
  is on OneDrive), `close_left` and the answer memory (`remember`).
- `sources` +0.07 s.

The saved workbook, read with openpyxl on a copy after the run, holds 61 worksheets: 28 Answers, 28 Data, 3 Lines
(b16), Sources and the hidden `_svmeta`.

**Hypothesis, not measured per call.** `read_book` and save_spec's `_spec_row` each count as one COM call, but
inside it they walk every worksheet of the workbook. `read_book` reads each sheet's name, visibility, tab colour
and tables; `_spec_row` reads each sheet's name to find `_svmeta`, then all of `_svmeta`'s used range (one row
per question). Each property is a cross-process round trip, so both grow with the sheet count. The save grows with
the file.

The answers without a lines sheet that took over 3 s:

- round 2: Q11 (3.08 s) and Q16 (3.24 s);
- round 3: every one but Q21 (2.84 s), so from Q22 on all of them.

## Per answer (from the run log, `config.RUN_LOG`)

| answer | case | round | excel job s | open target s | build s | move in s | finish s | read back s | sources s | arrange s | draw outside its stages s | save s | COM calls |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Q1 | b01 (warm-up) | warm-up | 7.05 | 2.52 | 0.17 | 0.98 | 0.30 | 0.03 | 2.17 | 0.02 | 0.47 | 0.31 | 68 |
| Q2 | b01 | 1 | 2.17 | 0.11 | 0.09 | 0.69 | 0.26 | 0.05 | 0.11 | 0.20 | 0.21 | 0.30 | 36 |
| Q3 | b02 | 1 | 2.05 | 0.09 | 0.06 | 0.69 | 0.25 | 0.03 | 0.20 | 0.01 | 0.31 | 0.38 | 34 |
| Q4 | b05 | 1 | 1.89 | 0.09 | 0.08 | 0.88 | 0.11 | 0.01 | 0.14 | 0.02 | 0.24 | 0.30 | 34 |
| Q5 | b06 | 1 | 1.97 | 0.11 | 0.08 | 0.89 | 0.12 | 0.01 | 0.12 | 0.01 | 0.26 | 0.34 | 34 |
| Q6 | b07 | 1 | 2.03 | 0.09 | 0.08 | 0.72 | 0.25 | 0.03 | 0.11 | 0.03 | 0.27 | 0.38 | 36 |
| Q7 | b09 | 1 | 2.38 | 0.12 | 0.12 | 0.97 | 0.28 | 0.03 | 0.12 | 0.02 | 0.29 | 0.41 | 36 |
| Q8 | b13 | 1 | 2.88 | 0.09 | 0.08 | 0.69 | 0.11 | 0.12 | 0.23 | 0.06 | 0.76 | 0.69 | 34 |
| Q9 | b15 | 1 | 2.02 | 0.09 | 0.09 | 0.69 | 0.11 | 0.02 | 0.14 | 0.01 | 0.35 | 0.41 | 34 |
| Q10 | b16 | 1 | 4.19 | 0.12 | 0.64 | 1.73 | 0.25 | 0.05 | 0.14 | 0.03 | 0.41 | 0.61 | 34 |
| Q11 | b01 | 2 | 3.08 | 0.11 | 0.09 | 1.02 | 0.31 | 0.05 | 0.16 | 0.02 | 0.43 | 0.61 | 36 |
| Q12 | b02 | 2 | 2.47 | 0.11 | 0.08 | 0.89 | 0.11 | 0.01 | 0.16 | 0.02 | 0.39 | 0.53 | 34 |
| Q13 | b05 | 2 | 2.72 | 0.11 | 0.08 | 0.92 | 0.12 | 0.03 | 0.17 | 0.03 | 0.46 | 0.61 | 34 |
| Q14 | b06 | 2 | 2.64 | 0.11 | 0.11 | 0.95 | 0.20 | 0.01 | 0.17 | 0.03 | 0.48 | 0.55 | 34 |
| Q15 | b07 | 2 | 2.74 | 0.11 | 0.06 | 0.94 | 0.22 | 0.03 | 0.16 | 0.03 | 0.45 | 0.56 | 36 |
| Q16 | b09 | 2 | 3.24 | 0.11 | 0.12 | 1.03 | 0.31 | 0.05 | 0.14 | 0.03 | 0.49 | 0.66 | 36 |
| Q17 | b13 | 2 | 2.74 | 0.09 | 0.08 | 0.91 | 0.09 | 0.02 | 0.17 | 0.03 | 0.48 | 0.64 | 34 |
| Q18 | b15 | 2 | 2.78 | 0.11 | 0.09 | 0.94 | 0.12 | 0.03 | 0.20 | 0.03 | 0.56 | 0.67 | 34 |
| Q19 | b16 | 2 | 4.11 | 0.11 | 0.53 | 1.38 | 0.25 | 0.03 | 0.17 | 0.02 | 0.64 | 0.84 | 34 |
| Q20 | b01 | 3 | 3.41 | 0.11 | 0.09 | 0.81 | 0.50 | 0.03 | 0.22 | 0.03 | 0.62 | 0.76 | 36 |
| Q21 | b02 | 3 | 2.84 | 0.19 | 0.08 | 0.73 | 0.12 | 0.03 | 0.17 | 0.03 | 0.68 | 0.78 | 34 |
| Q22 | b05 | 3 | 3.17 | 0.12 | 0.09 | 0.92 | 0.12 | 0.03 | 0.19 | 0.03 | 0.67 | 0.77 | 34 |
| Q23 | b06 | 3 | 3.30 | 0.11 | 0.09 | 0.99 | 0.20 | 0.02 | 0.19 | 0.05 | 0.66 | 0.80 | 34 |
| Q24 | b07 | 3 | 3.27 | 0.11 | 0.06 | 0.95 | 0.25 | 0.03 | 0.19 | 0.05 | 0.75 | 0.86 | 36 |
| Q25 | b09 | 3 | 3.92 | 0.12 | 0.12 | 1.00 | 0.31 | 0.03 | 0.30 | 0.05 | 0.88 | 0.88 | 36 |
| Q26 | b13 | 3 | 3.28 | 0.12 | 0.09 | 0.95 | 0.11 | 0.03 | 0.20 | 0.03 | 0.82 | 0.86 | 34 |
| Q27 | b15 | 3 | 3.66 | 0.11 | 0.08 | 0.91 | 0.12 | 0.03 | 0.23 | 0.05 | 1.02 | 0.88 | 34 |
| Q28 | b16 | 3 | 5.19 | 0.12 | 0.58 | 1.67 | 0.25 | 0.05 | 0.22 | 0.05 | 0.98 | 1.06 | 34 |

(1, False) n 8 median xl 2.040 outside 0.28 save 0.38 over3 0
(1, True) n 1 median xl 4.190 outside 0.41 save 0.61 over3 1
(2, False) n 8 median xl 2.740 outside 0.47 save 0.61 over3 2
(2, True) n 1 median xl 4.110 outside 0.64 save 0.84 over3 1
(3, False) n 8 median xl 3.290 outside 0.71 save 0.83 over3 7
(3, True) n 1 median xl 5.190 outside 0.98 save 1.06 over3 1

## The bench's own output

### Baseline v2 answers, 2026-11-20 (the heading the script prints)

Engine 0.2.27; render file; 3 rounds of 9 cases; medians over the ok rounds.

| case | ok runs | _total s | prepare.fetch s | prepare.drill s | excel job s | COM calls | errors |
|---|---|---|---|---|---|---|---|
| b01 | 3/3 | 3.9 | 0.8 | 0.0 | 3.1 | 36 |  |
| b02 | 3/3 | 3.3 | 0.8 | 0.0 | 2.5 | 34 |  |
| b05 (no chart block) | 3/3 | 3.2 | 0.5 | 0.0 | 2.7 | 34 |  |
| b06 | 3/3 | 5.3 | 2.6 | 0.0 | 2.6 | 34 |  |
| b07 | 3/3 | 3.5 | 0.9 | 0.0 | 2.7 | 36 |  |
| b09 (no chart block) | 3/3 | 6.5 | 3.0 | 0.0 | 3.2 | 36 |  |
| b13 | 3/3 | 3.7 | 0.8 | 0.0 | 2.9 | 34 |  |
| b15 (no chart block) | 3/3 | 4.9 | 2.1 | 0.0 | 2.8 | 34 |  |
| b16 | 3/3 | 12.5 | 0.2 | 7.8 | 4.2 | 34 |  |
| **all cases** (median of all ok runs) | 27/27 | 4.1 | 0.8 | 0.0 | 2.8 | 34 |  |
| warm-up b01 (target 'new', not in the medians) | 1/1 | 10.4 | 3.3 | 0.0 | 7.0 | 68 |  |

#### Per round (ok runs without a lines block)

| round | runs | _total s | excel job s | draw s | draw.data sheet s | draw.answer s | draw.build s | draw.move in s | draw.finish s | draw.read back s | draw.sources s | draw.arrange s | save s | prepare s | COM calls |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | 8 | 3.4 | 2.0 | 1.5 | - | - | 0.1 | 0.7 | 0.2 | 0.0 | 0.1 | 0.0 | 0.4 | 1.1 | 34 |
| 2 | 8 | 3.7 | 2.7 | 1.9 | - | - | 0.1 | 0.9 | 0.2 | 0.0 | 0.2 | 0.0 | 0.6 | 0.8 | 34 |
| 3 | 8 | 4.2 | 3.3 | 2.3 | - | - | 0.1 | 0.9 | 0.2 | 0.0 | 0.2 | 0.0 | 0.8 | 0.8 | 34 |

#### Stages

| stage | runs | median s | sum s | share of _total |
|---|---|---|---|---|
| _total | 27 | 4.14 | 167.9 | 100% |
| prepare | 27 | 0.88 | 87.7 | 52% |
| excel job._total | 27 | 2.84 | 80.1 | 48% |
| excel job.draw | 27 | 1.97 | 56.3 | 34% |
| prepare.drill | 27 | 0.00 | 52.1 | 31% |
| drills.direct drill (1) | 3 | 7.50 | 50.8 | 30% |
| prepare.fetch | 27 | 0.80 | 35.6 | 21% |
| excel job.draw.move in | 27 | 0.92 | 25.9 | 15% |
| excel job.save | 27 | 0.64 | 17.1 | 10% |
| drills.browser sign-in (1) | 1 | 7.40 | 7.4 | 4% |
| excel job.draw.finish | 27 | 0.20 | 5.5 | 3% |
| excel job.draw.sources | 27 | 0.17 | 4.7 | 3% |
| excel job.draw.build | 27 | 0.09 | 3.8 | 2% |
| excel job.open target | 27 | 0.11 | 3.0 | 2% |
| drills.browser start (1) | 1 | 2.20 | 2.2 | 1% |
| excel job.draw.arrange | 27 | 0.03 | 1.0 | 1% |
| drills.http drill link (1) | 3 | 0.20 | 1.0 | 1% |
| excel job.draw.read back | 27 | 0.03 | 0.9 | 1% |
| excel job.open target.ready + bridge | 27 | 0.02 | 0.5 | 0% |
| excel job.open target.find Excel windows | 27 | 0.00 | 0.2 | 0% |
| excel job.open target.open workbook | 27 | 0.01 | 0.2 | 0% |
| excel job.open target.open workbook.wait for that Excel (another agent) | 27 | 0.00 | 0.0 | 0% |

Excel: 16.0 (build 20527.0)
