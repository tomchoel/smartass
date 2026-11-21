# The first answer of a new workbook: profile, spike, hidden and live figures (2026-11-20)

Light and fast phase 3c, Tasks 1 and 1b (the first answer of a new workbook is its built file, Excel only opens it)
and Task 2 (the save after the reply). Seconds and counts only; no figure from the finance system. "Live" is the
owner's Excel; "hidden" is a hidden Excel a script started itself (`DispatchEx`) on "zz probe" files in the OneDrive
analysis folder, closed and deleted at the end.

## 0.4.0: where a new workbook's first answer went (run log, owner's Excel, 17:28-17:31)

Total 10.6-11.3 s:

| stage | s |
|---|---|
| open target | 4.8-5.2 (add from template 1.1-1.2, save as 3.2-3.4, ready + bridge 0.2-0.4) |
| look | 0.59-0.64 |
| move in | 0.93 |
| finish | 0.26 |
| paint | 0.16 |
| close out | 2.9-3.0 (print setup 1.4-1.5, sources 0.75, end 0.57, spec 0.08) |
| save | 0.61-0.66 |

## The spike (controller, `seed_open.py`, owner's Excel, 3 rounds)

- Today's way: `Workbooks.Add` 0.87-1.36 s, then name, label and `SaveAs` 1.36-3.19 s: 2.4-4.3 s.
- A copy of the labelled seed opened with `Workbooks.Open`: 1.55-1.67 s, with AutoSave on, Saved, labelled, 47 styles
  against 101, one sheet "Sources".
- The first Save after an edit: 0.5-0.65 s.

## Task 1 in a hidden Excel (OneDrive folder, through the add-in)

- Probe 3, whole first answers (`excel_job`), final code, three rounds:

| round | today's way s | built as its file s | built: draw.open s | built: look s |
|---|---|---|---|---|
| 0 | 16.22 (the template built first) | 2.44 | 1.47 | 0.02 |
| 1 | 8.05 | 2.06 | 1.01 | 0.02 |
| 2 | 11.55 | 3.43 | 1.50 | 0.02 |

  An earlier run of four rounds: built 1.92-2.46 s against 7.05-13.72 s; `draw.open` 1.06-1.83 s; look 0.01-0.08 s.
  In the ledger's summary: 2.0-3.4 s in the Excel job built as its file against 8.0-11.6 s made from the template.
- `check_insert.py --parts newbook`: 72/72 item by item against a workbook made today's way (78/78 after Task 2); in
  a local folder the first answer 0.65-0.72 s built against 4.5-4.9 s today's way.
- Every built workbook: no repair, no Protected View, AutoSave on and Saved after the Open, the label recognised,
  its Save writes, the base equal to the template's.

## Live, Task 1 (20.11.2026 22:16-22:18, frozen checkout at f9004f9, a fresh Excel, nothing else open)

`check_followups.py` twice, exit 0, every step "checks OK". The save was still in the job (before Task 2).

| new workbook's first answer | total s | open target s | of it "new workbook" s | draw.open s | look s | build s | move in s | finish s | paint s | close out s | save s |
|---|---|---|---|---|---|---|---|---|---|---|---|
| run 1 | 6.06 | 2.64 (ready + bridge 0.56) | 1.98 | 1.08 | 1.14 | 0.14 | 0.11 | 0.13 | 0.16 | 0.14 | 0.27 |
| run 2 | 5.27 | 2.11 | 2.00 | 1.05 | 0.83 | 0.14 | 0.09 | 0.13 | 0.16 | 0.11 | 0.24 |

"New workbook" had no Excel sub-stage in it. Second answer 1.23-1.61 s; follow-ups 1.53-2.14 s; `fin_refresh`
3.4-4.1 s. Against the goal (5 s or under): 5.3-6.1 s, not yet.

**What the window captures showed** (frames of the screen during the runs): the workbook opens showing the formatted,
empty Answer with "Internal" in the title, no repair prompt, no Protected View bar; the whole answer with its chart is
on screen about 1.4 s later; no black half-paint.

**The label check.** The kept workbook: AutoSave True, Saved True, Sources frozen (row 1), filter on, print area
`$A$1:$N$1048575`, 48 styles. On disk `docMetadata/LabelInfo.xml` carries our label id, enabled, method Privileged,
not removed: identical to a workbook made the old way. The API's `LabelName` is empty for both; the title briefly
showed "Select a Label" during the next job (a transient).

## 0.4.0 in the same fresh Excel (22:19-22:22)

- `check_followups.py`: all OK. New workbook 7.52 s (save as 2.75, add from template 0.50, close out 1.95 of which
  print setup 0.94 and sources 0.55); second answer 1.19 s; follow-ups 1.62-1.95 s (the sixth 3.48 s, with a 1.17 s
  save); `fin_refresh` 3.20 s.
- `bench_answers.py`, 3 rounds: 1.3 / 1.9 / 2.0 s without a lines sheet, 3.6 / 3.1 / 3.6 s with one. The gate (3 s /
  5 s) held in every round in a fresh Excel. The afternoon's 2.3 / 3.2 / 3.6 s came from an Excel process that had
  run for 18 hours.

## The first attempt (21:10-21:12): Excel closed by a think-cell update

The first live run made its first answer (12.84 s: the old way once, which stores the machine's base) and its second,
and then the owner's Excel closed. The Windows Application log shows MsiInstaller installing think-cell 14.0.38.836 at
21:11:57-21:12:55 (`tcupdate.exe`, Restart Manager sessions, "machine restart is required"): the update shut Excel
down. No Application Error or hang event for EXCEL.EXE. Not caused by smartass. The run after it found no Excel and
took the no-Excel path; that Excel stopped at think-cell's "running into problems ... disable?" prompt.

## Task 1b: the seconds that were not Excel's

**"New workbook": 2.0 s.** The controller's profile (cProfile, 22:33): `target._reserve` -> `_new_path` ->
`session.project_dir` -> `session.info` -> `psutil.Process.parents()`, which on Windows takes a snapshot of every
process on the machine for each ancestor: 8 calls of `ppid_map` at 188 ms = 1.5 s, once per server process (the
first answer of every session). `info()` now walks the parents from one snapshot. Plain script on this PC (531
processes), a fresh process each time:

| | before s | after s |
|---|---|---|
| no relay ancestor (8 parents) | 1.75 / 1.80 / 1.97 | 0.20 / 0.21 / 0.21 |
| under a relay (`run_server.py`, 10 parents) | 2.15 | 0.20 |

The same answer before and after, in one process, in both cases.

Other process-table reads (plain script, this PC):

| call | where | s |
|---|---|---|
| `sso.global_watcher_pids` (`process_iter` name + command line) | every `resolve_target` (`sso.ensure_global_watcher`) | 0.05-0.10 |
| `session.set_workbook` (command line of every process) | once a new workbook holds its answer | 0.04 |
| `instances.excel_instances` | every job | 0.03 |
| `hostproc._excel_pids` (`process_iter` name) | the hidden Excel's start and reset | 0.002 |
| `dialogs._webview2_rows` (`process_iter` with `ppid`: a snapshot per process) | a Smart View sign-in stuck or fallen back to Internet Explorer | 104 |
| `host._do_hard_reset` with no pid yet (`process_iter` with `create_time`) | resetting the hidden Excel while `DispatchEx` has not returned | 47 |

The last two are recovery paths, not a normal job. They are not changed here.

**The look: 0.8-1.1 s live against 0.03 s.** In a hidden Excel, on "zz probe" files in the OneDrive folder, with the
look split into the add-in's load, the batch's wall time, its time inside Excel and Excel's rejections:

| round | look s | the add-in's load s | the batch inside Excel s | rejected calls |
|---|---|---|---|---|
| a just-opened built workbook, add-in loaded in this process | 0.01-0.02 (once 0.09) | 0 | 0.004-0.009 | 0 |
| the same, a new process meeting the Excel that holds the add-in (before) | 0.45-0.49 | 0.37-0.38 | 0.03-0.05 | 0 |
| the same (after) | 0.04-0.09 | 0.03-0.07 | 0.005-0.006 | 0 |
| a fresh Excel, the first load (today's way, before and after) | 1.08-1.09 | 0.96-1.01 | 0.06-0.09 | 0 |

So the look's reads are not the cost: they take milliseconds in a just-opened workbook too. The cost was the first
batch of a server process in that Excel, where the add-in was registered again (`RegisterXLL`: 0.18-0.19 s alone,
0.37-0.38 s right after a file opened) although an earlier process had loaded the same build there; a new workbook's
first answer is usually a new session's first batch. The add-in is now asked for its ping first and used as it is
when it answers. The live figure after both cuts is owed.

## Task 2: the save after the reply

Measured in a visible Excel of the builder's own (20.11.2026): with the 0.3 s delay the answer's last picture change
came 0.2-0.35 s before the save started and nothing changed during or after it. In the owner's Excel it is not yet
measured (owed: the paint, `check_followups.py`, the bench).

## Live in the owner's Excel, 21.11.2026 (controller runs, frozen checkout)

Seconds are the Excel job as the user waits for it; `scripts/check_followups.py`, every step "checks OK" unless said.

| run | route | new workbook's first answer | second answer | follow-ups | `fin_refresh` (2 answers) | save after the reply |
|---|---|---|---|---|---|---|
| 05:05, head b63e198, Excel fresh since 21:11 | add-in | 4.36 (builds and loads the add-in) | 1.09 | 1.25-1.70 | 2.60 | 0.29-0.46 |
| 05:06, same | add-in | 2.94 | 1.12 | 1.38-1.91 | 2.93 | 0.34-0.45 |
| 13:43, head 9dbe8f9, Excel just started | add-in | 6.25 (builds and loads the add-in) | 1.31 | 1.55-2.00 | 3.25 | 0.27-0.58 |
| 13:44, same | add-in | 2.72 | 1.14 | 1.42-2.52 | 3.16 | 0.24-0.33 |
| 13:45, same Excel | COM (`SV_INPROC=0`) | 4.88 | 2.12 | 2.81-4.42 | 5.31 | in the job |
| 20.11 22:19, 0.4.0, for comparison | add-in | 7.52 | 1.19 | 1.62-1.95 | 3.20 | in the job (0.25-0.31) |

The goal of 5 s or under for a new workbook's first answer is met once the add-in is loaded in that Excel (2.7-2.9 s);
the first answer after Excel starts pays the add-in's load once (4.4-6.3 s).

**The bench** (`scripts/bench_answers.py --rounds 3`, 05:16, head b63e198): 27 of 27 ok; the Excel job 1.2 / 1.3 / 1.3 s
without a lines sheet and 2.9 / 2.9 / 2.5 s with one, the saves 0.3-0.6 s after the reply. It ran in a second Excel
the engine had started for itself (the owner's was frozen, see below), a brand-new instance: not comparable one to
one with 0.4.0's 1.3 / 1.9 / 2.0 s and 3.6 / 3.1 / 3.6 s in the owner's main Excel.

**The freeze of 05:08.** After the two add-in runs of 05:05-05:07 the same script ran with `SV_INPROC=0` in the same
Excel: new answer, second answer and six follow-ups passed (2.5-3.1 s); then, between the job that reads the stored
specs for `fin_refresh` and the first redraw job, Excel stopped answering for good (no CPU use, no dialog). The
script's job thread was found releasing a COM object in a call that never returned; nothing of the add-in's was
pending (its last save record: ok, 05:06:50). The controller ended the Excel process; it held only the script's test
workbook. That Excel instance had started at 21:11:31 the evening before, in the middle of a think-cell update
(Windows log: MsiInstaller 21:11:57-21:12:55, "machine restart is required"), and had shown Excel's "problems with
the 'thinkcell.addin' add-in" dialog at its start. The cause was not found: about 115 runs of the sequence in hidden
Excels (a stand-in, then the engine's real job path; with and without the add-in; the old and the new code) did not
freeze. What changed because of it is in the CHANGELOG ("After Excel froze in the live check"): a job's COM objects
are released on its own thread before the thread ends, a busy Excel never shows up as an AttributeError, an
abandoned job no longer keeps the next one waiting in silence, and `scripts/check_insert.py --parts mixed` runs both
executors in one Excel. On that code the same sequence (add-in twice, then COM, one Excel) passed at 13:43-13:46 -
the rows above. That shows "not reproduced", not a proven fix.
