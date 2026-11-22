r"""Live check of the follow-ups (light and fast phase 3, Tasks 7 and 9): one new answer
(a test-set case, through the server's own new-question run, as scripts/bench_answers.py runs them) in ONE new
workbook of the engine's own (target 'new'), and a second answer after it (so Q1's sheets are not side by side and a
follow-up has to move a sheet to its old place), then six follow-ups on Q1 in a row through the server's own
follow-up path - server.fin_analyze(patch=..., q=1, target=<that workbook>): read_stored, find_question,
merge_patch, _prepare and the redraw job - with a harmless patch that changes the answer (the question's title, then
the period of its queries, by turns), then fin_refresh's path once (server.fin_refresh on that workbook: both
answers).

Per step it prints the seconds, the Excel job's seconds, its batches (xl/inproc: one Excel run each), COM-call
count and seconds per stage, sheet = model (OK / DIFF in n cells) and the workbook checks a fake Excel cannot prove: no 'old ...' sheet and no '..._old' table left,
the sheet order as after the second answer, the last drawn Answer active, every Sources link of the question still
naming its sheet (Excel does not rewrite a hyperlink's SubAddress, nor a HYPERLINK formula's text, when the sheet is
parked under another name), each HYPERLINK link showing its text (no click needed: the formula and what the cell
shows are compared), Excel still answering, and (Task 3) a summary sheet the script adds after the second answer:
its cell reference, table reference, cross-sheet SUM, reference to the Answer, defined name and chart series pointing
at Q1, and a cell on Q2's Answer pointing at Q1 - after every follow-up and fin_refresh each still reads as written
(no #REF!) and shows its target's value in the new answer (compared, never printed), and the workbook has no link to
another workbook. The sheet order is taken again once the summary sheet has joined it, before the first follow-up.
Never a figure from EPM: the replies are dropped. The first failure stops the run (exit 2) with its message, amounts
masked.

Usage (from the repo root; Excel open, signed in to EPM; SV_INPROC=0 runs every batch over COM):
    engine\.venv\Scripts\python.exe scripts\check_followups.py [--case b01] [--second b07]
        [--keep] [--settle 10]
At the end it says which executor ran the batches (the in-process add-in, or COM). Per step it also watches the
file's modification time (the job switches AutoSave off in its look; with the add-in, Excel saves the file right after
the reply and switches AutoSave back on then - xl/inproc/later.py -, over COM or with SV_SAVE=sync the job's save batch
does both; either way the Save is the job's own write, and a write during the batches is printed as a note, as
OneDrive stamps the file of its own too). Every step waits for the saves after the reply it caused (their seconds are
printed as "saved after"; a failed one stops the run) before it reads Workbook.Saved (True); after fin_refresh it
reads the saved file itself with openpyxl (read-only): every answer's sheets there and the last answer's row on its
Sources log.
Only the workbook this script created is touched, and closed at the end (without saving: the engine saved it after
every answer) unless --keep; the user's own workbooks and Excel are never saved or closed (the engine has no Excel of
its own to shut down).
"""
import argparse
import asyncio
import copy
import json
import os
import re
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

from smartass.core import config
from smartass.testset import runner
from smartass.xl.inproc import later

FOLLOW_UPS = 6
OTHER_PERIOD = "Jan 26..Jun 26"         # the period a follow-up switches to (and back from) by turns
COM = "excel job.com calls."            # the Excel job's COM-call counts, as server._draw_job logs them
BATCHES = "excel job.batches."          # ... and its batches (xl/inproc runs)
MOVE_IN = "excel job.draw.move in"      # a stage every answer's Excel job has
STAGES = ("open target", "draw.build", "draw.open", "draw.look", "draw.move in", "draw.finish", "draw.paint",
          "draw.close out", "save")                # draw.open: the first answer of a new workbook (its built file opened)
HYPERLINK = re.compile(r'^=HYPERLINK\("#((?:[^"]|"")*)","((?:[^"]|"")*)"\)$')


class Writes:
    """How often a file is written while a step runs: its modification time watched every 10 ms on a thread of its
    own, reported per step and between the steps (controller ruling R3's proof, 20.11.2026: with AutoSave left on the
    in-process job wrote the file only in its save batch, but the COM executor's job had a write in the middle of its
    batches, so the job keeps switching AutoSave off while it writes; and the file is also stamped 2-11 s after a
    save at times of OneDrive's own, with AutoSave on or off - so a write is a note here, not a failure)."""

    def __init__(self, path: str | None):
        self.path, self.count, self._stop = path, 0, threading.Event()
        self.at: list[float] = []                       # when (time.monotonic) each write was seen
        self._t = None

    def _mtime(self):
        try:
            return os.stat(self.path).st_mtime_ns
        except OSError:
            return None

    def _watch(self) -> None:
        last = self._mtime()
        while not self._stop.wait(0.01):
            now = self._mtime()
            if now != last:
                self.count += 1
                self.at.append(time.monotonic())
                last = now

    def __enter__(self):
        if self.path:
            self._t = threading.Thread(target=self._watch, daemon=True)
            self._t.start()
        return self

    def __exit__(self, *exc):
        self._stop.set()
        if self._t:
            self._t.join()
            time.sleep(0.05)
            self._stop.clear()


def on_disk(path: str, q_last: int, qs: list[int]) -> list[str]:
    """What the saved file holds, read with openpyxl (read-only, from the local path): every answer's sheets and the
    last answer's row on the Sources log. Problems, never a value."""
    import openpyxl
    problems = []
    book = openpyxl.load_workbook(path, read_only=True)
    try:
        names = book.sheetnames
        for q in qs:
            for n in (f"Q{q} Answer", f"Q{q} Data"):
                if n not in names:
                    problems.append(f"the saved file has no '{n}'")
        if "Sources" not in names:
            problems.append("the saved file has no Sources sheet")
        else:
            rows = [r[1] for r in book["Sources"].iter_rows(min_row=2, max_col=2, values_only=True) if r[0]]
            if not rows or rows[-1] != f"Q{q_last}":
                problems.append(f"the saved file's last Sources row is not Q{q_last}'s")
    finally:
        book.close()
    return problems


class Stop(Exception):
    """A step failed or a check did not hold: the run stops here."""


def scrub(text) -> str:
    """An error as one short line: amounts masked (any number with a separator, or of 4+ digits)."""
    s = " ".join(str(text).split())
    return re.sub(r"-?\d+(?:[ ,.]\d+)+|\d{4,}", "#", s)[:300]


def com_calls(timings: dict, prefix: str = COM) -> int | None:
    """The Excel job's COM calls (prefix BATCHES: its batches): its top-level stages summed (nested stages already
    count in their parents)."""
    tops = [v for k, v in timings.items() if k.startswith(prefix) and "." not in k[len(prefix):]]
    return int(sum(tops)) if tops else None


def executor() -> str:
    """Which executor ran the batches in this process: the add-in in which Excel, or COM (and why not)."""
    from smartass.core import config
    from smartass.xl.inproc import runner
    if not config.INPROC:
        return "COM (SV_INPROC=0)"
    ready = [f"pid {ex.pid}" for ex in runner._EXCELS.values() if ex.ready]
    other = [f"pid {ex.pid}: {ex.ready}" for ex in runner._EXCELS.values() if not ex.ready]
    return (f"in-process add-in ({', '.join(ready)})" if ready else "COM") + (f"; not loaded: {other}" if other else "")


def run_line(offset: int, tool: str) -> dict | None:
    """The last run-log line of this tool appended after byte offset."""
    try:
        with open(config.RUN_LOG, "rb") as f:
            f.seek(offset)
            data = f.read()
    except FileNotFoundError:
        return None
    found = None
    for raw in data.decode("utf-8", "replace").splitlines():
        try:
            line = json.loads(raw)
        except ValueError:
            continue                                    # a half-written or foreign line
        if line.get("tool") == tool:
            found = line
    return found


def _inspect(h, path: str, q: int) -> dict:
    """One short user-Excel job: the workbook's sheets, tables, active sheet and its Sources links, by names only."""
    from smartass import exceljob
    from smartass.xl import jobs
    t = exceljob._target(h, path, None)

    def read():
        names = [str(s.Name) for s in t.wb.Worksheets]
        tables = [str(lo.Name) for s in t.wb.Worksheets for lo in s.ListObjects]
        active = (str(t.app.ActiveSheet.Name)
                  if str(t.app.ActiveWorkbook.FullName).casefold() == str(t.wb.FullName).casefold() else None)
        links, unshown, logged = [], 0, []
        if "Sources" in names:
            ws = t.wb.Worksheets("Sources")
            for link in ws.Hyperlinks:                     # the rows an earlier version logged: real links
                r = int(link.Range.Row)
                if ws.Cells(r, 2).Value == f"Q{q}":
                    links.append(str(link.SubAddress))
            last = int(ws.Cells(ws.Rows.Count, 1).End(-4162).Row)
            if last > 1:                                   # this version's rows: HYPERLINK formulas in M and N
                formulas = ws.Range(f"B2:N{last}").Formula
                shown = ws.Range(f"B2:N{last}").Value
                for row, seen in zip(formulas, shown):
                    if row[0] != f"Q{q}":
                        continue
                    logged.append(str(seen[10] or ""))                 # Status (L): OK / DIFF, never blank
                    for f, v in zip(row[11:], seen[11:]):
                        m = HYPERLINK.match(str(f or ""))
                        if m:
                            links.append(m.group(1).replace('""', '"'))
                            unshown += v != m.group(2).replace('""', '"')
        return {"names": names, "tables": tables, "active": active, "links": links, "unshown": unshown,
                "statuses": logged,
                "ready": bool(t.app.Ready), "saved": bool(t.wb.Saved)}
    return jobs.com(read)


def _summarise(h, path: str) -> dict:
    """(Task 3) A summary sheet of the user's own with formulas pointing at Q1, and a cell on Q2's Answer pointing at
    Q1: {where: (the formula as Excel keeps it, what its value must equal: ("cell", sheet, address) |
    ("sum", sheet, range) | None)}. The defined name and the chart series are compared by their text."""
    from smartass import exceljob
    from smartass.xl import jobs
    t = exceljob._target(h, path, None)

    def make():
        wb = t.wb
        data = wb.Worksheets("Q1 Data")
        lo = data.ListObjects("Data_Q1")
        heads = [str(x) for x in lo.HeaderRowRange.Value[0]]
        k = next(i for i in range(1, len(heads) + 1)
                 if isinstance(lo.DataBodyRange.Cells(1, i).Value2, (int, float)))
        cell, column = str(lo.DataBodyRange.Cells(1, k).Address), str(lo.ListColumns(k).DataBodyRange.Address)
        answer = next(a for a in ("A1", "B1", "A2", "B2") if wb.Worksheets("Q1 Answer").Range(a).Value2 is not None)
        col = "".join("'" + ch if ch in "'[]#" else ch for ch in heads[k - 1])
        answers = [str(x.Name) for x in wb.Worksheets if re.fullmatch(r"Q\d+ Answer", str(x.Name))]
        s = wb.Worksheets.Add(None, wb.Worksheets(answers[-1]))   # where the sheet order keeps a sheet of the user's:
        s.Name = "Summary"                                        # after the answers, before the Data sheets
        wb.Names.Add("Q1First", f"='Q1 Data'!{cell}")
        refs = {"Summary!A1": (f"='Q1 Data'!{cell}", ("cell", "Q1 Data", cell)),
                "Summary!A2": (f"=SUM(Data_Q1[{col}])", ("sum", "Q1 Data", column)),
                "Summary!A3": (f"=SUM('Q1 Data'!{column})", ("sum", "Q1 Data", column)),
                "Summary!A4": (f"='Q1 Answer'!{answer}", ("cell", "Q1 Answer", answer)),
                "Summary!A5": ("=Q1First", ("cell", "Q1 Data", cell)),
                "Q2 Answer!Z1": (f"='Q1 Data'!{cell}", ("cell", "Q1 Data", cell))}
        for where, (formula, _) in refs.items():
            sheet, addr = where.split("!")
            wb.Worksheets(sheet).Range(addr).Formula2 = formula
        s.ChartObjects().Add(300, 10, 300, 200).Chart.SetSourceData(data.Range(column))
        out = {where: (str(wb.Worksheets(where.split("!")[0]).Range(where.split("!")[1]).Formula2), target)
               for where, (_, target) in refs.items()}
        out["the name Q1First"] = (str(wb.Names("Q1First").RefersTo), None)
        out["the chart series"] = (str(s.ChartObjects(1).Chart.SeriesCollection(1).Formula), None)
        return out
    return jobs.com(make)


def _pointing(h, path: str, refs: dict, q2: bool) -> list[str]:
    """(Task 3) What is wrong with the formulas _summarise made: one that no longer reads as written (#REF!), a value
    that is not its target's (compared here, never printed), a link to another workbook. q2 False: Q2's own redraw
    replaced its Answer and the cell on it."""
    from smartass import exceljob
    from smartass.xl import jobs
    t = exceljob._target(h, path, None)

    def read():
        wb, bad = t.wb, []
        for where, (formula, target) in refs.items():
            if where.startswith("Q2 ") and not q2:
                continue
            if where == "the name Q1First":
                now, value = str(wb.Names("Q1First").RefersTo), None
            elif where == "the chart series":
                now, value = str(wb.Worksheets("Summary").ChartObjects(1).Chart.SeriesCollection(1).Formula), None
            else:
                sheet, addr = where.split("!")
                now, value = str(wb.Worksheets(sheet).Range(addr).Formula2), wb.Worksheets(sheet).Range(addr).Value2
            if now != formula:
                bad.append(f"{where} no longer reads as written" + (" (#REF!)" if "#REF!" in now else ""))
                continue
            if target is None:
                continue
            kind, sheet, addr = target
            got = wb.Worksheets(sheet).Range(addr).Value2
            if kind == "sum":
                got = sum(v for (v,) in got if isinstance(v, (int, float)) and not isinstance(v, bool))
            num = all(isinstance(x, (int, float)) and not isinstance(x, bool) for x in (value, got))
            if not (abs(value - got) <= 1e-9 * max(1.0, abs(got)) if num else value == got):
                bad.append(f"{where} does not show its target's value")
        if wb.LinkSources(1):
            bad.append("the workbook links to another workbook")
        return bad
    return jobs.com(read)


def _close(h, path: str) -> None:
    """Close the workbook this script created (only that one), without saving."""
    from smartass import exceljob
    from smartass.xl import jobs
    t = exceljob._target(h, path, None)
    name = jobs.com(lambda: str(t.wb.Name))       # FullName is the https address of a OneDrive file: path and name
    if str(t.path).casefold() != path.casefold() or name.casefold() != Path(path).name.casefold():
        raise Stop(f"refusing to close {name}: it is not the workbook this script created")
    jobs.com(lambda: t.wb.Close(False))


def _patch(i: int, first: dict) -> dict:
    """Follow-up i (1-based) on the first answer's stored spec: odd ones change the title, even ones the period of
    every query (to OTHER_PERIOD, then back to the first one, by turns; a list is replaced whole by a patch)."""
    if i % 2:
        return {"question": f"{first['question']} (follow-up {i})"}
    queries = copy.deepcopy(first["queries"])
    if i % 4 == 2:
        for qq in queries:
            qq["periods"] = OTHER_PERIOD
    return {"queries": queries}


async def check(case: dict, second: dict, keep: bool, settle: float = 0.0) -> int:
    from smartass import exceljob, server
    from smartass.core import serverkit as kit
    from smartass.model.spec import AnalysisSpec

    statuses: list[str] = []
    real_redraw = server.redraw_job
    saves: list[tuple[float, float]] = []                  # (start, end) of every save batch, for the writes
    looks: list[float] = []                                # when every job's look batch began
    real_save, real_look = exceljob._save, exceljob._look

    def look(job):
        looks.append(time.monotonic())
        return real_look(job)
    exceljob._look = look

    def save(job):
        t0 = time.monotonic()
        try:
            return real_save(job)
        finally:
            saves.append((t0, time.monotonic()))
    exceljob._save = save

    notes: list[tuple[int, int]] = []                      # (layout findings, warnings) of each redraw's reply
    texts: set[str] = set()                                # their texts, amounts masked

    def redraw_job(h, pre, q, target):                     # the server's own job; its reply's check is noted
        out = real_redraw(h, pre, q, target)
        statuses.append(out["checks"]["sheet_equals_model"])
        notes.append((len(out.get("layout") or []), len(out.get("warnings") or [])))
        texts.update(scrub(x)[:120] for x in [*(out.get("layout") or []), *(out.get("warnings") or [])])
        return out
    server.redraw_job = redraw_job

    async def fin_analyze(spec=None, patch=None, q=None, target="session"):   # the names runlog keys on
        if spec is not None:
            return await server._new_question(spec, target)
        return await server.fin_analyze(patch=patch, q=q, target=target)

    async def fin_refresh(target):
        return await server.fin_refresh(target=target)
    analyze, refresh = kit.guard(fin_analyze), kit.guard(fin_refresh)
    user = kit._call_user
    path, q, order, failed, disk_check = None, None, None, None, False
    refs = None                                            # (Task 3) the summary sheet's formulas pointing at Q1
    between, last_end = Writes(None), [0.0]               # every write of the file, also between the steps

    async def step(label: str, tool: str, call, want_followup: bool, baseline: bool = False) -> None:
        nonlocal path, q, order, refs
        offset = config.RUN_LOG.stat().st_size if config.RUN_LOG.exists() else 0
        before = len(statuses)
        if path is not None and between.path is None:       # the file's writes watched from the first answer on
            between.path = path
            between.__enter__()
        if settle:
            await asyncio.sleep(settle)                    # --settle: a quiet spell before the step
        gap = [a for a in between.at if a > last_end[0]]   # writes since the step before ended
        t0 = time.monotonic()
        try:
            with Writes(path) as writes:
                out = await call()
                secs = time.monotonic() - t0
                saved = await asyncio.to_thread(later.wait, 60)    # Excel's save after the reply, before any read
        except Exception as e:   # noqa: BLE001
            raise Stop(f"{label} failed: {scrub(e)}") from None
        last_end[0] = time.monotonic()
        if gap:
            print(f"{'':<14} the file was written {len(gap)} time(s) between the steps (at "
                  f"{', '.join(f'-{t0 - a:.2f}' for a in gap)} s before this one)", flush=True)
        line = run_line(offset, tool) or {}
        timings = line.get("timings") or {}
        if tool == "fin_refresh":
            done = (out or {}).get("refreshed") or []
            if not done or not all(d.get("ok") for d in done):
                raise Stop(f"{label}: " + "; ".join(scrub(d.get("error")) for d in done if not d.get("ok")))
            if len(statuses) - before != len(done):
                raise Stop(f"{label}: {len(done)} answers refreshed, {len(statuses) - before} by redraw_job")
            status = "; ".join(dict.fromkeys(statuses[before:]))          # 'OK' when every answer is OK
            drawn = max(d["q"] for d in done)
            found = tuple(sum(x) for x in zip(*notes[before:])) if notes[before:] else (0, 0)
        else:
            status = out["checks"]["sheet_equals_model"]
            if path is None:
                path, q = out["workbook"], out["q"]
            drawn = out["q"]
            found = (len(out.get("layout") or []), len(out.get("warnings") or []))
            texts.update(scrub(x)[:120] for x in [*(out.get("layout") or []), *(out.get("warnings") or [])])
            if want_followup and (out.get("followup") is not True or len(statuses) == before):
                raise Stop(f"{label}: the reply did not come from redraw_job")
        if MOVE_IN not in timings:
            raise Stop(f"{label}: no '{MOVE_IN}' stage in the run log - not this engine's Excel job")
        seen = await user(lambda h: _inspect(h, path, q), what="check_followups: inspect")
        if baseline:
            order = seen["names"]
        problems = [f"the save after the reply failed: {scrub(s.get('error'))}" for s in saved if not s["ok"]]
        if refs is not None:
            problems += await user(lambda h: _pointing(h, path, refs, tool != "fin_refresh"),
                                   what="check_followups: formulas")
        if status != "OK":
            problems.append(f"sheet = model {status}")
        if [n for n in seen["names"] if n.startswith("old ")]:
            problems.append("an 'old ...' sheet is left")
        if [n for n in seen["tables"] if re.search(r"_old\d*$", n)]:
            problems.append("an '..._old' table is left")
        if seen["names"] != order:
            problems.append("the sheet order changed")
        if seen["active"] != f"Q{drawn} Answer":
            problems.append(f"the active sheet is {seen['active']!r}, not 'Q{drawn} Answer'")
        stale = [s for s in seen["links"] if s.strip("'").split("'!")[0].replace("''", "'") not in seen["names"]]
        if stale:
            problems.append(f"{len(stale)} Sources link(s) of Q{q} name a sheet that is gone")
        if seen["unshown"]:
            problems.append(f"{seen['unshown']} Sources link(s) of Q{q} do not show their text")
        if [s for s in seen["statuses"] if s not in ("OK", "DIFF")]:
            problems.append(f"a Sources row of Q{q} has no OK/DIFF status")
        if not seen["ready"]:
            problems.append("Excel is not ready")
        if not seen["saved"]:
            problems.append("Workbook.Saved is False after the job")
        windows = [(a, b) for a, b in saves if a >= t0]
        jobs_here = list(zip([a for a in looks if a >= t0], windows))     # (look began, its save batch)
        during = [a for a in writes.at if any(lk <= a < s - 0.05 for lk, (s, _) in jobs_here)]
        if writes.path and (during or not writes.count):     # reported, not a failure: OneDrive stamps the file
            when = ", ".join(f"+{a - t0:.2f}" for a in writes.at)     # at times of its own (see Writes)
            spans = ", ".join(f"+{lk - t0:.2f}..+{s - t0:.2f}" for lk, (s, _) in jobs_here)
            print(f"{'':<14} note: the file was written during the job's batches ({len(during)} of {writes.count} "
                  f"writes, at {when} s; the batches from the look to the save at {spans} s)", flush=True)
        com = com_calls(timings)
        runs = com_calls(timings, BATCHES)
        stages = "  ".join(f"{s.split('.')[-1]} {timings.get('excel job.' + s, 0):.2f}" for s in STAGES)
        counts = "/".join(f"{k} {seen['statuses'].count(k)}" for k in ("OK", "DIFF"))
        after = (f"{sum(s['seconds'] for s in saved if s.get('seconds') is not None):4.2f} s"
                 f"{'' if len(saved) < 2 else f' ({len(saved)} saves)'}" if saved else "-")
        print(f"{label:<14} {secs:6.1f} s  writes {'-' if not writes.path else writes.count}  excel job "
              f"{timings.get('excel job._total', 0):5.2f} s  saved after {after}  batches "
              f"{'-' if runs is None else runs:>2}  COM calls {'-' if com is None else com:>4}  [{stages}]  sheet = model {status}  Sources links "
              f"{len(seen['links'])}  status {counts}  layout {found[0]} warnings {found[1]}  "
              f"{'checks OK' if not problems else 'DIFF: ' + '; '.join(problems)}", flush=True)
        inside = "  ".join(
            f"{s.split('.')[-1]} [" + " ".join(f"{k.rsplit('.excel.', 1)[1]} {v:.2f}" for k, v in timings.items()
                                               if k.startswith(f"excel job.{s}.excel.")) + "]"
            for s in ("draw.move in", "draw.finish", "draw.close out"))
        print(f"{'':<14} inside Excel: {inside}", flush=True)
        if problems:
            raise Stop(f"{label}: " + "; ".join(problems))

    print(f"check_followups {datetime.now():%Y-%m-%d %H:%M}, cases {case['id']} and "
          f"{second['id']}", flush=True)
    try:
        spec = AnalysisSpec.model_validate(case["spec"])
        await step("new answer", "fin_analyze", lambda: analyze(spec=spec, target="new"), False, baseline=True)
        spec2 = AnalysisSpec.model_validate(second["spec"])
        await step("second answer", "fin_analyze", lambda: analyze(spec=spec2, target=path), False, baseline=True)
        refs = await user(lambda h: _summarise(h, path), what="check_followups: summary")
        order = (await user(lambda h: _inspect(h, path, q), what="check_followups: order"))["names"]   # with it
        stored, _, _ = await user(lambda h: exceljob.read_stored(h, path), what="check_followups: stored")
        first = stored[q]["spec"]["v2"]
        for i in range(1, FOLLOW_UPS + 1):
            patch = _patch(i, first)
            await step(f"follow-up {i}", "fin_analyze", lambda p=patch: analyze(patch=p, q=q, target=path), True)
        await step("fin_refresh", "fin_refresh", lambda: refresh(target=path), True)
        disk_check = True                       # read once Excel has closed it (Excel holds it open)
    except Stop as e:
        failed = str(e)
    finally:
        server.redraw_job = real_redraw
        exceljob._save, exceljob._look = real_save, real_look
        if between.path is not None:
            between.__exit__()
        if path is not None and not keep:
            try:
                await user(lambda h: _close(h, path), what="check_followups: close")
            except Exception as e:   # noqa: BLE001
                print(f"could not close the workbook this script created: {scrub(e)}", flush=True)
                disk_check = False
            if disk_check:
                try:
                    disk = on_disk(path, 2, [1, 2])     # fin_refresh redraws Q1, then Q2
                except OSError as e:
                    disk = [f"the saved file could not be read ({scrub(e)})"]
                print(f"{'saved file':<14} {'checks OK' if not disk else 'DIFF: ' + '; '.join(disk)}", flush=True)
                if disk and not failed:
                    failed = "the saved file: " + "; ".join(disk)
    for x in sorted(texts):
        print(f"  reply note: {x}", flush=True)
    print(f"executor: {executor()}", flush=True)
    if failed:
        print(f"\nSTOPPED: {failed}", flush=True)
        return 2
    print("\nall steps OK", flush=True)
    return 0


def main(argv=None) -> int:
    cases = {c["id"]: c for c in runner.load_cases()}
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--case", default="b01", help=f"the test-set case of the new answer (known: {' '.join(cases)})")
    ap.add_argument("--second", default="b07", help="the case of the second answer, drawn before the follow-ups")
    ap.add_argument("--keep", action="store_true", help="leave the workbook this script created open")
    ap.add_argument("--settle", type=float, default=0.0, help="seconds to wait before each step (writes of the file "
                    "that come late after a save then fall between the steps)")
    args = ap.parse_args(argv)
    for c in (args.case, args.second):
        if c not in cases:
            ap.error(f"unknown case id {c} (known: {' '.join(cases)})")
    if not cases[args.case]["spec"].get("queries"):
        ap.error(f"case {args.case} has no queries: a period follow-up needs one")
    sys.stdout.reconfigure(encoding="utf-8")
    return asyncio.run(check(cases[args.case], cases[args.second], args.keep, args.settle))


if __name__ == "__main__":
    sys.exit(main())
