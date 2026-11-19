r"""Live check of the file renderer's follow-ups (light and fast phase 3, Task 7). With SV_RENDER=file: one new answer
(a test-set case, through the server's own new-question run, as scripts/bench_answers.py runs them) in ONE new
workbook of the engine's own (target 'new'), and a second answer after it (so Q1's sheets are not side by side and a
follow-up has to move a sheet to its old place), then six follow-ups on Q1 in a row through the server's own
follow-up path - server.fin_analyze(patch=..., q=1, target=<that workbook>): read_stored, find_question,
merge_patch, _prepare and the redraw job - with a harmless patch that changes the answer (the question's title, then
the period of its queries, by turns), then fin_refresh's path once (server.fin_refresh on that workbook: both
answers).

Per step it prints the seconds, the Excel job's seconds, COM-call count and seconds per stage, sheet = model (OK /
DIFF in n cells) and the workbook checks a fake Excel cannot prove: no 'old ...' sheet and no '..._old' table left,
the sheet order as after the second answer, the last drawn Answer active, every Sources link of the question still
naming its sheet (Excel does not rewrite a hyperlink's SubAddress, nor a HYPERLINK formula's text, when the sheet is
parked under another name), each HYPERLINK link showing its text (no click needed: the formula and what the cell
shows are compared), and Excel still answering. Never a figure from
EPM: the replies are dropped. The first failure stops the run (exit 2) with its message, amounts masked.

Usage (from the repo root; Excel open, signed in to EPM):
    $env:SV_RENDER = "file"; engine\.venv\Scripts\python.exe scripts\check_followups.py [--case b01] [--second b07]
        [--keep]
Only the workbook this script created is touched, and closed at the end (without saving: the engine saved it after
every answer) unless --keep; the user's own workbooks and Excel are never saved or closed. At the end only the
engine's own hidden Excel and browser are shut down, as the server's lifespan does.
"""
import argparse
import asyncio
import copy
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path

from smartass.core import config
from smartass.testset import runner

FOLLOW_UPS = 6
OTHER_PERIOD = "Jan 26..Jun 26"         # the period a follow-up switches to (and back from) by turns
COM = "excel job.com calls."            # the Excel job's COM-call counts, as server._draw_job logs them
MOVE_IN = "excel job.draw.move in"      # a stage only the file renderer has
STAGES = ("draw.build", "draw.move in", "draw.finish", "draw.read back", "draw.sources", "draw.arrange", "save")
HYPERLINK = re.compile(r'^=HYPERLINK\("#((?:[^"]|"")*)","((?:[^"]|"")*)"\)$')


class Stop(Exception):
    """A step failed or a check did not hold: the run stops here."""


def scrub(text) -> str:
    """An error as one short line: amounts masked (any number with a separator, or of 4+ digits)."""
    s = " ".join(str(text).split())
    return re.sub(r"-?\d+(?:[ ,.]\d+)+|\d{4,}", "#", s)[:300]


def com_calls(timings: dict) -> int | None:
    """The Excel job's COM calls: its top-level stages summed (nested stages already count in their parents)."""
    tops = [v for k, v in timings.items() if k.startswith(COM) and "." not in k[len(COM):]]
    return int(sum(tops)) if tops else None


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
    from smartass import pipeline
    from smartass.xl import jobs
    t = pipeline._target(h, path, None)

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
                "ready": bool(t.app.Ready)}
    return jobs.com(read)


def _close(h, path: str) -> None:
    """Close the workbook this script created (only that one), without saving."""
    from smartass import pipeline
    from smartass.xl import jobs
    t = pipeline._target(h, path, None)
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


async def check(case: dict, second: dict, keep: bool) -> int:
    from smartass import pipeline, server
    from smartass.core import serverkit as kit
    from smartass.model.spec import AnalysisSpec

    statuses: list[str] = []
    real_redraw = server.file_redraw_job

    notes: list[tuple[int, int]] = []                      # (layout findings, warnings) of each redraw's reply
    texts: set[str] = set()                                # their texts, amounts masked

    def file_redraw_job(h, pre, q, owned, target):       # the server's own job; its reply's check is noted
        out = real_redraw(h, pre, q, owned, target)
        statuses.append(out["checks"]["sheet_equals_model"])
        notes.append((len(out.get("layout") or []), len(out.get("warnings") or [])))
        texts.update(scrub(x)[:120] for x in [*(out.get("layout") or []), *(out.get("warnings") or [])])
        return out
    server.file_redraw_job = file_redraw_job

    async def fin_analyze(spec=None, patch=None, q=None, target="session"):   # the names runlog keys on
        if spec is not None:
            return await server._new_question(spec, target)
        return await server.fin_analyze(patch=patch, q=q, target=target)

    async def fin_refresh(target):
        return await server.fin_refresh(target=target)
    analyze, refresh = kit.guard(fin_analyze), kit.guard(fin_refresh)
    user = kit._call_user
    path, q, order, failed = None, None, None, None

    async def step(label: str, tool: str, call, want_followup: bool, baseline: bool = False) -> None:
        nonlocal path, q, order
        offset = config.RUN_LOG.stat().st_size if config.RUN_LOG.exists() else 0
        before = len(statuses)
        t0 = time.monotonic()
        try:
            out = await call()
        except Exception as e:   # noqa: BLE001
            raise Stop(f"{label} failed: {scrub(e)}") from None
        secs = time.monotonic() - t0
        line = run_line(offset, tool) or {}
        timings = line.get("timings") or {}
        if tool == "fin_refresh":
            done = (out or {}).get("refreshed") or []
            if not done or not all(d.get("ok") for d in done):
                raise Stop(f"{label}: " + "; ".join(scrub(d.get("error")) for d in done if not d.get("ok")))
            if len(statuses) - before != len(done):
                raise Stop(f"{label}: {len(done)} answers refreshed, {len(statuses) - before} by file_redraw_job")
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
                raise Stop(f"{label}: the reply did not come from file_redraw_job")
        if MOVE_IN not in timings:
            raise Stop(f"{label}: no '{MOVE_IN}' stage in the run log - not the file renderer")
        seen = await user(lambda h: _inspect(h, path, q), what="check_followups: inspect")
        if baseline:
            order = seen["names"]
        problems = []
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
        com = com_calls(timings)
        stages = "  ".join(f"{s.split('.')[-1]} {timings.get('excel job.' + s, 0):.2f}" for s in STAGES)
        counts = "/".join(f"{k} {seen['statuses'].count(k)}" for k in ("OK", "DIFF"))
        print(f"{label:<14} {secs:6.1f} s  excel job {timings.get('excel job._total', 0):5.2f} s  COM calls "
              f"{'-' if com is None else com:>4}  [{stages}]  sheet = model {status}  Sources links "
              f"{len(seen['links'])}  status {counts}  layout {found[0]} warnings {found[1]}  "
              f"{'checks OK' if not problems else 'DIFF: ' + '; '.join(problems)}", flush=True)
        if problems:
            raise Stop(f"{label}: " + "; ".join(problems))

    print(f"check_followups {datetime.now():%Y-%m-%d %H:%M}, render {config.RENDER}, cases {case['id']} and "
          f"{second['id']}", flush=True)
    try:
        spec = AnalysisSpec.model_validate(case["spec"])
        await step("new answer", "fin_analyze", lambda: analyze(spec=spec, target="new"), False, baseline=True)
        spec2 = AnalysisSpec.model_validate(second["spec"])
        await step("second answer", "fin_analyze", lambda: analyze(spec=spec2, target=path), False, baseline=True)
        stored, _, _ = await user(lambda h: pipeline.read_stored(h, path), what="check_followups: stored")
        first = stored[q]["spec"]["v2"]
        for i in range(1, FOLLOW_UPS + 1):
            patch = _patch(i, first)
            await step(f"follow-up {i}", "fin_analyze", lambda p=patch: analyze(patch=p, q=q, target=path), True)
        await step("fin_refresh", "fin_refresh", lambda: refresh(target=path), True)
    except Stop as e:
        failed = str(e)
    finally:
        server.file_redraw_job = real_redraw
        if path is not None and not keep:
            try:
                await user(lambda h: _close(h, path), what="check_followups: close")
            except Exception as e:   # noqa: BLE001
                print(f"could not close the workbook this script created: {scrub(e)}", flush=True)
        _shutdown(kit)
    for x in sorted(texts):
        print(f"  reply note: {x}", flush=True)
    if failed:
        print(f"\nSTOPPED: {failed}", flush=True)
        return 2
    print("\nall steps OK", flush=True)
    return 0


def _shutdown(kit) -> None:
    """As the server's lifespan ends: the engine's own hidden Excel and browser only."""
    try:
        if kit._host is not None:
            kit._host.shutdown()
    finally:
        from smartass.epm import drill_browser
        drill_browser.shutdown_shared()


def main(argv=None) -> int:
    cases = {c["id"]: c for c in runner.load_cases()}
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--case", default="b01", help=f"the test-set case of the new answer (known: {' '.join(cases)})")
    ap.add_argument("--second", default="b07", help="the case of the second answer, drawn before the follow-ups")
    ap.add_argument("--keep", action="store_true", help="leave the workbook this script created open")
    args = ap.parse_args(argv)
    for c in (args.case, args.second):
        if c not in cases:
            ap.error(f"unknown case id {c} (known: {' '.join(cases)})")
    if config.RENDER != "file":
        ap.error(f"SV_RENDER is {config.RENDER!r}: set it to 'file' (this checks the file renderer's follow-ups)")
    if not cases[args.case]["spec"].get("queries"):
        ap.error(f"case {args.case} has no queries: a period follow-up needs one")
    sys.stdout.reconfigure(encoding="utf-8")
    return asyncio.run(check(cases[args.case], cases[args.second], args.keep))


if __name__ == "__main__":
    sys.exit(main())
