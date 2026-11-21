r"""Baseline timings of today's engine (light and fast, phase 0): every test-set case answered live through the MCP
server's own new-question run, under the tool name fin_analyze so core/runlog.py logs it, and the run-log lines it
caused (config.RUN_LOG) read back into a markdown table of medians. Only times, COM-call counts and stage names are
printed: the replies carry figures from EPM and are never printed or kept.

Usage (from the repo root; Excel open, signed in to EPM):
    engine\.venv\Scripts\python.exe scripts\bench_answers.py --rounds 3 [--case b01 ...]
The per-round table has a column per stage of the Excel job (exceljob.excel_job: open target, build, open, look,
move in, finish, close out, save; 'open' only for the first answer of a new workbook, whose built file is opened, and
its build comes before the look) and its "runs": the batches it ran (xl/inproc: one Excel run each); a stage the run
did not have shows "-". SV_INPROC=0 runs every batch over COM. With the add-in, our file is saved by Excel after the
reply (xl/inproc/later.py): "excel job s" and "save s" are what the user waits for (the save step only asks for it),
and "saved after s" is that save's own seconds, from its outcome - the bench waits for it after every answer, so the
next one starts with nothing pending ("-": saved in the job, SV_SAVE=sync or over COM).

One warm-up answer (the first case, target 'new') creates ONE new workbook in the analyser folder; every measured
round then lands in that session workbook (target 'session'). Rounds run outer and cases inner, so the growing
workbook weighs on every case alike. The warm-up has its own row and is not in the medians. A failed run is listed
with its error (amounts masked), never hidden; when the warm-up and the first measured run both fail the bench stops
(the cause is the environment, not a case). A case that guard_chart refuses (no chart block) runs the same steps as
server._new_question minus that one guard, still logged as fin_analyze, and its row is marked '(no chart block)';
the case files are not changed. Never saves or closes a workbook of the user's; at the end only the
engine's own hidden Excel is shut down, as the server's lifespan does.
"""
import argparse
import asyncio
import json
import re
import statistics
import sys
import time
from datetime import datetime
from pathlib import Path

from smartass import __version__
from smartass.core import config
from smartass.testset import runner
from smartass.xl.inproc import later

COM = "com calls."                  # keys with this in them are COM-call counts, not seconds
BATCHES = "batches."                # ... and these batch counts (one Excel run each)
JOB = "excel job."                  # the Excel job's Timer summary, as server._draw_job adds it
PREVIOUS = JOB + "previous save"    # the save the answer before left to Excel (its own row has it as "saved after")
COLUMNS = (("_total", "_total s"), ("prepare.fetch", "prepare.fetch s"), ("prepare.drill", "prepare.drill s"),
           ("excel job._total", "excel job s"))


def com_calls(timings: dict, what: str = COM) -> int | None:
    """The run's COM calls (what BATCHES: its batches): the counts of the Excel job's top-level stages (open target,
    draw, save) summed; nested stages already count in their parents. None when the job logged no counts (it did not
    run)."""
    pre = JOB + what
    tops = [v for k, v in timings.items() if k.startswith(pre) and "." not in k[len(pre):]]
    return sum(tops) if tops else None


def scrub(text) -> str:
    """An error as one short table cell: amounts masked (any number with a separator, or of 4+ digits), no pipes,
    in a code span (the advice in it has <placeholders>)."""
    s = " ".join(str(text).split())
    s = re.sub(r"-?\d+(?:[ ,.]\d+)+|\d{4,}", "#", s).replace("|", "/").replace("`", "'")
    return f"`{s[:120]}...`" if len(s) > 120 else f"`{s}`"


def read_lines(path: Path, offset: int, start: datetime) -> list[dict]:
    """The fin_analyze lines appended to the run log after byte offset, with ts at or after start."""
    try:
        with open(path, "rb") as f:
            f.seek(offset)
            data = f.read()
    except FileNotFoundError:
        return []
    out = []
    for raw in data.decode("utf-8", "replace").splitlines():
        try:
            line = json.loads(raw)
            if line.get("tool") == "fin_analyze" and datetime.fromisoformat(line["ts"]) >= start:
                out.append(line)
        except (ValueError, KeyError, TypeError):
            continue                                    # a half-written or foreign line
    return out


def _median(values) -> float | None:
    values = [v for v in values if v is not None]
    return statistics.median(values) if values else None


def _secs(v) -> str:
    return "-" if v is None else f"{v:.1f}"


def _count(v) -> str:
    return "-" if v is None else f"{v:g}"


def _row(label: str, runs: list[dict], errors: bool = True) -> str:
    ok = [r for r in runs if r["ok"]]
    cells = [label, f"{len(ok)}/{len(runs)}"]
    cells += [_secs(_median(r["timings"].get(k) for r in ok)) for k, _ in COLUMNS]
    cells.append(_secs(_median(r.get("saved") for r in ok)))
    cells.append(_count(_median(com_calls(r["timings"]) for r in ok)))
    cells.append(_count(_median(com_calls(r["timings"], BATCHES) for r in ok)))
    failed = [r["error"] for r in runs if not r["ok"]]
    cells.append("; ".join(dict.fromkeys(failed)) if errors else (f"{len(failed)} failed" if failed else ""))
    return "| " + " | ".join(cells) + " |"


def table(warm: dict | None, runs: list[dict]) -> str:
    """Per case the medians over its ok rounds; then the medians over every ok run of every case; then the warm-up."""
    head = ["case", "ok runs"] + [title for _, title in COLUMNS] + ["saved after s", "COM calls", "runs", "errors"]
    rows = ["| " + " | ".join(head) + " |", "|" + "---|" * len(head)]
    by_case: dict[str, list[dict]] = {}
    for r in runs:
        by_case.setdefault(r["case"], []).append(r)
    rows += [_row(case, rs) for case, rs in by_case.items()]
    rows.append(_row("**all cases** (median of all ok runs)", runs, errors=False))
    if warm is not None:
        rows.append(_row(f"warm-up {warm['case']} (target 'new', not in the medians)", [warm]))
    return "\n".join(rows)


ROUND_COLUMNS = (("_total", "_total s"), ("excel job._total", "excel job s"), ("excel job.open target", "open target s"),
                 ("excel job.draw", "draw s"), ("excel job.draw.build", "draw.build s"),
                 ("excel job.draw.open", "draw.open s"), ("excel job.draw.look", "draw.look s"),
                 ("excel job.draw.move in", "draw.move in s"), ("excel job.draw.finish", "draw.finish s"),
                 ("excel job.draw.paint", "draw.paint s"), ("excel job.draw.close out", "draw.close out s"),
                 ("excel job.save", "save s"),
                 ("prepare", "prepare s"))


def _has_lines(timings: dict) -> bool:
    """A run with a lines block: a drill timed in drills.*."""
    return any(k.startswith("drills.") for k in timings)


def round_table(runs: list[dict], lines: bool = False) -> str:
    """Per round the medians over the ok runs without a lines block (no drill; lines: those with one), to show how
    the answers change as the one workbook grows; the COM calls and batches show whether the work itself changed."""
    head = ["round", "answers"] + [title for _, title in ROUND_COLUMNS] + ["saved after s", "COM calls", "runs"]
    rows = ["| " + " | ".join(head) + " |", "|" + "---|" * len(head)]
    for n in sorted({r["round"] for r in runs}):
        rs = [r for r in runs if r["round"] == n and r["ok"] and _has_lines(r["timings"]) == lines]
        cells = [str(n), str(len(rs))] + [_secs(_median(r["timings"].get(k) for r in rs)) for k, _ in ROUND_COLUMNS]
        cells.append(_secs(_median(r.get("saved") for r in rs)))
        cells.append(_count(_median(com_calls(r["timings"]) for r in rs)))
        cells.append(_count(_median(com_calls(r["timings"], BATCHES) for r in rs)))
        rows.append("| " + " | ".join(cells) + " |")
    return "\n".join(rows)


def stage_table(runs: list[dict]) -> str:
    """Every timed key over the ok runs, by its share of the bench's seconds. Stages nest (draw holds draw.move in),
    so the shares of a stage and its parts overlap; drills.* are summed over concurrent drills."""
    ok = [r for r in runs if r["ok"]]
    whole = sum(r["timings"].get("_total", 0) for r in ok) or 1
    vals: dict[str, list[float]] = {}
    for r in ok:
        for k, v in r["timings"].items():
            if COM not in k and BATCHES not in k and k != PREVIOUS:    # the case before's save: not this case's time
                vals.setdefault(k, []).append(v)
    rows = ["| stage | runs | median s | sum s | share of _total |", "|---|---|---|---|---|"]
    for k, vs in sorted(vals.items(), key=lambda kv: -sum(kv[1])):
        rows.append(f"| {k} | {len(vs)} | {statistics.median(vs):.2f} | {sum(vs):.1f} | {sum(vs) / whole:.0%} |")
    return "\n".join(rows)


def _has_chart(spec) -> bool:
    """False when guard_chart refuses the spec (a new question without a chart block)."""
    from smartass.model.validate import GuardError, guard_chart
    try:
        guard_chart(spec)
    except GuardError:
        return False
    return True


async def _run(tool, case: dict, target: str, start: datetime, round_no: int = 0) -> dict:
    """One answer; its run-log line (the reply itself is dropped: it carries figures)."""
    from smartass.model.spec import AnalysisSpec
    path = config.RUN_LOG
    offset = path.stat().st_size if path.exists() else 0
    t0 = time.monotonic()
    error = None
    label = case["id"]
    status = None
    try:
        spec = AnalysisSpec.model_validate(case["spec"])
        chart = _has_chart(spec)
        label = case["id"] if chart else f"{case['id']} (no chart block)"
        out = await tool(spec, target, chart=chart)
        status = ((out or {}).get("checks") or {}).get("sheet_equals_model")     # 'OK' / 'DIFF in n cells' only
    except Exception as e:   # noqa: BLE001 - a failed case is listed with its error, the bench goes on
        error = e
    wall = time.monotonic() - t0                       # the answer as the user waits for it
    saves = await asyncio.to_thread(later.wait, 60)     # the save after the reply: its own seconds, then on
    lines = read_lines(path, offset, start)
    question = case["spec"].get("question")
    line = next((x for x in reversed(lines) if x.get("question") in (question, None)), None)
    rec = {"case": label, "round": round_no, "wall": wall,
           "timings": (line or {}).get("timings") or {}, "status": status,
           "saved": sum(s["seconds"] for s in saves if s.get("seconds") is not None) if saves else None}
    failed = [s for s in saves if not s["ok"]]
    if error is not None:
        rec.update(ok=False, error=scrub(error))
    elif failed:
        rec.update(ok=False, error=scrub(f"the save after the reply failed: {failed[0].get('error')}"))
    elif line is None:
        rec.update(ok=False, error="no run-log line")
    else:
        rec.update(ok=bool(line.get("ok")), error=None if line.get("ok") else scrub(line.get("error")))
    return rec


def _progress(label: str, rec: dict) -> None:
    state = f"ok, sheet = model {rec.get('status')}" if rec["ok"] else f"FAILED: {rec['error']}"
    if rec.get("saved") is not None:
        state += f", saved after the reply in {rec['saved']:.2f}s"
    print(f"<!-- {label} {rec['case']}: {rec['wall']:.1f}s {state} -->", flush=True)


def _excel_build(h) -> str:
    """Application.Version and Build of the Excel the session workbook (the answers) is in."""
    from smartass import exceljob
    from smartass.xl import jobs
    t = exceljob._target(h, "session", None)
    return jobs.com(lambda: f"{t.app.Version} (build {t.app.Build})")


async def _build(kit) -> str:
    from smartass.core import session
    if not session.workbook():
        return "unknown (no answer workbook)"      # 'session' would make a new one
    try:
        return await kit._call_user(_excel_build, what="bench: Excel build")
    except Exception as e:   # noqa: BLE001 - the table is printed already
        return f"unknown ({scrub(e)})"


def _shutdown(kit) -> None:
    """As the server's lifespan ends: the engine's own hidden Excel only."""
    if kit._host is not None:
        kit._host.shutdown()


async def bench(cases: list[dict], rounds: int) -> int:
    from smartass import server
    from smartass.core import serverkit as kit

    async def fin_analyze(spec, target, chart=True):    # the name runlog.begin keys on (serverkit.RUN_LOGGED)
        if chart:
            return await server._new_question(spec, target)
        # server._new_question's own steps minus guard_chart (the case has no chart block; the draw adds one)
        server.guard_key_points(spec)
        pre, notes = await server._prepare(spec.model_dump(), server._provider(), spec, target, complete=True)
        return server._with_notes(await server._draw_job(lambda h: server.excel_job(h, pre, target), "fin_analyze",
                                                         spec.question), notes)

    tool = kit.guard(fin_analyze)
    start = datetime.now().astimezone().replace(microsecond=0)
    runs: list[dict] = []
    blocked = False
    try:
        warm = await _run(tool, cases[0], "new", start)
        _progress("warm-up", warm)
        for n in range(1, rounds + 1):
            for case in cases:
                rec = await _run(tool, case, "session", start, n)
                runs.append(rec)
                _progress(f"round {n}", rec)
                blocked = len(runs) == 1 and not warm["ok"] and not rec["ok"]
                if blocked:
                    break
            if blocked:
                break
        print(f"\n## Baseline v2 answers, {start:%Y-%m-%d}\n")
        print(f"Engine {__version__}; {rounds} rounds of {len(cases)} cases; medians over the "
              "ok rounds.\n")
        print(table(warm, runs))
        print("\n### Per round (ok runs without a lines block)\n")
        print(round_table(runs))
        print("\n### Per round (ok runs with a lines block)\n")
        print(round_table(runs, lines=True))
        print("\n### Stages\n")
        print(stage_table(runs), flush=True)
        print(f"\nExcel: {await _build(kit)}", flush=True)
    finally:
        _shutdown(kit)
    if blocked:
        print("\nSTOPPED: the warm-up and the first measured run both failed (see the errors above).")
        return 2
    return 0


def main(argv=None) -> int:
    cases = runner.load_cases()
    known = [c["id"] for c in cases]
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--rounds", type=int, default=3, help="measured rounds per case (default 3)")
    ap.add_argument("--case", nargs="+", default=None, help=f"case ids (default all: {' '.join(known)})")
    args = ap.parse_args(argv)
    if args.rounds < 1:
        ap.error("--rounds must be 1 or more")
    unknown = sorted(set(args.case or ()) - set(known))
    if unknown:
        ap.error(f"unknown case ids: {' '.join(unknown)} (known: {' '.join(known)})")
    if args.case:
        cases = [c for c in cases if c["id"] in args.case]
    sys.stdout.reconfigure(encoding="utf-8")
    return asyncio.run(bench(cases, args.rounds))


if __name__ == "__main__":
    sys.exit(main())
