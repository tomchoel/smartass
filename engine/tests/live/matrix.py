"""Live audit matrix: runs the registered scenarios (scenarios.py) over stdio MCP and scores them.
Usage (from engine): uv run python tests\\live\\matrix.py [--only TAGS] [--repeat N] [--fresh] [--keep] [--out NAME]
Writes docs\\audit\\<date>-<out>.json and .md (or <date>-<out>-HHMM if that exists), with a before column from the
newest earlier *-matrix*.json (a same-day file counts).
Exit code 1 if any scenario FAILs."""
from __future__ import annotations

import argparse
import json
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

import helpers
import scenarios
from mcpclient import ENGINE, McpClient

from smartview_mcp import __version__, config
from smartview_mcp.auditlib import error_class, merge_before, verdict

AUDIT = ENGINE.parent / "docs" / "audit"


def parse(argv: list[str] | None = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Live audit matrix")
    ap.add_argument("--only", default="", help="comma-separated tags (or scenario ids)")
    ap.add_argument("--repeat", type=int, default=1, help="run each selected scenario N times")
    ap.add_argument("--fresh", action="store_true", help="run only the fresh-user scenarios")
    ap.add_argument("--keep", action="store_true", help="leave the scratch workbooks and _matrix folder")
    ap.add_argument("--out", default="matrix", help="output name: docs\\audit\\<date>-<out>.json/.md")
    return ap.parse_args(argv)


def select(only: str, fresh: bool) -> list[scenarios.Scenario]:
    want = {t.strip() for t in only.split(",") if t.strip()}
    return [s for s in scenarios.REGISTRY.values()
            if s.fresh == fresh and (not want or want & set(s.tags) or s.id in want)]


def attempt(sc: scenarios.Scenario, ctx: scenarios.Ctx) -> tuple[bool, float, str | None, dict]:
    t0 = time.monotonic()
    try:
        out = sc.fn(ctx) or {}
        return True, time.monotonic() - t0, None, out if isinstance(out, dict) else {}
    except Exception as e:
        return False, time.monotonic() - t0, f"{type(e).__name__}: {e}", {}


def run_one(sc: scenarios.Scenario, ctx: scenarios.Ctx, run: int) -> dict:
    ok, secs, err, out = attempt(sc, ctx)
    flaky = False
    if not ok and err.startswith("Limit:"):        # an injection that cannot run safely here: no re-run
        return {"id": sc.id, "tags": sc.tags, "result": "LIMIT", "secs": round(secs, 1), "error": err,
                "error_class": error_class(err), "notes": "", "limit": err, "flaky": False, "run": run}
    if not ok:
        first_err = err
        ok, secs, err, out = attempt(sc, ctx)
        if ok:
            flaky, err = True, first_err
    budget = out.get("budget", sc.budget)
    return {"id": sc.id, "tags": sc.tags, "result": verdict(ok, secs, budget, sc.limit), "secs": round(secs, 1),
            "error": err, "error_class": error_class(err) if err else "", "notes": out.get("notes", ""),
            "limit": sc.limit if not ok else None, "flaky": flaky, "run": run}


def add_dialogs(r: dict, dialogs: list[dict]) -> dict:
    r["dialogs"] = [dict(d) for d in dialogs]
    seen = [f"DIALOG: {d['title']}: {d['text']}" for d in dialogs]
    if seen:
        r["notes"] = "; ".join(([r["notes"]] if r.get("notes") else []) + seen)
    return r


def atlas_age(client: McpClient) -> str:
    body, _, err = client.call("sv_atlas_status", {}, timeout=60)
    if err or not body:
        return f"unknown ({error_class(err or 'no answer')})"
    age = body.get("oldest_age_days")
    return "not built" if age is None else f"{age} days" + (" (stale)" if body.get("stale") else "")


def previous() -> tuple[Path | None, list[dict]]:
    files = sorted(AUDIT.glob("*-matrix*.json"), key=lambda p: (p.name[:10], p.stat().st_mtime))
    if not files:
        return None, []
    try:
        return files[-1], json.loads(files[-1].read_text(encoding="utf-8"))["results"]
    except (OSError, ValueError, KeyError):
        return files[-1], []


def render_md(head: dict, results: list[dict]) -> str:
    lines = [f"# Live matrix {head['date']} ({head['out']})", "",
             f"- Time: {head['time']}",
             f"- Engine: {head['version']}",
             f"- Atlas age: {head['atlas_age']}",
             f"- Options: {head['options']}",
             f"- Before: {head['before_file'] or 'none'}",
             f"- Totals: {head['totals']}", "",
             "| ID | Tags | Result | Before | s | Error class |",
             "|---|---|---|---|---:|---|"]
    for r in results:
        res = r["result"] + (" (flaky)" if r.get("flaky") else "")
        rid = r["id"] + (f" #{r['run']}" if head["repeat"] > 1 else "")
        cls = (r["error_class"] or (f"limit: {r['limit']}" if r.get("limit") else "")).replace("|", "\\|")
        lines.append(f"| {rid} | {', '.join(r['tags'])} | {res} | {r.get('before', '')} | {r['secs']} | {cls} |")
    counts = Counter(r["error_class"] for r in results if r["error_class"])
    lines += ["", "## Error classes", ""]
    if counts:
        lines += ["| Count | Error class | Scenarios |", "|---:|---|---|"]
        for cls, n in counts.most_common():
            ids = sorted({r["id"] for r in results if r["error_class"] == cls})
            lines.append(f"| {n} | {cls.replace('|', chr(92) + '|')} | {', '.join(ids)} |")
    else:
        lines.append("None.")
    dialogs = [(r, d) for r in results for d in r.get("dialogs", [])]
    lines += ["", "## Dialogs seen", ""]
    if dialogs:
        lines += ["| Scenario | s into scenario | PID | Title | Text | Clicked | Killed |", "|---|---:|---:|---|---|---|---|"]
        for r, d in dialogs:
            cells = [r["id"], d["seconds_into_scenario"], d["pid"], d["title"], d["text"], d.get("clicked") or "",
                     "yes" if d.get("killed") else ""]
            lines.append("| " + " | ".join(str(c).replace("|", "\\|").replace("\n", " ") for c in cells) + " |")
    else:
        lines.append("None.")
    return "\n".join(lines) + "\n"


def write_report(args: argparse.Namespace, started: datetime, atlas: str, results: list[dict]) -> Path:
    AUDIT.mkdir(parents=True, exist_ok=True)
    date = started.strftime("%Y-%m-%d")
    base = AUDIT / f"{date}-{args.out}"
    for stamp in ("%H%M", "%H%M%S"):
        if not (base.with_suffix(".json").exists() or base.with_suffix(".md").exists()):
            break
        base = AUDIT / f"{date}-{args.out}-{started.strftime(stamp)}"
    before_file, before = previous()
    merged = merge_before(results, before)
    totals = dict(Counter(r["result"] for r in merged))
    head = {"date": date, "out": args.out, "time": started.isoformat(timespec="seconds"), "version": __version__,
            "atlas_age": atlas, "repeat": args.repeat, "totals": totals,
            "options": {"only": args.only, "repeat": args.repeat, "fresh": args.fresh, "keep": args.keep},
            "before_file": before_file.name if before_file else None}
    base.with_suffix(".json").write_text(json.dumps({**head, "results": merged}, ensure_ascii=False, indent=1),
                                         encoding="utf-8")
    base.with_suffix(".md").write_text(render_md(head, merged), encoding="utf-8")
    return base


def main(argv: list[str] | None = None) -> int:
    args = parse(argv)
    started = datetime.now()
    restored = helpers.restore_matrix_bak()      # a crashed S8.sso run must not leave the drill Edge signed out
    if restored:
        print(f"restored leftover *.matrix-bak: {restored}", flush=True)
    resumed = helpers.resume_leftover_suspended()   # a crashed S8.sleep run must not leave servers suspended
    if resumed:
        print(f"resumed leftover suspended server(s): {resumed}", flush=True)
    (Path(config.HOME) / "matrix-hkcu-options.json").unlink(missing_ok=True)   # retired S9 HKCU marker (582b6fb)
    chosen = select(args.only, args.fresh)
    if not chosen:
        print("no scenarios match", file=sys.stderr)
        return 2
    scratch = Path(config.ANALYSER_DIR) / "_matrix"
    scratch.mkdir(parents=True, exist_ok=True)
    results: list[dict] = []
    atlas = "unknown"
    client = None
    ctx = scenarios.Ctx(client=None, scratch=scratch, keep=args.keep)
    try:
        client = McpClient(cwd=scratch)          # project '_matrix': target 'new' lands in the scratch folder
        ctx.client = client
        helpers._CLIENTS.append(client)          # its servers own the engine Excels DialogWatch may touch
        atlas = atlas_age(client)
        for run in range(1, args.repeat + 1):
            for sc in chosen:
                with helpers.DialogWatch(scratch) as watch:
                    ctx.watch = watch
                    try:
                        r = run_one(sc, ctx, run)
                    finally:
                        ctx.watch = None
                add_dialogs(r, watch.dialogs)
                results.append(r)
                print(f"{r['result']:5} {r['secs']:7.1f}s  {sc.id}" + (f"  {r['error_class']}" if r["error"] else "")
                      + (f"  [{len(r['dialogs'])} dialog(s)]" if r["dialogs"] else ""), flush=True)
    finally:
        try:
            helpers.cleanup(ctx)
        except Exception as e:
            print(f"cleanup failed: {e}", file=sys.stderr)
        if client is not None:
            helpers._CLIENTS.remove(client)
            client.close()
        base = write_report(args, started, atlas, results)
        print(f"report: {base.with_suffix('.md')}")
    return 1 if any(r["result"] == "FAIL" for r in results) else 0


if __name__ == "__main__":
    sys.exit(main())
