"""smartass command line: doctor, serve (HTTP API + update at login), setup (installer), update, atlas, xl (task pane),
knowledge."""
from __future__ import annotations

import argparse
import json
import os
import sys

# Before core.config is imported: a CLI run gets its own browser profile slots and no warm browser, so it never
# collides with the running servers' slots (an explicit setting wins).
os.environ.setdefault("SV_DRILL_PROFILE_SLOTS", "6")
os.environ.setdefault("SV_DRILL_KEEP_WARM", "0")


def utf8_stdio() -> None:
    """Print UTF-8 whatever the console code page ('·', 'å' in labels); a stream that cannot be reconfigured
    (pythonw, a test capture) is left as it is."""
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except Exception:   # noqa: BLE001
            pass


def main(argv: list[str] | None = None) -> int:
    utf8_stdio()
    ap = argparse.ArgumentParser(prog="smartass")
    sub = ap.add_subparsers(dest="cmd", required=True)
    d = sub.add_parser("doctor", help="health checks (cheap)")
    d.add_argument("--bundle", action="store_true", help="zip logs, redacted config and versions for the maintainer")
    d.add_argument("--json", action="store_true")
    sub.add_parser("serve", help="localhost HTTP API (started at login)")
    sub.add_parser("setup", help="per-user setup (run by installer\\install.ps1)")
    u = sub.add_parser("update", help="install the release from the release folder")
    u.add_argument("--force", action="store_true")
    at = sub.add_parser("atlas", help="the local member atlas: status, or build it from Smart View")
    at_sub = at.add_subparsers(dest="atlas_cmd", required=True)
    at_sub.add_parser("status", help="what the atlas holds and whether it is stale")
    b = at_sub.add_parser("build", help="rebuild from Smart View (own hidden Excel; read-only; a few minutes)")
    b.add_argument("--cube", action="append", choices=["FinMax", "FinLite", "FinPlan"],
                   help="only this cube (repeatable; default all three)")
    b.add_argument("--dimension", action="append", help="only this dimension (repeatable; no purge of the cube)")
    kn = sub.add_parser("knowledge", help="the semantic layer: compile knowledge.db from the semantic files + atlas")
    kn_sub = kn.add_subparsers(dest="knowledge_cmd", required=True)
    kb = kn_sub.add_parser("build", help="build knowledge.db and print the build report (errors keep the old db)")
    kb.add_argument("--semantic", help="semantic folder (default: knowledge/semantic)")
    kb.add_argument("--atlas", help="atlas.db (default: the home folder's)")
    kb.add_argument("--out", help="knowledge.db to write (default: the home folder's)")
    kr = kn_sub.add_parser("refresh", help="monthly: rebuild the atlas (live, a few minutes), diff it, draft facts, "
                           "rebuild knowledge.db, run the offline test set, write the impact report")
    kr.add_argument("--no-atlas", action="store_true", help="skip the atlas rebuild (use the local atlas)")
    kr.add_argument("--month", help="the closed month YYYY-MM (default: the latest an answer saw, else last month)")
    ka = kn_sub.add_parser("approve", help="approve drafts (status approved, approved = {by, on}), then rebuild")
    ka.add_argument("ids", nargs="+", help="metric or fact ids, or 'all'")
    ka.add_argument("--by", default="tom")
    kn_sub.add_parser("status", help="last build, drafts awaiting approval, last report")
    xl = sub.add_parser("xl", help="the Excel task pane's actions on open workbooks (one JSON object on stdout)")
    xl_sub = xl.add_subparsers(dest="xl_cmd", required=True)
    xc = xl_sub.add_parser("changes", help="the snapshot/change log (kept 7 days), newest first")
    xc.add_argument("--workbook")
    xc.add_argument("--limit", type=int, default=30)
    xr = xl_sub.add_parser("revert", help="undo a change from its snapshot (default: the workbook's latest)")
    xr.add_argument("--id")
    xr.add_argument("--workbook")
    xr.add_argument("--force", action="store_true")
    xk = xl_sub.add_parser("check", help="tn-formelstil check of the live formulas")
    xk.add_argument("--workbook")
    xk.add_argument("--sheet")
    a = ap.parse_args(argv)

    if a.cmd == "xl":
        return _xl(a)

    if a.cmd == "doctor":
        from . import doctor
        if a.bundle:
            print(doctor.bundle())
            return 0
        rs = doctor.run()
        print(json.dumps(doctor.as_dict(rs), indent=2) if a.json else doctor.to_text(rs))
        return 1 if doctor.overall(rs) == "FAIL" else 0
    if a.cmd == "serve":
        from ..core import config, logsetup
        from . import httpapi, update
        logsetup.configure(config.HOME / "logs" / "engine.log")
        logsetup.add_secret(config.HTTP_SECRET)
        if update.start_if_available():
            return 0                                  # the installer restarts `serve` when it is done
        return httpapi.serve()
    if a.cmd == "setup":
        from . import setup_user
        for line in setup_user.run():
            print(line)
        return 0
    if a.cmd == "atlas":
        from . import atlas_build
        if a.atlas_cmd == "status":
            print(json.dumps(atlas_build.status(), indent=2))
            return 0
        from ..core import config, logsetup
        logsetup.configure(config.HOME / "logs" / "atlas.log")
        done = atlas_build.run_build(a.cube, a.dimension)
        failed = [d for d in done if d.get("error")]
        print(f"atlas: {len(done) - len(failed)} built, {len(failed)} failed ({config.ATLAS_DB})")
        return 1 if failed else 0
    if a.cmd == "knowledge":
        return _knowledge(a)
    if a.cmd == "update":
        from . import update
        print(update.apply(force=a.force))
        return 0
    return 2


def _user_job(fn):
    """Run fn as v2 runs a user-Excel job (server._user -> host.call(user_excel=True)): COM on this thread, a job
    token bound, the per-Excel turns taken while resolving the workbook released afterwards."""
    import pythoncom

    from ..xl import jobs
    pythoncom.CoInitialize()
    jobs.bind_job()
    try:
        return fn()
    finally:
        jobs.release_turns()


def _xl_changes(workbook, limit):
    from ..xl.workbench import history
    return history.changes(workbook, limit)


def _xl_revert(change_id, workbook, force):
    from .. import workbench
    return workbench.revert(change_id, workbook, force)


def _xl_check(workbook, sheet):
    from .. import workbench
    return workbench._check(workbook, sheet)


def _xl(a) -> int:
    """{"ok": true, "result": ...} or {"ok": false, "error": ...} - one line, what the task pane parses."""
    try:
        if a.xl_cmd == "changes":
            res = _xl_changes(a.workbook, a.limit)                  # files only: no Excel turn
        elif a.xl_cmd == "revert":
            res = _user_job(lambda: _xl_revert(a.id, a.workbook, a.force))
        else:
            res = _user_job(lambda: _xl_check(a.workbook, a.sheet))
        print(json.dumps({"ok": True, "result": res}, ensure_ascii=False, default=str))
        return 0
    except Exception as e:   # noqa: BLE001 - the pane shows the message
        print(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False))
        return 1


def _refresh_month(text: str | None, home) -> tuple[int, int]:
    import datetime as dt

    from ..knowledge.refresh import due
    from ..core import config
    text = text or due(home, config.KNOWLEDGE_DB)["closed_seen"]
    if text:
        return int(text[:4]), int(text[5:7])
    first = dt.date.today().replace(day=1) - dt.timedelta(days=1)
    return first.year, first.month


def _knowledge(a) -> int:
    from pathlib import Path

    from ..core import config
    from ..knowledge import refresh
    from ..knowledge.build import build
    sem, adb, kdb, home = Path(config.SEMANTIC_DIR), Path(config.ATLAS_DB), Path(config.KNOWLEDGE_DB), Path(config.HOME)
    if a.knowledge_cmd == "build":
        r = build(Path(a.semantic or sem), Path(a.atlas or adb), Path(a.out or kdb))
        print(r.to_text())
        return 0 if r.ok else 1
    if a.knowledge_cmd == "refresh":
        from ..testset import runner
        rebuild = None
        if not a.no_atlas:
            from ..core import logsetup
            from . import atlas_build
            logsetup.configure(config.HOME / "logs" / "atlas.log")
            rebuild = atlas_build.run_build
        res = refresh.refresh(sem, adb, kdb, home, _refresh_month(a.month, home), rebuild_atlas=rebuild,
                              run_testset=lambda: runner.run_offline(runner.load_cases()))
        ok = sum(1 for t in res.testset if t.get("ok") is True)
        print(f"diff {res.diff.counts()}{' (first refresh: baseline saved)' if res.first else ''}; "
              f"{len(res.drafts)} draft(s) written; test set {ok}/{len(res.testset)}")
        print(res.build.to_text())
        print(f"report: {res.report}")
        return 0 if res.ok else 1
    if a.knowledge_cmd == "approve":
        try:
            done, r = refresh.approve(sem, adb, kdb, a.ids, by=a.by)
        except KeyError as e:
            print(e.args[0])
            return 1
        print(f"approved {len(done)}: {', '.join(done) or '-'}")
        print(r.to_text())
        return 0 if r.ok else 1
    st = refresh.status(sem, kdb, home)
    print(json.dumps(st, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
