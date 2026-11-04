"""smartass command line: doctor, serve (HTTP API + update at login), setup (installer), update, atlas,
knowledge."""
from __future__ import annotations

import argparse
import json
import sys


def main(argv: list[str] | None = None) -> int:
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
    a = ap.parse_args(argv)

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
