"""smartass command line: doctor, serve (HTTP API + update at login), setup (installer), update, atlas."""
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
    if a.cmd == "update":
        from . import update
        print(update.apply(force=a.force))
        return 0
    return 2


if __name__ == "__main__":
    sys.exit(main())
