"""smartass command line: doctor, serve (HTTP API + update at login), setup (installer), update."""
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
    if a.cmd == "update":
        from . import update
        print(update.apply(force=a.force))
        return 0
    return 2


if __name__ == "__main__":
    sys.exit(main())
