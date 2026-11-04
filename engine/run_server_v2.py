# engine/run_server_v2.py
"""Stable stdio entry for smartass2: v1's relay (run_server.py) pointed at child_server_v2.py, with its own log
and a fingerprint that also covers engine/smartass/. Register with:
claude mcp add smartass2 -- <engine>\\.venv\\Scripts\\python.exe <engine>\\run_server_v2.py"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import run_server as rs  # noqa: E402

rs.CHILD = rs.HERE / "child_server_v2.py"
rs.LOG = rs.LOG.with_name("proxy_v2.log")
_v1_fingerprint = rs.fingerprint


def fingerprint() -> tuple:
    extra = []
    for f in sorted((rs.HERE / "smartass").rglob("*.py")):
        try:
            st = f.stat()
            extra.append((str(f.relative_to(rs.HERE)), st.st_mtime_ns, st.st_size))
        except OSError:
            pass
    return _v1_fingerprint() + tuple(extra)


rs.fingerprint = fingerprint

if __name__ == "__main__":
    rs.Relay().run()
