"""Blocks commits that contain data or secrets. Default: staged files (git index). --all: every tracked file (CI).
A line containing `allow-secret` is skipped (test fixtures)."""
from __future__ import annotations

import re
import subprocess
import sys

DB = re.compile(r"\.(db|sqlite3?|db-wal|db-shm)$", re.I)
WORKBOOK = re.compile(r"\.(xlsx|xlsm|xlsb|xls)$", re.I)
WORKBOOK_OK = ("knowledge/", "packs/")
TOKENS = [
    re.compile(r"sso_token=[A-Za-z0-9._~%-]{16,}"),
    re.compile(r"eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}"),
    re.compile(r"(?i)\b(secret|password|passwd|api_key|apikey|client_secret)\s*[=:]\s*['\"][^'\"{}\s]{16,}['\"]"),
]


def problems(path: str, data: bytes) -> list[str]:
    p = path.replace("\\", "/")
    if DB.search(p):
        return [f"{p}: database files never go in git"]
    if WORKBOOK.search(p):
        return [] if any(m in p for m in WORKBOOK_OK) else [f"{p}: workbooks only under knowledge/ or packs/"]
    out = []
    for n, line in enumerate(data.decode("utf-8", errors="ignore").splitlines(), 1):
        if "allow-secret" in line:
            continue
        if any(t.search(line) for t in TOKENS):
            out.append(f"{p}:{n}: looks like a token or secret")
    return out


def _git(*a: str) -> bytes:
    return subprocess.run(["git", *a], capture_output=True, check=True).stdout


def main(argv: list[str]) -> int:
    if "--all" in argv:
        files = _git("ls-files", "-z").decode().split("\0")
        read = lambda f: open(f, "rb").read()  # noqa: E731
    else:
        files = _git("diff", "--cached", "--name-only", "--diff-filter=ACMR", "-z").decode().split("\0")
        read = lambda f: _git("show", f":{f}")  # noqa: E731
    found = [m for f in files if f for m in problems(f, read(f))]
    for m in found:
        print("BLOCKED", m)
    if found:
        print("Remove the data/secret, or mark a deliberate test fixture line with `allow-secret`.")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
