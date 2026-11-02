"""Command line for the Excel task pane (svkit addin): the deterministic xl_* actions without a Claude round trip.

    python -m smartview_mcp.xlcli check   [--workbook W] [--sheet S]
    python -m smartview_mcp.xlcli changes [--workbook W] [--limit N]
    python -m smartview_mcp.xlcli revert  [--id ID] [--workbook W] [--force]
    python -m smartview_mcp.xlcli workbooks

Prints one JSON object on stdout: {"ok": true, "result": ...} or {"ok": false, "error": "..."}.
Same engine, rules and snapshots as the MCP tools (livexl); takes the same per-Excel turn.
"""
from __future__ import annotations

import argparse
import json
import sys


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="xlcli")
    ap.add_argument("command", choices=["check", "changes", "revert", "workbooks"])
    ap.add_argument("--workbook")
    ap.add_argument("--sheet")
    ap.add_argument("--id")
    ap.add_argument("--limit", type=int, default=20)
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")
    import pythoncom
    pythoncom.CoInitialize()
    from . import livexl
    from .excel import jobs
    try:
        if a.command == "check":
            res = livexl.check(a.workbook, a.sheet)
        elif a.command == "changes":
            res = livexl.changes(a.workbook, a.limit)
        elif a.command == "revert":
            res = livexl.revert(a.id, a.workbook, a.force)
        else:
            res = livexl.workbooks()
        print(json.dumps({"ok": True, "result": res}, ensure_ascii=False, default=str))
        return 0
    except Exception as e:   # noqa: BLE001 - the pane shows the message
        print(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False))
        return 1
    finally:
        jobs.release_turns()


if __name__ == "__main__":
    raise SystemExit(main())
