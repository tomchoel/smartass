"""Talk to the MCP server over stdio, as a client does. No Excel or browser is started: listing tools touches neither."""
from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
from pathlib import Path

ENGINE = Path(__file__).resolve().parents[1]


def list_tools(home: Path, timeout: float = 60) -> list[dict]:
    env = {**os.environ, "SMARTASS_HOME": str(home), "PYTHONUTF8": "1", "PYTHONUNBUFFERED": "1"}
    p = subprocess.Popen([sys.executable, "-B", str(ENGINE / "child_server.py")], stdin=subprocess.PIPE,
                         stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, env=env, cwd=str(ENGINE))
    lines: queue.Queue = queue.Queue()
    threading.Thread(target=lambda: [lines.put(raw) for raw in p.stdout], daemon=True).start()

    def send(msg: dict):
        p.stdin.write((json.dumps(msg) + "\n").encode())
        p.stdin.flush()

    def answer(mid: int) -> dict:
        while True:
            msg = json.loads(lines.get(timeout=timeout))
            if msg.get("id") == mid:
                return msg

    try:
        send({"jsonrpc": "2.0", "id": 1, "method": "initialize",
              "params": {"protocolVersion": "2025-06-18", "capabilities": {},
                         "clientInfo": {"name": "smoke", "version": "0"}}})
        assert "result" in answer(1)
        send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        send({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
        return answer(2)["result"]["tools"]
    finally:
        p.stdin.close()
        try:
            p.wait(30)
        except subprocess.TimeoutExpired:
            p.kill()
