"""Stdio MCP client for live tests: starts run_server.py exactly like Copilot/Claude do."""
from __future__ import annotations

import json
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path

ENGINE = Path(__file__).resolve().parents[2]


class ToolFailed(Exception):
    pass


class McpClient:
    def __init__(self, python: str | None = None, env: dict | None = None):
        self.p = subprocess.Popen([python or sys.executable, "-B", str(ENGINE / "run_server.py")], cwd=ENGINE,
                                  stdin=subprocess.PIPE, stdout=subprocess.PIPE, env=env)
        self.inbox: queue.Queue = queue.Queue()
        self.notes: list[dict] = []
        self._id = 0
        threading.Thread(target=self._read, daemon=True).start()
        self.request("initialize", {"protocolVersion": "2025-06-18", "capabilities": {},
                                    "clientInfo": {"name": "matrix", "version": "1"}}, timeout=120)
        self._send({"jsonrpc": "2.0", "method": "notifications/initialized"})

    def _read(self):
        for line in self.p.stdout:
            if line.strip():
                m = json.loads(line)
                (self.notes.append(m) if "id" not in m else self.inbox.put(m))

    def _send(self, msg: dict):
        self.p.stdin.write((json.dumps(msg) + "\n").encode())
        self.p.stdin.flush()

    def request(self, method: str, params: dict | None = None, timeout: float = 600) -> dict:
        self._id += 1
        i = self._id
        self._send({"jsonrpc": "2.0", "id": i, "method": method, "params": params or {}})
        t0 = time.monotonic()
        while time.monotonic() - t0 < timeout:
            try:
                m = self.inbox.get(timeout=1)
            except queue.Empty:
                continue
            if m.get("id") == i:
                return m
        raise TimeoutError(f"{method} gave no answer within {timeout:.0f}s")

    def call(self, tool: str, args: dict, timeout: float = 600) -> tuple[dict | None, float, str | None]:
        """Returns (body, seconds, error). body is the parsed JSON result, or {"_text": ...} for plain text."""
        t0 = time.monotonic()
        r = self.request("tools/call", {"name": tool, "arguments": args}, timeout)
        secs = time.monotonic() - t0
        if "error" in r:
            return None, secs, json.dumps(r["error"])
        res = r["result"]
        text = res["content"][0]["text"] if res.get("content") else ""
        if res.get("isError"):
            return None, secs, text
        try:
            return json.loads(text), secs, None
        except ValueError:
            return {"_text": text}, secs, None

    def must(self, tool: str, args: dict, timeout: float = 600) -> tuple[dict, float]:
        body, secs, err = self.call(tool, args, timeout)
        if err:
            raise ToolFailed(f"{tool}: {err}")
        return body, secs

    def tools(self) -> list[str]:
        return [t["name"] for t in self.request("tools/list")["result"]["tools"]]

    def close(self):
        try:
            self.p.stdin.close()
            self.p.wait(120)
        except Exception:
            self.p.kill()
