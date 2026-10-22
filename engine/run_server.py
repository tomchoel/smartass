"""Stable stdio entry point for Claude Code. Registered with:
claude mcp add --scope user smartview -- "%LOCALAPPDATA%\\svkit\\venv\\Scripts\\python.exe" "<this file>"

Claude Code never restarts a stdio server by itself, so a code change used to need '/mcp' reconnect. This file is a
small relay instead: it runs the real server (child_server.py) as a child process and passes MCP JSON-RPC lines
through unchanged. Before forwarding a request it compares the code on disk (smartview_mcp/*.py, vba/*.bas,
child_server.py) with what the child was started from. If it changed and nothing is in flight, it stops the child
gracefully (its Excel closes and Smart View options are restored), starts a new one, replays the client's
'initialize', and sends notifications/tools/list_changed so Claude Code refreshes the tool list. A crashed child is
restarted the same way. Only an edit to THIS file still needs a reconnect - keep it small.
"""
import itertools
import json
import os
import subprocess
import sys
import threading
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
CHILD = HERE / "child_server.py"
LOG = Path(os.environ.get("LOCALAPPDATA", str(HERE))) / "svkit" / "logs" / "proxy.log"
STOP_GRACE = 60          # s for the child to close its Excel on shutdown
INIT_TIMEOUT = 60        # s for a restarted child to answer the replayed 'initialize'
STALE_INFLIGHT = 1800    # s after which an unanswered request no longer blocks a reload


_log_lock = threading.Lock()


def log(msg: str):
    try:
        LOG.parent.mkdir(parents=True, exist_ok=True)
        with _log_lock, LOG.open("a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} [{os.getpid()}] {msg}\n")
    except OSError:
        pass


def fingerprint() -> tuple:
    files = [CHILD, *sorted((HERE / "smartview_mcp").glob("*.py")), *sorted((HERE / "vba").glob("*.bas"))]
    out = []
    for f in files:
        try:
            st = f.stat()
            out.append((f.name, st.st_mtime_ns, st.st_size))
        except OSError:
            pass
    return tuple(out)


class Relay:
    def __init__(self):
        self.out_lock = threading.Lock()
        self.lock = threading.RLock()
        self.child: subprocess.Popen | None = None
        self.fp: tuple = ()
        self.init_req: dict | None = None
        self.init_done = threading.Event()
        self.inflight: dict = {}            # request id -> time forwarded (to the current child)
        self.seq = itertools.count(1)

    # ------------------------------------------------------------ output to Claude Code
    def to_client(self, data: bytes):
        with self.out_lock:
            sys.stdout.buffer.write(data.rstrip(b"\r\n") + b"\n")
            sys.stdout.buffer.flush()

    def notify_client(self, method: str):
        self.to_client(json.dumps({"jsonrpc": "2.0", "method": method}).encode())

    # ------------------------------------------------------------ child lifecycle
    def start_child(self, replay: bool):
        env = {**os.environ, "SVKIT_PROXY": "1", "PYTHONUNBUFFERED": "1"}
        self.fp = fingerprint()
        self.child = subprocess.Popen([sys.executable, "-B", str(CHILD)], stdin=subprocess.PIPE,
                                      stdout=subprocess.PIPE, stderr=None, env=env, cwd=str(HERE))
        self.inflight.clear()
        self.init_done.clear()
        threading.Thread(target=self.pump_child, args=(self.child,), daemon=True).start()
        log(f"child started pid={self.child.pid} replay={replay}")
        if replay and self.init_req:
            req = {**self.init_req, "id": f"relay-init-{next(self.seq)}"}
            self.write_child(json.dumps(req).encode())
            if not self.init_done.wait(INIT_TIMEOUT):
                log("restarted child did not answer initialize in time")
            self.write_child(json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized"}).encode())

    def stop_child(self, reason: str):
        c, self.child = self.child, None
        if c is None:
            return
        log(f"stopping child pid={c.pid}: {reason}")
        try:
            c.stdin.close()                  # EOF -> the server's lifespan shuts its Excel down
        except OSError:
            pass
        try:
            c.wait(STOP_GRACE)
        except subprocess.TimeoutExpired:
            log(f"child pid={c.pid} did not exit in {STOP_GRACE}s - killing its process tree")
            try:
                import psutil
                p = psutil.Process(c.pid)
                for k in p.children(recursive=True) + [p]:
                    k.kill()
            except Exception as e:   # noqa: BLE001
                log(f"kill failed: {e}")
                c.kill()

    def write_child(self, data: bytes) -> bool:
        c = self.child
        if c is None or c.poll() is not None:
            return False
        try:
            c.stdin.write(data.rstrip(b"\r\n") + b"\n")
            c.stdin.flush()
            return True
        except OSError:
            return False

    def pump_child(self, c: subprocess.Popen):
        """Child stdout -> Claude Code. Drops the replayed initialize answer; tracks answered requests."""
        for raw in c.stdout:
            line = raw.strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
            except ValueError:
                self.to_client(line)
                continue
            mid = msg.get("id") if isinstance(msg, dict) else None
            if isinstance(mid, str) and mid.startswith("relay-init-"):
                self.init_done.set()
                continue
            if mid is not None and "method" not in msg:
                with self.lock:
                    if c is self.child:
                        self.inflight.pop(mid, None)
                if self.init_req is not None and mid == self.init_req.get("id") and "result" in msg:
                    # announce tools.listChanged: Claude Code ignores notifications/tools/list_changed unless the
                    # server declared it at connect time - without it, reloaded code ran but NEW tools/params
                    # stayed invisible until /mcp reconnect
                    caps = msg["result"].setdefault("capabilities", {})
                    caps.setdefault("tools", {})["listChanged"] = True
                    line = json.dumps(msg, ensure_ascii=False).encode()
            self.to_client(line)
        # child exited: fail whatever it still owed, so Claude Code does not wait forever
        with self.lock:
            if c is self.child or self.child is None:
                for mid in list(self.inflight):
                    self.to_client(json.dumps({"jsonrpc": "2.0", "id": mid, "error": {
                        "code": -32603, "message": "The smartview server process stopped. Retry the call - it "
                                                   "restarts automatically."}}).encode())
                self.inflight.clear()
        log(f"child pid={c.pid} exited rc={c.poll()}")

    def maybe_reload(self):
        with self.lock:
            dead = self.child is None or self.child.poll() is not None
            now = time.monotonic()
            busy = [m for m, t in self.inflight.items() if now - t < STALE_INFLIGHT]
            if not dead and (busy or fingerprint() == self.fp):
                return
            if not dead:
                self.stop_child("code on disk changed")
            self.start_child(replay=True)
        self.notify_client("notifications/tools/list_changed")
        log("reloaded; sent tools/list_changed")

    # ------------------------------------------------------------ input from Claude Code
    def run(self):
        log("relay started")
        for raw in sys.stdin.buffer:
            line = raw.strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
            except ValueError:
                msg = None
            if isinstance(msg, dict):
                method, mid = msg.get("method"), msg.get("id")
                if method == "initialize":
                    self.init_req = msg
                    with self.lock:
                        if self.child is None:
                            self.start_child(replay=False)
                elif method == "notifications/cancelled":
                    with self.lock:
                        self.inflight.pop((msg.get("params") or {}).get("requestId"), None)
                elif method and mid is not None and self.init_req is not None:
                    self.maybe_reload()
                if method and mid is not None:
                    with self.lock:
                        self.inflight[mid] = time.monotonic()
            if not self.write_child(line) and isinstance(msg, dict) and msg.get("method") and msg.get("id") is not None:
                with self.lock:
                    self.inflight.pop(msg["id"], None)
                self.to_client(json.dumps({"jsonrpc": "2.0", "id": msg["id"], "error": {
                    "code": -32603, "message": "The smartview server process is not running. Retry the call."}}).encode())
        log("client closed stdin - shutting down")
        with self.lock:
            self.stop_child("client disconnected")


if __name__ == "__main__":
    Relay().run()
