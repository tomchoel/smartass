r"""Stable stdio entry point for Claude Code.
Registered as: <engine>\.venv\Scripts\python.exe <this file>   (MCP entry 'smartass')

Claude Code never restarts a stdio server by itself. This file is a small relay instead: it runs the real server
(child_server.py) as a child process and passes MCP JSON-RPC lines through unchanged. The child is replaced (told to
stop by the end of its input, it finishes and exits; one still running after STOP_GRACE has its process tree ended,
never an Excel in it - end_tree; a new one starts, the client's 'initialize' is replayed and
notifications/tools/list_changed is sent so Claude Code refreshes the tool list) when:
- engine/VERSION changed and nothing is in flight, or
- the client calls fin_status(restart=True) (an explicit restart; in-flight calls then fail with "retry"), or
- the child crashed.
Editing a .py file does NOT restart the child: half-saved edits and parallel agents' work in progress no longer
bounce the server. To roll a code change into a running session:
- `scripts\bump.ps1` (bumps engine/VERSION and pyproject.toml; the next request reloads), or
- `fin_status(restart=True)` (reloads the current code without a version bump).
Then run /mcp in Claude Code if a tool's parameters changed. Only an edit to THIS file still needs a reconnect -
keep it small.
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
LOG = Path(os.environ.get("SMARTASS_HOME") or os.environ.get("SVKIT_DIR")
           or Path(os.environ.get("LOCALAPPDATA", str(HERE))) / "smartass") / "logs" / "proxy.log"
STOP_GRACE = 60          # s for the child to finish and exit once told to stop
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
    """What a running child was started from: the content of engine/VERSION (only)."""
    try:
        return ((HERE / "VERSION").read_text(encoding="utf-8").strip(),)
    except OSError:
        return ("",)


def end_tree(root) -> list[int]:
    """End root (a psutil.Process: the server child) and every process under it, except each excel.exe and everything
    under that Excel: an Excel the server launched (xl/target._launch_excel: EXCEL.EXE /x <file>) is a child of the
    server process, and it is the USER's Excel, with the user's workbooks (final review B, I1). The tree is walked
    one level at a time and an Excel's subtree pruned (children(recursive=True) is a flat list: a filter on the name
    alone would still end Excel's own children). Read whole first, then ended: the processes under root first, root
    last. A process gone meanwhile (or not ours to read) is skipped and never stops the others. Returns the pids
    ended."""
    import psutil
    found, todo = [], [root]
    while todo:
        p = todo.pop()
        try:
            if p is not root and p.name().lower() == "excel.exe":
                continue                                 # the user's Excel and all under it: never ended
            todo.extend(p.children())
        except psutil.Error:
            continue
        found.append(p)
    ended = []
    for p in found[1:] + found[:1]:
        try:
            p.kill()
            ended.append(p.pid)
        except psutil.Error:
            pass
    return ended


RESTART_TOOL = "fin_status"


def wants_restart(msg: dict) -> bool:
    """A tools/call of fin_status with restart=true: the client asks for a fresh child (current code on disk)."""
    if not isinstance(msg, dict) or msg.get("method") != "tools/call":
        return False
    params = msg.get("params") or {}
    return params.get("name") == RESTART_TOOL and (params.get("arguments") or {}).get("restart") is True


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
            c.stdin.close()                  # EOF -> the server finishes and exits
        except OSError:
            pass
        try:
            c.wait(STOP_GRACE)
        except subprocess.TimeoutExpired:
            log(f"child pid={c.pid} did not exit in {STOP_GRACE}s - ending its process tree (never an Excel)")
            try:
                import psutil
                log(f"ended pids {end_tree(psutil.Process(c.pid))}")
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
                        "code": -32603, "message": "The smartass server process stopped. Retry the call - it "
                                                   "restarts automatically."}}).encode())
                self.inflight.clear()
        log(f"child pid={c.pid} exited rc={c.poll()}")

    def maybe_reload(self, force: bool = False):
        with self.lock:
            dead = self.child is None or self.child.poll() is not None
            now = time.monotonic()
            busy = [m for m, t in self.inflight.items() if now - t < STALE_INFLIGHT]
            if not dead and not force and (busy or fingerprint() == self.fp):
                return
            if not dead:
                self.stop_child("fin_status(restart=True)" if force else "engine/VERSION changed")
            self.start_child(replay=True)
        self.notify_client("notifications/tools/list_changed")
        log("reloaded; sent tools/list_changed")

    # ------------------------------------------------------------ input from Claude Code
    def run(self):
        try:                                 # proxy.log is appended per line with no rotation: cap it at startup
            if LOG.stat().st_size > 1_000_000:
                os.replace(LOG, LOG.with_suffix(".log.1"))
        except OSError:
            pass
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
                    self.maybe_reload(force=wants_restart(msg))
                if method and mid is not None:
                    with self.lock:
                        self.inflight[mid] = time.monotonic()
            if not self.write_child(line) and isinstance(msg, dict) and msg.get("method") and msg.get("id") is not None:
                with self.lock:
                    self.inflight.pop(msg["id"], None)
                self.to_client(json.dumps({"jsonrpc": "2.0", "id": msg["id"], "error": {
                    "code": -32603, "message": "The smartass server process is not running. Retry the call."}}).encode())
        log("client closed stdin - shutting down")
        with self.lock:
            self.stop_child("client disconnected")


if __name__ == "__main__":
    Relay().run()
