"""Localhost-only HTTP API for the Excel add-in: GET /health, POST /gate, POST /answer (the same pipeline calls as the
MCP tools sv_gate/sv_answer; JSON body = their arguments), behind a per-user shared secret (header
X-Smartass-Secret), one request at a time. Never binds anything but 127.0.0.1."""
from __future__ import annotations

import asyncio
import errno
import hmac
import json
import logging
import socket
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from importlib import import_module

from . import config
from .host import HostError, InputError

log = logging.getLogger("smartass.http")
MAX_BODY = 1_000_000
_DRAIN_MAX = 16 * MAX_BODY
_FIELDS = {"/gate": {"question", "answers"}, "/answer": {"spec", "followup", "q", "target", "allow_archive", "detail"}}


class _ExclusiveServer(ThreadingHTTPServer):
    # On Windows SO_REUSEADDR lets a second socket bind the same port; take the port exclusively instead.
    allow_reuse_address = False
    daemon_threads = True

    def server_bind(self):
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


def _default_health() -> dict:
    from . import doctor
    return doctor.as_dict(doctor.run(include_http=False))


def make_server(port: int, secret: str, health=None, gate=None, answer=None) -> ThreadingHTTPServer:
    """gate/answer default to tools_core.gate/answer (imported on the first POST: it loads the MCP tools)."""
    health = health or _default_health
    routes = {"/gate": gate, "/answer": answer}
    busy = threading.Lock()

    class Handler(BaseHTTPRequestHandler):
        server_version = "smartass"

        def _send(self, code: int, body: dict):
            data = json.dumps(body, default=str).encode()
            self.send_response(code)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def _drain(self, size: int):
            # Closing a socket with unread request bytes makes Windows send RST, which can beat the 413 to the client.
            self.close_connection = True
            left = min(size, _DRAIN_MAX)
            while left > 0:
                chunk = self.rfile.read(min(left, 65536))
                if not chunk:
                    break
                left -= len(chunk)

        def _authorised(self, size: int = 0) -> bool:
            ok = hmac.compare_digest(self.headers.get("X-Smartass-Secret", "").encode(), secret.encode())
            if not ok:
                self._drain(size)
            return ok or self._send(401, {"error": "missing or wrong X-Smartass-Secret"})

        def do_GET(self):  # noqa: N802
            if not self._authorised():
                return
            if self.path.split("?")[0] != "/health":
                return self._send(404, {"error": "unknown path"})
            return self._send(200, health())

        def do_POST(self):  # noqa: N802
            path = self.path.split("?")[0]
            size = int(self.headers.get("Content-Length") or 0)
            if not self._authorised(size):
                return
            if path not in routes or size > MAX_BODY:
                self._drain(size)
                if path not in routes:
                    return self._send(404, {"error": "unknown path"})
                return self._send(413, {"error": f"body larger than {MAX_BODY} bytes"})
            try:
                body = json.loads(self.rfile.read(size) or b"{}")
            except ValueError:
                return self._send(400, {"error": "body is not JSON"})
            if not isinstance(body, dict) or set(body) - _FIELDS[path]:
                extra = sorted(set(body) - _FIELDS[path]) if isinstance(body, dict) else "not an object"
                return self._send(400, {"error": f"unknown fields: {extra}; allowed: {sorted(_FIELDS[path])}"})
            if not busy.acquire(blocking=False):
                return self._send(409, {"error": "busy: one request at a time"})
            try:
                fn = routes[path] or getattr(import_module(".tools_core", __package__), path[1:])
                self._send(200, asyncio.run(fn(**body)))
            except InputError as e:
                self._send(400, {"error": str(e)})
            except HostError as e:
                self._send(503, {"error": str(e)})
            except Exception:   # noqa: BLE001 - the client gets a short answer, the log the traceback
                log.exception("POST %s failed", path)
                self._send(500, {"error": "internal error; see server.log"})
            finally:
                busy.release()

        def log_message(self, fmt, *args):
            log.info("%s %s", self.address_string(), fmt % args)

    return _ExclusiveServer(("127.0.0.1", port), Handler)


def serve() -> int:
    if not config.HTTP_SECRET:
        log.error("no http.secret in config.toml - run the installer")
        return 2
    try:
        srv = make_server(config.HTTP_PORT, config.HTTP_SECRET)
    except OSError as e:
        if e.errno in (errno.EADDRINUSE, 10048):
            log.info("port %s already in use - another serve is running", config.HTTP_PORT)
            return 0
        raise
    log.info("serving on 127.0.0.1:%s", config.HTTP_PORT)
    srv.serve_forever()
    return 0
