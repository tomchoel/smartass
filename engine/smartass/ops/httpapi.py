"""Localhost-only HTTP API for the Excel add-in and scripts: GET /health (the doctor checks) and POST /analyze
(body {"spec": <AnalysisSpec v2>, "target": "session"}: the same run as the MCP tool fin_analyze with a new spec),
behind a per-user shared secret (header X-Smartass-Secret, from secret.bin via config), one request at a time.
Never binds anything but 127.0.0.1."""
from __future__ import annotations

import asyncio
import errno
import hmac
import json
import logging
import socket
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from ..core import config
from ..epm import oauth
from ..xl.host import HostError, InputError

log = logging.getLogger("smartass.http")
MAX_BODY = 1_000_000
_DRAIN_MAX = 16 * MAX_BODY
_FIELDS = {"/analyze": {"spec", "target"}}


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


async def _default_analyze(spec, target: str) -> dict:
    """fin_analyze's own new-question run (imported on the first POST: it loads the MCP server module)."""
    from ..server import _new_question
    return await _new_question(spec, target)


def _call_analyze(analyze, body: dict) -> dict:
    """Validate the body (InputError -> 400), then run the analysis."""
    from pydantic import ValidationError

    from ..model.spec import AnalysisSpec
    if "spec" not in body:
        raise InputError("body needs 'spec' (an AnalysisSpec v2)")
    target = body.get("target", "session")
    if not isinstance(target, str):
        raise InputError("'target' is 'session', 'new' or 'active'")
    try:
        spec = AnalysisSpec.model_validate(body["spec"])
    except ValidationError as e:
        raise InputError(f"invalid spec: {e}") from None
    return asyncio.run(analyze(spec, target))


def make_server(port: int, secret: str, health=None, analyze=None) -> ThreadingHTTPServer:
    """analyze(spec: AnalysisSpec, target) defaults to the MCP server's _new_question (fin_analyze with a spec)."""
    from ..model.validate import GuardError
    health = health or _default_health
    analyze = analyze or _default_analyze
    routes = {"/analyze": analyze}
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
                self._send(200, _call_analyze(routes[path], body))
            except InputError as e:
                self._send(400, {"error": str(e)})
            except GuardError as e:
                self._send(400, {"error": f"{e.reason} - {e.fix}"})
            except HostError as e:
                self._send(503, {"error": str(e)})
            except oauth.NotSignedIn as e:               # the same page and code the MCP tools answer with
                log.info("POST %s: EPM sign-in needed", path)
                self._send(503, {"error": oauth.sign_in_reply(e)})
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
        log.error("no HTTP API secret (secret.bin) - run `smartass setup`")
        return 2
    try:
        srv = make_server(config.HTTP_PORT, config.HTTP_SECRET)
    except OSError as e:
        if e.errno in (errno.EADDRINUSE, 10048):
            log.info("port %s already in use - another serve is running", config.HTTP_PORT)
            return 0
        raise
    log.info("serving on 127.0.0.1:%s", config.HTTP_PORT)
    from ..epm import oauth
    oauth.start_keep_fresh()                         # the bearer route's sign-ins are renewed while this runs
    srv.serve_forever()
    return 0
