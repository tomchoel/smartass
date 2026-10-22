"""Localhost-only HTTP API for the future Excel add-in. Sub-project 1 ships only GET /health, behind a per-user
shared secret (header X-Smartass-Secret). Never binds anything but 127.0.0.1."""
from __future__ import annotations

import errno
import hmac
import json
import logging
import socket
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from . import config

log = logging.getLogger("smartass.http")


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


def make_server(port: int, secret: str, health=None) -> ThreadingHTTPServer:
    health = health or _default_health

    class Handler(BaseHTTPRequestHandler):
        server_version = "smartass"

        def _send(self, code: int, body: dict):
            data = json.dumps(body).encode()
            self.send_response(code)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def do_GET(self):  # noqa: N802
            given = self.headers.get("X-Smartass-Secret", "")
            if not hmac.compare_digest(given.encode(), secret.encode()):
                return self._send(401, {"error": "missing or wrong X-Smartass-Secret"})
            if self.path.split("?")[0] != "/health":
                return self._send(404, {"error": "unknown path"})
            return self._send(200, health())

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
