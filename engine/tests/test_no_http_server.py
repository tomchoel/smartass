"""The engine is an MCP server over stdio and nothing else (light-fast phase 4 Task 1): the localhost HTTP API
(`smartass serve`, ops/httpapi.py) with its secret, the Excel task pane (addin/tn-claude) and the pane's CLI twins
(`smartass xl changes|revert|check`) are deleted. No module or name of them is left in the engine except where an older
release's leftovers are found and removed; no engine source opens a listening socket or holds an HTTP server; the docs
a user or an agent reads name none of them as a current thing. Read from the sources (as test_neutral_imports does
for the hidden Excel), not by running anything."""
import ast
import importlib.util

import pytest
from test_neutral_imports import DOCS, ENGINE, _rel, _sources, doc_findings

GONE = ["smartass.ops.httpapi"]
NAMES = ["httpapi", "HTTP_SECRET", "SECRET_FILE", "read_secret", "load_http_secret", "protect_secret", "secret_step",
         "HTTP_PORT", "make_shortcut", "start_if_available", "_xl_changes", "_xl_revert", "_xl_check", "_user_job",
         "X-Smartass-Secret", "SMARTASS_HTTP", "POST /analyze", "tn-claude", "server.mjs"]
LEFTOVER_USES = {
    ("smartass/core/logsetup.py", "X-Smartass-Secret"): "an older release's log lines may still carry the header: the "
                                                        "scrub keeps it (doctor --bundle)",
    ("smartass/ops/pane_leftovers.py", "tn-claude"): "the pane's Startup entry is found by what it names",
    ("smartass/ops/pane_leftovers.py", "server.mjs"): "the pane's Startup entry is found by what it names",
}


def test_the_http_api_and_the_pane_are_gone():
    assert [m for m in GONE if importlib.util.find_spec(m) is not None] == []
    assert not (ENGINE.parent / "addin").exists()


@pytest.mark.parametrize("cmd", [["serve"], ["xl", "changes"], ["xl", "revert"], ["xl", "check"]])
def test_the_command_line_has_no_serve_and_no_xl(cmd, capsys):
    from smartass.ops import cli
    with pytest.raises(SystemExit) as e:
        cli.main(cmd)
    assert e.value.code == 2 and "invalid choice" in capsys.readouterr().err


def test_no_name_of_the_deleted_parts_is_left_in_the_engine():
    found = [f"{_rel(p)}: {n}" for p in _sources() for n in NAMES
             if n in p.read_text(encoding="utf-8") and (_rel(p), n) not in LEFTOVER_USES]
    assert found == [] and all(LEFTOVER_USES.values())


SERVER_MODULES = ("http.server", "socketserver", "wsgiref.simple_server", "xmlrpc.server", "aiohttp", "uvicorn",
                  "starlette", "hypercorn", "flask", "fastapi", "tornado", "twisted", "websockets")
SERVER_CALLS = {"listen", "bind", "serve_forever", "start_server", "start_unix_server", "create_server",
                "create_unix_server", "make_server", "HTTPServer", "ThreadingHTTPServer", "TCPServer",
                "ThreadingTCPServer", "UDPServer", "run_sse_async", "run_streamable_http_async", "sse_app",
                "streamable_http_app"}


def server_sites(source: str, rel: str) -> list[str]:
    """'rel: what' for every way source could listen: an import of a server module (http.server, socketserver,
    asyncio's servers' frameworks ...), a call named like a socket's listen / bind or a server's start, and a call
    that passes a transport other than "stdio" (FastMCP.run(transport="sse") serves HTTP)."""
    out = []
    for node in ast.walk(ast.parse(source)):
        mods = []
        if isinstance(node, ast.Import):
            mods = [a.name for a in node.names]
        elif isinstance(node, ast.ImportFrom) and node.module:
            mods = [node.module, *(f"{node.module}.{a.name}" for a in node.names)]
        hit = next((m for m in mods if any(m == s or m.startswith(s + ".") for s in SERVER_MODULES)), None)
        if hit:
            out.append(f"{rel}: import {hit}")
        if isinstance(node, ast.Call):
            f = node.func
            name = f.attr if isinstance(f, ast.Attribute) else f.id if isinstance(f, ast.Name) else ""
            if name in SERVER_CALLS:
                out.append(f"{rel}: {name}()")
            out += [f"{rel}: {name}(transport=...)" for k in node.keywords if k.arg == "transport"
                    and not (isinstance(k.value, ast.Constant) and k.value.value == "stdio")]
    return out


@pytest.mark.parametrize("source, found", [
    ("from http.server import ThreadingHTTPServer\ns = ThreadingHTTPServer(('127.0.0.1', 0), H)\n",
     ["x.py: import http.server", "x.py: ThreadingHTTPServer()"]),
    ("import socketserver\n", ["x.py: import socketserver"]),
    ("from http import server\n", ["x.py: import http.server"]),
    ("import socket\ns = socket.socket()\ns.bind(('127.0.0.1', 0))\ns.listen(1)\n", ["x.py: bind()", "x.py: listen()"]),
    ("import asyncio\nasync def go(h):\n    await asyncio.start_server(h, '127.0.0.1', 0)\n", ["x.py: start_server()"]),
    ("def go(mcp):\n    mcp.run(transport='sse')\n", ["x.py: run(transport=...)"]),
    ("def go(mcp):\n    mcp.run()\n    mcp.run(transport='stdio')\n", []),
    ("import uvicorn\n", ["x.py: import uvicorn"]),
    ("import urllib.request\nimport http.client\nfrom smartass.xl import jobs\njobs.bind_job()\n", []),
], ids=["http-server", "socketserver", "from-http-import-server", "socket-listen", "asyncio-server", "mcp-sse",
        "mcp-stdio", "uvicorn", "a-client-is-not-a-server"])
def test_the_server_guard_sees_each_way_to_listen(source, found):
    assert server_sites(source, "x.py") == found


def test_the_engine_opens_no_listening_socket_and_holds_no_http_server():
    assert [s for p in _sources() for s in server_sites(p.read_text(encoding="utf-8"), _rel(p))] == []


DOC_NAMES = ["smartass serve", "HTTP API", "X-Smartass-Secret", "httpapi", "http_api", "POST /analyze", "tn-claude",
             "server.mjs", "task pane", "xl changes", "xl revert", "xl check", "http.port", "secret_step"]
DOC_LEFTOVERS = {
    ("docs/troubleshooting.md", "smartass serve"): "the leftovers section: what an older release ran, and its removal",
    ("docs/troubleshooting.md", "HTTP API"): "the leftovers section",
    ("docs/troubleshooting.md", "task pane"): "the leftovers section and its manual steps in Excel",
    ("docs/troubleshooting.md", "server.mjs"): "the leftovers section: the pane's server, if it still runs",
    ("docs/DESIGN.md", "task pane"): "ops/pane_leftovers: an older release's task pane in Office's settings",
    ("skills/smartass/reference/operations.md", "HTTP API"): "the home's files: secret.bin left from before 0.8.0",
    ("skills/smartass/reference/operations.md", "smartass serve"): "setup stops an older release's serve",
    ("skills/smartass/reference/operations.md", "task pane"): "doctor --fix reports an old task pane's registration",
}


def test_the_docs_name_no_part_of_the_http_api_or_the_pane_as_a_current_thing():
    assert doc_findings(ENGINE.parent, DOCS, DOC_NAMES, DOC_LEFTOVERS, {}) == []
    assert all(DOC_LEFTOVERS.values())
