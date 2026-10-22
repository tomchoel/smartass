import json
import threading
import urllib.error
import urllib.request

import pytest

from smartview_mcp import httpapi

SECRET = "s" * 32


@pytest.fixture
def srv():
    s = httpapi.make_server(0, SECRET, health=lambda: {"status": "OK"})
    threading.Thread(target=s.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{s.server_address[1]}"
    s.shutdown()
    s.server_close()


def _get(url, secret=None):
    req = urllib.request.Request(url, headers={"X-Smartass-Secret": secret} if secret else {})
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    return opener.open(req, timeout=5)


def test_health_needs_the_secret(srv):
    for bad in (None, "wrong"):
        with pytest.raises(urllib.error.HTTPError) as e:
            _get(srv + "/health", bad)
        assert e.value.code == 401
    assert json.loads(_get(srv + "/health", SECRET).read()) == {"status": "OK"}


def test_unknown_path_is_404(srv):
    with pytest.raises(urllib.error.HTTPError) as e:
        _get(srv + "/query", SECRET)
    assert e.value.code == 404


def test_binds_loopback_only():
    s = httpapi.make_server(0, SECRET)
    try:
        assert s.server_address[0] == "127.0.0.1"
    finally:
        s.server_close()


def test_port_cannot_be_bound_twice():
    s = httpapi.make_server(0, SECRET)
    try:
        with pytest.raises(OSError):
            httpapi.make_server(s.server_address[1], SECRET).server_close()
    finally:
        s.server_close()


def test_serve_returns_0_when_port_taken(monkeypatch):
    s = httpapi.make_server(0, SECRET)
    try:
        monkeypatch.setattr(httpapi.config, "HTTP_SECRET", SECRET)
        monkeypatch.setattr(httpapi.config, "HTTP_PORT", s.server_address[1])
        assert httpapi.serve() == 0
    finally:
        s.server_close()


def test_serve_refuses_without_secret(monkeypatch):
    monkeypatch.setattr(httpapi.config, "HTTP_SECRET", "")
    assert httpapi.serve() == 2
