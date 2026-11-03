import json
import threading
import time
import urllib.error
import urllib.request

import pytest

from smartview_mcp import httpapi
from smartview_mcp.host import HostError, InputError

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


def _post(url, body, secret=SECRET, raw=None):
    data = raw if raw is not None else json.dumps(body).encode()
    req = urllib.request.Request(url, data=data, method="POST",
                                 headers={"Content-Type": "application/json",
                                          **({"X-Smartass-Secret": secret} if secret else {})})
    return urllib.request.build_opener(urllib.request.ProxyHandler({})).open(req, timeout=5)


@pytest.fixture
def pipe():
    seen = []

    async def gate(question):
        seen.append(("gate", question))
        if not question:
            raise InputError("question is the user's question")
        return {"outcome": "clear", "question": question}

    async def answer(**kw):
        seen.append(("answer", kw))
        if kw.get("spec") == "slow":
            time.sleep(0.5)
        if kw.get("spec") == "down":
            raise HostError("Excel is gone")
        return {"sheet": "Q1 Answer"}

    s = httpapi.make_server(0, SECRET, health=lambda: {}, gate=gate, answer=answer)
    threading.Thread(target=s.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{s.server_address[1]}", seen
    s.shutdown()
    s.server_close()


def test_gate_and_answer_run_the_pipeline_functions(pipe):
    url, seen = pipe
    assert json.loads(_post(url + "/gate", {"question": "EBITDA Sep 26?"}).read())["outcome"] == "clear"
    assert json.loads(_post(url + "/answer", {"spec": {"version": 1}, "target": "new"}).read()) == {"sheet": "Q1 Answer"}
    assert seen == [("gate", "EBITDA Sep 26?"), ("answer", {"spec": {"version": 1}, "target": "new"})]


def test_http_gate_accepts_answers():
    seen = []

    async def gate(question, answers=None):
        seen.append(answers)
        return {"outcome": "clear"}
    s = httpapi.make_server(0, SECRET, health=lambda: {}, gate=gate, answer=None)
    threading.Thread(target=s.serve_forever, daemon=True).start()
    try:
        url = f"http://127.0.0.1:{s.server_address[1]}"
        _post(url + "/gate", {"question": "How are we doing on costs?", "answers": {"measure": "opex"}})
    finally:
        s.shutdown()
        s.server_close()
    assert seen == [{"measure": "opex"}]


def test_post_needs_the_secret_valid_json_and_known_fields(pipe):
    url, _ = pipe
    for path, kwargs, code in (("/gate", {"body": {"question": "x"}, "secret": None}, 401),
                               ("/gate", {"body": None, "raw": b"{not json"}, 400),
                               ("/gate", {"body": {"question": "x", "evil": 1}}, 400),
                               ("/gate", {"body": None, "raw": b"{" + b" " * (httpapi.MAX_BODY + 1) + b"}"}, 413),
                               ("/query", {"body": {}}, 404),
                               ("/gate", {"body": {"question": ""}}, 400),
                               ("/answer", {"body": {"spec": "down"}}, 503)):
        with pytest.raises(urllib.error.HTTPError) as e:
            _post(url + path, **kwargs)
        assert e.value.code == code


def test_one_request_at_a_time(pipe):
    url, _ = pipe
    t = threading.Thread(target=lambda: _post(url + "/answer", {"spec": "slow"}))
    t.start()
    time.sleep(0.1)
    with pytest.raises(urllib.error.HTTPError) as e:
        _post(url + "/gate", {"question": "x"})
    assert e.value.code == 409
    t.join()