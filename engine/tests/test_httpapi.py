import json
import threading
import time
import urllib.error
import urllib.request

import pytest

from smartass.ops import httpapi
from smartass.xl.host import HostError, InputError

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


SPEC = {
    "version": 2, "question": "Business opex vs last year?",
    "plan": {"depth": "quick", "purpose": "Opex vs LY"},
    "queries": [{"id": "fm", "cube": "FinMax", "metrics": ["Opex"], "entity": "un.SY.2BUS", "periods": "closed",
                 "compare": ["ly"]}],
    "blocks": [{"type": "table", "query": "fm", "title": "Opex"}],
}


@pytest.fixture
def pipe():
    seen = []

    async def analyze(spec, target):
        seen.append((spec, target))
        if spec.question == "slow":
            time.sleep(0.5)
        if spec.question == "down":
            raise HostError("Excel is gone")
        if spec.question == "guard":
            from smartass.model.errors import GuardError
            raise GuardError("unknown entity", "find the member with fin_lookup")
        if spec.question == "bad":
            raise InputError("no such member")
        return {"sheet": "Q1 Answer", "key_values": {"Opex": -1}}

    s = httpapi.make_server(0, SECRET, health=lambda: {}, analyze=analyze)
    threading.Thread(target=s.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{s.server_address[1]}", seen
    s.shutdown()
    s.server_close()


def test_analyze_runs_the_spec_like_fin_analyze(pipe):
    from smartass.model.spec import AnalysisSpec
    url, seen = pipe
    out = json.loads(_post(url + "/analyze", {"spec": SPEC, "target": "session"}).read())
    assert out == {"sheet": "Q1 Answer", "key_values": {"Opex": -1}}
    (spec, target), = seen
    assert isinstance(spec, AnalysisSpec) and spec.question == SPEC["question"] and target == "session"
    _post(url + "/analyze", {"spec": SPEC})
    assert seen[-1][1] == "session"                       # target defaults to the session workbook


def test_default_analyze_is_the_servers_new_question(monkeypatch):
    import asyncio

    from smartass import server
    from smartass.model.spec import AnalysisSpec
    calls = []

    async def fake(spec, target):
        calls.append((spec.question, target))
        return {"ok": True}
    monkeypatch.setattr(server, "_new_question", fake)
    out = asyncio.run(httpapi._default_analyze(AnalysisSpec.model_validate(SPEC), "new"))
    assert out == {"ok": True} and calls == [(SPEC["question"], "new")]


def test_post_needs_the_secret_valid_json_and_known_fields(pipe):
    url, _ = pipe

    def spec(question):
        return {"spec": {**SPEC, "question": question}}
    for path, kwargs, code in (("/analyze", {"body": {"spec": SPEC}, "secret": None}, 401),
                               ("/analyze", {"body": None, "raw": b"{not json"}, 400),
                               ("/analyze", {"body": {"spec": SPEC, "evil": 1}}, 400),
                               ("/analyze", {"body": {}}, 400),                       # no spec
                               ("/analyze", {"body": {"spec": {"version": 2}}}, 400),  # invalid spec
                               ("/analyze", {"body": {"spec": SPEC, "target": 3}}, 400),
                               ("/analyze", {"body": None, "raw": b"{" + b" " * (httpapi.MAX_BODY + 1) + b"}"}, 413),
                               ("/gate", {"body": {"question": "x"}}, 404),            # v1 routes are gone
                               ("/answer", {"body": {"spec": SPEC}}, 404),
                               ("/analyze", {"body": spec("bad")}, 400),
                               ("/analyze", {"body": spec("guard")}, 400),
                               ("/analyze", {"body": spec("down")}, 503)):
        with pytest.raises(urllib.error.HTTPError) as e:
            _post(url + path, **kwargs)
        assert e.value.code == code, (path, kwargs.get("body"))


def test_invalid_spec_says_what_is_wrong(pipe):
    url, _ = pipe
    with pytest.raises(urllib.error.HTTPError) as e:
        _post(url + "/analyze", {"spec": {**SPEC, "blocks": [{"type": "nonsense"}]}})
    assert e.value.code == 400 and "blocks" in json.loads(e.value.read())["error"]


def test_one_request_at_a_time(pipe):
    url, _ = pipe
    t = threading.Thread(target=lambda: _post(url + "/analyze", {"spec": {**SPEC, "question": "slow"}}))
    t.start()
    time.sleep(0.1)
    with pytest.raises(urllib.error.HTTPError) as e:
        _post(url + "/analyze", {"spec": SPEC})
    assert e.value.code == 409
    t.join()
