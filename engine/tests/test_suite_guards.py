"""tests/conftest.py's guards: the suite cannot start Excel or reach the network, whatever a default is. Every test here
puts a sentinel ONE LEVEL BELOW the guard (the COM call that would create Excel, the process start), so even a broken
guard could not start Excel or open a connection: what is tested is that the guard stops the call first."""
import asyncio
import socket
import urllib.request
from types import SimpleNamespace

import pytest


@pytest.fixture
def below(monkeypatch):
    """Sentinels under the guarded start points: the COM instance creation DispatchEx calls, process starts, and the
    SSO watcher's first step (its list of running watchers) and its kill - so a broken guard of ensure_global_watcher
    could not even reach the code that kills a stale watcher (sso.py)."""
    import pythoncom
    import subprocess

    from smartass.epm import sso
    reached = []

    def sentinel(*args, **kwargs):
        reached.append(args[:1])
        raise RuntimeError("reached below the guard")
    monkeypatch.setattr(pythoncom, "CoCreateInstanceEx", sentinel)
    monkeypatch.setattr(subprocess, "Popen", sentinel)
    monkeypatch.setattr(sso, "global_watcher_pids", sentinel)
    monkeypatch.setattr(sso, "_heartbeat_age", sentinel)
    monkeypatch.setattr(sso, "psutil", SimpleNamespace(Process=sentinel, Error=Exception))   # sso's own psutil
    return reached


def test_excel_cannot_be_dispatched_and_the_users_excel_cannot_be_launched(below):
    from smartass.xl import host, target
    with pytest.raises(AssertionError, match="never start Excel"):
        host.wc.DispatchEx("Excel.Application")
    with pytest.raises(AssertionError, match="never start Excel"):
        target._launch_excel("x.xlsx")
    with pytest.raises(AssertionError, match="never start Excel"):
        host.sso.ensure_global_watcher()
    with pytest.raises(AssertionError, match="never start Excel"):
        host.ExcelHost._start(SimpleNamespace())                 # the hidden Excel's own start (run_build's route)
    assert below == []


@pytest.mark.parametrize("call", [
    lambda: socket.create_connection(("login.example.com", 443), timeout=1),
    lambda: socket.create_connection(("203.0.113.7", 443), timeout=1),
    lambda: socket.create_connection(("127.evil.example", 443), timeout=1),
    lambda: socket.socket().connect_ex(("203.0.113.7", 443)),
    lambda: socket.getaddrinfo("plan1.example.com", 443),
    lambda: urllib.request.urlopen("https://plan1.example.com/HyperionPlanning/", timeout=1),
    lambda: asyncio.run(asyncio.wait_for(asyncio.open_connection("203.0.113.7", 443), 2)),
])
def test_the_network_cannot_be_reached(call):
    with pytest.raises(AssertionError, match="never reach the network"):
        call()


@pytest.mark.parametrize("host", ["127.0.0.1", "localhost"])
def test_this_machine_can_still_be_reached(host):
    srv = socket.socket()
    srv.bind(("127.0.0.1", 0))
    srv.listen(1)
    try:
        with socket.create_connection((host, srv.getsockname()[1]), timeout=2):
            pass
    finally:
        srv.close()


def test_loopback_is_parsed_not_matched_by_prefix():
    from conftest import _loopback
    assert _loopback("127.0.0.1") and _loopback("::1") and _loopback("localhost") and _loopback("0:0:0:0:0:0:0:1")
    assert not any(_loopback(h) for h in ("127.evil.example", "127.0.0.1.example", "10.0.0.1", "plan1.example", ""))
