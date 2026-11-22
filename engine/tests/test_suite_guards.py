"""tests/conftest.py's guards: the suite cannot start Excel or reach the network, whatever a default is. Every test here
puts a sentinel ONE LEVEL BELOW the guard (the COM call that would create Excel, the process start), so even a broken
guard could not start Excel or open a connection: what is tested is that the guard stops the call first."""
import asyncio
import socket
import urllib.request

import pytest


@pytest.fixture
def below(monkeypatch):
    """Sentinels under the guarded start points: the COM instance creation Dispatch / DispatchEx call, process starts
    and process kills - so a broken guard could not start Excel or end a process."""
    import pythoncom
    import subprocess

    import psutil
    import win32api
    reached = []

    def sentinel(*args, **kwargs):
        reached.append(args[:1])
        raise RuntimeError("reached below the guard")
    monkeypatch.setattr(pythoncom, "CoCreateInstanceEx", sentinel)
    monkeypatch.setattr(pythoncom, "CoCreateInstance", sentinel)
    monkeypatch.setattr(subprocess, "Popen", sentinel)
    monkeypatch.setattr(win32api, "TerminateProcess", sentinel)
    monkeypatch.setattr(psutil, "Process", sentinel)
    return reached


def test_excel_cannot_be_dispatched_and_the_users_excel_cannot_be_launched(below):
    import win32com.client as wc

    from smartass.xl import target
    for start in (wc.Dispatch, wc.DispatchEx):
        with pytest.raises(AssertionError, match="never start Excel"):
            start("Excel.Application")
    with pytest.raises(AssertionError, match="never start Excel"):
        target._launch_excel("x.xlsx")
    assert below == []


def test_no_process_can_be_killed(below):
    """The sweep of an older release's hidden Excels and setup's stop of its SSO watcher are the engine's only kills."""
    from smartass.ops import setup_user
    from smartass.xl import hostproc
    with pytest.raises(AssertionError, match="never kill a process"):
        hostproc.kill_ours(4242, 1.0)
    with pytest.raises(AssertionError, match="never kill a process"):
        setup_user.stop_old_watcher(found=lambda: [(4242, True)])
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
