"""tests/conftest.py's guards: the suite cannot start Excel or reach the network, whatever a default is. Every test here
puts a sentinel ONE LEVEL BELOW the guard (the COM call that would create Excel, the process start), so even a broken
guard could not start Excel or open a connection: what is tested is that the guard stops the call first."""
import asyncio
import socket
import types
import urllib.request

import pytest
from conftest import VIOLATIONS, GuardViolation


@pytest.fixture(autouse=True)
def _ours():
    """The violations these tests make on purpose are taken back after each one, so they do not fail the session
    (conftest.pytest_sessionfinish); any other stays."""
    before = len(VIOLATIONS)
    yield
    del VIOLATIONS[before:]


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
        with pytest.raises(GuardViolation, match="never start Excel"):
            start("Excel.Application")
    with pytest.raises(GuardViolation, match="never start Excel"):
        target._launch_excel("x.xlsx")
    assert below == []


def test_no_process_can_be_killed(below):
    """The sweep of an older release's hidden Excels and setup's stop of its SSO watcher and of its `smartass serve`
    are the engine package's only kills. One more site ends processes, outside the package: the relay's
    run_server.end_tree (a server that did not exit in time, with its tree, never an Excel in it or what runs under
    that Excel: tests/test_relay_stop.py)."""
    from smartass.ops import setup_user
    from smartass.xl import hostproc
    with pytest.raises(GuardViolation, match="never kill a process"):
        hostproc.kill_ours(4242, 1.0)
    with pytest.raises(GuardViolation, match="never kill a process"):
        setup_user.stop_old_watcher(found=lambda: [(types.SimpleNamespace(pid=4242, is_running=lambda: True), True)])
    with pytest.raises(GuardViolation, match="never kill a process"):
        found = [setup_user.Serve(types.SimpleNamespace(pid=4242, is_running=lambda: True), True, ())]
        setup_user.stop_old_serve(lambda s: None, found=lambda: (found, 0))
    assert below == []


def test_a_violation_is_not_swallowed_by_a_broad_except_and_is_recorded():
    """Engine code that catches Exception (a clean-up, a fallback) cannot hide a guarded call: the guard raises a
    BaseException, and records it so the session fails even if something caught that too."""
    from smartass.xl import target
    before = len(VIOLATIONS)
    with pytest.raises(GuardViolation, match="never start Excel"):
        try:
            target._launch_excel("x.xlsx")
        except Exception:   # noqa: BLE001 - what a broad engine handler does
            pass
    assert VIOLATIONS[before:] == ["the test suite must never start Excel"]


def test_a_recorded_violation_fails_the_session_at_its_end():
    import types

    import conftest
    said = []
    reporter = types.SimpleNamespace(write_line=lambda text, **kw: said.append(text))
    session = types.SimpleNamespace(exitstatus=pytest.ExitCode.OK, config=types.SimpleNamespace(
        pluginmanager=types.SimpleNamespace(get_plugin=lambda name: reporter)))
    conftest.pytest_sessionfinish(session, pytest.ExitCode.OK)
    assert session.exitstatus == pytest.ExitCode.OK and said == []          # nothing recorded: the session stands
    VIOLATIONS.append("the test suite must never reach the network")          # taken back after the test (_ours)
    conftest.pytest_sessionfinish(session, pytest.ExitCode.OK)
    assert session.exitstatus == pytest.ExitCode.TESTS_FAILED and "never reach the network" in said[0]


@pytest.mark.parametrize("status", [pytest.ExitCode.INTERRUPTED, pytest.ExitCode.INTERNAL_ERROR])
def test_a_recorded_violation_keeps_a_worse_exit_status(status):
    """Task 10 (final review, Minor 3): an interrupted run or an internal error keeps its own status; the guard's line
    is still written."""
    import types

    import conftest
    said = []
    reporter = types.SimpleNamespace(write_line=lambda text, **kw: said.append(text))
    session = types.SimpleNamespace(exitstatus=status, config=types.SimpleNamespace(
        pluginmanager=types.SimpleNamespace(get_plugin=lambda name: reporter)))
    VIOLATIONS.append("the test suite must never reach the network")          # taken back after the test (_ours)
    conftest.pytest_sessionfinish(session, status)
    assert session.exitstatus == status and "never reach the network" in said[0]


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
    with pytest.raises(GuardViolation, match="never reach the network"):
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
