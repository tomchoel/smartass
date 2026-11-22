import os

import pytest


@pytest.hookimpl(trylast=True)
def pytest_configure(config):
    # Keep the suite out of the user's real state folder (%LOCALAPPDATA%\smartass): set before smartass.core.config is imported.
    os.environ["SMARTASS_HOME"] = str(config._tmp_path_factory.mktemp("smartass_home"))


LOOPBACK = frozenset({"127.0.0.1", "::1"})


def _loopback(host) -> bool:
    """Exactly this machine: 'localhost', 127.0.0.1 or ::1 (parsed: '127.evil.example' or '0x7f.1' is not)."""
    import ipaddress
    if isinstance(host, bytes):
        host = host.decode("ascii", "replace")
    if host is None or str(host).lower() == "localhost":
        return True
    try:
        return str(ipaddress.ip_address(str(host).split("%")[0])) in LOOPBACK
    except ValueError:
        return False


def _address_ok(sock, address) -> bool:
    import socket
    if getattr(socket, "AF_UNIX", None) is not None and sock.family == socket.AF_UNIX:
        return True                                     # a path on this machine
    return isinstance(address, tuple) and _loopback(address[0])


@pytest.fixture(autouse=True, scope="session")
def _no_excel_no_network():
    """The suite never starts Excel and never reaches the network, whatever a default is (2026-11-22: a test reached
    atlas_build.run_build's hidden-Excel route and started the engine's Excel). From the first test's setup to the end
    of the session - every test and every fixture, module- and session-scoped ones included - the lowest start points
    in this process fail with an AssertionError: Dispatch / DispatchEx of Excel.Application (win32com: the engine has
    no Excel of its own since Task 6, so any is a fault), the user's Excel (xl/target._launch_excel, the one place the
    engine starts Excel), a kill of a process (xl/hostproc.kill_ours: an older release's hidden Excel at the sweep;
    ops/setup_user._kill: an older release's SSO watcher); a socket connect,
    a name lookup or an asyncio connect (the Windows proactor's ConnectEx too) to anything but this machine. A test
    that fakes one of them sets its own fake over this one (monkeypatch restores the guard after it). Out of reach:
    CHILD PROCESSES a test starts (tests/stdio_util.py's server, the scripts some tests run) carry none of these
    patches; and code a test module runs when it is imported (collection comes before this fixture's setup) - keep
    start points out of module level."""
    import asyncio
    import socket

    import win32com.client as wc

    from smartass.ops import setup_user
    from smartass.xl import hostproc, target

    def no_excel(*args, **kwargs):
        raise AssertionError("the test suite must never start Excel")

    def no_kill(*args, **kwargs):
        raise AssertionError("the test suite must never kill a process")
    real_dispatch, real_dispatch_ex = wc.Dispatch, wc.DispatchEx

    def dispatch(progid, *args, **kwargs):
        if "excel" in str(progid).lower():
            no_excel()
        return real_dispatch(progid, *args, **kwargs)

    def dispatch_ex(progid, *args, **kwargs):
        if "excel" in str(progid).lower():
            no_excel()
        return real_dispatch_ex(progid, *args, **kwargs)

    def no_network(*args, **kwargs):
        raise AssertionError("the test suite must never reach the network")
    real_connect, real_connect_ex = socket.socket.connect, socket.socket.connect_ex
    real_getaddrinfo = socket.getaddrinfo

    def connect(self, address):
        return real_connect(self, address) if _address_ok(self, address) else no_network()

    def connect_ex(self, address):
        return real_connect_ex(self, address) if _address_ok(self, address) else no_network()

    def getaddrinfo(host_, *args, **kwargs):
        return real_getaddrinfo(host_, *args, **kwargs) if _loopback(host_) else no_network()
    mp = pytest.MonkeyPatch()
    mp.setattr(wc, "Dispatch", dispatch)
    mp.setattr(wc, "DispatchEx", dispatch_ex)
    mp.setattr(target, "_launch_excel", no_excel)
    mp.setattr(hostproc, "kill_ours", no_kill)
    mp.setattr(setup_user, "_kill", no_kill)
    mp.setattr(socket.socket, "connect", connect)
    mp.setattr(socket.socket, "connect_ex", connect_ex)
    mp.setattr(socket, "getaddrinfo", getaddrinfo)
    proactor = getattr(asyncio, "IocpProactor", None)
    if proactor is not None:
        real_proactor_connect = proactor.connect

        def proactor_connect(self, conn, address):
            return real_proactor_connect(self, conn, address) if _address_ok(conn, address) else no_network()
        mp.setattr(proactor, "connect", proactor_connect)
    yield
    mp.undo()


@pytest.fixture(autouse=True)
def _no_live_drill_routes(monkeypatch):
    """Tests never reach EPM: the HTTP drill link/values are off unless a test turns them on (drillstep then has no
    route: a drill link is an error, split values are not read; tests that drill fake drillstep._http_link), and the
    token drill raises unless a test that drills sets its own fake."""
    from smartass.core import config
    from smartass.epm import drill_token

    def _no_epm():
        raise AssertionError("tests never reach EPM")
    monkeypatch.setattr(drill_token, "shared", _no_epm)
    monkeypatch.setattr(config, "ANSWER_PREFETCH", False)        # tests never reach the live HTTP route
    monkeypatch.setattr(config, "INPROC", False)                 # batches run over COM: no add-in build, no compiler
    monkeypatch.setenv("SV_TEST_NO_HTTP_DRILL", "1")
