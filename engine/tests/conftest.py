import os

import pytest


@pytest.hookimpl(trylast=True)
def pytest_configure(config):
    # Keep the suite out of the user's real state folder (%LOCALAPPDATA%\smartass): set before smartass.core.config is imported.
    os.environ["SMARTASS_HOME"] = str(config._tmp_path_factory.mktemp("smartass_home"))


LOOPBACK = ("localhost", "127.0.0.1", "::1", "")


def _loopback(address) -> bool:
    if not isinstance(address, tuple):                  # an AF_UNIX path, a named pipe
        return True
    host = str(address[0]).lower()
    return host in LOOPBACK or host.startswith("127.")


@pytest.fixture(autouse=True)
def _no_excel_no_network(monkeypatch):
    """The suite never starts Excel and never reaches the network, whatever a default is (2026-11-22: a test reached
    atlas_build.run_build's hidden-Excel route and started the engine's Excel). Blocked at the lowest start points:
    DispatchEx of Excel.Application (xl/host.py, the hidden Excel), the SSO dialog watcher that start launches, and
    the user's Excel (xl/target._launch_excel); a socket connect or a name lookup to anything but this machine. A test
    that fakes one of them installs its own fake over this one, as before."""
    import socket

    from smartass.xl import host, target

    def no_excel(*args, **kwargs):
        raise AssertionError("the test suite must never start Excel")
    real_dispatch_ex = host.wc.DispatchEx

    def dispatch_ex(progid, *args, **kwargs):
        if "excel" in str(progid).lower():
            no_excel()
        return real_dispatch_ex(progid, *args, **kwargs)
    monkeypatch.setattr(host.wc, "DispatchEx", dispatch_ex)
    monkeypatch.setattr(host.sso, "ensure_global_watcher", no_excel)
    monkeypatch.setattr(target, "_launch_excel", no_excel)

    real_connect, real_connect_ex, real_getaddrinfo = socket.socket.connect, socket.socket.connect_ex, socket.getaddrinfo

    def connect(self, address):
        if not _loopback(address):
            raise AssertionError("the test suite must never reach the network")
        return real_connect(self, address)

    def connect_ex(self, address):
        if not _loopback(address):
            raise AssertionError("the test suite must never reach the network")
        return real_connect_ex(self, address)

    def getaddrinfo(host_, *args, **kwargs):
        if not _loopback((host_ if not isinstance(host_, bytes) else host_.decode(), 0)):
            raise AssertionError("the test suite must never reach the network")
        return real_getaddrinfo(host_, *args, **kwargs)
    monkeypatch.setattr(socket.socket, "connect", connect)
    monkeypatch.setattr(socket.socket, "connect_ex", connect_ex)
    monkeypatch.setattr(socket, "getaddrinfo", getaddrinfo)


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
