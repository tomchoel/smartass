"""tests/conftest.py's guards: the suite cannot start Excel or reach the network, whatever a default is. Nothing here
reaches Excel, EPM or the identity provider - that is what is tested."""
import socket
import threading
import urllib.request

import psutil
import pytest

from smartass.ops import atlas_build


def _excels() -> set[int]:
    return {p.pid for p in psutil.process_iter(["name"]) if (p.info["name"] or "").lower() == "excel.exe"}


def test_the_hidden_excel_route_of_the_atlas_build_cannot_start_excel():
    """2026-11-22: run_build reached its hidden-Excel route in a test and started the engine's Excel. Now the start
    fails at once, and no Excel process appears."""
    before = _excels()
    got: list = []

    def run():
        try:
            got.append(atlas_build.run_build(["FinMax"], excel=True))
        except BaseException as e:   # noqa: BLE001 - the guard's assertion, however the host reports it
            got.append(e)
    t = threading.Thread(target=run, daemon=True)
    t.start()
    t.join(60)
    assert not t.is_alive() and got
    assert "the test suite must never start Excel" in repr(got[0]) + "".join(
        str(d.get("error")) for d in (got[0] if isinstance(got[0], list) else []))
    assert _excels() <= before


def test_excel_cannot_be_dispatched_and_the_users_excel_cannot_be_launched():
    from smartass.xl import host, target
    with pytest.raises(AssertionError, match="never start Excel"):
        host.wc.DispatchEx("Excel.Application")
    with pytest.raises(AssertionError, match="never start Excel"):
        target._launch_excel("x.xlsx")
    with pytest.raises(AssertionError, match="never start Excel"):
        host.sso.ensure_global_watcher()


@pytest.mark.parametrize("call", [
    lambda: socket.create_connection(("login.example.com", 443), timeout=1),
    lambda: socket.create_connection(("203.0.113.7", 443), timeout=1),
    lambda: urllib.request.urlopen("https://plan1.example.com/HyperionPlanning/", timeout=1),
])
def test_the_network_cannot_be_reached(call):
    with pytest.raises(AssertionError, match="never reach the network"):
        call()


def test_this_machine_can_still_be_reached():
    srv = socket.socket()
    srv.bind(("127.0.0.1", 0))
    srv.listen(1)
    try:
        with socket.create_connection(srv.getsockname(), timeout=2):
            pass
    finally:
        srv.close()
