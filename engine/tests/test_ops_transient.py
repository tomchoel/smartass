import pytest

from smartass.core import config
from smartass.epm import ops
from smartass.xl.host import HostError, ProviderEmpty


class _H:
    pid = 4242

    def check_signin(self):
        pass


def test_a_752_that_survives_the_retry_raises_provider_empty(monkeypatch):
    monkeypatch.setattr(config, "TRANSIENT_RETRY_WAIT_S", 0)
    monkeypatch.setattr(ops, "ensure_grid_sheet", lambda h, conn, fresh=False: None)

    def fail(*a, **k):
        raise HostError("Retrieve failed rc=752: Empty response from provider")
    monkeypatch.setattr(ops, "_query_once", fail)
    with pytest.raises(ProviderEmpty) as e:
        ops.query(_H(), "FinMax", {}, [], [], True, True)
    assert e.value.pid == 4242


def test_a_752_that_clears_on_the_retry_returns(monkeypatch):
    monkeypatch.setattr(config, "TRANSIENT_RETRY_WAIT_S", 0)
    monkeypatch.setattr(ops, "ensure_grid_sheet", lambda h, conn, fresh=False: None)
    calls = []

    def once(*a, **k):
        calls.append(1)
        if len(calls) == 1:
            raise HostError("rc=752 Empty response from provider")
        return "grid"
    monkeypatch.setattr(ops, "_query_once", once)
    assert ops.query(_H(), "FinMax", {}, [], [], True, True) == "grid"


def test_a_connect_that_keeps_answering_empty_raises_provider_empty(monkeypatch):
    monkeypatch.setattr(ops.time, "sleep", lambda s: None)

    class H(_H):
        def connect_sheet(self, sheet, conn):
            raise HostError("Connect failed (rc=752 Empty response from provider)")
    with pytest.raises(ProviderEmpty):
        ops._connect_with_retry(H(), "Q_FinMax", "FinMax")
