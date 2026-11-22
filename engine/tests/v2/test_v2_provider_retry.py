# engine/tests/v2/test_v2_provider_retry.py
import pytest
from smartass.epm.svhttp import SvHttpError

from smartass.data import provider as prov
from smartass.data.provider import SvHttpProvider
from smartass.model.requests import GridRequest

REQ = GridRequest("fm", "FinMax", ("ac.C30",), ("E1",), ("Sep 26",), "Actual", "", "Periodic")


class Parsed:
    columns = ["All Accounts", "All Entities", "Sep 26"]
    rows = [["ac.C30", "E1", 5.0]]
    unknown_members = []


class SV:
    def __init__(self, fails, text="lost the request (TimeoutError: signal timed out)"):
        self.fails, self.text, self.calls = fails, text, 0

    def fetch_rest(self, *a):                   # the second engine fails too (tests/test_second_engine.py: when not)
        self.rest_calls = getattr(self, "rest_calls", 0) + 1
        raise SvHttpError("Planning REST at x: answered HTTP 503")

    def retrieve(self, *a):
        self.calls += 1
        if self.calls <= self.fails:
            raise SvHttpError(f"Smart View provider at x: {self.text}")
        return None, Parsed(), 0.1


def test_one_lost_request_is_retried(monkeypatch):
    monkeypatch.setattr(prov, "RETRY_WAIT_S", 0)
    sv = SV(fails=1)
    got = SvHttpProvider(sv).fetch(REQ)
    assert sv.calls == 2 and [f.value for f in got.facts] == [5.0]
    assert any("retried" in w for w in got.warnings)


def test_two_lost_requests_fail_and_other_errors_fail_at_once(monkeypatch):
    monkeypatch.setattr(prov, "RETRY_WAIT_S", 0)
    sv = SV(fails=2)
    with pytest.raises(SvHttpError, match="lost the request.*could not fetch it either"):
        SvHttpProvider(sv).fetch(REQ)
    assert sv.calls == 2 and sv.rest_calls == 1
    sv = SV(fails=1, text="rejected members: ac.X")
    with pytest.raises(SvHttpError):
        SvHttpProvider(sv).fetch(REQ)
    assert sv.calls == 1 and sv.rest_calls == 1
