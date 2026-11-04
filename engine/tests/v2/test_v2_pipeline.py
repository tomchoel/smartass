# engine/tests/v2/test_v2_pipeline.py
import pytest
from v2util import facts, resolver, spec_dict

from smartass.data.provider import FetchResult
from smartass.model.validate import GuardError
from smartass.pipeline import prepare, reply


class FakeProvider:
    def __init__(self, frame, warn=None):
        self.frame, self.warn, self.asked = frame, warn, []

    def closed_months(self):
        return {"FinMax": (2026, 9), "FinLite": (2026, 9), "FinPlan": (2026, 9)}

    def fetch(self, req):
        self.asked.append(req)
        got = [f for f in self.frame.facts if f.query == req.query and f.period in req.periods
               and f.balance == req.balance and f.account in req.accounts and f.entity in req.entities]
        return FetchResult(req, got, [self.warn] if self.warn else [], "http", 0.01)


def test_prepare_fetches_every_request_and_computes():
    prov = FakeProvider(facts())
    pre = prepare(spec_dict(), resolver(), prov)
    assert len(prov.asked) == len(pre.requests) >= 2
    assert pre.result.key_values["Key figures · EBITDA"]["Actual"] == 65.0
    assert pre.closed == {"FinMax": (2026, 9), "FinLite": (2026, 9), "FinPlan": (2026, 9)}
    assert pre.routes == {"http": len(prov.asked)}


def test_prepare_guards_before_fetching():
    prov = FakeProvider(facts())
    d = spec_dict()
    d["queries"][0]["entity"] = "un.SY.ZZ"
    with pytest.raises(GuardError):
        prepare(d, resolver(), prov)
    assert prov.asked == []


def test_reply_shape_and_fetch_warnings():
    pre = prepare(spec_dict(), resolver(), FakeProvider(facts(), warn="FinMax x: #NoAccess (left out)"))
    out = reply(pre, q=3, workbook="C:/x.xlsx", sheets=["Q3 Answer", "Q3 Data"], mismatches=[], lint=[])
    assert out["q"] == 3 and out["sheets"] == ["Q3 Answer", "Q3 Data"]
    assert "Key figures · EBITDA" in out["key_values"]
    assert any("#NoAccess" in w for w in out["warnings"])
    assert out["checks"] == {"sheet_equals_model": "OK"}
    assert out["saved_user_workbook"] is False


def test_lines_block_is_refused_in_1b():
    d = spec_dict(blocks=[{"type": "lines", "cell": {"All Accounts": "ac.C60"}}])
    with pytest.raises(GuardError, match="lines"):
        prepare(d, resolver(), FakeProvider(facts()))
