# engine/tests/v2/test_v2_pipeline.py
import pytest
from v2util import facts, resolver, spec_dict

from smartass.data.provider import FetchResult
from smartass.model.validate import GuardError
from smartass import pipeline
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


LINE_CELL = {"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Yr 26"}


def lines_only_spec():
    d = spec_dict(blocks=[{"type": "lines", "cell": LINE_CELL, "group_by": "supplier", "top": 2}])
    d["queries"] = []
    return d


def test_lines_only_spec_needs_no_grid_requests():
    from smartass.model.lines import LineSet
    prov = FakeProvider(facts())
    rows = [{"AP_CUSTOMER_NAME": "Bouvet", "AMOUNT": -30.0}, {"AP_CUSTOMER_NAME": "Accenture", "AMOUNT": -70.0}]
    ls = LineSet("b0", rows, ["AP_CUSTOMER_NAME", "AMOUNT"], cube_value=-100.0, cell=LINE_CELL)
    pre = prepare(lines_only_spec(), resolver(), prov, lines={"b0": ls})
    assert prov.asked == [] and pre.requests == []
    assert pre.frame.lines["b0"] is ls
    assert pre.result.outs[0].type == "lines"
    assert pre.result.key_values["Ledger lines by supplier · Ledger lines total"]["Amount"] == -100.0


def test_spec_needs_a_query_or_a_lines_block():
    d = spec_dict()
    d["queries"] = []
    with pytest.raises(Exception, match="query"):
        prepare(d, resolver(), FakeProvider(facts()))


def test_lines_block_without_its_line_set_is_refused():
    with pytest.raises(GuardError, match="lines"):
        prepare(lines_only_spec(), resolver(), FakeProvider(facts()))


def test_empty_line_set_is_refused_before_excel():
    from smartass.model.lines import LineSet
    ls = LineSet("b0", [], [], cube_value=-100.0, cell=LINE_CELL, warnings=["no drill report"])
    with pytest.raises(GuardError, match="no ledger lines"):
        prepare(lines_only_spec(), resolver(), FakeProvider(facts()), lines={"b0": ls})


class TStored:
    path = "C:/a/v2/x.xlsx"
    wb = type("WB", (), {"ActiveSheet": type("S", (), {"Name": "Q2 Answer"})()})()


def _stored_wire(monkeypatch, session_wb="C:/a/v2/x.xlsx"):
    from smartview_mcp.excel import sheets
    calls = []
    monkeypatch.setattr(pipeline, "_target", lambda h, target, title: calls.append(target) or TStored())
    monkeypatch.setattr(pipeline, "_session_workbook", lambda: session_wb)
    monkeypatch.setattr(sheets, "_stored_specs", lambda t: {2: {}})
    monkeypatch.setattr("smartview_mcp.excel.jobs.com", lambda call, *a, **k: call())
    return calls


def test_read_stored_returns_the_absolute_workbook_path(monkeypatch):
    calls = _stored_wire(monkeypatch)
    stored, active, path = pipeline.read_stored(None, "session")
    assert (stored, active, path) == ({2: {}}, "Q2 Answer", "C:/a/v2/x.xlsx") and calls == ["session"]


def test_followup_refuses_target_new_before_touching_excel(monkeypatch):
    from smartass.model.validate import GuardError
    calls = _stored_wire(monkeypatch)
    with pytest.raises(GuardError, match="new"):
        pipeline.read_stored(None, "new")
    assert calls == []


def test_followup_without_a_session_workbook_creates_no_file(monkeypatch):
    from smartass.model.validate import GuardError
    calls = _stored_wire(monkeypatch, session_wb=None)
    with pytest.raises(GuardError, match="no session workbook yet"):
        pipeline.read_stored(None, "session")
    assert calls == []
