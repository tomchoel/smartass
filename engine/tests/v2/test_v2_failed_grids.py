from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass.data.provider import facts_from_parsed
from smartass.model.requests import GridRequest
from smartass.pipeline import prepare
from smartview_mcp.svhttp import SvHttpError

REQ = GridRequest("fm", "FinMax", ("ac.C30",), ("E1",), ("Sep 26",), "Actual", "", "Periodic")


class Parsed:
    columns = ["All Accounts", "All Entities", "Sep 26"]
    rows = [["ac.C30", "E1", 5.0], ["ac.C30", "E2", None]]
    unknown_members = []


def test_facts_from_parsed_is_shared():
    out, warnings = facts_from_parsed(REQ, Parsed(), {"Sep 26": "Sep 26"})
    assert [(f.entity, f.value) for f in out] == [("E1", 5.0)] and warnings == []


class FailingProvider(FakeProvider):
    def fetch(self, req):
        if req.periods == ("Sep 25",):
            raise SvHttpError("Smart View provider at x: lost the request (TimeoutError)")
        return super().fetch(req)


def test_a_failed_grid_is_recorded_not_raised():
    pre = prepare(spec_dict(), resolver(), FailingProvider(facts()))
    assert [r.periods for r in pre.failed] == [("Sep 25",)]
    assert pre.failed_errors and "lost the request" in pre.failed_errors[0]
