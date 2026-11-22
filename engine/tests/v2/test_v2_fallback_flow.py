"""A grid the provider route could not fetch: the second engine (Planning REST) or a warning - never Excel (Task 3b
deleted the Excel fallback: data/adhoc.py, xl/pull_adhoc.py and _prepare's ad hoc round)."""
import asyncio
import importlib.util

import pytest
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass import server
from smartass.core import serverkit
from smartass.data import provider as data_provider
from smartass.data.provider import FetchResult
from smartass.epm import drillstep
from smartass.pipeline import prepare, reply
from smartass.epm.svhttp import SvHttpError


class Flaky(FakeProvider):
    def fetch(self, req):
        if req.periods == ("Sep 25",):
            raise SvHttpError("lost the request")
        return super().fetch(req)


@pytest.fixture
def no_excel(monkeypatch):
    """Every way a tool call reaches an Excel (the hidden one, the user's) fails the test."""
    jobs = []

    async def excel(*a, **k):
        jobs.append(k.get("what") or (a[1] if len(a) > 1 else "?"))
        raise AssertionError("Excel was used")
    for mod, name in ((serverkit, "_call"), (serverkit, "_call_user"), (server, "_user"), (drillstep, "_call")):
        monkeypatch.setattr(mod, name, excel)
    monkeypatch.setattr(server, "run_prepare",
                        lambda raw, prov, lines=None, **kw: prepare(raw, resolver(), prov, lines=lines, **kw))
    return jobs


def test_the_excel_fallback_is_gone():
    assert importlib.util.find_spec("smartass.data.adhoc") is None
    assert importlib.util.find_spec("smartass.xl.pull_adhoc") is None
    assert not hasattr(server, "adhoc_fetch") and not hasattr(data_provider, "ComboProvider")


def test_a_grid_neither_engine_could_fetch_is_a_warning_and_no_excel_job(no_excel):
    pre, _ = asyncio.run(server._prepare(spec_dict(), Flaky(facts())))
    assert no_excel == [] and len(pre.failed) == 1
    assert any("could not be fetched" in w and "lost the request" in w for w in pre.warnings)


def test_a_failed_grids_warning_carries_no_session_key_or_query(no_excel):
    """The warning is a reply field (it never passes the tool guard): the engine's one scrub is applied to it."""
    import secrets
    x = secrets.token_urlsafe(24)

    class Leaky(FakeProvider):
        def fetch(self, req):
            if req.periods == ("Sep 25",):
                raise SvHttpError(f"lost the request (sessionId: {x}; see https://h.example/p?code={x})")
            return super().fetch(req)
    pre, _ = asyncio.run(server._prepare(spec_dict(), Leaky(facts())))
    said = " ".join(pre.warnings)
    assert "could not be fetched" in said and "lost the request" in said and x not in said


def test_a_member_that_cannot_be_left_out_reaches_the_reply_by_code_with_what_to_do(no_excel):
    """Ruling (a): the request's only entity refused on both engines - the request fails, the reply says which
    member and what to do, and no Excel is asked to rescue it."""
    why = ("Smart View provider: Grid is Invalid. - the second engine (Planning REST) could not fetch it either "
           "(query fm): Planning REST refused the FinMax grid (Periodic / NOK / Actual): All Entities member "
           "un.SY.2BUS does not exist there or you have no access to it; as the only All Entities member asked it "
           "cannot be left out - check the member or rebuild the atlas (`smartass atlas build`)")

    class Refused(FakeProvider):
        def fetch(self, req):
            raise SvHttpError(why)
    pre, _ = asyncio.run(server._prepare(spec_dict(), Refused(facts())))
    out = reply(pre, q=1, workbook="C:/x.xlsx", sheets=[], mismatches=[], lint=[])
    assert no_excel == [] and len(pre.failed) == len(pre.requests)
    said = [w for w in out["warnings"] if "un.SY.2BUS" in w]
    assert said and "smartass atlas build" in said[0]


class _Rest(FakeProvider):
    """The pipeline's fake provider, with some requests answered by the second engine."""

    def __init__(self, frame, over_rest):
        super().__init__(frame)
        self.over_rest = over_rest

    def fetch(self, req):
        got = super().fetch(req)
        if self.over_rest(req):
            return FetchResult(req, got.facts, [], "rest", 0.01, fallback="Smart View provider: lost the request")
        return got


def test_the_reply_names_the_rest_route_and_one_warning_says_the_second_engine_was_used():
    pre = prepare(spec_dict(), resolver(), _Rest(facts(), lambda req: True))
    n = len(pre.requests)
    assert n >= 2 and pre.routes == {"rest": n}
    said = [w for w in pre.warnings if "second engine" in w]
    assert len(said) == 1 and f"{n} grid(s)" in said[0] and "lost the request" in said[0]
    assert "cannot be told apart from an empty one" in said[0]          # #NoAccess over REST (Oracle documents none)
    assert pre.result.key_values["Key figures · EBITDA"]["Actual"] == 65.0          # the same numbers
    plain = prepare(spec_dict(), resolver(), _Rest(facts(), lambda req: False))
    assert "rest" not in plain.routes and not any("second engine" in w for w in plain.warnings)


def test_a_grid_rest_fetched_needs_no_excel_job(no_excel):
    pre, _ = asyncio.run(server._prepare(spec_dict(), _Rest(facts(), lambda req: req.periods[0] == "Sep 25")))
    assert no_excel == [] and pre.routes.get("rest") and pre.failed == []
