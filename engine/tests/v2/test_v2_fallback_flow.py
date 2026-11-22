import asyncio

from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass import server
from smartass.data.provider import ComboProvider, FetchResult
from smartass.pipeline import prepare
from smartass.epm.svhttp import SvHttpError


def test_combo_provider_prefers_the_fixed_result_and_delegates_the_rest():
    base = FakeProvider(facts())
    from smartass.model.requests import GridRequest
    a = GridRequest("fm", "FinMax", ("ac.C30",), ("un.SY.B1",), ("Sep 26",), "Actual", "", "Periodic")
    b = GridRequest("fm", "FinMax", ("ac.C30",), ("un.SY.B1",), ("Aug 26",), "Actual", "", "Periodic")
    fixed = {a: FetchResult(a, [], ["fixed"], "adhoc")}
    combo = ComboProvider(base, fixed)
    assert combo.fetch(a).route == "adhoc" and base.asked == []
    assert combo.fetch(b).route == "http" and base.asked == [b]
    assert combo.closed_months() == base.closed_months()


class Flaky(FakeProvider):
    def fetch(self, req):
        if req.periods == ("Sep 25",):
            raise SvHttpError("lost the request")
        return super().fetch(req)


def _wire(monkeypatch, provider, jobs):
    monkeypatch.setattr(server, "run_prepare",
                        lambda raw, prov, lines=None, **kw: prepare(raw, resolver(), prov, lines=lines, **kw))

    async def user(fn, what):
        jobs.append(what)
        return fn(None)
    monkeypatch.setattr(server, "_user", user)


def test_failed_grids_fall_back_to_adhoc_once(monkeypatch):
    jobs, pulled = [], []
    provider = Flaky(facts())
    _wire(monkeypatch, provider, jobs)

    def fake_adhoc(h, reqs, target):
        pulled.append((list(reqs), target))
        return {r: FetchResult(r, FakeProvider(facts()).fetch(r).facts, [], "adhoc") for r in reqs}
    monkeypatch.setattr(server, "adhoc_fetch", fake_adhoc)
    pre, notes = asyncio.run(server._prepare(spec_dict(), provider, target="active"))
    assert jobs == ["ad hoc fallback"] and len(pulled) == 1 and pulled[0][1] == "active"
    assert pre.routes["adhoc"] == 1 and pre.failed == []


def test_no_failures_no_adhoc_job(monkeypatch):
    jobs = []
    _wire(monkeypatch, None, jobs)
    monkeypatch.setattr(server, "adhoc_fetch", lambda *a: (_ for _ in ()).throw(AssertionError("no fallback")))
    pre, _ = asyncio.run(server._prepare(spec_dict(), FakeProvider(facts())))
    assert jobs == [] and "adhoc" not in pre.routes


def test_a_second_round_failure_stays_in_the_warnings(monkeypatch):
    jobs = []
    _wire(monkeypatch, None, jobs)
    monkeypatch.setattr(server, "adhoc_fetch", lambda h, reqs, target: {})   # nothing fixed: the live retry fails again
    pre, _ = asyncio.run(server._prepare(spec_dict(), Flaky(facts())))
    assert jobs == ["ad hoc fallback"] and pre.failed
    assert any("could not be fetched over HTTP" in w for w in pre.warnings)


def test_the_fallback_round_reuses_the_grids_http_already_fetched(monkeypatch):
    class Counting(Flaky):
        def fetch(self, req):
            self.asked.append(req)
            if req.periods == ("Sep 25",):
                raise SvHttpError("lost the request")
            return FakeProvider(self.frame).fetch(req)

    jobs = []
    provider = Counting(facts())
    _wire(monkeypatch, provider, jobs)
    monkeypatch.setattr(server, "adhoc_fetch", lambda h, reqs, target: {
        r: FetchResult(r, FakeProvider(facts()).fetch(r).facts, [], "adhoc") for r in reqs})
    pre, _ = asyncio.run(server._prepare(spec_dict(), provider))
    assert len(provider.asked) == len(set(provider.asked)) == len(pre.requests)     # each request asked exactly once
    assert pre.failed == [] and pre.routes["adhoc"] == 1 and pre.routes["http"] == len(pre.requests) - 1


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


def test_a_grid_rest_fetched_needs_no_excel_job(monkeypatch):
    jobs = []
    monkeypatch.setattr(server, "run_prepare",
                        lambda raw, prov, lines=None, **kw: prepare(raw, resolver(), prov, lines=lines, **kw))

    async def user(fn, what):
        jobs.append(what)
        raise AssertionError("Excel was used")
    monkeypatch.setattr(server, "_user", user)
    pre, _ = asyncio.run(server._prepare(spec_dict(), _Rest(facts(), lambda req: req.periods[0] == "Sep 25")))
    assert jobs == [] and pre.routes.get("rest") and pre.failed == []
