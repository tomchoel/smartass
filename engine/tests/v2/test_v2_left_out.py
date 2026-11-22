"""A member EPM refuses over Planning REST (phase 2, Task 3b), from the fetch to the tool reply: the second engine
leaves it out of the request (tests/test_second_engine.py), and the answer says so - one warning per member, by code,
in the reply's warnings and as a caveat (the Answer's Caveats, the Sources sheet's Known limits). The member takes
the path of a missing fact: no Fact. Every member and figure is invented."""
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass.data.provider import FetchResult
from smartass.pipeline import prepare, reply


class LeavesOut(FakeProvider):
    """The pipeline's fake provider, where Planning REST refused some members: each request that asked one was
    fetched over REST without it (no fact for it) and says which (FetchResult.dropped)."""

    def __init__(self, frame, *refused):
        super().__init__(frame)
        self.refused = refused

    def fetch(self, req):
        got = super().fetch(req)
        hit = [a for a in req.accounts if a in self.refused]
        if not hit:
            return got
        return FetchResult(req, [f for f in got.facts if f.account not in hit], [], "rest", 0.01,
                           fallback="Smart View provider: Grid is Invalid.",
                           dropped=[("FinMax", "All Accounts", a) for a in hit])


def _reply(pre):
    return reply(pre, q=1, workbook="C:/x.xlsx", sheets=["Q1 Answer", "Q1 Data"], mismatches=[], lint=[])


def _said(out, member):
    return [w for w in out["warnings"] if f"member {member} " in w]


def test_a_total_over_several_members_is_the_rest_and_the_same_reply_says_it_is_incomplete():
    """Opex = ac.C50S + ac.C60: with ac.C60 refused the cell shows ac.C50S alone - the #Missing path - and the very
    same reply carries the warning that every total with ac.C60 is INCOMPLETE (asserted together: never apart)."""
    full = prepare(spec_dict(), resolver(), FakeProvider(facts()))
    pre = prepare(spec_dict(), resolver(), LeavesOut(facts(), "ac.C60"))
    out = _reply(pre)
    c50s = sum(f.value for f in facts().facts if f.account == "ac.C50S" and f.entity == "un.SY.2BUS"
               and f.period == "Sep 26")
    assert out["key_values"]["P&L · Opex"]["Actual"] == c50s != full.result.key_values["P&L · Opex"]["Actual"]
    said = _said(out, "ac.C60")
    assert len(said) == 1 and "INCOMPLETE" in said[0] and "Opex" in said[0]
    assert "smartass atlas build" in said[0] and "missing from this answer" in said[0]


def test_a_cell_whose_only_member_was_refused_is_empty_not_zero_and_the_reply_says_why():
    pre = prepare(spec_dict(), resolver(), LeavesOut(facts(), "ac.C30"))
    out = _reply(pre)
    assert out["key_values"]["Key figures · Revenue"]["Actual"] is None                        # empty, not 0
    assert not any(f.account == "ac.C30" for f in pre.frame.facts)                             # absent, no 0 fact
    assert pre.result.key_values["Key figures · EBITDA"]["Actual"] == 65.0                     # its own member
    assert len(_said(out, "ac.C30")) == 1


def test_one_warning_and_one_caveat_per_member_however_many_requests_left_it_out():
    pre = prepare(spec_dict(), resolver(), LeavesOut(facts(), "ac.C60", "ac.C30"))
    out = _reply(pre)
    assert sum(len(r.dropped) for r in pre.fetched.values()) >= 4                  # actual and last year: two each
    assert len(_said(out, "ac.C60")) == 1 and len(_said(out, "ac.C30")) == 1
    cav = [c for c in out["caveats"] if c["id"].startswith("left-out:")]
    assert sorted(c["id"] for c in cav) == ["left-out:FinMax:ac.C30", "left-out:FinMax:ac.C60"]
    assert all(c["severity"] == "warn" and "INCOMPLETE" in c["text"] for c in cav)


def test_no_member_left_out_no_such_warning():
    out = _reply(prepare(spec_dict(), resolver(), FakeProvider(facts())))
    assert not any("left out" in w for w in out["warnings"])
    assert not any(c["id"].startswith("left-out:") for c in out["caveats"])
