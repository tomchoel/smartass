# engine/tests/v2/test_v2_critic_brief.py
"""Plan 2B task 2: the critic brief - material lines backed by a drivers/lines block or a data driver, gaps for
unbacked lines, unmentioned warn/block caveats and a P&L query without last year; only for standard/deep."""
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass.model.critic import brief
from smartass.pipeline import prepare, reply

CAVEATS = [{"id": "opex_note", "text": "Opex includes the IT reclassification from June.", "severity": "warn"},
           {"id": "bus_block", "text": "2BUS history before the reorg is not comparable.", "severity": "block"},
           {"id": "g910_info", "text": "G910 moved under 2BUS.", "severity": "info"}]
CHECKS = {"sheet_equals_model": "OK"}


def _pre(d=None):
    return prepare(d or spec_dict(), resolver(), FakeProvider(facts()))


def _line(b, key):
    return next(m for m in b["material"] if m["key"] == key)


def test_drivers_block_backs_the_opex_line():
    pre = _pre()
    b = brief(pre.spec, pre.result, CAVEATS, CHECKS)
    opex = _line(b, "b1.m1")
    assert opex["line"] == "Opex" and opex["variance"] == -8.0
    assert opex["backed"] is True and any("b4" in e for e in opex["evidence"])
    assert not any("'Opex'" in g and "no driver" in g for g in b["gaps"])
    assert b["purpose"] == "Explain Business opex vs LY" and b["depth"] == "standard"
    assert b["hypotheses"] == ["Consultants up"]
    assert b["comparisons"] == ["ly"] and b["checks"] == CHECKS


def test_without_a_drivers_block_the_line_is_a_gap_until_a_data_driver_names_it():
    d = spec_dict()
    d["blocks"] = [x for x in d["blocks"] if x["type"] != "drivers"]
    pre = _pre(d)
    b = brief(pre.spec, pre.result, [], CHECKS)
    assert _line(b, "b1.m1")["backed"] is False
    assert "material line 'Opex' has no driver" in b["gaps"]
    d["blocks"][-1] = {"type": "commentary", "key_points": ["Opex up"],
                       "drivers": [{"line": "Opex", "text": "Consultants", "basis": "data", "evidence": "b1.m1.2"}]}
    pre = _pre(d)
    b = brief(pre.spec, pre.result, [], CHECKS)
    opex = _line(b, "b1.m1")
    assert opex["backed"] is True and "b1.m1.2" in opex["evidence"]
    assert b["commentary"] == {"key_points": 1, "drivers": [{"line": "Opex", "basis": "data"}]}


def test_a_hypothesis_driver_does_not_back_a_line():
    d = spec_dict()
    d["blocks"] = [x for x in d["blocks"] if x["type"] != "drivers"]
    d["blocks"][-1] = {"type": "commentary",
                       "drivers": [{"line": "Opex", "text": "Probably consultants", "basis": "hypothesis"}]}
    pre = _pre(d)
    assert _line(brief(pre.spec, pre.result, [], CHECKS), "b1.m1")["backed"] is False


def test_no_commentary_gives_a_gap_per_warn_or_block_caveat():
    pre = _pre()
    b = brief(pre.spec, pre.result, CAVEATS, CHECKS)
    assert [c["id"] for c in b["caveats"]] == ["opex_note", "bus_block"]           # warn/block only
    cav = [g for g in b["gaps"] if g.startswith(("warn caveat", "block caveat"))]
    assert cav == ["warn caveat 'opex_note' not mentioned", "block caveat 'bus_block' not mentioned"]


def test_a_caveat_named_in_the_commentary_is_no_gap():
    d = spec_dict()
    d["blocks"][-1] = {"type": "commentary", "key_points": [
        "Opex includes the IT reclassification from June, so LY is not like-for-like."]}
    pre = _pre(d)
    b = brief(pre.spec, pre.result, CAVEATS, CHECKS)
    gaps = [g for g in b["gaps"] if "caveat" in g]
    assert gaps == ["block caveat 'bus_block' not mentioned"]
    d["blocks"][-1] = {"type": "commentary", "key_points": ["See bus_block: history not comparable."]}
    pre = _pre(d)
    assert not any("bus_block" in g for g in brief(pre.spec, pre.result, CAVEATS, CHECKS)["gaps"])


def test_a_pl_query_without_last_year_is_a_gap():
    d = spec_dict()
    d["queries"][0]["compare"] = []
    d["blocks"] = [{"type": "table", "query": "fm", "columns": ["value"]}]
    pre = _pre(d)
    b = brief(pre.spec, pre.result, [], CHECKS)
    assert b["comparisons"] == []
    assert "query 'fm' (P&L) has no last-year comparison" in b["gaps"]


def test_reply_carries_the_brief_for_standard_and_deep_only():
    pre = _pre()
    out = reply(pre, q=1, workbook="W", sheets=[], mismatches=[], lint=[])
    assert out["critic_brief"]["purpose"] == "Explain Business opex vs LY"
    assert out["critic_brief"]["checks"] == out["checks"]
    d = spec_dict()
    d["plan"] = {"depth": "quick"}
    out = reply(_pre(d), q=1, workbook="W", sheets=[], mismatches=[], lint=[])
    assert "critic_brief" not in out


def test_refresh_adds_one_brief_per_question(monkeypatch):
    import asyncio

    from smartass import server
    stored = {q: {"spec": {"v2": {"question": f"q{q}"}}, "ids": {}, "owned": None} for q in (2, 3)}

    async def user(fn, what):
        return fn(None)
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "_provider", lambda: "provider")
    monkeypatch.setattr(server, "read_stored", lambda h, target="session": (stored, None, "C:/a/x.xlsx"))
    monkeypatch.setattr(server, "run_prepare", lambda raw, provider, lines=None: raw)
    monkeypatch.setattr(server, "redraw_job", lambda h, pre, q, target="session": {
        "q": q, "workbook": "W", "closed": {}, **({"critic_brief": {"purpose": f"p{q}"}} if q == 2 else {})})
    out = asyncio.run(server.fin_refresh())
    r = out["refreshed"]
    assert r[0]["critic_brief"] == {"purpose": "p2"} and "critic_brief" not in r[1]
