"""Phase 4: semantic layer, period parser, learned defaults, the advisor gate (benchmark table) and commentary."""
from __future__ import annotations

import pytest

from smartview_mcp.core import semantic


def test_measure_synonyms_resolve_english_and_norwegian():
    assert [m.key for m in semantic.measures("EBITDA for Business")] == ["ebitda"]
    assert [m.key for m in semantic.measures("driftskostnader i IT")] == ["opex"]
    assert [m.key for m in semantic.measures("Cost per FTE in IT")] == ["opex_per_fte"]   # longest match wins
    vague = semantic.measures("How are we doing on costs?")
    assert {m.key for m in vague} == {"costs", "numbers"} and all(m.vague for m in vague)
    assert [m.key for m in semantic.measures("How much did we spend on consultants?")] == ["consultants"]


def test_grain_feasibility_and_proxy():
    g = semantic.grains("EBITDA per product")[0]
    ok, proxy = semantic.feasible("ebitda", g)
    assert g.key == "product" and not ok and "contribution margin" in proxy
    assert semantic.feasible("revenue", g) == (True, None)
    assert not semantic.feasible("ebitda_margin", semantic.grains("per customer")[0])[0]


def test_currency_only_nok():
    assert semantic.currency("Revenue in EUR for Technology") == ("EUR", False)
    assert semantic.currency("Revenue for Technology") == ("NOK", True)


def test_capabilities_reference_only_known_concepts():
    assert semantic.problems() == []


def test_describe_states_the_definition():
    assert "C693" in semantic.describe(semantic.measure("consultants"))
    assert semantic.describe(semantic.measure("opex_per_fte")).startswith("Opex per FTE")

from smartview_mcp.core import period  # noqa: E402


@pytest.mark.parametrize("text,kind,label,gap", [
    ("Business EBITDA Sep 26 vs Sep 25", "month", "Sep 26", None),
    ("EBITDA YTD September 2026", "ytd", "YTD Sep 26", None),
    ("Q3 numbers for Consumer", "quarter", "Q3 26", None),
    ("full-year EBITDA latest forecast", "fy", "FY 26", None),
    ("Opex for Wholesale this year", "this_year", None, "YTD to Sep 26, or full year (actual + forecast)?"),
    ("Cost per FTE in IT", "none", "YTD Sep 26", None),
    ("Forecast for October 26 actuals", "month", "Oct 26", None),
    ("Driftskostnader hittil i år", "ytd", "YTD Sep 26", None),
    ("What may we expect in mars 2026?", "month", "Mar 26", None),
    ("EBITDA per month 2026 vs 2025", "year", "2026", None),
])
def test_period_parse(text, kind, label, gap):
    p = period.parse(text, "Sep 26")
    assert (p.kind, p.label, p.gap) == (kind, label, gap)


def test_open_month_is_flagged():
    assert period.parse("Oct 26 actuals", "Sep 26").open and not period.parse("Sep 26", "Sep 26").open
    assert period.parse("Cost per FTE", "Sep 26").convention


def test_remember_then_lookup_roundtrip(tmp_path, monkeypatch):
    from smartview_mcp.core import prefs
    monkeypatch.setattr(prefs.config, "PREFERENCES", tmp_path / "p.toml")
    assert prefs.lookup("measure", "costs") is None
    prefs.remember("measure", "Costs", "opex")
    assert prefs.lookup("measure", "costs") == "opex"
    prefs.remember("measure", "costs", "personnel")          # replaces, never duplicates
    assert prefs.lookup("measure", "COSTS") == "personnel"
    assert (tmp_path / "p.toml").read_text(encoding="utf-8").count("[[default]]") == 1


def test_remember_rejects_unknown_item_or_value(tmp_path, monkeypatch):
    from smartview_mcp.core import prefs
    from smartview_mcp.host import InputError
    monkeypatch.setattr(prefs.config, "PREFERENCES", tmp_path / "p.toml")
    with pytest.raises(InputError):
        prefs.remember("colour", "x", "y")
    with pytest.raises(InputError):
        prefs.remember("measure", "costs", "not_a_measure")
    with pytest.raises(InputError):
        prefs.remember("measure", "costs", "numbers")          # a vague measure is no answer
    with pytest.raises(InputError):
        prefs.remember("period", "this year", "sometimes")
    assert not (tmp_path / "p.toml").exists()


from smartview_mcp.core import scope  # noqa: E402

GATE_CASES = [
    ("B09", "What do we earn on mobile subscriptions?", "clear"),
    ("B10", "EBITDA per product for Mobile Consumer", "infeasible"),
    ("B11", "How are we doing on costs?", "ambiguous"),
    ("B13", "EBITDA YTD for unit un.SY.5OK_W", "clear"),
    ("B14", "Business full-year EBITDA: latest forecast vs budget", "clear"),
    ("T01", "What's our budget for Business?", "ambiguous"),
    ("T02", "Opex for Wholesale this year", "ambiguous"),
    ("T03", "EBITDA margin per customer", "infeasible"),
    ("T04", "Revenue in EUR for Technology", "infeasible"),
    ("T05", "Profit per product", "infeasible"),
    ("T06", "Business EBITDA Sep 26 vs Sep 25", "clear"),
    ("T07", "Q3 numbers for Consumer", "ambiguous"),
    ("T08", "Cost per FTE in IT", "clear"),
    ("T09", "How much did we spend on consultants?", "clear"),
    ("T10", "Forecast for October 26 actuals", "infeasible"),
]
_NONE = lambda i, p: None  # noqa: E731


@pytest.mark.parametrize("cid,q,want", GATE_CASES, ids=[c[0] for c in GATE_CASES])
def test_gate_outcome(cid, q, want):
    """The advisor benchmark: if a case fails, fix the rule or capabilities.toml, never the expected outcome."""
    g = scope.gate(q, "Sep 26", learned=_NONE)
    assert g["suggested_outcome"] == want, g
    scope.to_contract(g)                                       # always a valid contract Gate


def test_gate_learned_default_turns_an_ask_into_a_stated_assumption():
    d = {("measure", "costs"): "opex", ("period", "this year"): "ytd"}
    g = scope.gate("How are we doing on costs this year?", "Sep 26", learned=lambda i, p: d.get((i, p)))
    assert g["suggested_outcome"] == "clear"
    assert g["checklist"]["measure"]["source"] == g["checklist"]["period"]["source"] == "learned"
    assert g["checklist"]["period"]["value"] == "YTD Sep 26"


def test_gate_details():
    t02 = scope.gate("Opex for Wholesale this year", "Sep 26", learned=_NONE)
    assert len(t02["questions"]) == 1 and "full year" in t02["questions"][0]
    b13 = scope.gate("EBITDA YTD for unit un.SY.5OK_W", "Sep 26", learned=_NONE)
    assert any("last year only" in w for w in b13["warnings"]) and b13["entities"][0]["finplan"] is None
    t10 = scope.gate("Forecast for October 26 actuals", "Sep 26", learned=_NONE)
    assert "Oct 26" in t10["reason"] and "forecast" in t10["proxy"]
    t08 = scope.gate("Cost per FTE in IT", "Sep 26", learned=_NONE)
    assert t08["checklist"]["measure"]["value"].startswith("Opex per FTE")     # the definition is stated
    t09 = scope.gate("How much did we spend on consultants?", "Sep 26", learned=_NONE)
    assert "C693" in t09["checklist"]["measure"]["value"] and "temporary" in t09["checklist"]["measure"]["value"]
    b09 = scope.gate("What do we earn on mobile subscriptions?", "Sep 26", learned=_NONE)
    assert b09["checklist"]["measure"]["source"] == "catalog" and b09["concepts"]
    assert b09["checklist"]["entity"]["source"] == b09["checklist"]["period"]["source"] == "convention"
    t06 = scope.gate("Business EBITDA Sep 26 vs Sep 25", "Sep 26", learned=_NONE)
    assert "last year" in t06["checklist"]["scenario"]["value"]
    b14 = scope.gate("Business full-year EBITDA: latest forecast vs budget", "Sep 26", learned=_NONE)
    assert b14["checklist"]["scenario"]["value"] == "Latest forecast vs Target (budget)"
