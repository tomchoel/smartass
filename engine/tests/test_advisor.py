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
