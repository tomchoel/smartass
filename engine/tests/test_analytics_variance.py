"""analytics.variance and analytics.materiality: sign conventions, zero comparisons, the 10 % / NOK 500k floor /
NOK 5m rule and the always-lines."""
import dataclasses

import pytest

from smartview_mcp import catalog
from smartview_mcp.analytics.facts import Fact, missing
from smartview_mcp.analytics.materiality import assess, flags
from smartview_mcp.analytics.variance import variance

LINES = {x.key: x for x in catalog.load().pnl["FinMax"]}
M = catalog.load().materiality


def v(line, cur, cmp, unit="nok"):
    return variance(LINES[line], Fact("c", cur, unit, "cur"), Fact("p", cmp, unit, "cmp"), line)


def test_sign_convention_costs_negative_higher_is_better():
    assert v("rev", 120.0, 100.0).favourable is True
    cost = v("oth", -120.0, -100.0)                       # costs up by 20 (more negative)
    assert cost.abs.value == -20.0 and cost.favourable is False and cost.pct.value == pytest.approx(-0.2)
    assert v("fte", 12.0, 10.0, "fte").favourable is None  # FTE has no better direction
    lower = variance(dataclasses.replace(LINES["oth"], better="lower"), Fact("c", 5.0, "nok", ""),
                     Fact("p", 7.0, "nok", ""), "x")
    assert lower.favourable is True


def test_zero_comparison_has_no_pct_and_missing_parts_are_explicit():
    z = v("oth", -1000.0, 0.0)
    assert z.abs.value == -1000.0 and z.pct.value is None and "zero" in z.pct.reason
    gone = variance(LINES["rev"], Fact("c", 1.0, "nok", ""), missing("p", "nok", "", "not read"), "rev")
    assert gone.abs.value is None and gone.pct.value is None and gone.favourable is None


@pytest.mark.parametrize("cur, cmp, material", [
    (-1_200_000.0, -1_000_000.0, False),    # 20 % but only NOK 200k: under the 500k floor
    (-6_000_000.0, -5_000_000.0, True),     # 20 % and NOK 1m: the % rule
    (-106_000_000.0, -100_000_000.0, True),  # 6 % but NOK 6m: the NOK 5m rule
    (-104_000_000.0, -100_000_000.0, False),  # 4 % and NOK 4m: neither
    (-5_000_000.0, 0.0, True),              # zero comparison: only the NOK rule, and 5m is enough
    (-400_000.0, 0.0, False),
    (-1_100_000.0, -1_000_000.0, False),    # 10 % of 1m but NOK 100k: under the floor
    (-5_500_000.0, -5_000_000.0, True),     # exactly 10 % and exactly the NOK 500k floor
    (-5_499_999.0, -5_000_000.0, False),    # one krone under both
    (-7_000_000.0, -1_000_000.0, True),     # NOK 6m and 600 %
    (-450_000.0, -300_000.0, False),        # base under NOK 500k counts as zero: NOK 150k, no %
    (-5_300_000.0, -300_000.0, True),       # small base, but the NOK 5m rule still applies
])
def test_materiality_rule_either_with_the_floor(cur, cmp, material):
    assert assess(v("oth", cur, cmp), M).material is material


def test_boundaries_and_small_base_reason():
    assert assess(v("oth", -105_000_000.0, -100_000_000.0), M).material is True       # exactly NOK 5m
    assert assess(v("oth", -104_999_999.0, -100_000_000.0), M).material is False
    small = assess(v("oth", -450_000.0, -300_000.0), M)
    assert small.material is False and "counts as zero" in small.reason
    pct_only = assess(v("oth", -6_000_000.0, -5_000_000.0), M)
    assert "at least" not in pct_only.reason and "20% of the comparison" in pct_only.reason
    nok_only = assess(v("oth", -106_000_000.0, -100_000_000.0), M)
    assert "% of the comparison" not in nok_only.reason and nok_only.basis == "ytd"


def test_month_swing_flags_on_its_own():
    ytd, month = v("oth", -1_100_000.0, -1_000_000.0), v("oth", -5_000_000.0, 0.0)
    f = assess(ytd, M, month)
    assert f.material and f.basis == "month" and "single month" in f.reason
    assert assess(ytd, M, v("oth", -4_999_999.0, 0.0)).material is False
    assert assess(v("oth", -6_000_000.0, -5_000_000.0), M, month).basis == "ytd+month"


def test_flags_rank_material_by_the_larger_of_ytd_and_month():
    months = {"rev": v("rev", 20e6, 0.0)}
    fs = flags([v("rev", 101.0, 100.0), v("oth", -9e6, -1e6)], M, months)
    assert [f.line for f in fs] == ["rev", "oth"] and fs[0].size == 20e6 and fs[0].basis == "month"


def test_both_rule_zero_comparison_is_consistent():
    both = dataclasses.replace(M, rule="both")
    assert assess(v("oth", -50_000_000.0, 0.0), both).material is False        # no % possible: both cannot hold


def test_flags_skip_subtotals_but_keep_always_lines():
    vs = [v("gm", -9e6, -1e6), v("ebitda", 1.0, 0.0), v("boi", -9e6, -1e6), v("oth", -9e6, -1e6)]
    kept = [f.line for f in flags(vs, M, lines=catalog.load().pnl["FinMax"])]
    assert sorted(kept) == ["ebitda", "oth"]
    assert {f.line for f in flags(vs, M)} == {"gm", "ebitda", "boi", "oth"}


def test_lower_is_better_flips_the_variance():
    kpi = next(k for k in catalog.load().kpis if k.better == "lower")
    up = variance(kpi, Fact("c", 130.0, "nok_per_fte", ""), Fact("p", 120.0, "nok_per_fte", ""), "k")
    assert up.abs.value == -10.0 and up.favourable is False and up.pct.value == pytest.approx(-10 / 120)
    down = variance(kpi, Fact("c", 120.0, "nok_per_fte", ""), Fact("p", 130.0, "nok_per_fte", ""), "k")
    assert down.abs.value == 10.0 and down.favourable is True
    assert "comparison - current" in up.abs.definition


def test_rule_both_needs_both_and_always_lines_are_commented():
    both = dataclasses.replace(M, rule="both")
    assert assess(v("oth", -6_000_000.0, -5_000_000.0), both).material is False   # 1m < 5m
    assert assess(v("oth", -60_000_000.0, -50_000_000.0), both).material is True
    small = assess(v("rev", 101.0, 100.0), M)
    assert small.material is False and small.always is True and small.comment is True
    assert assess(v("fte", 30.0, 10.0, "fte"), M).reason == "not an amount: no NOK rule"


def test_flags_rank_by_size():
    fs = flags([v("rev", 101.0, 100.0), v("oth", -9e6, -1e6), v("cogs", -3e6, -1e6)], M)
    assert [f.line for f in fs] == ["oth", "cogs", "rev"] and fs[0].size == 8e6
