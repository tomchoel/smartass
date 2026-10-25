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
])
def test_materiality_rule_either_with_the_floor(cur, cmp, material):
    assert assess(v("oth", cur, cmp), M).material is material


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
