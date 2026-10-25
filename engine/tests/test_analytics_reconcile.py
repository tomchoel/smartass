"""analytics.reconcile: identities per entity and month, months = YTD, areas vs the company total."""
from analytics_util import E, frame

from smartview_mcp import catalog
from smartview_mcp.analytics import reconcile

CHECKS = catalog.load().checks["FinMax"]
GOOD = {"rev": 100.0, "cogs": -40.0, "gm": 60.0, "pers": -20.0, "oth": -10.0, "opex": -30.0, "boi": 30.0,
        "c73": -5.0, "ebitda": 25.0}


def test_identities_pass_and_name_the_broken_cell():
    v = {(E, k, "Sep 26", "actual"): x for k, x in GOOD.items()}
    v.update({(E, k, "Aug 26", "actual"): x for k, x in {**GOOD, "gm": 70.0}.items()})
    n, problems = reconcile.identities(frame(values=v), CHECKS)
    assert n == 8 and len(problems) == 2                 # Aug 26: C30 + C40S = C48 and C48 + C70 = EBITDA_BOI
    assert problems[0].startswith(f"FinMax {E} Aug 26 actual: C30 + C40S = C48 is off by")
    partial = {(E, "rev", "Jul 26", "actual"): 1.0}
    assert reconcile.identities(frame(values=partial), CHECKS) == (0, [])


def test_months_vs_ytd():
    key = lambda p, b, a="ac.C30": ("FinMax", "Actual", "", E, a, p, b)  # noqa: E731
    periodic = {key("Jan 26", "Periodic"): 10.0, key("Feb 26", "Periodic"): None, key("Mar 26", "Periodic"): 5.0,
                key("Jan 26", "Periodic", "ac.C60"): 1.0}
    acc = {key("Mar 26", "Accumulated"): 15.0, key("Mar 26", "Accumulated", "ac.C60"): 9.0}
    n, problems = reconcile.months_vs_ytd(periodic, acc, ["Jan 26", "Feb 26", "Mar 26"])
    assert n == 1 and problems == []                      # ac.C60 lacks Feb and Mar: not checked
    acc[key("Mar 26", "Accumulated")] = 16.0
    assert "differs from the sum of Jan 26..Mar 26 by 1" in reconcile.months_vs_ytd(
        periodic, acc, ["Jan 26", "Feb 26", "Mar 26"])[1][0]


def test_areas_vs_total_is_information():
    v = {("A", "ebitda", "Sep 26", "actual"): 10.0, ("B", "ebitda", "Sep 26", "actual"): 5.0,
         ("un.SY", "ebitda", "Sep 26", "actual"): 16.0}
    info = reconcile.areas_vs_total(frame(values=v), ["A", "B"], "un.SY", "ebitda", "Sep 26", "actual")
    assert "less the 2 areas is 1 " in info
    assert reconcile.areas_vs_total(frame(values=v), ["A", "C"], "un.SY", "ebitda", "Sep 26", "actual") is None
