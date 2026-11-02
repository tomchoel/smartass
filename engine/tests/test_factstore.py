"""factstore.py: the in-memory facts of one pnl recipe call. Every figure is invented."""
from smartview_mcp.factstore import OPEN, FactStore, Fill, Row


def k(entity="un.SY.2WHS", account="ac.C30", period="Sep 26", balance="Periodic", cube="FinMax", scenario="Actual",
      version=""):
    return (cube, scenario, version, entity, account, period, balance)


def r(key, value, closed="Sep 26", t=1000.0):
    return Row(key, value, "http", t, closed)


def test_put_get_and_upsert():
    store = FactStore()
    assert store.put([r(k(), 100.0), r(k(account="ac.C40S"), -40.0), r(k(period="Aug 26"), None)]) == 3
    got = store.get("FinMax", ["un.SY.2WHS"], ["ac.C30", "ac.C40S"], ["Sep 26", "Aug 26"], "Actual", "", "Periodic")
    assert got == {k(): 100.0, k(account="ac.C40S"): -40.0, k(period="Aug 26"): None}
    store.put([r(k(), 120.0)])                                     # same member tuple: replaced
    assert store.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic") == {k(): 120.0}
    assert store.get("FinMax", ["un.SY.2BUS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic") == {}
    assert store.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Accumulated") == {}
    assert store.put([r(k(), 1.0), r(k(), 2.0)]) == 1                # duplicate keys in one put: the last wins
    assert store.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic") == {k(): 2.0}


def test_open_rows_are_kept_out_of_closed_history():
    store = FactStore()
    store.put([r(k(period="Oct 26"), 5.0, closed=OPEN), r(k(), 1.0)])
    ask = ("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26", "Oct 26"], "Actual", "", "Periodic")
    assert set(store.get(*ask)) == {k(), k(period="Oct 26")}
    assert set(store.get(*ask, closed_only=True)) == {k()}


def test_fills_and_freshness():
    store = FactStore()
    store.put([r(k(), 1.0, t=10.0), r(k(account="ac.C40S"), 2.0, t=20.0, closed=OPEN)])
    f = store.freshness("FinMax", ["un.SY.2WHS"])
    assert (f["rows"], f["oldest"], f["newest"], f["open_rows"], f["closed_months"]) == (2, 10.0, 20.0, 1, ["Sep 26"])
    store.record_fill(Fill("live", "FinMax Actual", 5.0, 1.5, 2, "ok", {"problems": []}))
    assert store.fills[0].as_dict()["job"] == "live"


def test_an_open_row_never_replaces_a_closed_period():
    store = FactStore()
    store.put([r(k(), 100.0, closed="Sep 26"), r(k(period="Aug 26"), 5.0, closed=OPEN)])
    wrote = store.put([r(k(), 999.0, closed=OPEN), r(k(period="Aug 26"), 6.0, closed=OPEN),
                       r(k(period="Oct 26"), 7.0, closed=OPEN)])
    assert wrote == 2                                              # the refused row is not counted
    got = store.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26", "Aug 26", "Oct 26"], "Actual", "", "Periodic",
                    closed_only=True)
    assert got == {k(): 100.0}                                     # closed value and label kept
    assert store.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Aug 26", "Oct 26"], "Actual", "", "Periodic") == {
        k(period="Aug 26"): 6.0, k(period="Oct 26"): 7.0}          # open over open, and a new open key, are stored
    store.put([r(k(), 101.0, closed="Sep 26")])                    # a closed read replaces anything
    assert store.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic",
                     closed_only=True) == {k(): 101.0}


def test_plan_months_after_the_closed_month_stay_replaceable():
    store = FactStore()
    plan = k(cube="FinPlan", scenario="Forecast Q326", version="Final", period="Dec 26")
    store.put([r(plan, 100.0, closed="Sep 26")])                   # a forecast month read when Sep 26 was closed
    assert store.put([r(plan, 777.0, closed=OPEN)]) == 1
    assert store.get("FinPlan", ["un.SY.2WHS"], ["ac.C30"], ["Dec 26"], "Forecast Q326", "Final", "Periodic") == {
        plan: 777.0}
