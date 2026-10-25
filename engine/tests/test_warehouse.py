"""warehouse.py on a temporary DuckDB file, and MemoryWarehouse with the same semantics. Every figure is invented."""
import pytest

from smartview_mcp import config, warehouse as wh
from smartview_mcp.warehouse import OPEN, Fill, MemoryWarehouse, Row, Warehouse, WarehouseUnavailable


def k(entity="un.SY.2WHS", account="ac.C30", period="Sep 26", balance="Periodic", cube="FinMax", scenario="Actual",
      version=""):
    return (cube, scenario, version, entity, account, period, balance)


def r(key, value, closed="Sep 26", t=1000.0):
    return Row(key, value, "http", t, closed)


@pytest.fixture(params=["duckdb", "memory"])
def store(request, tmp_path):
    return Warehouse(tmp_path / "w.duckdb") if request.param == "duckdb" else MemoryWarehouse()


def test_put_get_and_upsert(store):
    assert store.put([r(k(), 100.0), r(k(account="ac.C40S"), -40.0), r(k(period="Aug 26"), None)]) == 3
    got = store.get("FinMax", ["un.SY.2WHS"], ["ac.C30", "ac.C40S"], ["Sep 26", "Aug 26"], "Actual", "", "Periodic")
    assert got == {k(): 100.0, k(account="ac.C40S"): -40.0, k(period="Aug 26"): None}
    store.put([r(k(), 120.0)])                                     # same member tuple: replaced
    assert store.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic") == {k(): 120.0}
    assert store.get("FinMax", ["un.SY.2BUS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic") == {}
    assert store.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Accumulated") == {}


def test_open_rows_are_kept_out_of_closed_history(store):
    store.put([r(k(period="Oct 26"), 5.0, closed=OPEN), r(k(), 1.0)])
    ask = ("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26", "Oct 26"], "Actual", "", "Periodic")
    assert set(store.get(*ask)) == {k(), k(period="Oct 26")}
    assert set(store.get(*ask, closed_only=True)) == {k()}


def test_state_fills_and_freshness(store):
    assert store.state() == {}
    store.set_state(closed_month="Sep 26", forecast="Forecast Q326")
    store.set_state(closed_month="Oct 26")
    assert store.state() == {"closed_month": "Oct 26", "forecast": "Forecast Q326"}
    store.put([r(k(), 1.0, t=10.0), r(k(account="ac.C40S"), 2.0, t=20.0, closed=OPEN)])
    f = store.freshness("FinMax", ["un.SY.2WHS"])
    assert (f["rows"], f["oldest"], f["newest"], f["open_rows"], f["closed_months"]) == (2, 10.0, 20.0, 1, ["Sep 26"])
    store.record_fill(Fill("base", "FinMax Actual", 5.0, 1.5, 2, "ok", {"problems": []}))
    store.record_fill(Fill("top_up", "forecast", 6.0, 0.2, 1, "warn", {"problems": ["x"]}))
    last = store.last_fills(1)[0]
    assert (last["job"], last["result"], last["detail"]) == ("top_up", "warn", {"problems": ["x"]})
    st = store.status()
    assert st["rows"] == 2 and st["state"]["closed_month"] == "Oct 26" and len(st["fills"]) == 2


def test_a_missing_file_reads_as_unavailable_and_status_says_so(tmp_path):
    w = Warehouse(tmp_path / "none.duckdb")
    with pytest.raises(WarehouseUnavailable, match="no warehouse yet"):
        w.get("FinMax", [], [], [], "Actual", "", "Periodic")
    assert w.state() == {} and w.status()["exists"] is False


def test_another_writer_makes_it_unavailable(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "WAREHOUSE_LOCK_WAIT_S", 0.2)
    w = Warehouse(tmp_path / "w.duckdb")
    w.put([r(k(), 1.0)])
    other = wh.Warehouse(tmp_path / "w.duckdb")
    monkeypatch.setattr(other._lock, "acquire", lambda: False)    # another process holds the lock file
    with pytest.raises(WarehouseUnavailable, match="another smartass process"):
        other.put([r(k(), 2.0)])
    assert w.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic") == {k(): 1.0}


def test_duplicate_keys_in_one_put_last_wins(store):
    assert store.put([r(k(), 1.0), r(k(), 2.0)]) in (1, 2)
    assert store.get("FinMax", ["un.SY.2WHS"], ["ac.C30"], ["Sep 26"], "Actual", "", "Periodic") == {k(): 2.0}


def test_put_in_chunks_keeps_every_row(store, monkeypatch):
    monkeypatch.setattr(wh, "PUT_CHUNK", 3)
    rows = [r(k(account=f"ac.{i}"), float(i)) for i in range(8)]
    assert store.put(rows) == 8
    got = store.get("FinMax", ["un.SY.2WHS"], [f"ac.{i}" for i in range(8)], ["Sep 26"], "Actual", "", "Periodic")
    assert len(got) == 8 and got[k(account="ac.7")] == 7.0


def test_a_foreign_file_is_unavailable_at_once_and_status_never_throws(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "WAREHOUSE_LOCK_WAIT_S", 30)       # a retry loop would blow the test's time
    p = tmp_path / "bad.duckdb"
    p.write_bytes(b"not a duckdb file" * 50)
    w = Warehouse(p)
    with pytest.raises(WarehouseUnavailable, match="cannot open"):
        w.get("FinMax", [], [], [], "Actual", "", "Periodic")
    st = w.status()
    assert st["exists"] is True and "cannot open" in st["error"]


def test_a_valid_file_without_tables_is_unavailable_and_status_says_so(tmp_path):
    import duckdb
    p = tmp_path / "empty.duckdb"
    duckdb.connect(str(p)).close()
    w = Warehouse(p)
    for call in (lambda: w.get("FinMax", [], [], [], "Actual", "", "Periodic"), w.freshness, w.last_fills, w.state):
        with pytest.raises(WarehouseUnavailable, match="cannot read"):
            call()
    st = w.status()
    assert st["exists"] is True and "cannot read" in st["error"]
