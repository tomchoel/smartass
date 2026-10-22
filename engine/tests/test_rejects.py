"""rc=1000 handling: exact code match and bisection of rejected members (fake retrieve, no Excel)."""
from smartview_mcp import ops
from smartview_mcp.grid import Axis
from smartview_mcp.host import HostError


def test_is_rc1000():
    assert ops.is_rc1000("ERR|1000|Unsupported Error|H&yperion")
    assert ops.is_rc1000("Retrieve failed (rc=1000 Unsupported Error; ...)")
    assert not ops.is_rc1000("ERR|100000|Unable to connect to the provider.")
    assert not ops.is_rc1000("OK")


def test_bisect_finds_rejected(monkeypatch):
    bad = {"Other", "SDx Revenue"}
    calls = []

    def fake(h, conn, pov, rows, cols, sm, sz, _retried=False):
        calls.append(1)
        if any(m in bad for a in rows + cols for m in a.members):
            raise HostError("Retrieve failed (rc=1000 Unsupported Error)")
        return "lay", "parsed", 0.1

    monkeypatch.setattr(ops, "_query_once", fake)
    rows = [Axis("Account", ["Copper", "Other", "TV", "Fiber", "SDx Revenue", "HFC", "Mobile"])]
    cols = [Axis("Years", ["FY25", "FY26"])]
    found = ops._rejected_members(None, "FinPlan", {}, rows, cols)
    assert found == [("Account", "Other"), ("Account", "SDx Revenue")]
    assert len(calls) <= ops.MAX_BISECT_RETRIEVES


def test_query_drops_and_warns(monkeypatch):
    class P:
        pov_problems: list

    def fake(h, conn, pov, rows, cols, sm, sz, _retried=False):
        if any(m == "Other" for a in rows for m in a.members):
            raise HostError("Retrieve failed (rc=1000 Unsupported Error)")
        p = P()
        p.pov_problems = []
        return [a.members for a in rows], p, 0.1

    monkeypatch.setattr(ops, "_query_once", fake)
    monkeypatch.setattr(ops, "_explain_reject", lambda c, d, m: f"{d}: {m}")
    lay, p, _ = ops.query(None, "FinPlan", {}, [Axis("Account", ["Copper", "Other"])], [Axis("Years", ["FY25"])],
                          True, False)
    assert lay == [["Copper"]]
    assert "Account: Other" in p.pov_problems[0]


def test_connection_error_is_not_bisected(monkeypatch):
    def fake(*a, **k):
        raise HostError("ERR|100000|Unable to connect to the provider.")

    monkeypatch.setattr(ops, "_query_once", fake)
    try:
        ops.query(None, "FinPlan", {}, [Axis("Account", ["A", "B"])], [Axis("Years", ["FY25"])], True, False)
    except HostError as e:
        assert "100000" in str(e)
    else:
        raise AssertionError("expected HostError")
