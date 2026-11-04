from smartass.data import adhoc
from smartass.model.requests import GridRequest

REQ = GridRequest("fm", "FinMax", ("ac.C30",), ("E1",), ("Sep 26",), "Actual", "", "Periodic")


class Parsed:
    columns = ["All Accounts", "All Entities", "Sep 26"]
    rows = [["ac.C30", "E1", 7.0]]
    unknown_members = []


class Pull:
    sheet = "Q4 FinMax"
    parsed = Parsed()


def test_adhoc_fetch_pulls_parses_and_removes_the_sheet(monkeypatch):
    log = []
    monkeypatch.setattr(adhoc, "_target", lambda h, target: "T")
    monkeypatch.setattr(adhoc, "_next_q", lambda t: 4)
    monkeypatch.setattr(adhoc, "_pull", lambda h, t, q, conn, pov, rows, cols: log.append(("pull", conn)) or Pull())
    monkeypatch.setattr(adhoc, "_drop", lambda t, name: log.append(("drop", name)))
    got = adhoc.adhoc_fetch(None, [REQ], "session")
    assert [(f.value, f.period) for f in got[REQ].facts] == [(7.0, "Sep 26")]
    assert got[REQ].route == "adhoc"
    assert log == [("pull", "FinMax"), ("drop", "Q4 FinMax")]


def test_one_failed_adhoc_grid_does_not_stop_the_rest(monkeypatch):
    monkeypatch.setattr(adhoc, "_target", lambda h, target: "T")
    monkeypatch.setattr(adhoc, "_next_q", lambda t: 4)

    def boom(*a):
        raise RuntimeError("Smart View not installed")
    monkeypatch.setattr(adhoc, "_pull", boom)
    monkeypatch.setattr(adhoc, "_drop", lambda t, name: None)
    got = adhoc.adhoc_fetch(None, [REQ], "session")
    assert got[REQ].facts == [] and "Smart View not installed" in got[REQ].warnings[0]
