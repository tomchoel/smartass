# engine/tests/v2/test_v2_workbench.py
from smartass import workbench


def test_snapshot_captures_without_mutating(monkeypatch):
    calls = []

    class B:
        name = "Book1.xlsx"
        app = object()

        def sheet(self, name):
            calls.append(("sheet", name))
            return "WS"

    monkeypatch.setattr(workbench, "_find_book", lambda wb: B())
    monkeypatch.setattr(workbench, "_wait_ready", lambda app: calls.append("ready"))
    monkeypatch.setattr(workbench, "_range", lambda ws, address: ("RNG", address or "A1:F30"))
    monkeypatch.setattr(workbench, "_capture", lambda ws, rng: {"snap": 1})
    monkeypatch.setattr(workbench, "_record", lambda b, ws, address, snap, label: {"id": "c1", "address": address})
    out = workbench.snapshot("Book1", "Sheet1", None, label="before Excel MCP edit")
    assert out == {"change_id": "c1", "workbook": "Book1.xlsx", "sheet": "Sheet1", "address": "A1:F30"}
    assert ("sheet", "Sheet1") in calls and "ready" in calls


def test_snapshot_without_sheet_names_the_active_sheet(monkeypatch):
    class B:
        name = "Book1.xlsx"
        app = object()

        def sheet(self, name):
            return "WS"

    monkeypatch.setattr(workbench, "_find_book", lambda wb: B())
    monkeypatch.setattr(workbench, "_wait_ready", lambda app: None)
    monkeypatch.setattr(workbench, "_range", lambda ws, address: ("RNG", "A1:C3"))
    monkeypatch.setattr(workbench, "_capture", lambda ws, rng: {})
    monkeypatch.setattr(workbench, "_record", lambda b, ws, address, snap, label: {"id": "c2"})
    monkeypatch.setattr(workbench, "_sheet_name", lambda ws: "Active")
    assert workbench.snapshot(None, None, None)["sheet"] == "Active"


def test_revert_delegates(monkeypatch):
    monkeypatch.setattr(workbench, "_history_revert", lambda cid, wb, force: {"reverted": cid, "force": force})
    assert workbench.revert("c1", "Book1", force=True) == {"reverted": "c1", "force": True}


def test_server_tools_route_through_user_excel(monkeypatch):
    import asyncio

    from smartass import server

    async def user(fn, what):
        return {"what": what, "out": fn(None)}
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(workbench, "snapshot", lambda wb, sh, rg: ("snap", wb, sh, rg))
    monkeypatch.setattr(workbench, "revert", lambda cid, wb, force: ("rev", cid, wb, force))
    assert asyncio.run(server.xl_snapshot("B", "S", "A1")) == {"what": "xl_snapshot", "out": ("snap", "B", "S", "A1")}
    assert asyncio.run(server.xl_revert("c1", "B", True)) == {"what": "xl_revert", "out": ("rev", "c1", "B", True)}
