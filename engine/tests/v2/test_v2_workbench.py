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

    monkeypatch.setattr(workbench, "_find_book", lambda wb, allow_archive=False: B())
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

    monkeypatch.setattr(workbench, "_find_book", lambda wb, allow_archive=False: B())
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
    monkeypatch.setattr(workbench, "snapshot", lambda wb, sh, rg, allow_archive=False: ("snap", wb, sh, rg))
    monkeypatch.setattr(workbench, "revert", lambda cid, wb, force: ("rev", cid, wb, force))
    assert asyncio.run(server.xl_snapshot("B", "S", "A1")) == {"what": "xl_snapshot", "out": ("snap", "B", "S", "A1")}
    assert asyncio.run(server.xl_revert("c1", "B", True)) == {"what": "xl_revert", "out": ("rev", "c1", "B", True)}


def test_tidy_checks_and_optionally_styles(monkeypatch):
    log = []
    monkeypatch.setattr(workbench, "_format_sheet", lambda wb, sh, last_col, rows, allow_archive=False:
                        log.append("style") or {"change_id": "c9"})
    monkeypatch.setattr(workbench, "_check", lambda wb, sh: {"ok": True, "block": [], "warn": [], "counts": {}})
    assert workbench.tidy("B", "S")["check"]["ok"] is True and log == []
    out = workbench.tidy("B", "S", rows={"1": "title"}, last_col="F")
    assert log == ["style"] and out["styled"]["change_id"] == "c9"


def test_xl_tidy_tool_routes_through_user_excel(monkeypatch):
    import asyncio

    from smartass import server

    async def user(fn, what):
        return {"what": what, "out": fn(None)}
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(workbench, "tidy", lambda wb, sh, rows, last_col, allow_archive=False:
                        ("tidy", wb, sh, rows, last_col))
    assert asyncio.run(server.xl_tidy("B", "S", {"1": "title"}, "F")) == {
        "what": "xl_tidy", "out": ("tidy", "B", "S", {"1": "title"}, "F")}


def test_snapshot_passes_on_book_notes(monkeypatch):
    class B:
        name = "Book1.xlsx"
        app = object()
        notes = ["AutoSave was switched off in Book1.xlsx"]

        def sheet(self, name):
            return "WS"

    monkeypatch.setattr(workbench, "_find_book", lambda wb, allow_archive=False: B())
    monkeypatch.setattr(workbench, "_wait_ready", lambda app: None)
    monkeypatch.setattr(workbench, "_range", lambda ws, address: ("RNG", "B1"))
    monkeypatch.setattr(workbench, "_capture", lambda ws, rng: {})
    monkeypatch.setattr(workbench, "_record", lambda b, ws, address, snap, label: {"id": "c2"})
    assert workbench.snapshot("Book1", "S", "B1")["notes"] == ["AutoSave was switched off in Book1.xlsx"]


def test_row_kind_aliases_map_before_validation(monkeypatch):
    import pytest

    from smartass.xl.host import InputError
    from smartass.xl.workbench import books, fmt
    assert fmt._expand_rows({"1": "heading", "2-3": "subtotal", "4": "Total", "5": "line"}) == {
        1: "section", 2: "sum", 3: "sum", 4: "grand", 5: "line"}

    class Validated(Exception):
        pass

    class Book:
        app = None

        def sheet(self, name):
            return object()
    monkeypatch.setattr(books, "find_book", lambda *a, **k: Book())
    monkeypatch.setattr(fmt, "_wait_ready", lambda app: None)

    def com(fn, **k):
        raise Validated                                  # the first COM call comes after the kind validation
    monkeypatch.setattr(fmt, "com", com)
    with pytest.raises(Validated):
        fmt.format_sheet(None, "S", "F", {"4": "header", "5": "subtotal", "6": "total"})
    with pytest.raises(InputError, match="subtotal"):
        fmt.format_sheet(None, "S", "F", {"5": "bogus"})

def test_allow_archive_reaches_find_book_and_format_sheet(monkeypatch):
    """The archive refusal tells the caller to pass allow_archive=true; xl_snapshot and xl_tidy must accept it
    and hand it on (17.11.2026: the tools had no such parameter, so an explicitly asked archive edit was stuck)."""
    import asyncio

    from smartass import server

    seen = {}

    async def user(fn, what):
        return fn(None)
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(workbench, "snapshot", lambda wb, sh, rg, allow_archive=False: seen.update(snap=allow_archive))
    monkeypatch.setattr(workbench, "tidy", lambda wb, sh, rows, last_col, allow_archive=False:
                        seen.update(tidy=allow_archive))
    asyncio.run(server.xl_snapshot("B", "S", None, allow_archive=True))
    asyncio.run(server.xl_tidy("B", "S", {"1": "title"}, "F", allow_archive=True))
    assert seen == {"snap": True, "tidy": True}


def test_tidy_passes_allow_archive_to_format_sheet(monkeypatch):
    got = {}
    monkeypatch.setattr(workbench, "_format_sheet", lambda wb, sh, last_col, rows, allow_archive=False:
                        got.update(allow=allow_archive) or {"change_id": "c1"})
    monkeypatch.setattr(workbench, "_check", lambda wb, sh: {"ok": True, "block": [], "warn": [], "counts": {}})
    workbench.tidy("B", "S", rows={"1": "title"}, last_col="F", allow_archive=True)
    assert got == {"allow": True}
