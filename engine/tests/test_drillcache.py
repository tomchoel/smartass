"""Ledger lines of closed periods kept on disk (core/drillcache) and their use in drillstep.drill."""
import asyncio
import json

import pytest

from smartview_mcp import config
from smartview_mcp.core import drillcache, drillstep

CELL = {"All Accounts": "ac.C60090", "All Entities": "un.SY", "All Time Periods": "Yr 25"}
VIEW = {"items": [{"viewName": "DrillFilter", "viewDefinition": json.dumps([{"label": "AMOUNT", "value": "AMOUNTX"}])}]}
PIECES = [{"cell": CELL, "cube_value": -2.0, "items": [{"AMOUNTX": "-1"}, {"AMOUNTX": "-1"}], "view": VIEW,
           "secs": 1.0}]


@pytest.mark.parametrize("period,closed,ok", [
    ("Yr 25", "Oct 26", True), ("Yr 26", "Oct 26", False), ("Yr 26", "Dec 26", True),
    ("Q3 26", "Oct 26", True), ("Q4 26", "Oct 26", False), ("Oct 26", "Oct 26", True), ("Nov 26", "Oct 26", False),
    ("YTD Oct 26", "Oct 26", False), ("Yr 25", None, False)])
def test_only_closed_periods_are_cacheable(period, closed, ok):
    assert drillcache.closed_period({**CELL, "All Time Periods": period}, closed) is ok


def test_save_then_load_and_ttl(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    drillcache.save("FinLite", CELL, "Oct 26", PIECES, -2.0)
    hit = drillcache.load("FinLite", CELL, "Oct 26")
    assert hit["pieces"] == PIECES and hit["parent"] == -2.0
    assert drillcache.load("FinLite", {**CELL, "All Accounts": "ac.C60080"}, "Oct 26") is None
    assert drillcache.load("FinLite", {**CELL, "All Time Periods": "Yr 26"}, "Oct 26") is None   # open period
    monkeypatch.setattr(config, "DRILL_CACHE_DAYS", 0)
    assert drillcache.load("FinLite", CELL, "Oct 26") is None                                    # too old


def _fakes(monkeypatch, tmp_path, cube):
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    calls = {"cell": 0, "link": 0}

    async def fake_cell(conn, cell, max_lines, split):
        calls["cell"] += 1
        return [dict(p) for p in PIECES], -2.0

    async def fake_call(fn, what="call"):
        return fn(None)

    def fake_link(h, conn, cell):
        calls["link"] += 1
        return {"url": "u", "cube_value": cube}
    monkeypatch.setattr(drillstep, "_drill_cell", fake_cell)
    monkeypatch.setattr(drillstep, "_call", fake_call)
    monkeypatch.setattr(drillstep.ops, "drill_link", fake_link)
    return calls


def test_drill_reuses_a_closed_period_after_the_cube_recheck(monkeypatch, tmp_path):
    calls = _fakes(monkeypatch, tmp_path, -2.0)
    first = asyncio.run(drillstep.drill(CELL, say=lambda t: None, closed="Oct 26"))
    again = asyncio.run(drillstep.drill(CELL, say=lambda t: None, closed="Oct 26"))
    assert calls == {"cell": 1, "link": 1} and len(again.rows) == 2 and again.reconciliation["gap"] == 0
    assert "Cached" not in first.source and "Cached" in again.source and "re-checked" in again.source


def test_a_changed_cube_value_drills_again(monkeypatch, tmp_path):
    calls = _fakes(monkeypatch, tmp_path, -2.0)
    asyncio.run(drillstep.drill(CELL, say=lambda t: None, closed="Oct 26"))
    monkeypatch.setattr(drillstep.ops, "drill_link", lambda h, c, cell: calls.update(link=9) or
                        {"url": "u", "cube_value": -3.0})
    res = asyncio.run(drillstep.drill(CELL, say=lambda t: None, closed="Oct 26"))
    assert calls["cell"] == 2 and "Cached" not in res.source


def test_an_open_period_is_never_cached(monkeypatch, tmp_path):
    calls = _fakes(monkeypatch, tmp_path, -2.0)
    cell = {**CELL, "All Time Periods": "Yr 26"}
    for _ in range(2):
        asyncio.run(drillstep.drill(cell, say=lambda t: None, closed="Oct 26"))
    assert calls == {"cell": 2, "link": 0}


def test_combine_keeps_the_cache_note():
    def r(src):
        return drillstep.DrillResult(CELL, [], [], -1.0, {"lines": 0, "gap": 0.0}, None, src, 1.0)
    both = drillstep.combine([r(drillstep.SOURCE), r(drillstep.SOURCE + " Cached 01.11.2026 (closed period)")])
    assert "Cached 01.11.2026" in both.source and both.source.startswith(drillstep.SOURCE)
