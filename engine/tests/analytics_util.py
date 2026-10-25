"""Synthetic frames for the analytics tests (imported like svxml_util; every figure is invented)."""
from smartview_mcp import catalog
from smartview_mcp.analytics import periods
from smartview_mcp.analytics.frame import Frame

E = "un.SY.2WHS"


def frame(cube="FinMax", values=None):
    return Frame(cube, catalog.load().pnl[cube], dict(values or {}))


def months(values: dict, line: str, scenario: str, start: str, amounts: list, entity: str = E) -> dict:
    """values with one entry per month from `start` (inclusive)."""
    for i, v in enumerate(amounts):
        values[(entity, line, periods.shift(start, i), scenario)] = v
    return values
