"""A synthetic Wholesale warehouse for the fin_pnl tests: a base fill and a top-up over EchoTransport (svhttp's fake
provider). Every figure is invented: the pnl_value P&L, x1.1 in 2026, Target x0.9, the forecast x1.05."""
from svxml_util import EchoTransport, atlas_with_dimensions, pnl_value

from smartview_mcp import catalog, extract, refresh, svhttp
from smartview_mcp.fin_pnl import Scope
from smartview_mcp.warehouse import MemoryWarehouse

CAT = catalog.load()
WHS = next(e for e in CAT.entities if e.key == "wholesale")
FC = "Forecast Q326"
UNITS = (("un.SY.U1", "Unit one"), ("un.SY.U2", "Unit two"))


def value(rows, cols, pov):
    v = pnl_value(rows, cols, pov)
    year = 1.1 if any(c.endswith(" 26") for c in cols) or "FY26" in pov or "FY26" in cols else 1.0
    plan = 0.9 if "Target" in pov or "Target" in cols else (1.05 if FC in pov or FC in cols else 1.0)
    return v * year * plan


def http(tmp_path, value_fn=value):
    tmp_path.mkdir(parents=True, exist_ok=True)
    return svhttp.SvHttp(EchoTransport(value_fn), home="https://plan1.example",
                         db_path=atlas_with_dimensions(tmp_path / "a.db"))


def filled(tmp_path, period="Sep 26", closed="Sep 26", forecast=FC):
    """(store, retrieve, scope) after a base fill and the question's top-up."""
    store, retrieve = MemoryWarehouse(), extract.route(http=http(tmp_path))
    extract.fill(store, retrieve, CAT, extract.base_grids(CAT, closed, {26: forecast}.get), "base", closed)
    refresh.top_up(store, retrieve, CAT, WHS, period, closed, forecast, [u for u, _ in UNITS])
    return store, retrieve, Scope(WHS, period, closed, forecast, UNITS)
