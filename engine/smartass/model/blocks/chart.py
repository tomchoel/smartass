# engine/smartass/model/blocks/chart.py
"""chart: a chart on a block already on the sheet (a BlockOut of type 'chart' with no lines; chart_of names the
block) or on its own data (built as a table and laid out as 'Chart data: <title>', evaluated and read back like
any table). implicit_chart gives an answer without a chart one (every path but a new question's guard)."""
from __future__ import annotations

from ..context import Ctx
from ..spec import ChartBlock
from .common import BlockOut
from .table import build_table

IMPLICIT_KEY = "chart0"     # the implicit chart's out: not a spec block index (the spec stays as given)

DEFAULT_SERIES: dict[str, list[str]] = {
    "waterfall": ["amount"], "line": ["value", "ly"], "column": ["value", "ly"], "bar": ["value"],
    "stacked": ["value"], "stacked100": ["value"], "combo": ["value", "var_pct"], "variance": ["var"]}


def default_series(kind: str, cols: list[str]) -> list[str]:
    """The kind's default series; line and column compare against the base the data has (ly, else budget, else
    forecast)."""
    if kind in ("line", "column"):
        base = next((b for b in ("ly", "budget", "forecast") if b in cols), None)
        return ["value", base] if base else ["value"]
    return list(DEFAULT_SERIES[kind])


def build_chart(b: ChartBlock, ctx: Ctx, key: str) -> BlockOut:
    table = b.data_table()
    if table is None:
        return BlockOut(key, "chart", b.title, [], chart=b.kind, chart_of=f"b{b.source.block}")
    out = build_table(table, ctx, key)
    out.type = "chart_data"
    out.chart = b.kind
    return out


CHARTABLE = ("bridge", "trend", "table")     # the implicit chart's preference: waterfall, line, column
IMPLICIT_KIND = {"bridge": "waterfall", "trend": "line", "ranking": "bar", "lines": "bar", "drivers": "bar"}


def implicit_chart(outs: list[BlockOut]) -> BlockOut | None:
    """A chart on the first bridge (waterfall), else the first trend (line), else the first table (column), else the
    first block with lines (a bar for ranking, ledger lines and drivers, else a column); None when no block has
    data."""
    data = [o for o in outs if o.type not in ("chart", "chart_data") and o.lines]
    pick = next((o for t in CHARTABLE for o in data if o.type == t), data[0] if data else None)
    if pick is None:
        return None
    return BlockOut(IMPLICIT_KEY, "chart", pick.title, [], chart=IMPLICIT_KIND.get(pick.type, "column"),
                    chart_of=pick.key)
