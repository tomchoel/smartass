# engine/smartass/model/blocks/__init__.py
"""Block builders: each typed block becomes lines of expression trees. 'commentary' is not computed here (text the
agent writes); 'lines' reads the line sets the drill put on the frame."""
from .bridge import build_bridge
from .drivers import build_drivers
from .lines import build_lines
from .ranking import build_ranking
from .table import build_kpi, build_table
from .trend import build_matrix, build_trend

BUILDERS = {"table": build_table, "kpi_tiles": build_kpi, "trend": build_trend, "bridge": build_bridge,
            "ranking": build_ranking, "matrix": build_matrix, "drivers": build_drivers, "lines": build_lines}
