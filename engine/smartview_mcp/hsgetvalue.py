"""Oracle Smart View HsGetValue formula constructor."""

from __future__ import annotations

from typing import Mapping

FINMAX_DEFAULTS: dict[str, str] = {
    "Scenarios": "Actual",
    "Currencies": "NOK",
    "Functions": "fu.NONE",
    "Products": "pd.NONE",
    "Projects": "pj.NONE",
    "Reserves": "re.NONE",
    "Systems": "sy.NONE",
    "GLSources": "gl.NONE",
    "Counterparts": "cp.NONE",
    "LegalEntities": "le.NONE",
}


def build_hsgetvalue_formula(connection: str, coordinates: Mapping[str, str]) -> str:
    """Build an Oracle Smart View =HsGetValue(...) formula string.

    Args:
        connection: Smart View connection name (e.g. "FinMax").
        coordinates: Mapping of Dimension -> Member (e.g. {"Accounts": "ac.EBITDA"}).
    """
    merged = dict(FINMAX_DEFAULTS)
    merged.update(coordinates)

    pairs = [f'"{dim}#{member}"' for dim, member in sorted(merged.items())]
    inner = ", ".join([f'"{connection}"'] + pairs)
    return f"=HsGetValue({inner})"
