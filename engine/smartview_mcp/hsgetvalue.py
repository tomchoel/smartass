"""Oracle Smart View HsGetValue formula constructor."""

from __future__ import annotations

from typing import Mapping

FINMAX_DEFAULTS: dict[str, str] = {
    "All Scenarios": "Actual",
    "All Currencies": "NOK",
    "All Balance Types": "Periodic",
    "All Functions": "All Functions",
    "All Products": "All Products",
    "All Projects": "All Projects",
    "All Reserves": "All Reserves",
    "All Systems": "All Systems",
    "All GL Sources": "All GL Sources",
    "All Counterparts": "All Counterparts",
    "All Legal Entities": "All Legal Entities",
}

CANONICAL_FINMAX_DIMS: dict[str, str] = {
    "accounts": "All Accounts",
    "all accounts": "All Accounts",
    "entities": "All Entities",
    "all entities": "All Entities",
    "time periods": "All Time Periods",
    "all time periods": "All Time Periods",
    "period": "All Time Periods",
    "periods": "All Time Periods",
    "scenarios": "All Scenarios",
    "all scenarios": "All Scenarios",
    "scenario": "All Scenarios",
    "currencies": "All Currencies",
    "all currencies": "All Currencies",
    "currency": "All Currencies",
    "balance types": "All Balance Types",
    "all balance types": "All Balance Types",
    "balance": "All Balance Types",
    "functions": "All Functions",
    "all functions": "All Functions",
    "products": "All Products",
    "all products": "All Products",
    "projects": "All Projects",
    "all projects": "All Projects",
    "reserves": "All Reserves",
    "all reserves": "All Reserves",
    "systems": "All Systems",
    "all systems": "All Systems",
    "gl sources": "All GL Sources",
    "glsources": "All GL Sources",
    "all gl sources": "All GL Sources",
    "counterparts": "All Counterparts",
    "all counterparts": "All Counterparts",
    "legal entities": "All Legal Entities",
    "legalentities": "All Legal Entities",
    "all legal entities": "All Legal Entities",
}


def build_hsgetvalue_formula(connection: str, coordinates: Mapping[str, str]) -> str:
    """Build an Oracle Smart View =HsGetValue(...) formula string.

    Args:
        connection: Smart View connection name (e.g. "FinMax").
        coordinates: Mapping of Dimension -> Member (e.g. {"All Accounts": "ac.EBITDA"}).
    """
    merged = dict(FINMAX_DEFAULTS)
    for dim, member in coordinates.items():
        canonical_dim = CANONICAL_FINMAX_DIMS.get(dim.strip().lower(), dim)
        merged[canonical_dim] = member

    pairs = [f'"{dim}#{member}"' for dim, member in sorted(merged.items())]
    inner = ", ".join([f'"{connection}"'] + pairs)
    return f"=HsGetValue({inner})"
