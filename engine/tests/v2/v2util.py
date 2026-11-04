"""Shared v2 fixtures. Every figure is invented."""
import copy

CLOSED = (2026, 9)

BASE_SPEC = {
    "version": 2,
    "question": "How is Business opex developing vs last year?",
    "plan": {"depth": "standard", "purpose": "Explain Business opex vs LY", "hypotheses": ["Consultants up"]},
    "queries": [{"id": "fm", "cube": "FinMax", "metrics": ["Revenue", "Opex", "EBITDA"],
                 "entity": "un.SY.2BUS", "periods": "closed", "compare": ["ly"]}],
    "blocks": [
        {"type": "kpi_tiles", "query": "fm", "metrics": ["EBITDA", "Revenue"]},
        {"type": "table", "query": "fm", "title": "P&L"},
        {"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]},
        {"type": "ranking", "query": "fm", "metric": "Opex", "top": 1},
        {"type": "drivers", "query": "fm", "metric": "Opex", "depth": 2},
        {"type": "commentary"},
    ],
}


def spec_dict(**over) -> dict:
    d = copy.deepcopy(BASE_SPEC)
    d.update(over)
    return d
