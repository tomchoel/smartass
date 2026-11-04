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


from smartass.model.frame import Fact, Frame  # noqa: E402
from smartass.model.resolver import MetricDef, StaticResolver  # noqa: E402

TREE = {"un.SY.2BUS": ("un.SY.4BUS_B", "un.SY.G910"), "un.SY.4BUS_B": ("un.SY.B1", "un.SY.B2")}
LEAVES = {  # period -> entity -> account -> value
    "Sep 26": {"un.SY.B1": {"ac.C30": 100.0, "ac.C50S": -40.0, "ac.C60": -20.0},
               "un.SY.B2": {"ac.C30": 30.0, "ac.C50S": -10.0, "ac.C60": -5.0},
               "un.SY.G910": {"ac.C30": 20.0, "ac.C50S": -8.0, "ac.C60": -2.0}},
    "Sep 25": {"un.SY.B1": {"ac.C30": 90.0, "ac.C50S": -35.0, "ac.C60": -15.0},
               "un.SY.B2": {"ac.C30": 32.0, "ac.C50S": -12.0, "ac.C60": -6.0},
               "un.SY.G910": {"ac.C30": 18.0, "ac.C50S": -8.0, "ac.C60": -1.0}},
}


def resolver(**kw) -> StaticResolver:
    return StaticResolver(
        metrics=[MetricDef("Revenue", {"FinMax": ("ac.C30",), "FinPlan": ("Revenue",)}, sign="income"),
                 MetricDef("Opex", {"FinMax": ("ac.C50S", "ac.C60"), "FinPlan": ("OPEX",)}, sign="cost",
                           materiality=(5.0, 0.05)),
                 MetricDef("EBITDA", {"FinMax": ("ac.EBITDA",), "FinPlan": ("EBITDA",)}, sign="income"),
                 MetricDef("FTE", {"FinMax": ("ac.STA20Y",)}, balance="Accumulated", aggregation="last"),
                 MetricDef("Draft thing", {"FinMax": ("ac.C99",)}, status="draft")],
        tree=TREE, labels={"un.SY.2BUS": "2BUS Business Norway"}, **kw)


def facts() -> Frame:
    """Leaves plus parents (EPM returns parent values), EBITDA = C30 + C50S + C60 per entity."""
    out = []
    for per, ents in LEAVES.items():
        vals = {e: dict(a) for e, a in ents.items()}
        for parent in ("un.SY.4BUS_B", "un.SY.2BUS"):
            vals[parent] = {acc: sum(vals[c][acc] for c in TREE[parent]) for acc in ("ac.C30", "ac.C50S", "ac.C60")}
        for ent, accs in vals.items():
            accs["ac.EBITDA"] = accs["ac.C30"] + accs["ac.C50S"] + accs["ac.C60"]
            out += [Fact("FinMax", "fm", acc, ent, per, "Actual", "", v) for acc, v in accs.items()]
    return Frame(out)


def monthly_facts() -> Frame:
    """Opex for B1, B2, G910 and parents, Jul-Sep 26 and Jul-Sep 25: leaf value = -(month) per account."""
    out = []
    for yy in (26, 25):
        for mon, lab in ((7, "Jul"), (8, "Aug"), (9, "Sep")):
            per = f"{lab} {yy}"
            leaf = {e: {"ac.C50S": -float(mon), "ac.C60": -1.0} for e in ("un.SY.B1", "un.SY.B2", "un.SY.G910")}
            leaf["un.SY.4BUS_B"] = {a: leaf["un.SY.B1"][a] + leaf["un.SY.B2"][a] for a in ("ac.C50S", "ac.C60")}
            leaf["un.SY.2BUS"] = {a: leaf["un.SY.4BUS_B"][a] + leaf["un.SY.G910"][a] for a in ("ac.C50S", "ac.C60")}
            out += [Fact("FinMax", "fm", a, e, per, "Actual", "", v) for e, accs in leaf.items() for a, v in accs.items()]
    return Frame(out)
