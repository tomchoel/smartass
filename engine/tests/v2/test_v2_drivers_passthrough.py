# engine/tests/v2/test_v2_drivers_passthrough.py
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.blocks.drivers import build_drivers
from smartass.model.context import Ctx
from smartass.model.frame import Fact, Frame
from smartass.model.spec import AnalysisSpec

# 2BUS -> PASS (one child) -> 4BUS_B, G910 ; 4BUS_B -> B1, B2
TREE = {"un.SY.2BUS": ("un.SY.PASS",), "un.SY.PASS": ("un.SY.4BUS_B", "un.SY.G910"),
        "un.SY.4BUS_B": ("un.SY.B1", "un.SY.B2")}


def passthrough_frame() -> Frame:
    out = []
    for f in facts().facts:
        out.append(f)
        if f.entity == "un.SY.2BUS":                       # PASS carries the same numbers as its only parent
            out.append(Fact(f.cube, f.query, f.account, "un.SY.PASS", f.period, f.scenario, f.version, f.value,
                            f.balance))
    return Frame(out)


def test_single_child_node_is_skipped_and_costs_no_depth():
    frame = passthrough_frame()                            # before the resolver: facts() reads the shared tree
    r = resolver()
    r.tree = dict(TREE)                                    # own copy: never mutate the shared v2util tree
    r.__post_init__()
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "drivers", "query": "fm", "metric": "Opex",
                                                       "depth": 2}]))
    out = build_drivers(s.blocks[0], Ctx(s, r, CLOSED, frame), "b0")
    labels = [ln.label for ln in out.lines]
    assert "un.SY.PASS" not in labels
    assert labels[:2] == ["2BUS Business Norway", "un.SY.4BUS_B"]
    assert "un.SY.B1" in labels                            # depth 2 still reaches the grandchildren
