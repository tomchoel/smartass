"""Plan 2D task 5: variance % is blank on a tiny base. When |base| is below the metric's percentage floor
(materiality abs / 10 for a NOK metric; a stock or ratio metric only with its own floor) the model gives None and
Excel "" - the floor sits in a labelled cell on the answer sheet, never as a constant in the formula. Every figure
is invented."""
import pytest
from v2util import CLOSED, spec_dict

from smartass.model.analyze import compute
from smartass.model.blocks.common import ck
from smartass.model.expr import Param, Ratio, Ref, Sub
from smartass.model.frame import Fact, Frame
from smartass.model.resolver import MetricDef, StaticResolver
from smartass.model.spec import AnalysisSpec
from smartass.render.layout import col_letter, layout
from smartass.render.readback import compare


def res_():
    return StaticResolver(
        metrics=[MetricDef("Revenue", {"FinMax": ("ac.C30",)}, sign="income"),           # 5 M -> floor 500 k
                 MetricDef("Opex", {"FinMax": ("ac.C60",)}, sign="cost", materiality=(1e6, 0.1)),  # floor 100 k
                 MetricDef("FTE", {"FinMax": ("ac.STA20Y",)}, unit="FTE", balance="Accumulated",
                           aggregation="last"),                                            # no floor
                 MetricDef("Heads", {"FinMax": ("ac.STA30Y",)}, unit="FTE", balance="Accumulated",
                           aggregation="last", pct_floor=5.0)],                            # its own floor
        tree={"un.SY.2BUS": ("un.SY.B1", "un.SY.B2")}, extra_members={"ac.C99"})


# ---- the floor per metric

def test_the_floor_of_a_metric():
    r = res_()
    assert r.pct_floor("Revenue") == 500_000
    assert r.pct_floor("opex") == 100_000
    assert r.pct_floor("FTE") is None                 # a stock without its own floor: no floor
    assert r.pct_floor("Heads") == 5.0
    assert r.pct_floor("ac.C99") is None              # a raw member: no metric, no floor


# ---- the expression

def test_ratio_with_a_floor_is_empty_below_it():
    f = Frame([])

    class C:
        def __init__(self, d):
            self.d = d

        def get(self, k):
            return self.d[k]

    e = Ratio(Sub(Ref("cur"), Ref("base")), Ref("base"), abs_base=True, floor=Param("floor.Revenue", 500_000.0))
    assert e.evaluate(f, C({"cur": 900_000.0, "base": 300_000.0})) is None
    assert e.evaluate(f, C({"cur": 900_000.0, "base": -499_999.0})) is None
    assert e.evaluate(f, C({"cur": 1_200_000.0, "base": 600_000.0})) == pytest.approx(1.0)
    assert e.evaluate(f, C({"cur": 1_200_000.0, "base": -600_000.0})) == pytest.approx(3.0)
    addr = {"cur": "C10", "base": "D10", "floor.Revenue": "$C$40"}.__getitem__
    assert e.to_excel("Data_Q1", addr) == 'IF(ABS(D10)<$C$40,"",(C10-D10)/ABS(D10))'
    plain = Ratio(Sub(Ref("cur"), Ref("base")), Ref("base"), abs_base=True)
    assert plain.to_excel("Data_Q1", addr) == 'IF(D10=0,"",(C10-D10)/ABS(D10))'      # no floor: as before


# ---- the model

def frame(rev=(900_000.0, 300_000.0), opex=(-2_000_000.0, -1_500_000.0), fte=(10.0, 2.0), heads=(10.0, 2.0)):
    out = []
    for i, per in enumerate(("Sep 26", "Sep 25")):
        for acc, v in (("ac.C30", rev[i]), ("ac.C60", opex[i]), ("ac.STA20Y", fte[i]), ("ac.STA30Y", heads[i])):
            bal = "Accumulated" if acc.startswith("ac.STA") else "Periodic"
            out.append(Fact("FinMax", "fm", acc, "un.SY.2BUS", per, "Actual", "", v, balance=bal))
    return Frame(out)


def spec():
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "title": "P&L"},
                          {"type": "kpi_tiles", "query": "fm", "metrics": ["Revenue"]}])
    d["queries"][0]["metrics"] = ["Revenue", "Opex", "FTE", "Heads"]
    return AnalysisSpec.model_validate(d)


def test_variance_pct_is_none_below_the_floor_and_kept_above_it():
    s = spec()
    res = compute(s, res_(), CLOSED, frame())
    pct = lambda b, line: res.cells.get(ck(b, line, 3))           # noqa: E731
    assert pct("b0", "m0") is None                                   # Revenue base 300 k < 500 k
    assert pct("b1", "k0") is None                                   # the kpi tile too
    assert res.cells.get(ck("b0", "m0", 2)) == pytest.approx(600_000)   # the variance itself stays
    assert pct("b0", "m1") == pytest.approx(-0.5 / 1.5)              # Opex base 1.5 M >= 100 k
    assert pct("b0", "m2") == pytest.approx(4.0)                     # FTE: no floor
    assert pct("b0", "m3") is None                                   # Heads: base 2 < its floor 5
    assert res.key_values["P&L · Revenue"]["Variance %"] is None
    assert res.floors == {"Revenue": 500_000, "Opex": 100_000, "Heads": 5.0}


def test_at_the_floor_the_pct_is_shown():
    res = compute(spec(), res_(), CLOSED, frame(rev=(1_000_000.0, 500_000.0)))
    assert res.cells.get(ck("b0", "m0", 3)) == pytest.approx(1.0)


# ---- the sheet

def _plan(rev=(900_000.0, 300_000.0)):
    s = spec()
    res = compute(s, res_(), CLOSED, frame(rev=rev))
    return res, layout(1, s, res, stock={"FTE", "Heads"}, subtitle="x", const={})


def test_the_floor_sits_in_a_labelled_thresholds_cell():
    res, p = _plan()
    sections = {p.cells[(r["row"], 2)]: r["row"] for r in p.rows if r["kind"] == "section"}
    assert "Thresholds" in sections
    labels = {v: row for (row, col), v in p.cells.items() if col == 3 and isinstance(v, str)}
    row = labels["Variance % floor: Revenue"]
    assert p.cells[(row, 4)] == 500_000 and row > sections["Thresholds"]
    assert p.cells[(labels["Variance % floor: Heads"], 4)] == 5.0
    assert "Variance % floor: FTE" not in labels
    f = p.cells[p.at(ck("b0", "m0", 3))]
    r0, c0 = p.at(ck("b0", "m0", 1))
    assert f == f'=IF(ABS({col_letter(c0)}{r0})<$D${row},"",({col_letter(c0 - 1)}{r0}-{col_letter(c0)}{r0})' \
                f'/ABS({col_letter(c0)}{r0}))'
    assert p.cells[p.at(ck("b0", "m2", 3))].startswith("=IF(")                            # FTE: as before
    assert "=0," in p.cells[p.at(ck("b0", "m2", 3))]
    assert all("floor." not in k for k in p.addr)              # the readback compares model cells only


def test_the_floor_layout_has_no_block_findings():
    from smartass.xl.vendor.formelsjekk import check_cells
    _, p = _plan()
    cells = [("Q1 Answer", f"{col_letter(col)}{row}", v if isinstance(v, str) else repr(v))
             for (row, col), v in p.cells.items() if col >= 4]
    assert [f for f in check_cells(cells) if f.severity == "block"] == []


def test_no_floor_used_no_floor_rows():
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "metrics": ["FTE"], "material": False}])
    d["queries"][0]["metrics"] = ["FTE"]
    s = AnalysisSpec.model_validate(d)
    res = compute(s, res_(), CLOSED, frame())
    p = layout(1, s, res, stock={"FTE"}, subtitle="x", const={})
    assert not any(k.startswith("floor.") for k in p.params)                  # plan 5: Thresholds always shows
    assert set(p.params) == {"unit_mnok", "tolerance"}                       # the unit and the check tolerance


def test_sheet_equals_model_with_a_blank_pct():
    """The readback: Excel's "" for the blank % equals the model's None."""
    res, p = _plan()
    ra, ca = p.at(ck("b0", "m0", 3))
    grid = [[None] * 10 for _ in range(p.last_row)]
    for key, ad in p.addr.items():
        r, c = p.at(key)
        v = res.cells.get(key)
        grid[r - 1][c - 1] = "" if v is None else v
    assert grid[ra - 1][ca - 1] == ""
    assert compare(p.addr, res.cells, tuple(tuple(r) for r in grid), p.fractions) == []
