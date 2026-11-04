"""Plan 2C task 2: the product dimension end to end - spec, requests, provider grids, facts, Pick, constant
columns, the guard, and table/ranking rows by product. Every figure is invented."""
import pytest
from svxml_util import EchoTransport
from test_v2_pipeline import FakeProvider
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.data.provider import SvHttpProvider
from smartass.data.recorder import RecordingProvider, ReplayProvider
from smartass.epm import svhttp
from smartass.model.analyze import compute
from smartass.model.blocks.common import ck, evaluate
from smartass.model.blocks.ranking import build_ranking
from smartass.model.blocks.table import build_table
from smartass.model.context import Ctx
from smartass.model.frame import Fact, Frame
from smartass.model.requests import GridRequest, requests
from smartass.model.spec import AnalysisSpec
from smartass.model.validate import GuardError
from smartass.pipeline import prepare
from smartass.render.layout import constant_columns, layout

PRODUCTS = ("pd.SY.A", "pd.SY.B", "pd.SY.C")
VALUES = {"Sep 26": {"pd.SY.A": 50.0, "pd.SY.B": 30.0, "pd.SY.C": 5.0},
          "Sep 25": {"pd.SY.A": 40.0, "pd.SY.B": 35.0, "pd.SY.C": 4.0}}


def presolver():
    r = resolver(extra_members=set(PRODUCTS))
    r.labels.update({"pd.SY.A": "A Alpha", "pd.SY.B": "B Beta", "pd.SY.C": "C Gamma"})
    return r


def pfacts() -> Frame:
    return Frame([Fact("FinMax", "fm", "ac.C30", "un.SY.2BUS", per, "Actual", "", v, product=p)
                  for per, by in VALUES.items() for p, v in by.items()])


def pspec(blocks=None, products=PRODUCTS) -> dict:
    d = spec_dict(blocks=blocks or [{"type": "ranking", "query": "fm", "metric": "Revenue", "top": 2,
                                     "rows": "products"}])
    d["queries"][0].update({"metrics": ["Revenue"], "products": list(products)})
    return d


# ---- spec

def test_products_default_to_the_cube_total():
    s = AnalysisSpec.model_validate(spec_dict())
    assert s.queries[0].products == []


def test_rows_products_needs_products_in_the_query():
    with pytest.raises(ValueError, match="products"):
        AnalysisSpec.model_validate(pspec(products=[]))
    d = pspec([{"type": "table", "query": "fm", "rows": "products", "metric": "Revenue"}], products=[])
    with pytest.raises(ValueError, match="products"):
        AnalysisSpec.model_validate(d)


# ---- requests

def test_requests_carry_the_products_and_fetch_only_the_entity():
    reqs = requests(AnalysisSpec.model_validate(pspec()), presolver(), CLOSED)
    assert {r.products for r in reqs} == {PRODUCTS}
    assert {r.entities for r in reqs} == {("un.SY.2BUS",)}          # a product ranking needs no entity children


def test_requests_without_products_are_unchanged():
    reqs = requests(AnalysisSpec.model_validate(spec_dict()), resolver(), CLOSED)
    assert all(r.products == () for r in reqs)


def test_replay_key_is_unchanged_without_products():
    from smartass.data.recorder import _key
    req = GridRequest("fm", "FinMax", ("ac.C30",), ("E1",), ("Sep 26",), "Actual", "", "Periodic")
    assert "products" not in _key(req)
    assert "products" in _key(GridRequest("fm", "FinMax", ("ac.C30",), ("E1",), ("Sep 26",), "Actual", "",
                                          "Periodic", ("pd.SY.A",)))


# ---- provider grids (fake transport)

def _provider(tmp_path, value):
    from smartass.knowledge import atlas
    from svxml_util import FINMAX_DIMS, FINPLAN_DIMS
    path = tmp_path / "a.db"
    db = atlas.connect(path)
    for conn, dims in (("FinMax", FINMAX_DIMS + ["All Products"]), ("FinPlan", FINPLAN_DIMS + ["All Products"])):
        for d in dims:
            atlas.store(db, conn, d, [{"name": d, "label": d, "parent": None, "depth": 0, "order": 0}], 0.1)
    db.close()
    return SvHttpProvider(svhttp.SvHttp(EchoTransport(value), home="https://plan1.example", db_path=path))


def test_finmax_grid_with_products_adds_a_third_row_dimension(tmp_path):
    seen = []

    def value(rows, cols, pov):
        seen.append(rows)
        return {"pd.SY.A": 1.0, "pd.SY.B": 2.0}[rows[2]]

    p = _provider(tmp_path, value)
    req = GridRequest("fm", "FinMax", ("ac.C30",), ("E1", "E2"), ("Sep 26",), "Actual", "", "Periodic",
                      ("pd.SY.A", "pd.SY.B"))
    got = p.fetch(req)
    assert got.warnings == []
    assert {r[2] for r in seen} == {"pd.SY.A", "pd.SY.B"}
    assert sorted((f.entity, f.product, f.value) for f in got.facts) == [
        ("E1", "pd.SY.A", 1.0), ("E1", "pd.SY.B", 2.0), ("E2", "pd.SY.A", 1.0), ("E2", "pd.SY.B", 2.0)]
    assert {f.period for f in got.facts} == {"Sep 26"}


def test_finplan_grid_with_products(tmp_path):
    p = _provider(tmp_path, lambda rows, cols, pov: 3.0 if rows[2] == "Mobile" else None)
    req = GridRequest("fp", "FinPlan", ("EBITDA",), ("BANORWAY",), ("Sep 26",), "Actual", "Working", "Periodic",
                      ("Mobile", "Fixed"))
    got = p.fetch(req)
    assert [(f.product, f.period, f.value) for f in got.facts] == [("Mobile", "Sep 26", 3.0)]


def test_grid_without_products_has_no_product_dimension(tmp_path):
    seen = []
    p = _provider(tmp_path, lambda rows, cols, pov: seen.append(rows) or 1.0)
    got = p.fetch(GridRequest("fm", "FinMax", ("ac.C30",), ("E1",), ("Sep 26",), "Actual", "", "Periodic"))
    assert seen == [("ac.C30", "E1")]
    assert [f.product for f in got.facts] == [""]


# ---- facts, Pick and formulas

def test_frame_has_a_product_column_only_when_a_fact_has_a_product():
    assert "Product" not in facts().columns()
    fr = pfacts()
    cols = fr.columns()
    assert cols[cols.index("Entity") + 1] == "Product"
    assert fr.rows()[0][cols.index("Product")] == "pd.SY.A"
    assert fr.total({"Product": ("pd.SY.A", "pd.SY.B"), "Period": ("Sep 26",)}) == 80.0


def test_pick_filters_on_the_query_products():
    s = AnalysisSpec.model_validate(pspec())
    c = Ctx(s, presolver(), CLOSED, pfacts())
    e = c.pick(s.queries[0], "Revenue")
    assert c.value(e) == 85.0
    assert c.value(c.pick(s.queries[0], "Revenue", product="pd.SY.B")) == 30.0
    f = e.to_excel("Data_Q1", lambda k: k)
    assert 'Data_Q1[Product],{"pd.SY.A","pd.SY.B","pd.SY.C"}' in f


def test_a_single_product_is_left_out_of_the_formula_when_constant():
    s = AnalysisSpec.model_validate(pspec([{"type": "kpi_tiles", "query": "fm", "metrics": ["Revenue"]}],
                                          products=["pd.SY.A"]))
    fr = Frame([x for x in pfacts().facts if x.product == "pd.SY.A"])
    const = constant_columns(fr)
    assert const["Product"] == "pd.SY.A"
    c = Ctx(s, presolver(), CLOSED, fr)
    assert "Product" not in c.pick(s.queries[0], "Revenue").to_excel("Data_Q1", lambda k: k, const)


def test_formulas_without_products_never_name_the_product():
    s = AnalysisSpec.model_validate(spec_dict())
    res = compute(s, resolver(), CLOSED, facts())
    p = layout(1, s, res, stock=set(), subtitle="x", const=constant_columns(facts()))
    formulas = [v for v in p.cells.values() if isinstance(v, str) and v.startswith("=")]
    assert formulas and not any("Product" in f for f in formulas)


# ---- the guard

def test_unknown_product_is_refused():
    with pytest.raises(GuardError, match="pd.SY.NOPE"):
        prepare(pspec(products=["pd.SY.A", "pd.SY.NOPE"]), presolver(), FakeProvider(pfacts()))


# ---- blocks by product

def test_ranking_by_product_top_other_total_share():
    s = AnalysisSpec.model_validate(pspec())
    c = Ctx(s, presolver(), CLOSED, pfacts())
    out = build_ranking(s.blocks[0], c, "b0")
    assert [ln.key for ln in out.lines] == ["p0", "p1", "other", "total"]
    assert [ln.label for ln in out.lines] == ["A Alpha", "B Beta", "Other", "Total products"]
    cells = evaluate([out], c.frame)
    assert cells.get(ck("b0", "p0", 0)) == 50.0
    assert cells.get(ck("b0", "p1", 1)) == 35.0                     # B last year
    assert cells.get(ck("b0", "other", 0)) == 5.0
    assert cells.get(ck("b0", "total", 0)) == 85.0
    assert cells.get(ck("b0", "p0", 4)) == pytest.approx(50 / 85)


def test_table_rows_products():
    s = AnalysisSpec.model_validate(pspec([{"type": "table", "query": "fm", "rows": "products", "metric": "Revenue",
                                            "sort": "desc"}]))
    c = Ctx(s, presolver(), CLOSED, pfacts())
    out = build_table(s.blocks[0], c, "b0")
    cells = evaluate([out], c.frame)
    assert [ln.label for ln in out.lines] == ["A Alpha", "B Beta", "C Gamma", "Total products"]
    assert cells.get(ck("b0", "p2", 0)) == 5.0 and cells.get(ck("b0", "total", 0)) == 85.0


def test_ranking_by_product_end_to_end_records_and_replays(tmp_path):
    rec = RecordingProvider(FakeProvider(pfacts()))
    live = prepare(pspec(), presolver(), rec)
    assert all(r.products == PRODUCTS for r in live.requests)
    rec.save(tmp_path / "g.json")
    again = prepare(pspec(), presolver(), ReplayProvider(tmp_path / "g.json"))
    assert again.result.key_values == live.result.key_values
    assert live.result.key_values["Top 2 by Revenue · A Alpha"]["Actual"] == 50.0
