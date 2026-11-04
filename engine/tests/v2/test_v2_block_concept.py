"""Plan 2C task 4: the concept block - one section per concept block, one line per item (booked, attributed,
estimated, sum, missing), the key share in a helper column, the total, the checks as findings (never a stop), the
concept's grids and its notes as caveats. Every figure is invented."""
import re

import pytest
from test_v2_pipeline import FakeProvider
from v2util import CLOSED, resolver, spec_dict

from smartass.knowledge.semantic import Concept
from smartass.model.analyze import compute
from smartass.model.blocks.common import ck, evaluate
from smartass.model.blocks.concept import build_concept
from smartass.model.context import Ctx
from smartass.model.frame import Fact, Frame
from smartass.model.requests import requests
from smartass.model.spec import AnalysisSpec
from smartass.model.validate import GuardError
from smartass.pipeline import prepare
from smartass.render.layout import constant_columns, layout

TINY = Concept.model_validate({
    "id": "tiny", "title": "Tiny - direct contribution", "names": ["tiny offer"], "kind": "product_offer",
    "description": "A tiny offer.", "status": "approved", "notes": ["Tiny note."],
    "total": {"label": "Direct contribution Tiny", "of": ["rev", "cogs", "fin"]},
    "check": [{"label": "Revenue = the product's C30", "expr": "rev-$src",
               "source": {"account": "ac.C30", "products": ["pd.SY.T1"]}},
              {"label": "Total = the product's C30", "expr": "total-$src",
               "source": {"account": "ac.C30", "products": ["pd.SY.T1"]}}],
    "block": [
        {"title": "Revenue and COGS", "item": [
            {"key": "rev", "label": "Revenue (booked)", "kind": "booked", "account": "ac.C30",
             "products": ["pd.SY.T1"]},
            {"key": "cogs", "label": "COGS (estimated)", "kind": "estimated",
             "base": {"label": "COGS total", "account": "ac.C50S", "products": ["pd.SY.T9"]},
             "key_share": {"text": "Tiny's share of revenue", "num": {"label": "Tiny revenue", "expr": "rev"},
                           "den": {"label": "Revenue total", "account": "ac.C30", "products": ["pd.SY.T9"]}}},
            {"key": "gp", "label": "Gross profit", "kind": "sum", "of": ["rev", "cogs"]}]},
        {"title": "Financing", "item": [
            {"key": "fin", "label": "Interest income (attributed 100 %)", "kind": "attributed",
             "account": "ac.C80311", "entities": ["un.SY.B1"], "note": "Not product-tagged."},
            {"key": "gone", "label": "Financing cost", "kind": "missing", "note": "Not visible in FinMax."}]}]})

P, E = "fm:tiny:p", "fm:tiny:e"
VALUES = {  # period -> (query tag, account, entity, product) -> value
    "Sep 26": {(P, "ac.C30", "un.SY.2BUS", "pd.SY.T1"): 40.0, (P, "ac.C50S", "un.SY.2BUS", "pd.SY.T9"): -100.0,
               (P, "ac.C30", "un.SY.2BUS", "pd.SY.T9"): 200.0, (E, "ac.C80311", "un.SY.B1", ""): 7.0},
    "Sep 25": {(P, "ac.C30", "un.SY.2BUS", "pd.SY.T1"): 30.0, (P, "ac.C50S", "un.SY.2BUS", "pd.SY.T9"): -90.0,
               (P, "ac.C30", "un.SY.2BUS", "pd.SY.T9"): 150.0, (E, "ac.C80311", "un.SY.B1", ""): 5.0},
}


def cresolver():
    r = resolver(extra_members={"pd.SY.T1", "pd.SY.T9", "ac.C80311"}, concepts={"tiny": TINY})
    return r


def cfacts() -> Frame:
    return Frame([Fact("FinMax", tag, acc, ent, per, "Actual", "", v, product=prod)
                  for per, by in VALUES.items() for (tag, acc, ent, prod), v in by.items()])


def cspec(compare=("ly",), metrics=()) -> dict:
    d = spec_dict(blocks=[{"type": "concept", "query": "fm", "concept": "tiny"}])
    d["queries"][0].update({"metrics": list(metrics), "compare": list(compare)})
    return d


def _built(compare=("ly",)):
    s = AnalysisSpec.model_validate(cspec(compare))
    c = Ctx(s, cresolver(), CLOSED, cfacts())
    out = build_concept(s.blocks[0], c, "b0")
    return s, c, out, evaluate([out], c.frame)


def _col(out, header):
    return out.headers.index(header)


# ---- the spec

def test_a_concept_block_needs_a_finmax_or_finlite_query():
    d = cspec(compare=())
    d["queries"][0]["cube"] = "FinPlan"
    with pytest.raises(ValueError, match="FinMax or FinLite"):
        AnalysisSpec.model_validate(d)


def test_a_query_without_metrics_serves_only_concept_blocks():
    AnalysisSpec.model_validate(cspec())
    d = cspec()
    d["blocks"].append({"type": "kpi_tiles", "query": "fm", "metrics": ["Revenue"]})
    with pytest.raises(ValueError, match="metrics"):
        AnalysisSpec.model_validate(d)


# ---- lines, values and helper columns

def test_columns_follow_the_compare_and_add_the_key_helpers():
    _, _, out, _ = _built()
    assert out.headers == ["Actual", "Last year", "Variance", "Variance %", "Base", "Key share", "Base last year",
                           "Key share last year"]
    _, _, out, _ = _built(compare=())
    assert out.headers == ["Actual", "Base", "Key share"]


def test_one_line_per_item_with_block_headings_total_and_checks():
    _, _, out, _ = _built()
    assert [ln.key for ln in out.lines] == ["g0", "i_rev", "i_cogs", "i_gp", "g1", "i_fin", "i_gone", "total",
                                            "c0", "c1"]
    by = {ln.key: ln for ln in out.lines}
    assert by["g0"].label == "Revenue and COGS" and all(e is None for e in by["g0"].cells)
    assert by["i_gp"].kind == "sum" and by["total"].kind == "grand" and by["c0"].kind == "check"
    assert by["i_gone"].label == "Financing cost (not in FinMax)" and all(e is None for e in by["i_gone"].cells)
    assert out.title == "Tiny - direct contribution"


def test_values_booked_estimated_attributed_sum_total():
    _, _, out, cells = _built()
    v, ly, var = 0, 1, 2
    get = lambda line, j: cells.get(ck("b0", line, j))          # noqa: E731
    assert (get("i_rev", v), get("i_rev", ly), get("i_rev", var)) == (40.0, 30.0, 10.0)
    assert get("i_cogs", _col(out, "Base")) == -100.0
    assert get("i_cogs", _col(out, "Key share")) == pytest.approx(0.2)
    assert get("i_cogs", _col(out, "Key share last year")) == pytest.approx(0.2)
    assert get("i_cogs", v) == pytest.approx(-20.0) and get("i_cogs", ly) == pytest.approx(-18.0)
    assert get("i_gp", v) == pytest.approx(20.0)
    assert get("i_fin", _col(out, "Key share")) == 1.0 and get("i_fin", v) == 7.0 and get("i_fin", ly) == 5.0
    assert get("total", v) == pytest.approx(27.0) and get("total", ly) == pytest.approx(17.0)
    assert get("c0", v) == 0.0 and get("c1", v) == pytest.approx(-13.0)
    assert get("i_gone", v) is None


def test_formulas_keep_the_key_share_in_its_helper_column():
    s, c, out, _ = _built()
    res = compute(s, cresolver(), CLOSED, cfacts())
    p = layout(1, s, res, stock=set(), subtitle="x", const=constant_columns(c.frame))
    at = lambda line, header: p.addr[ck("b0", line, _col(out, header))]       # noqa: E731
    f = p.cells[p.at(ck("b0", "i_cogs", 0))]
    assert f == f"={at('i_cogs', 'Base')}*{at('i_cogs', 'Key share')}"
    key = p.cells[p.at(ck("b0", "i_cogs", _col(out, "Key share")))]
    assert key.startswith("=IF(") and at("i_rev", "Actual") in key and "Data_Q1[Product]" in key
    assert p.cells[p.at(ck("b0", "i_fin", _col(out, "Key share")))] == 1        # 100 %: an input, not a formula
    for (row, col), val in p.cells.items():
        if isinstance(val, str) and val.startswith("="):
            bare = re.sub(r'"[^"]*"|\$?[A-Z]{1,3}\$?\d+|Data_Q1\[\w+\]', "", val)
            bare = bare.replace("=0,", ",")              # the empty-at-zero guard IF(<base>=0,"",...) is house style
            assert not re.search(r"\d", bare), (row, col, val)                    # no constants in formulas
    assert p.cells[p.at(ck("b0", "i_gp", 0))] == f"=SUM({at('i_rev', 'Actual')}:{at('i_cogs', 'Actual')})"


def test_the_formula_check_has_no_block_findings():
    """tn-formelstil: the concept section passes the house formula check with zero BLOCK findings."""
    from smartass.render.layout import col_letter
    from smartass.xl.vendor.formelsjekk import check_cells
    s, c, _, _ = _built()
    res = compute(s, cresolver(), CLOSED, cfacts())
    p = layout(1, s, res, stock=set(), subtitle="x", const=constant_columns(c.frame))
    cells = [("Q1 Answer", f"{col_letter(col)}{row}", v if isinstance(v, str) else repr(v))
             for (row, col), v in p.cells.items() if col >= 3]
    blocks = [f for f in check_cells(cells) if f.severity == "block"]
    assert blocks == []


def test_every_seeded_concept_builds_and_passes_the_formula_check():
    import json

    from smartass.core import config
    from smartass.knowledge.seed import seed_concepts
    from smartass.render.layout import col_letter
    from smartass.xl.vendor.formelsjekk import check_cells
    concepts = seed_concepts(json.loads((config.KNOWLEDGE_DIR / "concepts.json").read_text(encoding="utf-8")))
    members = {x for c in concepts for xs in c.members().values() for x in xs}
    r = resolver(extra_members=members | {"un.SY"}, concepts={c.id: c for c in concepts})
    for c in concepts:
        d = spec_dict(blocks=[{"type": "concept", "query": "fm", "concept": c.id}])
        d["queries"][0].update({"metrics": [], "entity": "un.SY", "periods": "closed-ytd"})
        s = AnalysisSpec.model_validate(d)
        res = compute(s, r, CLOSED, Frame([]))
        assert len(res.outs[0].lines) == len(c.items()) + len(c.block) + 1 + len(c.check), c.id
        p = layout(1, s, res, stock=set(), subtitle="x", const={})
        cells = [("Q1 Answer", f"{col_letter(col)}{row}", v if isinstance(v, str) else repr(v))
                 for (row, col), v in p.cells.items() if col >= 3]
        assert [f for f in check_cells(cells) if f.severity == "block"] == [], c.id


def test_a_failing_check_is_a_finding_not_a_stop():
    s = AnalysisSpec.model_validate(cspec())
    res = compute(s, cresolver(), CLOSED, cfacts())
    hits = [f for f in res.findings if "Total = the product's C30" in f.text]
    assert len(hits) == 1 and hits[0].severity == "warning" and "-13" in hits[0].text
    assert not any("Revenue = the product's C30" in f.text for f in res.findings)


# ---- requests, guard, caveats

def test_requests_carry_the_concepts_accounts_and_products():
    reqs = requests(AnalysisSpec.model_validate(cspec()), cresolver(), CLOSED)
    p = [r for r in reqs if r.query == P]
    e = [r for r in reqs if r.query == E]
    assert [r.periods for r in p] == [("Sep 26",), ("Sep 25",)] and len(e) == 2
    assert set(p[0].accounts) == {"ac.C30", "ac.C50S"} and set(p[0].products) == {"pd.SY.T1", "pd.SY.T9"}
    assert p[0].entities == ("un.SY.2BUS",) and p[0].balance == "Periodic"
    assert e[0].accounts == ("ac.C80311",) and e[0].entities == ("un.SY.B1",) and e[0].products == ()
    assert not [r for r in reqs if r.query == "fm"]                  # a query without metrics fetches nothing itself


def test_the_query_metrics_keep_their_own_grid():
    reqs = requests(AnalysisSpec.model_validate(cspec(metrics=["Revenue"])), cresolver(), CLOSED)
    own = [r for r in reqs if r.query == "fm"]
    assert len(own) == 2 and all(r.products == () for r in own)


def test_unknown_concept_and_unknown_member_are_refused():
    d = cspec()
    d["blocks"][0]["concept"] = "nope"
    with pytest.raises(GuardError, match="nope"):
        prepare(d, cresolver(), FakeProvider(cfacts()))
    r = resolver(extra_members={"pd.SY.T1", "ac.C80311"}, concepts={"tiny": TINY})
    with pytest.raises(GuardError, match="pd.SY.T9"):
        prepare(cspec(), r, FakeProvider(cfacts()))


class ProductProvider(FakeProvider):
    def fetch(self, req):
        got = [f for f in self.frame.facts if f.query == req.query and f.period in req.periods
               and f.account in req.accounts and f.entity in req.entities
               and (f.product in req.products if req.products else not f.product)]
        from smartass.data.provider import FetchResult
        return FetchResult(req, got, [], "http", 0.01)


def test_prepare_end_to_end_with_caveats_from_the_notes():
    pre = prepare(cspec(), cresolver(), ProductProvider(cfacts()))
    kv = pre.result.key_values
    assert kv["Tiny - direct contribution · Direct contribution Tiny"]["Actual"] == pytest.approx(27.0)
    texts = [c["text"] for c in pre.caveats]
    assert "Tiny note." in texts
    assert "Financing cost: Not visible in FinMax." in texts
    assert any(t.startswith("COGS (estimated): estimated = base (COGS total) × key") for t in texts)
    assert all(c["severity"] == "info" for c in pre.caveats)


def test_a_concept_found_by_its_name():
    d = cspec()
    d["blocks"][0]["concept"] = "Tiny Offer"
    pre = prepare(d, cresolver(), ProductProvider(cfacts()))
    assert pre.result.outs[0].title == "Tiny - direct contribution"
