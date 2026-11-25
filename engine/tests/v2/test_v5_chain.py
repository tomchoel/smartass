# engine/tests/v2/test_v5_chain.py
"""Plan 5 Task 2: the P&L chain (completed in the spec on new questions, ordered and typed in build_table), the
code column, added lines last under the key_values cap, and the FY label."""
from v2util import CLOSED, chain_resolver, facts, resolver, spec_dict

from smartass.data.provider import FetchResult
from smartass.model.blocks.common import evaluate
from smartass.model.blocks.table import build_table
from smartass.model.chain import CHAIN, complete_chain
from smartass.model.context import Ctx
from smartass.model.findings import key_values
from smartass.model.periods import period_label
from smartass.model.spec import AnalysisSpec
from smartass.pipeline import prepare

ALL7 = ["Revenue", "COGS", "Gross margin", "Opex", "EBITDA before other items", "Other items", "EBITDA"]


def spec(metrics, table=None, cube="FinMax", mode="pack"):
    b = {"type": "table", "query": "fm", "title": "P&L", "metrics": metrics, **(table or {})}
    d = spec_dict(blocks=[b])
    d["plan"]["mode"] = mode
    d["queries"][0].update(cube=cube, metrics=list(metrics))
    if cube == "FinPlan":
        d["queries"][0]["compare"] = ["budget"]
        b["columns"] = ["value", "budget", "var"]
    return AnalysisSpec.model_validate(d)


def test_chain_names():
    assert CHAIN == ALL7


def test_complete_chain_adds_the_seven_to_table_and_query():
    s = spec(["EBITDA"])
    out = complete_chain(s, chain_resolver())
    b, q = out.blocks[0], out.queries[0]
    assert b.metrics[0] == "EBITDA" and sorted(b.metrics) == sorted(ALL7)
    assert set(q.metrics) == set(ALL7)
    assert b.chain_added == [m for m in ALL7 if m != "EBITDA"]
    assert s.blocks[0].metrics == ["EBITDA"] and s.queries[0].metrics == ["EBITDA"]     # not mutated
    assert complete_chain(out, chain_resolver()).model_dump() == out.model_dump()        # idempotent


def test_complete_chain_skips_what_the_cube_cannot_resolve():
    out = complete_chain(spec(["EBITDA"], cube="FinPlan"), chain_resolver())
    assert "Other items" not in out.blocks[0].metrics and "Other items" not in out.queries[0].metrics
    assert set(out.blocks[0].metrics) == set(ALL7) - {"Other items"}


def test_complete_chain_leaves_tables_without_ebitda_or_chain_false():
    r = chain_resolver()
    assert complete_chain(spec(["Revenue", "Opex"]), r).blocks[0].metrics == ["Revenue", "Opex"]
    assert complete_chain(spec(["EBITDA"], {"chain": False}), r).blocks[0].metrics == ["EBITDA"]


def test_complete_chain_on_a_table_defaulting_to_query_metrics():
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "title": "P&L"}])
    d["plan"]["mode"] = "pack"
    out = complete_chain(AnalysisSpec.model_validate(d), chain_resolver())
    assert out.blocks[0].metrics[:3] == ["Revenue", "Opex", "EBITDA"]                    # m0..m2 keep their index
    assert set(out.queries[0].metrics) == set(ALL7)


def test_adhoc_mode_keeps_explicit_metrics_unless_chain_explicit():
    r = chain_resolver()
    s = spec(["Revenue", "EBITDA"], mode="adhoc")
    out = complete_chain(s, r)
    assert out.blocks[0].metrics == ["Revenue", "EBITDA"]
    explicit = spec(["EBITDA"], table={"chain": True}, mode="adhoc")
    assert set(complete_chain(explicit, r).blocks[0].metrics) == set(ALL7)


class Prov:
    def __init__(self):
        self.asked = []

    def closed_months(self):
        return {"FinMax": CLOSED, "FinLite": CLOSED, "FinPlan": CLOSED}

    def fetch(self, req):
        self.asked.append(req)
        return FetchResult(req, [f for f in facts().facts if f.account in req.accounts and f.entity in req.entities
                                 and f.period in req.periods], [], "http", 0.0)


def test_completed_spec_fetches_the_chain_accounts():
    prov = Prov()
    s = complete_chain(spec(["EBITDA"]), chain_resolver())
    pre = prepare(s.model_dump(), chain_resolver(), prov)
    accs = {a for r in prov.asked for a in r.accounts}
    assert {"ac.C30", "ac.C40S", "ac.C48", "ac.C70", "ac.EBITDA_BOI", "ac.C73", "ac.EBITDA"} <= accs
    assert pre.raw["blocks"][0]["chain_added"]


def test_prepare_complete_stores_the_completed_spec():
    prov = Prov()
    pre = prepare(spec(["EBITDA"]).model_dump(), chain_resolver(), prov, complete=True)
    assert pre.raw["blocks"][0]["chain_added"] and set(pre.raw["queries"][0]["metrics"]) == set(ALL7)
    assert "ac.C48" in {a for r in prov.asked for a in r.accounts}
    plain = prepare(spec(["EBITDA"]).model_dump(), chain_resolver(), Prov())
    assert plain.raw["blocks"][0]["metrics"] == ["EBITDA"]                  # without complete: as given


def build(s, r=None):
    c = Ctx(s, r or chain_resolver(), CLOSED, facts())
    return c, build_table(s.blocks[0], c, "b0")


def test_build_table_orders_and_types_the_chain_with_stable_keys():
    s = complete_chain(spec(["EBITDA"]), chain_resolver())
    _, out = build(s)
    assert [ln.label for ln in out.lines] == ALL7
    assert [ln.kind for ln in out.lines] == ["line", "line", "sum", "line", "sum", "line", "grand"]
    assert [ln.code for ln in out.lines] == ["C30", "C40S", "C48", "C70", "EBITDA_BOI", "C73", "EBITDA"]
    assert out.lines[-1].key == "m0" and not out.lines[-1].added
    assert all(ln.added for ln in out.lines[:-1])


def test_given_metrics_keep_their_keys():
    s = complete_chain(spec(["EBITDA", "Revenue"]), chain_resolver())
    _, out = build(s)
    keys = {ln.label: ln.key for ln in out.lines}
    assert keys["EBITDA"] == "m0" and keys["Revenue"] == "m1"
    assert [ln.label for ln in out.lines] == ALL7
    assert not any(ln.added for ln in out.lines if ln.label in ("EBITDA", "Revenue"))


def test_non_chain_metrics_follow_the_chain():
    _, out = build(spec(["FTE", "EBITDA", "Revenue"]))
    assert [ln.label for ln in out.lines] == ["Revenue", "EBITDA", "FTE"]
    assert [ln.key for ln in out.lines] == ["m2", "m1", "m0"]


def test_chain_false_leaves_metrics_as_given():
    _, out = build(spec(["EBITDA", "Revenue"], {"chain": False}))
    assert [ln.label for ln in out.lines] == ["EBITDA", "Revenue"]
    assert [ln.kind for ln in out.lines] == ["line", "line"]


def test_stored_spec_without_completion_behaves_as_before():
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "table", "query": "fm", "title": "P&L"}]))
    c, out = build(s, resolver())
    assert [(ln.key, ln.label) for ln in out.lines] == [("m0", "Revenue"), ("m1", "Opex"), ("m2", "EBITDA")]
    assert list(key_values([out], evaluate([out], c.frame))) == ["P&L · Revenue", "P&L · Opex", "P&L · EBITDA"]


def test_entity_line_codes():
    s = AnalysisSpec.model_validate(spec_dict(blocks=[{"type": "table", "query": "fm", "rows": "entities",
                                                       "metric": "Opex"}]))
    _, out = build(s, resolver())
    codes = {ln.key: ln.code for ln in out.lines}
    assert {codes["e0"], codes["e1"]} == {"4BUS_B", "G910"}
    assert codes["total"] == "2BUS" and codes["check"] == ""


def test_added_lines_last_under_the_cap():
    s = complete_chain(spec(["EBITDA"]), chain_resolver())
    c, out = build(s)
    cells = evaluate([out], c.frame)
    assert list(key_values([out], cells, limit=1)) == ["P&L · EBITDA"]     # the given line wins the only slot
    assert list(key_values([out], cells, limit=3)) == ["P&L · Revenue", "P&L · COGS", "P&L · EBITDA"]


def test_fy_label():
    assert period_label([(2026, m) for m in range(1, 13)]) == "FY 26"
    assert period_label([(2026, m) for m in range(1, 12)]) == "YTD Nov 26"
