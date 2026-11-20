# engine/tests/v2/test_v5_material.py
"""Plan 5 Task 3: the live Material marker column (1/0 under the findings.material rule, thresholds in labelled
Param cells), Result.params, and the P&L chain completed on the follow-up patch path too. Every figure is
invented."""
import asyncio
import itertools


from v2util import CLOSED, chain_resolver, facts, resolver, spec_dict

from smartass import server
from smartass.followup import merge_patch
from smartass.model.analyze import compute
from smartass.model.blocks.common import ck
from smartass.model.chain import CHAIN
from smartass.model.expr import Material, Param, Ref
from smartass.model.findings import material
from smartass.model.frame import Frame
from smartass.model.spec import AnalysisSpec
from smartass.pipeline import prepare
from smartass.render.layout import MARK_CODES, col_letter, layout
from smartass.render.readback import compare


class C:
    def __init__(self, d):
        self.d = d

    def get(self, k):
        return self.d[k]


A, PCT = 100.0, 0.05
M = Material(Ref("v"), Ref("p"), Param("mat_abs.X", A), Param("mat_pct.X", PCT), Param("mat_small.X", A / 10))


def rule(v, p):
    """findings.material's rule for one cell."""
    if v is None:
        return 0.0
    return 1.0 if abs(v) >= A or (p is not None and abs(p) >= PCT and abs(v) >= A / 10) else 0.0


def test_evaluate_matches_the_material_rule_on_a_grid():
    for v, p in itertools.product([None, 0.0, 5.0, -10.0, 50.0, -99.9, 100.0, -250.0],
                                  [None, 0.0, 0.01, -0.05, 0.2, -3.0]):
        assert M.evaluate(Frame([]), C({"v": v, "p": p})) == rule(v, p), (v, p)
    assert M.evaluate(Frame([]), C({"v": 150.0, "p": None})) == 1.0              # above abs
    assert M.evaluate(Frame([]), C({"v": 20.0, "p": 0.1})) == 1.0                # pct, var above a/10
    assert M.evaluate(Frame([]), C({"v": 5.0, "p": 0.5})) == 0.0                 # pct, var below a/10
    assert M.evaluate(Frame([]), C({"v": 50.0, "p": 0.01})) == 0.0               # below both
    no_pct = Material(Ref("v"), None, M.a, M.pct_t, M.a10)
    assert no_pct.evaluate(Frame([]), C({"v": 99.0})) == 0.0
    assert no_pct.evaluate(Frame([]), C({"v": -100.0})) == 1.0


def test_to_excel_references_the_param_cells():
    addr = {"v": "E10", "p": "F10", "mat_abs.X": "$C$40", "mat_pct.X": "$C$41", "mat_small.X": "$C$42"}.__getitem__
    # N(): a blank variance % (Excel "") counts as 0 instead of making ABS() a #VALUE!
    assert M.to_excel("Data_Q1", addr) == "IF(OR(ABS(E10)>=$C$40,AND(ABS(N(F10))>=$C$41,ABS(E10)>=$C$42)),1,0)"
    no_pct = Material(Ref("v"), None, M.a, M.pct_t, M.a10)
    assert no_pct.to_excel("Data_Q1", addr) == "IF(ABS(E10)>=$C$40,1,0)"


def table_spec(table=None, metrics=("Revenue", "Opex", "EBITDA")):
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "title": "P&L", "metrics": list(metrics),
                           "chain": False, **(table or {})}])
    d["queries"][0]["metrics"] = list(metrics)
    return AnalysisSpec.model_validate(d)


def test_a_table_with_var_gets_the_material_column():
    s = table_spec()
    res = compute(s, resolver(), CLOSED, facts())
    o = res.outs[0]
    assert o.cols[-1] == "material" and o.headers[-1] == "Material"
    assert len(o.cols) == len(o.headers) == len(o.keys)
    j = len(o.cols) - 1
    hits = {h["key"] for h in material(res.outs, res.cells, resolver())}
    marks = {f"b0.{ln.key}": res.cells.get(ck("b0", ln.key, j)) for ln in o.lines}
    assert marks == {k: (1.0 if k in hits else 0.0) for k in marks}
    assert hits and any(v == 0.0 for v in marks.values())             # both outcomes occur in the fixture
    assert all("Material" not in kv for kv in res.key_values.values())
    assert res.params["mat_abs.Opex"] == ("Materiality: Opex (NOK)", 5.0)
    assert res.params["mat_pct.Opex"] == ("Materiality %: Opex", 0.05)
    assert res.params["mat_small.Opex"] == ("Materiality with %: Opex (NOK)", 0.5)


def test_material_false_leaves_it_out():
    res = compute(table_spec({"material": False}), resolver(), CLOSED, facts())
    assert "material" not in res.outs[0].cols
    assert not any(k.startswith("mat_") for k in res.params)


def test_no_var_column_no_material_column():
    res = compute(table_spec({"columns": ["value", "ly"]}), resolver(), CLOSED, facts())
    assert "material" not in res.outs[0].cols


def test_check_and_metricless_lines_get_no_marker():
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "rows": "entities", "metric": "Opex"}])
    s = AnalysisSpec.model_validate(d)
    res = compute(s, resolver(), CLOSED, facts())
    o = res.outs[0]
    j = o.cols.index("material")
    by = {ln.key: ln for ln in o.lines}
    assert by["check"].cells[j] is None
    assert isinstance(by["total"].cells[j], Material)


def test_column_groups_judge_the_last_group():
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "title": "P&L", "metrics": ["Opex"], "chain": False,
                           "periods": ["closed", "closed-ytd"]}])
    d["queries"][0]["periods"] = "closed-ytd"
    d["queries"][0]["metrics"] = ["Opex"]
    res = compute(AnalysisSpec.model_validate(d), resolver(), CLOSED, facts())
    o = res.outs[0]
    m = o.lines[0].cells[-1]
    gap = o.cols.index("gap")
    var_last = max(i for i, c in enumerate(o.cols) if c == "var")
    assert var_last > gap and m.var == Ref(ck("b0", "m0", var_last))
    assert o.cols.count("material") == 1 and o.keys[-1] == ""


def test_layout_writes_labelled_thresholds_and_the_live_formula():
    s = table_spec()
    res = compute(s, resolver(), CLOSED, facts())
    p = layout(3, s, res, stock=set(), subtitle="x", const={})
    labels = {v: row for (row, col), v in p.cells.items() if col == 3 and isinstance(v, str)}
    head = labels["Metric"]                               # the compact thresholds table: one row per metric
    cols = {p.cells[(head, c)]: c for c in range(4, 8)}
    assert list(cols) == ["Var % floor", "Materiality", "Materiality %", "Materiality with %"]
    rt = labels["Opex (NOK)"]
    ca, cp, cs = (cols[x] for x in ("Materiality", "Materiality %", "Materiality with %"))
    assert (p.cells[(rt, ca)], p.cells[(rt, cp)], p.cells[(rt, cs)]) == (5.0, 0.05, 0.5)
    a, pp, s = (f"${col_letter(c)}${rt}" for c in (ca, cp, cs))
    assert (p.params["mat_abs.Opex"], p.params["mat_pct.Opex"], p.params["mat_small.Opex"]) == (a, pp, s)
    row_num = next(x for x in p.rows if x["row"] == rt)["num"]
    assert (row_num[col_letter(cp)], row_num[col_letter(ca)], row_num[col_letter(cs)]) == ("prosent1", "heltall",
                                                                                           "heltall")
    o = res.outs[0]
    line = next(ln.key for ln in o.lines if ln.metric == "Opex")
    j = o.cols.index("material")
    r, c = p.at(ck("b0", line, j))
    v, pc = col_letter(c - 2), col_letter(c - 1)          # var, var_pct sit left of the marker
    assert p.cells[(r, c)] == (f"=IF(OR(ABS({v}{r})>={a},AND(ABS(N({pc}{r}))>={pp},"
                               f"ABS({v}{r})>={s})),1,0)")
    row = next(x for x in p.rows if x["row"] == r)
    assert "material" in MARK_CODES and row["num"][col_letter(c)] == "markor"


def test_readback_compares_the_markers():
    s = table_spec()
    res = compute(s, resolver(), CLOSED, facts())
    p = layout(3, s, res, stock=set(), subtitle="x", const={})
    grid = [[None] * 12 for _ in range(p.last_row)]
    for key in p.addr:
        r, c = p.at(key)
        v = res.cells.get(key)
        grid[r - 1][c - 1] = "" if v is None else v
    assert compare(p.addr, res.cells, tuple(tuple(x) for x in grid), p.fractions) == []
    o = res.outs[0]
    j = o.cols.index("material")
    key = next(ck("b0", ln.key, j) for ln in o.lines if res.cells.get(ck("b0", ln.key, j)) == 1.0)
    r, c = p.at(key)
    grid[r - 1][c - 1] = 0
    assert compare(p.addr, res.cells, tuple(tuple(x) for x in grid), p.fractions) != []


# ---- the chain on the patch path

class Prov:
    def closed_months(self):
        return {"FinMax": CLOSED, "FinLite": CLOSED, "FinPlan": CLOSED}

    def fetch(self, req):
        from smartass.data.provider import FetchResult
        return FetchResult(req, [f for f in facts().facts if f.account in req.accounts and f.entity in req.entities
                                 and f.period in req.periods], [], "http", 0.0)


def chain_spec():
    d = spec_dict(blocks=[{"type": "table", "query": "fm", "title": "P&L", "metrics": ["EBITDA"]}])
    d["queries"][0]["metrics"] = ["EBITDA"]
    return d


def test_patch_resending_queries_without_the_chain_is_completed_again():
    stored = prepare(chain_spec(), chain_resolver(), Prov(), complete=True).raw
    q = dict(stored["queries"][0], metrics=["EBITDA"])
    patched = merge_patch(stored, {"queries": [q]})
    pre = prepare(patched, chain_resolver(), Prov(), complete=True)
    assert set(pre.raw["queries"][0]["metrics"]) == set(CHAIN)
    assert set(pre.raw["blocks"][0]["metrics"]) == set(CHAIN)


def test_patch_path_completes_the_chain_and_refresh_does_not(monkeypatch):
    async def user(fn, what):
        return fn(None)
    seen = []
    stored = {2: {"spec": {"v2": {"question": "a"}}, "ids": {}, "owned": None}}
    monkeypatch.setattr(server, "_user", user)
    monkeypatch.setattr(server, "_provider", lambda: "provider")
    monkeypatch.setattr(server, "read_stored", lambda h, target="session": (stored, "Q2 Answer", "C:/a/x.xlsx"))
    monkeypatch.setattr(server, "run_prepare", lambda raw, provider, lines=None, **kw: seen.append(kw) or raw)
    monkeypatch.setattr(server, "redraw_job", lambda h, pre, q, target="session": {"q": q, "closed": {}})
    asyncio.run(server.fin_analyze(patch={"plan": {"purpose": "x"}}))
    assert seen == [{"complete": True}]
    seen.clear()
    asyncio.run(server.fin_refresh())
    assert seen == [{}]
