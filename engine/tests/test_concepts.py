"""concepts.py: the Splitt definition plans into consistent pulls + spec (every referenced key exists)."""
import re
import sys

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
from smartview_mcp import concepts, specs  # noqa: E402


def keys_of(spec):
    return {ln["key"] for s in spec["sections"] for ln in s["lines"] if ln.get("key")}


def test_splitt_plan_is_consistent():
    c = concepts.get("splitt")
    pulls, spec = concepts.plan(c, "Sep 26", "Sep 25")
    ids = {p["id"] for p in pulls}
    assert ids == {"p", "e"}
    prod = next(p for p in pulls if p["id"] == "p")
    assert "pd.SY.M4517" in prod["rows"][0]["members"] and "ac.C30" in prod["columns"][0]["members"]
    assert prod["columns"][1]["members"] == ["Sep 26", "Sep 25"]
    keys = keys_of(spec)
    for s in spec["sections"]:
        for ln in s["lines"]:
            for cell in ln.get("cells", []):
                if not isinstance(cell, dict):
                    continue
                for k in (cell.get("sum") or []) if isinstance(cell.get("sum"), list) else []:
                    assert k in keys, (ln["key"], k)
                if "expr" in cell:
                    for k in re.findall(r"[A-Za-z_][A-Za-z0-9_]*", cell["expr"]):
                        assert k in keys, (ln["key"], k)
                if "ref" in cell:
                    assert cell["ref"]["key"] in keys
                if "pull" in cell:
                    assert cell["pull"] in ids
    total = next(ln for s in spec["sections"] for ln in s["lines"] if ln["key"] == "total")
    assert "hw_cogs" in total["cells"][0]["sum"] and "fin_cost" not in total["cells"][0]["sum"]
    assert spec["checks"][0]["expr"] == "hw_rev-chk0"


def test_find():
    assert concepts.find("hva er lønnsomheten til Splitt?")[0]["id"] == "splitt"


def test_expand_passes_through():
    c = concepts.get("splitt")
    pulls, spec = concepts.plan(c, "Sep 26", "Sep 25")
    s = specs.expand(spec, {"p": "Q1 FinMax", "e": "Q1 FinMax e"}, {})
    cell = next(ln for sec in s["sections"] for ln in sec["lines"] if ln["key"] == "d1")["cells"][0]
    assert cell["pull"] == "Q1 FinMax"


def test_concept_checks_use_the_total_not_plus_chains():
    """tn-formelstil blocks 3+ cells added with + (PLUS_CHAIN): a check compares the total line instead."""
    import json
    import re
    from pathlib import Path
    data = json.loads((Path(__file__).parents[2] / "knowledge" / "concepts.json").read_text(encoding="utf-8"))
    concepts = data["concepts"] if isinstance(data, dict) else data
    for c in concepts:
        for ch in c.get("checks", []):
            assert len(re.findall(r"[+]", ch["expr"])) < 2, (c["id"], ch["expr"])


def test_family_concept_check_refers_to_existing_keys():
    c = concepts.get("fam_mobil_own_subscriptions_mobile_consumer")
    pulls, spec = concepts.plan(c, "Sep 26", "Sep 25")
    for ch in spec["checks"]:
        for k in re.findall(r"[A-Za-z_][A-Za-z0-9_]*", ch["expr"]):
            assert k in keys_of(spec), (k, ch["expr"])
