"""Plan 2C task 3: product concepts as semantic files - schema, the seed from knowledge/concepts.json, the build into
knowledge.db (aliases unique, never a metric's alias) and fin_lookup."""
import json
import sqlite3
import tomllib

import pytest
from kutil import knowledge

from smartass.core import config
from smartass.knowledge.build import build
from smartass.knowledge.seed import seed, seed_concepts, seed_metrics
from smartass.knowledge.semantic import Concept, SemanticError, concept_terms, load_concepts, to_toml

TINY = '''title = "Tiny offer - direct contribution"
names = ["tiny", "tiny offer"]
kind = "product_offer"
description = "A tiny offer."
status = "approved"
notes = ["Tiny note."]

[total]
label = "Direct contribution Tiny"
of = ["rev", "cogs"]

[[check]]
label = "Revenue = the product's C30"
expr = "rev-$src"

[check.source]
account = "ac.C30"
products = ["pd.SY.T1"]

[[block]]
title = "Revenue and COGS"

[[block.item]]
key = "rev"
label = "Revenue (C30)"
kind = "booked"
account = "ac.C30"
products = ["pd.SY.T1"]

[[block.item]]
key = "cogs"
label = "COGS"
kind = "estimated"

[block.item.base]
label = "COGS total"
account = "ac.C50S"
products = ["pd.SY.T9"]

[block.item.key_share]
text = "Tiny's share of revenue"

[block.item.key_share.num]
label = "Tiny revenue"
expr = "rev"

[block.item.key_share.den]
label = "Revenue total"
account = "ac.C30"
products = ["pd.SY.T9"]

[[block.item]]
key = "gone"
label = "Financing cost"
kind = "missing"
note = "Not visible in FinMax."
'''


def _concepts_json() -> dict:
    return json.loads((config.KNOWLEDGE_DIR / "concepts.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def seeded():
    return seed_concepts(_concepts_json())


# ---- schema

def test_a_concept_file_loads(tmp_path):
    d = tmp_path / "concepts"
    d.mkdir()
    (d / "tiny.toml").write_text(TINY, encoding="utf-8")
    (c,) = load_concepts(tmp_path)
    assert c.id == "tiny" and c.kind == "product_offer" and c.entity == "un.SY"
    cogs = c.block[0].item[1]
    assert cogs.kind == "estimated" and cogs.base.account == "ac.C50S"
    assert cogs.key_share.num.expr == "rev" and cogs.key_share.den.products == ["pd.SY.T9"]
    assert c.check[0].source.products == ["pd.SY.T1"] and c.total.of == ["rev", "cogs"]


@pytest.mark.parametrize("bad, match", [
    (TINY.replace('of = ["rev", "cogs"]', 'of = ["rev", "nope"]'), "nope"),
    (TINY.replace('expr = "rev-$src"', 'expr = "rev*2"'), "expr"),
    (TINY.replace('kind = "estimated"', 'kind = "guessed"'), "kind"),
    (TINY.replace('[block.item.base]\nlabel = "COGS total"\naccount = "ac.C50S"\nproducts = ["pd.SY.T9"]\n', ''),
     "estimated needs base"),
    (TINY.replace('expr = "rev"\n', 'expr = "rev"\naccount = "ac.C30"\n'), "either account"),
    (TINY.replace('key = "gone"', 'key = "rev"'), "duplicate"),
])
def test_bad_concepts_are_refused(tmp_path, bad, match):
    d = tmp_path / "concepts"
    d.mkdir()
    (d / "tiny.toml").write_text(bad, encoding="utf-8")
    with pytest.raises(SemanticError, match=match):
        load_concepts(tmp_path)


def test_concept_toml_round_trips(tmp_path):
    c = Concept.model_validate({**tomllib.loads(TINY), "id": "tiny"})
    d = tmp_path / "concepts"
    d.mkdir()
    (d / "tiny.toml").write_text(to_toml(c), encoding="utf-8")
    assert load_concepts(tmp_path) == [c]


# ---- the seed from concepts.json

def test_the_seed_gives_24_approved_concepts_that_load(seeded, tmp_path):
    assert len(seeded) == 24
    assert {c.kind for c in seeded} == {"product_offer", "product_family"}
    assert sum(c.id.startswith("fam_") for c in seeded) == 22
    assert all(c.status == "approved" and c.approved for c in seeded)
    d = tmp_path / "concepts"
    d.mkdir()
    for c in seeded:
        (d / f"{c.id}.toml").write_text(to_toml(c), encoding="utf-8")
    assert load_concepts(tmp_path) == sorted(seeded, key=lambda c: c.id)


def test_splitt_has_its_blocks_and_keys(seeded):
    s = next(c for c in seeded if c.id == "splitt")
    items = {it.key: it for b in s.block for it in b.item}
    assert len(s.block) == 6
    assert items["hw_apple"].kind == "booked" and items["hw_apple"].products[0] == "pd.SY.FV064"
    assert items["hw_rev"].kind == "sum" and items["hw_rev"].of == ["hw_apple", "hw_samsung", "hw_other", "hw_gsn",
                                                                    "hw_fin"]
    hw = items["hw_cogs"]
    assert hw.kind == "estimated" and hw.base.account == "ac.C40S" and hw.base.products == ["pd.SY.M5008",
                                                                                         "pd.SY.M5010"]
    assert hw.key_share.num.expr == "hw_rev-hw_fin"
    assert (hw.key_share.den.account, hw.key_share.den.products) == ("ac.C30", ["pd.SY.M411"])
    assert items["ins_cost"].key_share.num.expr == "ins_rev"                 # a 'ref' is a one-term expression
    trade = items["trade"]
    assert trade.key_share.same_as == "hw_cogs" and trade.base.entities[0] == "un.SY.C120"
    assert trade.base.account == "ac.EBITDA_BOI"
    fin = items["fin_int"]
    assert fin.kind == "attributed" and fin.entities == ["un.SY"] and fin.key_share is None and fin.note
    assert items["fin_cost"].kind == "missing" and "C81313" in items["fin_cost"].note
    assert s.check[0].expr == "hw_rev-$src" and s.check[0].source.products == ["pd.SY.M4517"]
    assert "trade" in s.total.of and "fin_cost" not in s.total.of
    assert "splitt" in [n.casefold() for n in s.names]


def test_no_item_is_left_missing_for_want_of_a_readable_key(seeded):
    """Every estimated item of concepts.json has a machine-readable key; only the source's own missing items are
    missing."""
    src = {it["key"]: it["kind"] for c in _concepts_json()["concepts"] for b in c["blocks"] for it in b["items"]
           if it["kind"] == "missing"}
    ours = [it for c in seeded for b in c.block for it in b.item if it.kind == "missing"]
    assert len(ours) == 24 and {it.key for it in ours} == set(src)


def test_streamix_and_a_family(seeded):
    by = {c.id: c for c in seeded}
    sx = by["streamix"]
    assert [b.title for b in sx.block][2].startswith("Memo")
    assert sx.total.of == ["rev", "cogs", "opex"] and sx.check[0].expr == "total-$src"
    fam = by["fam_mobil_own_subscriptions_mobile_consumer"]
    assert fam.kind == "product_family" and len(fam.block[0].item[0].products) == 13
    assert fam.title == "Mobil Own Subscriptions Mobile - direct contribution"


def test_concept_texts_are_english(seeded):
    norwegian = ("Inntekter", "Varekost", "Driftskostnader", "bidrag", "Ikke med", "Lønn", "Forsikring",
                 "Maskinvare", "Finansiering", "Innbytte", "bokført", "andel")
    for c in seeded:
        texts = [c.title, c.description, c.total.label, *c.notes, *(x.label for x in c.check)]
        for b in c.block:
            texts.append(b.title)
            for it in b.item:
                texts += [it.label, it.note]
                if it.key_share:
                    texts.append(it.key_share.text)
        for t in texts:
            assert not any(w in t for w in norwegian), (c.id, t)


def test_concept_aliases_are_unique_and_never_a_metric_alias(seeded):
    from smartass.knowledge.seed import _read
    metrics = seed_metrics(_read(config.KNOWLEDGE_DIR, "catalog.toml"), _read(config.KNOWLEDGE_DIR,
                                                                            "capabilities.toml"))
    metric_terms = {t.casefold() for m in metrics for t in (m.id.replace("_", " "), m.name, *m.aliases)}
    seen: dict[str, str] = {}
    for c in seeded:
        for t in concept_terms(c):
            assert t.casefold() not in metric_terms, (c.id, t)
            assert seen.setdefault(t.casefold(), c.id) == c.id, (c.id, t, seen[t.casefold()])


def test_the_repo_concept_files_are_the_seed(seeded):
    files = sorted((config.SEMANTIC_DIR / "concepts").glob("*.toml"))
    assert [p.stem for p in files] == sorted(c.id for c in seeded)
    by = {c.id: c for c in seeded}
    for p in files:
        assert p.read_text(encoding="utf-8") == to_toml(by[p.stem]), p.name


def test_seed_writes_the_concepts(tmp_path):
    seed(config.KNOWLEDGE_DIR, tmp_path)
    assert len(list((tmp_path / "concepts").glob("*.toml"))) == 24


# ---- build and lookup

def _with_tiny(tmp_path, monkeypatch, text=TINY):
    a, sem, k = knowledge(tmp_path, monkeypatch, build_db=False)
    (sem / "concepts").mkdir()
    (sem / "concepts" / "tiny.toml").write_text(text, encoding="utf-8")
    return a, sem, k


def test_concepts_and_their_aliases_go_into_knowledge_db(tmp_path, monkeypatch):
    a, sem, k = _with_tiny(tmp_path, monkeypatch)
    rep = build(sem, a, k)
    assert rep.ok and rep.concepts == 1
    db = sqlite3.connect(k)
    assert db.execute("SELECT id FROM concept").fetchall() == [("tiny",)]
    terms = {t for (t,) in db.execute("SELECT text FROM concept_alias WHERE concept_id='tiny'")}
    assert {"tiny", "tiny offer"} <= terms
    db.close()
    assert any("pd.SY.T1" in w for w in rep.warnings)       # concept members missing from the atlas: warned


def test_a_concept_alias_that_is_a_metric_alias_stops_the_build(tmp_path, monkeypatch):
    a, sem, k = _with_tiny(tmp_path, monkeypatch, TINY.replace('"tiny offer"', '"driftskostnader"'))
    rep = build(sem, a, k)
    assert not rep.ok and any("driftskostnader" in e for e in rep.errors)


def test_lookup_finds_concepts(tmp_path, monkeypatch):
    from smartass import server
    a, sem, k = _with_tiny(tmp_path, monkeypatch)
    assert build(sem, a, k).ok
    out = server.lookup("tiny", kind="concept")
    (hit,) = out["concepts"]
    assert hit["id"] == "tiny" and hit["kind"] == "product_offer" and hit["notes"] == ["Tiny note."]
    assert [b["title"] for b in hit["blocks"]] == ["Revenue and COGS"]
    assert out["metrics"] == [] and out["members"] == []
    assert server.lookup("tiny offer", kind="any")["concepts"][0]["id"] == "tiny"
    assert server.lookup("driftskostnader", kind="metric")["concepts"] == []


def test_the_resolver_returns_a_concept_by_id_or_alias(tmp_path, monkeypatch):
    from smartass.knowledge.resolver import KnowledgeResolver
    a, sem, k = _with_tiny(tmp_path, monkeypatch)
    assert build(sem, a, k).ok
    r = KnowledgeResolver(sqlite3.connect(k), sqlite3.connect(a), 26)
    try:
        assert r.concept("tiny").id == "tiny" and r.concept("Tiny Offer").id == "tiny"
        assert r.concept("nope") is None
    finally:
        r.close()
