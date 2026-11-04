import pytest

from smartass.knowledge.semantic import Fact, Metric, SemanticError, load_semantic

OPEX = """
name = "Opex"
aliases = ["opex", "driftskostnader", "C70"]
description = "Total operating expenses."
sign = "cost"
materiality = [5000000, 0.10]
status = "approved"
approved = { by = "tom", on = "2026-10-15" }
differs = "FinPlan groups accounts differently."

[finmax]
members = ["ac.C70"]

[finplan]
members = ["OPEX"]
"""

FACT = """
text = "Consumer Norway was abolished on 1 Jun 2026."
scope_members = { FinMax = ["un.SY.2CON"] }
valid_from = "2026-06"
severity = "warn"
source = "catalog.toml [[entity]] consumer (approved 2026-10-15)"
status = "approved"
"""


def _write(d, kind, stem, text):
    p = d / kind / f"{stem}.toml"
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding="utf-8")
    return p


def test_valid_metric_and_fact_load(tmp_path):
    _write(tmp_path, "metrics", "opex", OPEX)
    _write(tmp_path, "facts", "abolished_consumer", FACT)
    metrics, facts = load_semantic(tmp_path)
    (m,), (f,) = metrics, facts
    assert isinstance(m, Metric) and isinstance(f, Fact)
    assert m.id == "opex" and m.name == "Opex" and m.sign == "cost"
    assert m.finmax.members == ["ac.C70"] and m.finplan.members == ["OPEX"] and m.finlite is None
    assert m.materiality == (5_000_000.0, 0.10) and m.status == "approved"
    assert m.approved == {"by": "tom", "on": "2026-10-15"}
    assert f.id == "abolished_consumer" and f.scope_members == {"FinMax": ["un.SY.2CON"]}
    assert f.valid_from == "2026-06" and f.severity == "warn" and f.status == "approved"


def test_accumulated_with_sum_is_refused():
    with pytest.raises(ValueError, match="Accumulated"):
        Metric(id="fte", name="FTE", description="x", balance="Accumulated", aggregation="sum",
               finmax={"members": ["ac.STA20Y"]})
    m = Metric(id="fte", name="FTE", description="x", balance="Accumulated", aggregation="last",
               finmax={"members": ["ac.STA20Y"]})
    assert m.aggregation == "last"


def test_bad_valid_from_is_refused(tmp_path):
    with pytest.raises(ValueError, match="valid_from"):
        Fact(id="f", text="t", source="s", scope_metrics=["opex"], valid_from="2026-6")
    _write(tmp_path, "facts", "bad", FACT.replace('"2026-06"', '"June 2026"'))
    with pytest.raises(SemanticError, match="bad.toml"):
        load_semantic(tmp_path)


def test_metric_needs_a_cube_and_fact_needs_a_scope():
    with pytest.raises(ValueError, match="cube"):
        Metric(id="x", name="X", description="x")
    with pytest.raises(ValueError, match="scope"):
        Fact(id="f", text="t", source="s")


def test_unknown_keys_and_bad_ids_are_refused(tmp_path):
    with pytest.raises(ValueError):
        Metric(id="x", name="X", description="x", finmax={"members": ["a"]}, colour="red")
    with pytest.raises(ValueError, match="id"):
        Metric(id="Bad-Id", name="X", description="x", finmax={"members": ["a"]})


def test_duplicate_id_is_refused(tmp_path):
    _write(tmp_path, "metrics", "opex", OPEX)
    _write(tmp_path, "facts", "opex", FACT)
    with pytest.raises(SemanticError, match="duplicate id 'opex'"):
        load_semantic(tmp_path)


def test_invalid_toml_names_the_file(tmp_path):
    _write(tmp_path, "metrics", "broken", "name = ")
    with pytest.raises(SemanticError, match="broken.toml"):
        load_semantic(tmp_path)


def test_missing_dir_loads_nothing(tmp_path):
    assert load_semantic(tmp_path / "nope") == ([], [])
