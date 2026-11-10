"""Plan 2C task 5: the skill and its reference name the concept block and product rows; the test set has b09 on the
mobile subscriptions family concept."""
from pathlib import Path

from smartass.core import config
from smartass.knowledge.semantic import load_concepts
from smartass.model.spec import AnalysisSpec, ConceptBlock
from smartass.testset import runner

SKILL = Path(__file__).resolve().parents[3] / "skills" / "smartass"
REPO = SKILL.parents[1]


def test_skill_names_the_concept_block_and_product_rows():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    assert "`concept`" in text and 'rows: "products"' in text and "products" in text
    assert "reference/concepts.md" in text


def test_the_concepts_reference_explains_the_items_and_products():
    ref = (SKILL / "reference" / "concepts.md").read_text(encoding="utf-8")
    for word in ("booked", "attributed", "estimated", "missing", "Key share", "kind=\"concept\"", "splitt",
                 "All Products", "pd.SY.", 'rows": "products"', "finding"):
        assert word in ref, word


def test_b09_is_the_mobile_subscriptions_concept():
    case = next(c for c in runner.load_cases() if c["id"] == "b09")
    assert case["question"] == "What do we earn on mobile subscriptions?"
    spec = AnalysisSpec.model_validate(case["spec"])
    (block,) = [b for b in spec.blocks if isinstance(b, ConceptBlock)]
    concepts = {c.id for c in load_concepts(config.SEMANTIC_DIR)}
    assert block.concept == "fam_mobil_own_subscriptions_mobile_consumer" and block.concept in concepts
    q = spec.queries[0]
    assert (q.cube, q.entity, q.periods, q.compare) == ("FinMax", "un.SY", "closed-ytd", ["ly"])


def test_changelog_and_design_name_products_and_concepts():
    for f in (REPO / "CHANGELOG.md", REPO / "docs" / "DESIGN.md"):
        text = f.read_text(encoding="utf-8")
        assert "Plan 2C" in text and "concept" in text and "products" in text, f.name
