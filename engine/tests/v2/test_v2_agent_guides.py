# engine/tests/v2/test_v2_agent_guides.py
"""The agent's rubric (critic.md) and plan guide (plan.md): present, short, and the rubric names its 6 criteria."""
from pathlib import Path

AGENT = Path(__file__).resolve().parents[3] / "knowledge" / "agent"
CRITERIA = ("Right question", "Numbers reconciled", "Comparison basis", "Drivers backed by data",
            "Caveats respected", "Clear for a controller")


def test_both_guides_exist_and_are_short():
    for name in ("critic.md", "plan.md"):
        text = (AGENT / name).read_text(encoding="utf-8")
        assert 500 < len(text) < 4000, name


def test_critic_names_all_six_criteria_and_the_pass_mark():
    text = (AGENT / "critic.md").read_text(encoding="utf-8")
    for c in CRITERIA:
        assert c in text, c
    assert "at least 4" in text and "2 rounds" in text


def test_plan_guide_covers_purpose_bridge_drivers_lines_and_caveats():
    text = (AGENT / "plan.md").read_text(encoding="utf-8").casefold()
    for word in ("purpose", "hypotheses", "audience", "depth", "bridge", "other items", "drivers", "lines",
                 "caveat"):
        assert word in text, word
