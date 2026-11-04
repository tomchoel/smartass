"""The smartass skill names exactly the tools the server registers, stays short, and has no v1 leftovers."""
import re
from pathlib import Path

from smartass import server

SKILL = Path(__file__).resolve().parents[2] / "skills" / "smartass"
TOOL = re.compile(r"\b((?:fin|xl|sv)_[a-z_]+)\b")


def _text() -> str:
    return (SKILL / "SKILL.md").read_text(encoding="utf-8")


def test_skill_names_exactly_the_registered_tools():
    assert set(TOOL.findall(_text())) == set(server.TOOL_NAMES)


def test_registered_tools_match_tool_names():
    import asyncio
    names = {t.name for t in asyncio.run(server.mcp.list_tools())}
    assert names == set(server.TOOL_NAMES)


def test_skill_is_under_2000_tokens():
    assert len(_text()) / 3.5 < 2000          # conservative chars-per-token for markdown with code


def test_frontmatter_keeps_name_and_description():
    head = _text().split("---")[1]
    assert re.search(r"(?m)^name: smartass$", head) and re.search(r"(?m)^description: .*fin_\* tools", head)


def test_references_exist_and_name_no_v1_tools():
    text = _text()
    for link in re.findall(r"\]\((reference/[^)]+)\)", text):
        assert (SKILL / link).exists(), link
    assert not (SKILL / "reference" / "recipes.md").exists()
    for f in (SKILL / "reference").glob("*.md"):
        body = f.read_text(encoding="utf-8")
        assert not re.search(r"\bsv_(?!version)[a-z_]+\b", body), f.name
        assert set(TOOL.findall(body)) <= set(server.TOOL_NAMES), f.name


def test_skill_covers_caveats_aliases_and_the_monthly_refresh():
    text = _text()
    assert "caveats" in text and "alias" in text
    assert "smartass knowledge refresh" in text and "smartass knowledge approve" in text


def test_products_map_side_files_are_gone():
    knowledge = SKILL.parents[1] / "knowledge"
    assert (knowledge / "products_map.json").exists()
    assert not (knowledge / "products_map.json.leaves.json").exists()
    assert not (knowledge / "products_map.json.q4.json").exists()
