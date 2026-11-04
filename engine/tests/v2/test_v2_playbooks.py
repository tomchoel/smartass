# engine/tests/v2/test_v2_playbooks.py
"""Plan 2B task 4: the four playbooks + plan + critic as MCP prompts, the playbooks and conventions as resources."""
import asyncio
from pathlib import Path

from smartass import server

KNOWLEDGE = Path(__file__).resolve().parents[3] / "knowledge"
PLAYBOOKS = ("monthly-pack", "variance-commentary", "gl-extract", "prodreg-input")


def test_every_playbook_exists_and_is_under_6000_characters():
    for name in PLAYBOOKS:
        text = (KNOWLEDGE / "playbooks" / f"{name}.md").read_text(encoding="utf-8")
        assert 800 < len(text) < 6000, name


def test_variance_commentary_quotes_the_vendored_text_with_attribution():
    text = (KNOWLEDGE / "playbooks" / "variance-commentary.md").read_text(encoding="utf-8")
    assert "Apache" in text and "NOTICE.md" in text and "financial-services-plugins" in text
    vendored = (KNOWLEDGE / "advisor" / "variance-commentary.md").read_text(encoding="utf-8")
    quoted = "\n".join(ln[2:] if ln.startswith("> ") else ln[1:] for ln in text.splitlines() if ln.startswith(">"))
    for line in vendored.splitlines():
        if line.strip():
            assert line in quoted, line                       # quoted whole, never reworded
    assert "FinPlan" in text and "hypothesis" in text and "caveat" in text


def test_playbooks_name_their_method():
    read = {n: (KNOWLEDGE / "playbooks" / f"{n}.md").read_text(encoding="utf-8") for n in PLAYBOOKS}
    assert all(t in read["monthly-pack"] for t in ("fin_refresh", "fin_critic", "closed-ytd", "patch"))
    assert all(t in read["gl-extract"] for t in ("fin_drill", "lines", "Yr 26", "exact"))
    assert all(t in read["prodreg-input"] for t in ("fin_drill", "YTD", "reference/intervjuforberedelse.md"))


def test_prompts_are_listed():
    names = {p.name for p in asyncio.run(server.mcp.list_prompts())}
    assert names == {"plan", "critic", *PLAYBOOKS}


def test_a_prompt_returns_the_markdown_with_a_preface():
    res = asyncio.run(server.mcp.get_prompt("critic"))
    text = res.messages[0].content.text
    rubric = (KNOWLEDGE / "agent" / "critic.md").read_text(encoding="utf-8")
    assert text.endswith(rubric) and len(text.split("\n", 1)[0]) < 200 and text.split("\n", 1)[0] != ""


def test_resources_are_listed_and_readable():
    uris = {str(r.uri) for r in asyncio.run(server.mcp.list_resources())}
    assert uris == {*(f"smartass://playbooks/{n}" for n in PLAYBOOKS), "smartass://conventions"}
    got = list(asyncio.run(server.mcp.read_resource("smartass://playbooks/gl-extract")))
    assert "fin_drill" in got[0].content
    conv = list(asyncio.run(server.mcp.read_resource("smartass://conventions")))[0].content
    assert "FinPlan" in conv and "(1 000 000)" in conv and "costs" in conv
