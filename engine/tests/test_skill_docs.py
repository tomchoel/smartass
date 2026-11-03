import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
GONE = ("sv_query", "sv_ws_pull", "sv_ws_build", "sv_ws_answer", "sv_analyse", "sv_concept", "fin_pnl",
        "sv_scope", "sv_warehouse", "sv_refresh_workbook")


def _docs():
    yield ROOT / "skills" / "smartass" / "SKILL.md"
    yield from (ROOT / "skills" / "smartass" / "reference").glob("*.md")
    yield ROOT / "README.md"
    yield ROOT / "docs" / "DESIGN.md"


def test_docs_name_no_retired_tool():
    hits = [f"{p.name}: {t}" for p in _docs() for t in GONE
            if re.search(rf"\b{t}\b", p.read_text(encoding="utf-8"))]
    assert hits == []


def _flat(p: Path) -> str:
    return " ".join(p.read_text(encoding="utf-8").split()).lower()


SKILL = ROOT / "skills" / "smartass" / "SKILL.md"
GATE = ROOT / "knowledge" / "advisor" / "gate.md"


def test_server_instructions_lead_with_gate_then_answer():
    from smartview_mcp import server
    text = server.INSTRUCTIONS
    assert text.index("sv_gate") < text.index("sv_answer")
    assert not [t for t in GONE if re.search(rf"\b{t}\b", text)]
    flat = " ".join(text.split()).lower()
    for must in ("answers=", "lines_ranking", "key_values", "never ask whether to build a sheet", "deviation"):
        assert must in flat, must


def test_skill_says_always_build_sheets_and_never_asks():
    text = _flat(SKILL)
    assert "never ask whether to build a sheet" in text
    assert "several sheets" in text
    assert "lines_ranking" in text and "deviation" in text and "answers=" in text


def test_skill_says_quote_key_values():
    text = _flat(SKILL)
    assert "quote `key_values`" in text and "never compute" in text
    assert "followups" in text


def test_skill_and_gate_say_chat_in_user_language_sheets_english():
    for p in (SKILL, GATE):
        text = _flat(p)
        assert "user's language" in text, p.name
        assert "sheet text is english" in text, p.name
        assert "also when the user writes norwegian" not in text, p.name
