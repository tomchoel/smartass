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


def test_server_instructions_lead_with_gate_then_answer():
    from smartview_mcp import server
    text = server.INSTRUCTIONS
    assert text.index("sv_gate") < text.index("sv_answer")
    assert not [t for t in GONE if re.search(rf"\b{t}\b", text)]
