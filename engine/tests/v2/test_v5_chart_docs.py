# engine/tests/v2/test_v5_chart_docs.py
"""Plan 5 task 14: every chart example in skills/smartass/reference/charts.md is a valid ChartBlock in the answer
the reference documents, there is one example per kind, and the default month+YTD table in plan.md, SKILL.md and
spec.md is a valid table block. The docs cannot drift from the schema."""
import json
import re
from pathlib import Path
from typing import get_args

from smartass.model.spec import AnalysisSpec, ChartBlock, ChartKind, TableBlock

REPO = Path(__file__).resolve().parents[3]
CHARTS = REPO / "skills" / "smartass" / "reference" / "charts.md"
DEFAULT_TABLE = '{"type":"table","query":"q","periods":["closed","closed-ytd"]}'


def _fences() -> list[dict]:
    text = CHARTS.read_text(encoding="utf-8")
    return [json.loads(body) for body in re.findall(r"```json\n(.*?)```", text, re.S)]


def test_charts_md_has_the_context_answer_then_only_chart_blocks():
    first, *charts = _fences()
    assert first.get("version") == 2
    AnalysisSpec.model_validate(first)
    assert charts and all(c.get("type") == "chart" for c in charts)


def test_every_chart_example_validates_inside_the_context_answer():
    context, *charts = _fences()
    for c in charts:
        spec = AnalysisSpec.model_validate({**context, "blocks": [*context["blocks"], c]})
        assert isinstance(spec.blocks[-1], ChartBlock), c


def test_one_example_per_chart_kind():
    _, *charts = _fences()
    kinds = [c["kind"] for c in charts]
    assert sorted(kinds) == sorted(set(get_args(ChartKind))), kinds


def test_the_default_month_ytd_table_is_named_and_valid():
    TableBlock.model_validate(json.loads(DEFAULT_TABLE))
    for f in (REPO / "knowledge" / "agent" / "plan.md", REPO / "skills" / "smartass" / "SKILL.md",
              REPO / "skills" / "smartass" / "reference" / "spec.md"):
        assert DEFAULT_TABLE in f.read_text(encoding="utf-8"), f.name


def test_the_full_example_in_spec_md_validates():
    text = (REPO / "skills" / "smartass" / "reference" / "spec.md").read_text(encoding="utf-8")
    body = re.search(r"## A full example\s+```json\n(.*?)```", text, re.S).group(1)
    spec = AnalysisSpec.model_validate(json.loads(body))
    assert any(isinstance(b, ChartBlock) for b in spec.blocks)
