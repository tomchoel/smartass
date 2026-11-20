"""Numeric number formats reach Excel in the user's locale (live 05.11.2026: nb-NO refused '#,##0.0' set through
late-bound COM). Every engine NumberFormat write of a numeric code goes through styling.local_format."""
import pathlib

from smartass.xl import styling

PKG = pathlib.Path(styling.__file__).resolve().parents[1]


class _App:
    International = tuple([None, None, ",", " "] + [None] * 40)


class _Obj:
    Application = _App()


def test_local_format_swaps_separators_outside_quotes():
    assert styling.local_format(_Obj(), '#,##0.0;(#,##0.0);"–"') == '# ##0,0;(# ##0,0);"–"'
    assert styling.local_format(_Obj(), '[Color10]"▲ "0.0 %') == '[Color10]"▲ "0,0 %'
    assert styling.local_format(_Obj(), "#,##0.0,,") == "# ##0,0  "


def test_local_format_without_an_app_returns_the_code():
    assert styling.local_format(object(), "#,##0") == "#,##0"


def test_no_raw_numeric_number_format_writes():
    """Every NumberFormat write (attribute or setattr) of a numeric code goes through local_format; text '@' and
    'General' are locale-free."""
    ok = ("styling.local_format(", "local_format(", '"@"', "tn.to_local(")   # "General" is "Standard" under nb-NO
    bad = []
    for f in PKG.rglob("*.py"):
        if "vendor" in f.parts or "workbench" in f.parts:
            continue
        for i, line in enumerate(f.read_text(encoding="utf-8").splitlines(), 1):
            s = line.strip()
            if s.startswith("#"):
                continue
            if ".NumberFormat =" in s:
                rhs = s.split(".NumberFormat =", 1)[1].strip()
            elif '"NumberFormat",' in s:
                rhs = s.split('"NumberFormat",', 1)[1].strip()
            else:
                continue
            if not rhs.startswith(ok):
                bad.append(f"{f.name}:{i}: {s}")
    assert not bad, bad


def test_a_block_metric_missing_from_its_query_is_added_to_the_query():
    """Live 05.11.2026: a bridge over Revenue/COGS/Opex/Other items next to a chained EBITDA table was refused
    because the query only had EBITDA. The query now fetches every metric its blocks use."""
    from smartass.model.spec import AnalysisSpec
    spec = AnalysisSpec.model_validate({
        "version": 2, "question": "q", "plan": {"depth": "quick"},
        "queries": [{"id": "q", "cube": "FinMax", "metrics": ["EBITDA"], "entity": "un.SY.2BUS",
                     "periods": "closed-ytd", "compare": ["ly"]}],
        "blocks": [{"type": "bridge", "query": "q", "total": "EBITDA", "steps": ["Revenue", "COGS"]}]})
    assert spec.queries[0].metrics == ["EBITDA", "Revenue", "COGS"]


def test_series_with_gaps_become_a_union_reference():
    """Live 05.11.2026: Excel refuses a multi-area Range as Series.Values; a union reference text works."""
    from smartass.render.charts import union_ref
    assert union_ref("Q2 Answer", "D5:D7,D9") == "=('Q2 Answer'!$D$5:$D$7,'Q2 Answer'!$D$9)"
    assert union_ref("It's", "C1,C3") == "=('It''s'!$C$1,'It''s'!$C$3)"


def test_finplan_aliases_get_a_code_only_when_they_carry_one():
    from smartass.model.blocks.common import code_of
    assert code_of("2WHS Wholesale") == "2WHS"
    assert code_of("EBITDA before Other Items") == ""
    assert code_of("ac.C30") == "C30" and code_of("Gross_Profit") == "Gross_Profit"


def test_a_text_cell_starting_with_equals_is_written_as_text():
    """Live 05.11.2026: a wrapped caveat continued on a line starting '= 78 million …'; Excel refused it as a formula
    (E_OUTOFMEMORY) and the answer failed."""
    from smartass.render.excel import answer_grid
    from smartass.render.layout import SheetPlan
    p = SheetPlan()
    p.cells[(1, 2)] = "= 78 million against 8.7 million in stock)."
    p.cells[(2, 2)] = '="•  Revenue "&FIXED(D5,1)'
    p.cells[(3, 4)] = "=SUM(D1:D2)"
    p.literal.add((1, 2))
    _, _, g = answer_grid(p)
    assert g[0][0] == "'= 78 million against 8.7 million in stock)."
    assert g[1][0].startswith('="•') and g[2][2] == "=SUM(D1:D2)"


def test_a_waterfall_may_point_at_a_concept_block():
    """Live 05.11.2026: an explicit waterfall on a concept was refused although the implicit concept chart is one."""
    from smartass.model.spec import AnalysisSpec
    s = AnalysisSpec.model_validate({
        "version": 2, "question": "q", "plan": {"depth": "quick"},
        "queries": [{"id": "sx", "cube": "FinMax", "metrics": [], "entity": "un.SY", "periods": "closed-ytd"}],
        "blocks": [{"type": "concept", "query": "sx", "concept": "streamix"},
                   {"type": "chart", "kind": "waterfall", "title": "t", "source": {"block": 0}}]})
    assert s.blocks[1].kind == "waterfall"
