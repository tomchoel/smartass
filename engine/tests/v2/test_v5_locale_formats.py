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
    ok = ("styling.local_format(", "local_format(", "chart_format(", '"@"', "tn.to_local(")   # "General" is "Standard" under nb-NO
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


def test_chart_format_scales_with_no_break_spaces_and_keeps_the_sections():
    """Live 05.11.2026 (nb-NO): DataLabels.NumberFormat needs U+00A0 as the thousands/scaling character
    ('#\xa0##0,0\xa0\xa0' shows 3 811,4 for 3 811 374 124.3); plain spaces showed '3811374 124,3'."""
    from smartass.render.charts import label_format
    nb = "#\xa0##0,0\xa0\xa0"
    assert styling.chart_format(_Obj(), label_format("ac", True)) == f"{nb};-{nb};;"
    assert styling.chart_format(_Obj(), label_format("down", True)) == f"-{nb};-{nb};;"
    assert styling.chart_format(_Obj(), label_format("ac", False)) == "#\xa0##0,0;-#\xa0##0,0;;"
    assert styling.chart_format(_Obj(), label_format("pct", True)) == "0,0 %;-0,0 %;;"   # the space stays plain


def test_chart_format_keeps_en_us_and_fakes():
    class _Us:
        class Application:
            International = tuple([None, None, ".", ","] + [None] * 40)
    assert styling.chart_format(_Us(), "#,##0.0,,") == "#,##0.0,,"
    assert styling.chart_format(object(), "#,##0.0,,") == "#,##0.0,,"
