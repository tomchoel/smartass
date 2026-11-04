# engine/tests/v2/test_v2_layout.py
from v2util import CLOSED, facts, resolver, spec_dict

from smartass.model.analyze import compute
from smartass.model.spec import AnalysisSpec
from smartass.render.layout import FIRST_COL, layout


def plan(blocks=None):
    d = spec_dict() if blocks is None else spec_dict(blocks=blocks)
    s = AnalysisSpec.model_validate(d)
    res = compute(s, resolver(), CLOSED, facts())
    return s, res, layout(3, s, res, stock=set(), subtitle="NOK · FinMax · closed Sep 26")


def test_title_subtitle_and_sections():
    s, _, p = plan()
    assert p.cells[(1, 2)] == "Explain Business opex vs LY"
    assert p.cells[(2, 2)] == "NOK · FinMax · closed Sep 26"
    sections = [p.cells[(r["row"], 2)] for r in p.rows if r["kind"] == "section"]
    assert sections[:2] == ["Plan", "Key figures"]
    assert "EBITDA bridge" in sections and "Commentary" in sections


def test_every_model_cell_is_a_formula_on_the_data_table():
    _, res, p = plan()
    for o in res.outs:
        for ln in o.lines:
            for j, e in enumerate(ln.cells):
                key = f"{o.key}.{ln.key}.{j}"
                if e is None:
                    assert key not in p.addr
                    continue
                r, c = p.at(key)
                assert c == FIRST_COL + j
                assert str(p.cells[(r, c)]).startswith("=")
    first = next(v for v in p.cells.values() if isinstance(v, str) and "SUMIFS" in v)
    assert "Data_Q3[Value]" in first


def test_refs_point_at_their_rows():
    _, res, p = plan([{"type": "bridge", "query": "fm", "total": "EBITDA", "steps": ["Revenue", "Opex"]}])
    other = p.at("b0.other.0")
    end, start = p.at("b0.end.0"), p.at("b0.start.0")
    s0, s1 = p.at("b0.s0.0"), p.at("b0.s1.0")
    col = "C"
    assert p.cells[other] == f"=({col}{end[0]}-{col}{start[0]})-SUM({col}{s0[0]}:{col}{s1[0]})"


def test_number_formats_by_column_code():
    _, res, p = plan([{"type": "table", "query": "fm", "columns": ["value", "ly", "var", "var_pct"]}])
    line = next(r for r in p.rows if r["kind"] == "line")
    assert line["num"] == {"C": "heltall", "D": "heltall", "E": "heltall", "F": "prosent1"}
    header = next(r for r in p.rows if r["kind"] == "header")
    assert header["num_cols"] == ["C", "D", "E", "F"]
    assert p.last_col == "F"


def test_check_lines_and_findings_rows():
    _, res, p = plan([{"type": "ranking", "query": "fm", "metric": "Opex", "top": 1}])
    assert any(r["kind"] == "check" for r in p.rows)
    assert p.last_row == max(r["row"] for r in p.rows)
