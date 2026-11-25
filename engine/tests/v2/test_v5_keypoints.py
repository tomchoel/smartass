# engine/tests/v2/test_v5_keypoints.py
"""Plan 5 Task 6: the Key points - live number sentences on the main table's last column group, the cause row from
the commentary drivers, the agent's free points capped at 5 in total. Every figure is invented."""
import re
from types import SimpleNamespace

import pytest
from test_v5_layout import GROUPED, run
from v2util import spec_dict

from smartass.model.blocks.common import BlockOut, Line
from smartass.model.expr import Const
from smartass.model.spec import AnalysisSpec, Driver
from smartass.render.keypoints import key_points
from smartass.render.layout import HUNDRED, TEXT_COL, UNIT_MNOK
from smartass.xl.vendor import formelsjekk

COLS = ["value", "ly", "var", "var_pct", "gap", "value", "ly", "var", "var_pct", "material"]
HEADERS = ["AC Oct 26", "PY Oct 25", "ΔPY", "ΔPY %", "", "AC YTD Oct 26", "PY YTD Oct 25", "ΔPY", "ΔPY %",
           "Material"]


def _line(key, label, kind="line", metric=None):
    cells = [Const(0.0) if c != "gap" else None for c in COLS]
    return Line(key, label, cells, kind=kind, metric=metric or label)


def _main(lines=None):
    lines = lines or [_line("m0", "Revenue"), _line("m1", 'Opex "other"', metric="Opex"),
                      _line("m2", "EBITDA", kind="grand")]
    return BlockOut("b0", "table", "P&L", HEADERS, cols=list(COLS), lines=lines)


def _res(material):
    return SimpleNamespace(material=material)


def _mat(line, v, base="ly"):
    return {"block": "P&L", "line": line, "key": f"b0.{line}", "variance": v, "variance_pct": None, "base": base}


ROWS = {"m0": 10, "m1": 11, "m2": 12}


def addr(key: str) -> str:
    if key == UNIT_MNOK:
        return "$D$40"
    if key == HUNDRED:
        return "$D$42"
    _, line, col = key.split(".")
    return "DEFGHIJKLM"[int(col)] + str(ROWS[line])


def test_two_material_lines_in_abs_variance_order_exact_formulas():
    pts = key_points(_main(), _res([_mat("m0", 5.0e6), _mat("m1", -9.0e6)]), [], addr)
    assert [f for f, _ in pts] == [
        '="•  Opex ""other"" YTD Oct 26 "&SUBSTITUTE(FIXED(I11/$D$40,1),"-","−")&" MNOK; "&SUBSTITUTE(FIXED(K11/$D$40,1),"-","−")&" MNOK vs PY ("&'
        'IF(L11="","n/a",SUBSTITUTE(FIXED(L11*$D$42,1),"-","−")&" %")&")"',
        '="•  Revenue YTD Oct 26 "&SUBSTITUTE(FIXED(I10/$D$40,1),"-","−")&" MNOK; "&SUBSTITUTE(FIXED(K10/$D$40,1),"-","−")&" MNOK vs PY ("&'
        'IF(L10="","n/a",SUBSTITUTE(FIXED(L10*$D$42,1),"-","−")&" %")&")"',
    ]
    assert [c for _, c in pts] == [None, None]          # no driver: no placeholder cause row


def test_a_line_material_against_two_bases_is_one_point_and_at_most_four():
    lines = [_line(f"m{i}", f"L{i}") for i in range(6)]
    mat = [_mat(f"m{i}", float(i + 1)) for i in range(6)] + [_mat("m5", 99.0, base="budget")]
    rows = {f"m{i}": 10 + i for i in range(6)}
    pts = key_points(_main(lines), _res(mat), [], lambda k: k if "." not in k else "A" + str(rows[k.split(".")[1]]))
    assert [f.split(" YTD")[0] for f, _ in pts] == ['="•  L5', '="•  L4', '="•  L3', '="•  L2']


def test_causes_data_hypothesis_and_metric_match():
    drivers = [Driver(line="revenue", text="Price rise", basis="data", evidence="b2.r1.0"),
               Driver(line="Opex", text="Consultants", basis="hypothesis")]
    pts = key_points(_main(), _res([_mat("m0", 5.0e6), _mat("m1", -9.0e6)]), drivers, addr)
    assert [c for _, c in pts] == ["   Consultants (hypothesis — not verified in data)",
                                   "   Price rise (data: b2.r1.0)"]


def test_no_material_lines_gives_the_grand_line_without_cause():
    pts = key_points(_main(), _res([]), [], addr)
    assert len(pts) == 1 and pts[0][0].startswith('="•  EBITDA YTD Oct 26 "&SUBSTITUTE(FIXED(I12/$D$40,1),"-","−")') and pts[0][1] is None
    drv = [Driver(line="EBITDA", text="Lower opex", basis="hypothesis")]
    assert key_points(_main(), _res([]), drv, addr)[0][1] == "   Lower opex (hypothesis — not verified in data)"


def test_no_grand_line_gives_the_first_line():
    pts = key_points(_main([_line("m0", "Revenue"), _line("m1", "Opex")]), _res([]), [], addr)
    assert pts[0][0].startswith('="•  Revenue ')


def test_short_form_without_var_column():
    o = BlockOut("b0", "table", "P&L", ["AC YTD Oct 26"], cols=["value"],
                 lines=[Line("m0", "Revenue", [Const(0.0)], kind="grand", metric="Revenue")])
    pts = key_points(o, _res([]), [], lambda k: "$D$40" if k == UNIT_MNOK else "D10")
    assert pts == [('="•  Revenue YTD Oct 26 "&SUBSTITUTE(FIXED(D10/$D$40,1),"-","−")&" MNOK"', None)]


def test_stock_metric_has_no_mnok():
    o = BlockOut("b0", "table", "FTE", ["AC YTD Oct 26"], cols=["value"],
                 lines=[Line("m0", "FTE", [Const(0.0)], kind="grand", metric="FTE")])
    pts = key_points(o, _res([]), [], lambda k: "D10", stock={"FTE"})
    assert pts == [('="•  FTE YTD Oct 26 "&SUBSTITUTE(FIXED(D10,1),"-","−")&" FTE"', None)]


def test_stock_metric_with_variance_says_fte_twice_and_no_unit_division():
    o = _main([_line("m0", "FTE", kind="grand")])
    f = key_points(o, _res([]), [], addr, stock={"FTE"})[0][0]
    assert f == ('="•  FTE YTD Oct 26 "&SUBSTITUTE(FIXED(I10,1),"-","−")&" FTE; "&SUBSTITUTE(FIXED(K10,1),"-","−")&" FTE vs PY ("&'
                 'IF(L10="","n/a",SUBSTITUTE(FIXED(L10*$D$42,1),"-","−")&" %")&")"')
    assert "MNOK" not in f and "$D$40" not in f


TWO = ["value", "budget", "var", "forecast", "var", "var_pct"]       # as col_headers builds two bases


def _two(lines):
    heads = ["AC YTD Oct 26", "BU YTD Oct 26", "ΔBU", "FC YTD Oct 26", "ΔFC", "ΔFC %"]
    return BlockOut("b0", "table", "P&L", heads, cols=list(TWO), lines=lines)


def _tline(key, label, kind="line"):
    return Line(key, label, [Const(0.0)] * len(TWO), kind=kind, metric=label)


def test_the_sentence_uses_the_base_the_line_is_material_against():
    o = _two([_tline("m0", "Revenue"), _tline("m1", "Opex")])
    pts = key_points(o, _res([_mat("m0", 5.0, base="forecast"), _mat("m1", 3.0, base="budget"),
                              _mat("m1", 1.0, base="forecast")]), [], addr)
    rev, opex = (f for f, _ in pts)
    assert '&SUBSTITUTE(FIXED(H10/$D$40,1),"-","−")&" MNOK vs FC ("&IF(I10="","n/a"' in rev
    assert opex.endswith('&SUBSTITUTE(FIXED(F11/$D$40,1),"-","−")&" MNOK vs BU; "&SUBSTITUTE(FIXED(H11/$D$40,1),"-","−")&" MNOK vs FC"')          # no ΔBU % column: no bracket


def test_a_material_base_without_columns_falls_back_to_the_first_base():
    pts = key_points(_main(), _res([_mat("m0", 5.0, base="budget")]), [], addr)
    assert " MNOK vs PY (" in pts[0][0]


def test_fallback_point_uses_the_first_base():
    o = _two([_tline("m0", "EBITDA", kind="grand")])
    f = key_points(o, _res([]), [], addr)[0][0]
    assert f.endswith('&SUBSTITUTE(FIXED(F10/$D$40,1),"-","−")&" MNOK vs BU; "&SUBSTITUTE(FIXED(H10/$D$40,1),"-","−")&" MNOK vs FC"')


def test_no_main_table_gives_only_the_free_points():
    assert key_points(None, _res([]), [], addr, free=["a"]) == [("•  a", None)]


LONG = ("The consultants were moved from the project into running operations in the second quarter, so the "
        "cost now shows on opex instead of capex")


def test_long_causes_and_free_points_wrap_into_indented_rows():
    from smartass.render.keypoints import point_rows
    from smartass.render.layout import NOTE_WIDTH
    drivers = [Driver(line="Opex", text=LONG, basis="hypothesis")]
    pts = key_points(_main(), _res([_mat("m1", 9.0)]), drivers, addr, free=["=" + LONG])
    rows = point_rows(pts)
    assert rows[0] == pts[0][0]                                    # the formula is never wrapped
    cause = [r for r in rows[1:] if not r.startswith("•")]
    assert len(rows) > 3 and all(len(r) <= NOTE_WIDTH for r in rows[1:])
    assert all(r.startswith("   ") for r in cause)                 # cause and continuation rows stay indented
    assert not any(r.lstrip().startswith("=") for r in rows[1:] if r.startswith("   "))


def test_layout_reserves_the_wrapped_rows_and_the_lint_stays_clean():
    from functools import partial

    from fakexl import Workbook

    from smartass.render.excel import answer_grid, answer_roles
    from smartass.xl import styling
    kp = partial(key_points, free=[LONG])
    d = spec_dict(blocks=[GROUPED, {"type": "commentary", "drivers": [
        {"line": "EBITDA", "text": LONG, "basis": "hypothesis"}]}])
    d["queries"][0]["periods"] = "closed-ytd"
    from test_v5_headers import OCT, ytd_frame
    from v2util import resolver

    from smartass.model.analyze import compute
    from smartass.render.layout import layout
    s = AnalysisSpec.model_validate(d)
    res = compute(s, resolver(), OCT, ytd_frame())
    p = layout(3, s, res, stock=set(), subtitle="SUB", key_points=kp)
    z = dict(p.zones)
    rows = [p.cells.get((r, TEXT_COL)) for r in range(z["key_points"] + 1, z["charts"])]
    assert rows[-1] is None and all(rows[:-1])                     # every reserved row used, one blank row after
    assert len(rows) - 1 > 3
    wb = Workbook(["Q3 Answer"])
    ws = wb.order[0]
    r0, c0, g = answer_grid(p)
    for i, row in enumerate(g):
        for j, v in enumerate(row):
            ws.put(r0 + i, c0 + j, v)
    assert styling.apply_report(ws, answer_roles(p)) == []


def test_main_table_for_key_points_skips_a_bridge_first():
    from test_v5_layout import BRIDGE
    s, res, p = run([BRIDGE, GROUPED], key_points=key_points, scale=1e6)
    f = p.cells[(dict(p.zones)["key_points"] + 1, TEXT_COL)]
    assert f.startswith('="•  ') and " YTD Oct 26 " in f
    assert "MNOK vs PY" in f                                       # the P&L table, not the bridge's amount


def test_only_a_bridge_gives_only_free_points():
    from functools import partial

    from test_v5_layout import BRIDGE
    s, res, p = run([BRIDGE], key_points=partial(key_points, free=["Agent point"]))
    r = dict(p.zones)["key_points"]
    assert p.cells[(r + 1, TEXT_COL)] == "•  Agent point" and (r + 2, TEXT_COL) not in p.cells


def test_free_key_points_follow_and_cap_the_total_at_seven():
    free = ["one", "two", "three", "=four", "five"]
    warn: list[str] = []
    pts = key_points(_main(), _res([_mat("m0", 5.0), _mat("m1", 9.0), _mat("m2", 3.0)]), [], addr,
                     free=free, warnings=warn)
    assert len(pts) == 7
    assert pts[:2] == [key_points(_main(), _res([_mat("m0", 5.0), _mat("m1", 9.0)]), [], addr)[0],
                       key_points(_main(), _res([_mat("m0", 5.0), _mat("m1", 9.0)]), [], addr)[1]]
    assert pts[2:] == [("•  one", None), ("•  two", None), ("•  three", None), ("•  =four", None), ("•  five", None)]
    assert len(warn) == 1 and "truncated" in warn[0]
    pts = key_points(_main(), _res([]), [], addr, free=free[:4])
    assert pts[1:] == [("•  one", None), ("•  two", None), ("•  three", None), ("•  =four", None)]


def test_count_does_not_depend_on_addr():
    res = _res([_mat("m0", 5.0), _mat("m1", 9.0)])
    assert len(key_points(_main(), res, [], lambda k: "A1", free=["x"])) == \
        len(key_points(_main(), res, [], addr, free=["x"]))


def test_more_than_seven_key_points_refused_only_on_a_new_question():
    from smartass.model.validate import GuardError, guard_key_points
    from smartass.render.keypoints import free_points
    eight = ["a", "b", "c", "d", "e", "f", "g", "h"]
    d = spec_dict(blocks=[GROUPED, {"type": "commentary", "key_points": eight}])
    s = AnalysisSpec.model_validate(d)                 # a stored spec still loads
    with pytest.raises(GuardError):
        guard_key_points(s)
    pts, warnings = free_points(s)
    assert pts == eight[:7] and len(warnings) == 1 and "8" in warnings[0]
    d["blocks"][1]["key_points"] = eight[:7]
    s = AnalysisSpec.model_validate(d)
    guard_key_points(s)
    assert free_points(s) == (eight[:7], [])


def _balanced(f: str) -> bool:
    """Minimal Excel syntax: quotes pair up ("" is an escaped quote) and parentheses balance outside strings."""
    depth, i, inside = 0, 0, False
    while i < len(f):
        ch = f[i]
        if ch == '"':
            if inside and f[i + 1:i + 2] == '"':
                i += 2
                continue
            inside = not inside
        elif not inside and ch in "()":
            depth += 1 if ch == "(" else -1
            if depth < 0:
                return False
        i += 1
    return depth == 0 and not inside


def test_layout_end_to_end_formulas_are_valid_and_pass_formelsjekk():
    from functools import partial
    d_kp = partial(key_points, free=["Agent point"])
    s, res, p = run([GROUPED], key_points=d_kp, scale=1e6)
    assert HUNDRED in p.params and UNIT_MNOK in p.params
    r = dict(p.zones)["key_points"]
    f = p.cells[(r + 1, TEXT_COL)]
    assert f.startswith('="•  ') and ' YTD Oct 26 "&SUBSTITUTE(FIXED(' in f and p.params[UNIT_MNOK] in f and p.params[HUNDRED] in f
    shown = [p.cells.get((r + i, TEXT_COL)) for i in range(1, 12)]
    assert "•  Agent point" in shown                       # after the engine's points (MNOK-sized: material)
    for formula in (f, *(x for x, _ in key_points(_main(), _res([_mat("m1", 1.0)]), [], addr))):
        assert _balanced(formula), formula
        assert all(len(lit) <= 255 for lit in re.findall(r'"((?:[^"]|"")*)"', formula))
        assert [x for x in formelsjekk.check_formula(formula) if x.severity == "block"] == []


def test_key_points_pools_from_up_to_two_tables_and_respects_agent_priority():
    t1 = _main()
    t2 = BlockOut("b1", "table", "Cost breakdown", HEADERS, cols=list(COLS),
                  lines=[_line("m0", "Consultants"), _line("m1", "IT")])
    t3 = BlockOut("b2", "table", "Ignored third table", HEADERS, cols=list(COLS),
                  lines=[_line("m0", "Other")])
    mats = [
        {"block": "P&L", "line": "m0", "key": "b0.m0", "variance": 10.0, "variance_pct": None, "base": "ly"},
        {"block": "Cost breakdown", "line": "m1", "key": "b1.m1", "variance": 8.0, "variance_pct": None, "base": "ly"},
        {"block": "P&L", "line": "m1", "key": "b0.m1", "variance": 6.0, "variance_pct": None, "base": "ly"},
        {"block": "Cost breakdown", "line": "m0", "key": "b1.m0", "variance": 4.0, "variance_pct": None, "base": "ly"},
        {"block": "Ignored third table", "line": "m0", "key": "b2.m0", "variance": 99.0, "variance_pct": None, "base": "ly"},
    ]
    a2 = lambda k: "Z10" if k.startswith("b1.") else addr(k)        # noqa: E731
    warn: list[str] = []
    res = SimpleNamespace(material=mats, outs=[t1, t2, t3])
    pts = key_points(t1, res, [], a2, free=["f1", "f2", "f3", "f4"], warnings=warn)
    assert len(pts) == 7
    assert pts[3:] == [("•  f1", None), ("•  f2", None), ("•  f3", None), ("•  f4", None)]
    assert len(warn) == 1 and "truncated" in warn[0]
    engine_text = " ".join(f for f, _ in pts[:3])
    assert "Consultants" not in engine_text and "Other" not in engine_text
    assert "Revenue" in engine_text and "IT" in engine_text and 'Opex ""other""' in engine_text


def test_mixed_sign_lines_key_points_ignore_children_and_dont_treat_debit_as_pct():
    lines_blk = BlockOut(
        "b0", "lines", "Ledger lines by supplier",
        headers=["Amount", "Debit", "Credit", "Lines"],
        cols=["amount", "debit", "credit", "lines"],
        lines=[
            Line("g0", "Accenture AS", [Const(-90.0), Const(-120.0), Const(30.0), Const(2.0)], code="43200"),
            Line("g1", "Bouvet", [Const(-40.0), Const(-40.0), Const(0.0), Const(1.0)], code="43100"),
            Line("blank", "Lines without a supplier", [Const(-50.0), Const(-50.0), Const(0.0), Const(3.0)]),
            Line("blank_0", "HG 12", [Const(-40.0), Const(-40.0), Const(0.0), Const(2.0)], level=1, code="43100"),
            Line("total", "Ledger lines total", [Const(-180.0), Const(-210.0), Const(30.0), Const(6.0)], kind="sum"),
        ],
    )
    a = lambda k: {"b0.g0.0": "D10", "b0.g1.0": "D11", "b0.blank.0": "D12", "b0.blank_0.0": "D13",
                   "b0.total.0": "D14"}.get(k, "D15")   # noqa: E731
    pts = key_points(lines_blk, SimpleNamespace(material=[], outs=[lines_blk]), [], a)
    assert len(pts) == 3
    joined = " ".join(f for f, _ in pts)
    assert "HG 12" not in joined and "Lines without a supplier" not in joined
    assert "Accenture AS" in joined and "Bouvet" in joined and "Ledger lines total" in joined
    assert " %" not in joined and "*$D$" not in joined
