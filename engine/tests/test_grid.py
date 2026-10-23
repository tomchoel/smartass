from smartview_mcp.grid import Axis, Layout, full_pov, name_from_label, parse, validate_axes
from smartview_mcp.ops import tree_from_labels

DIMS = ["All Time Periods", "All Balance Types", "All Accounts", "All Entities", "All Scenarios", "All Currencies"]


def test_name_from_label():
    assert name_from_label("3WHS_W Wholesale (un.SY)") == "un.SY.3WHS_W"
    assert name_from_label("                C30 REVENUES (ac)") == "ac.C30"
    assert name_from_label("SY Telenor Norge AS (un)") == "un.SY"
    assert name_from_label("Sep 26") is None
    assert name_from_label("All Entities") is None


def test_layout_matrix_positions():
    rows = [Axis("All Entities", ["un.SY.2WHS"]), Axis("All Accounts", ["ac.C30", "ac.C40S"])]
    cols = [Axis("All Time Periods", ["Aug 26", "Sep 26"])]
    pov = full_pov(DIMS, {"All Scenarios": "Actual"}, rows, cols)
    lay = Layout(pov, rows, cols)
    m = lay.matrix()
    assert m[0][2:2 + len(pov)] == [mm for _, mm in pov]          # POV starts at column nr+1
    assert m[1][2:4] == ["Aug 26", "Sep 26"]
    assert m[2][:2] == ["un.SY.2WHS", "ac.C30"] and m[3][:2] == ["un.SY.2WHS", "ac.C40S"]
    assert ("All Scenarios", "Actual") in pov and ("All Balance Types", "All Balance Types") in pov


def test_validate():
    rows = [Axis("All Accounts", ["ac.C30"])]
    cols = [Axis("All Accounts", ["x"])]
    probs = validate_axes(DIMS, {"Nope": "x"}, rows, cols)
    assert any("unknown dimension 'Nope'" in p for p in probs)
    assert any("more than once" in p for p in probs)


def test_parse_real_spike_output():
    # captured from spike_query.py case A (2026-10-13), POV truncated
    rows = [Axis("All Entities", ["un.SY.2WHS", "un.SY.3WHS_W"]), Axis("All Accounts", ["ac.C30", "ac.C40S", "ac.NOPE99"])]
    cols = [Axis("All Time Periods", ["Aug 26", "Sep 26"])]
    lay = Layout([("All Scenarios", "Actual"), ("All Currencies", "NOK")], rows, cols)
    vals = (
        (None, None, "Actual", "NOK"),
        (None, None, "Aug 26", "Sep 26"),
        ("2WHS Wholesale (un.SY)", "C30 REVENUES (ac)", 1767304791.41, 1994151357.55),
        ("2WHS Wholesale (un.SY)", "C40S Total Cost (ac)", -44305197.4, -52330751.43),
        (None, "ac.NOPE99", None, None),
        ("3WHS_W Wholesale (un.SY)", "C30 REVENUES (ac)", 1767304791.41, 1994151357.55),
        ("3WHS_W Wholesale (un.SY)", "C40S Total Cost (ac)", -44305197.4, -52330751.43),
        (None, "ac.NOPE99", None, None),
    )
    p = parse(lay, vals, suppress_missing=True, suppress_zeros=False)
    assert p.unknown_members == ["All Accounts: ac.NOPE99"]
    assert len(p.rows) == 4
    assert p.rows[0] == ["un.SY.2WHS", "ac.C30", 1767304791.41, 1994151357.55]
    assert p.row_labels[0] == ["2WHS Wholesale (un.SY)", "C30 REVENUES (ac)"]
    assert p.columns == ["All Entities", "All Accounts", "Aug 26", "Sep 26"]
    assert not p.pov_problems


def test_parse_pov_problem_and_missing():
    rows = [Axis("All Accounts", ["ac.C30"])]
    cols = [Axis("All Time Periods", ["Sep 26"])]
    lay = Layout([("All Entities", "un.SY.BAD")], rows, cols)
    vals = ((None, "un.SY.BAD", "All Entities"), (None, "Sep 26", None), ("C30 REVENUES (ac)", "#Missing", None))
    p = parse(lay, vals, suppress_missing=False, suppress_zeros=False)
    assert p.pov_problems and "All Entities" in p.pov_problems[0]
    assert p.rows == [["ac.C30", None]]


def test_grid_labels_standard_and_odd_mode():
    from smartview_mcp.ops import _grid_labels
    fm = ((None, "All Accounts", "All Entities"), (None, "All Time Periods", None), ("All Balance Types", "#No Access", None))
    assert _grid_labels(fm, [1, 0, 2, 2]) == ["All Time Periods", "All Balance Types", "All Accounts", "All Entities"]
    fp = ((None, None, "Scenario", "Version"), (None, None, "Years", None), (None, None, "Period", None),
          ("Account", "Entity", 0.0, None))
    assert _grid_labels(fp, [1, 1, 0, 0, 2, 2]) == ["Years", "Period", "Account", "Entity", "Scenario", "Version"]
    # 2026-10-13 transient 'Ad hoc' mode: no in-sheet POV row -> must be rejected
    assert _grid_labels(((None, "TimePeriods"), ("BalanceTypes", "#No Access")), [2, 2, 0, 1]) is None


def test_markdown_two_column_dims_is_well_formed():
    from smartview_mcp.grid import to_markdown
    rows = [Axis("Entity", ["BANORWAY"])]
    cols = [Axis("Years", ["FY25", "FY26"]), Axis("Period", ["TotalYear"])]
    lay = Layout([("Scenario", "Target 2025")], rows, cols)
    vals = ((None, "Target 2025"), (None, "FY25", "FY26"), (None, "TotalYear", "TotalYear"), ("BANORWAY", 1.0, 2.0))
    p = parse(lay, vals, False, False)
    md = to_markdown(p.columns, p.rows).splitlines()
    assert p.columns == ["Entity", "FY25 / TotalYear", "FY26 / TotalYear"]
    assert md[0].count("|") == md[1].count("|") == md[2].count("|")


def test_tree_bottom_up():
    # ancestors-at-bottom ordering as Smart View returns it with the user's 'ancestor position' setting
    labels = ["          3WHS_W Wholesale (un.SY)", "     2WHS Wholesale (un.SY)", "     2BUS Business (un.SY)",
              "SY Telenor Norge AS (un)"]
    t = tree_from_labels(labels)
    by = {x["name"]: x for x in t}
    assert by["un.SY"]["depth"] == 0 and by["un.SY"]["parent"] is None
    assert by["un.SY.2WHS"]["parent"] == "un.SY"
    assert by["un.SY.3WHS_W"]["parent"] == "un.SY.2WHS"
    assert by["un.SY.2BUS"]["parent"] == "un.SY"


def test_parse_names_rows_suppressed_as_empty():
    rows = [Axis("All Entities", ["un.SY.2CON", "un.SY.2SSM"]), Axis("All Accounts", ["ac.C693"])]
    cols = [Axis("All Time Periods", ["Q3 26"])]
    vals = ((None, None, "Actual"), (None, None, "Q3 26"),
            ("2CON Consumer (un.SY)", "C693 Consultancy (ac)", "#Missing"),
            ("2SSM SSM (un.SY)", "C693 Consultancy (ac)", -25695850.0))
    lay = Layout([("All Scenarios", "Actual")], rows, cols)
    p = parse(lay, vals, suppress_missing=True, suppress_zeros=False)
    assert len(p.rows) == 1 and p.suppressed_rows == ["un.SY.2CON / ac.C693"]
