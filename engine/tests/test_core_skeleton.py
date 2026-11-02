"""The new core's walking skeleton (docs/plans/2026-11-02-walking-skeleton.md)."""
from types import SimpleNamespace

from smartview_mcp import userexcel
from smartview_mcp.grid import Axis, Layout


def test_fte_rows_are_the_sheet_rows_of_man_year_lines():
    lay = Layout([], [Axis("All Accounts", ["ac.C30", "ac.STA20Y"])], [Axis("All Time Periods", ["Sep 26"])])
    parsed = SimpleNamespace(row_labels=[["C30 REVENUES (ac)"], ["STA20Y Man-years (ac)"]])
    assert userexcel.fte_rows(lay, parsed) == [4]
