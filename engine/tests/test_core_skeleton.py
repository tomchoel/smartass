"""The new core's walking skeleton (docs/plans/2026-11-02-walking-skeleton.md)."""
import datetime as dt
from types import SimpleNamespace

import pytest

from smartview_mcp import svhttp, userexcel
from smartview_mcp.core import contract, source
from smartview_mcp.grid import Axis, Layout
from smartview_mcp.host import InputError


def test_fte_rows_are_the_sheet_rows_of_man_year_lines():
    lay = Layout([], [Axis("All Accounts", ["ac.C30", "ac.STA20Y"])], [Axis("All Time Periods", ["Sep 26"])])
    parsed = SimpleNamespace(row_labels=[["C30 REVENUES (ac)"], ["STA20Y Man-years (ac)"]])
    assert userexcel.fte_rows(lay, parsed) == [4]

PIN = contract.Pull.model_validate({"id": "m", "connection": "FinMax",
                                    "rows": [{"dimension": "All Accounts", "members": ["ac.EBITDA"]}],
                                    "columns": [{"dimension": "All Time Periods", "members": ["Sep 26"]}]})


def test_cells_keep_text_as_text_and_blanks_empty():
    assert source.cells([[None, "Jan 26"], ["EBITDA (ac)", 1.5, "x"]]) == (("", "'Jan 26", ""), ("'EBITDA (ac)", 1.5, "'x"))


def test_provenance_names_cube_pov_time_and_route():
    lay = Layout([("All Entities", "un.SY.2BUS"), ("All Scenarios", "All Scenarios")], [], [])
    p = dict(source.provenance("FinMax", lay, "http", dt.datetime(2026, 11, 2, 10, 5)))
    assert p["Source"] == "Oracle EPM Cloud · FinMax" and p["POV"] == "un.SY.2BUS"
    assert p["Fetched"] == "2026-11-02 10:05" and "re-fetch" in p["Route"]


def test_mismatches_compare_only_the_fetched_numbers():
    assert source.mismatches([[None, "a"], ["x", 1.0]], ((None, "a"), ("x", 1.0, "panel"))) == []
    assert source.mismatches([["x", 1.0]], (("x", 2.0),)) == ["row 1 column 2: fetched 1.0, sheet shows 2.0"]


def _adhoc(monkeypatch, seen):
    pull = SimpleNamespace(sheet="Q1 FinMax")
    monkeypatch.setattr(source.userexcel, "pull", lambda h, t, q, conn, pov, rows, cols, title: seen.append(title) or pull)
    return pull


def test_an_http_failure_falls_back_to_the_adhoc_grid_with_a_note(monkeypatch):
    def fail(*a):
        raise svhttp.SvHttpError("token expired")
    monkeypatch.setattr(source.svhttp, "shared", lambda: SimpleNamespace(fetch=fail))
    seen = []
    pull = _adhoc(monkeypatch, seen)
    s = source.pull(None, None, 1, 0, PIN)
    assert s.pull is pull and s.route == "adhoc" and "token expired" in s.note and seen == ["Q1 FinMax"]


def test_an_input_error_is_not_hidden_by_the_fallback(monkeypatch):
    def bad(*a):
        raise InputError("Invalid query: unknown dimension")
    monkeypatch.setattr(source.svhttp, "shared", lambda: SimpleNamespace(fetch=bad))
    with pytest.raises(InputError):
        source.pull(None, None, 1, 0, PIN)


def test_http_success_writes_the_sheet_and_names_later_pulls_by_id(monkeypatch):
    f = svhttp.Fetched(Layout([], [], []), [], 0.1)
    monkeypatch.setattr(source.svhttp, "shared", lambda: SimpleNamespace(fetch=lambda *a: f))
    monkeypatch.setattr(source, "write", lambda t, q, title, pin, fetched: (title, fetched))
    assert source.pull(None, None, 2, 1, PIN) == ("Q2 FinMax m", f)
