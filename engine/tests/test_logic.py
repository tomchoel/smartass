"""Unit tests for the analysis logic that needs no Excel."""
from types import SimpleNamespace

from smartview_mcp.grid import Axis, Layout
from smartview_mcp.server import fin
from smartview_mcp.templates import _clean, _fmt, _pct, months_to, same_period_ly, split_period
from smartview_mcp.userexcel import _verify_labels, is_archive, local_path


def test_fin_matches_sheet_format():
    assert fin(-1112834010.34) == "(1 112 834 010)"
    assert fin(59752923.67) == "59 752 924"
    assert fin(0) == "–" and fin(0.2) == "–"
    assert fin(544.2, "tall1") == "544,2"
    assert fin(0.13, "prosent1") == "+13,0 %" and fin(-0.044, "prosent1") == "-4,4 %"


def test_text_amounts_use_minus_never_double_parentheses():
    assert _fmt(-151328997) == "−151 328 997"
    assert "(" not in _fmt(-1) and _fmt(1234) == "1 234"
    assert _pct(954.0, 1000.0) == "−4,6 %"          # (current − comparison) / |comparison|
    assert _pct(-1112.8, -1066.4) == "−4,4 %"       # costs/EBITDA negative: worse = negative


def test_periods():
    assert split_period("Sep 26") == (9, 26)
    assert same_period_ly("Jan 26") == "Jan 25"
    assert months_to("Mar 26") == ["Jan 26", "Feb 26", "Mar 26"]


def test_clean_labels():
    assert _clean("C30 REVENUES (ac)") == "C30 REVENUES"
    assert _clean("EBITDA EBITDA (ac)") == "EBITDA"
    assert _clean("2TEC Technology (un.SY)") == "2TEC Technology"


def test_verify_labels_catches_shifted_rows():
    lay = Layout([], [Axis("All Accounts", ["ac.C30", "ac.C40S"])], [Axis("All Time Periods", ["Sep 26"])])
    ok = SimpleNamespace(row_labels=[["C30 REVENUES (ac)"], ["C40S Total Cost (ac)"]], col_labels=[["Sep 26"]])
    shifted = SimpleNamespace(row_labels=[["C40S Total Cost (ac)"], ["C48 Contribution (ac)"]], col_labels=[["Sep 26"]])
    assert _verify_labels(lay, ok) == []
    assert len(_verify_labels(lay, shifted)) == 2


def test_archive_guard_local_url_and_unsaved():
    assert is_archive(r"C:\Users\t992115\OneDrive - Telenor\Analyse (Telenor Norge AS) - Produktregnskapet\2025\x.xlsx")
    assert is_archive("https://telenorgroup.sharepoint.com/sites/AnalyseTelenorNorgeAS/Shared Documents/2025/x.xlsx")
    assert not is_archive(r"C:\Users\t992115\OneDrive - Telenor\Projects\prodreg2026\analyser\a.xlsx")
    assert not is_archive("Book1")


def test_local_path_maps_personal_onedrive():
    url = ("https://telenorgroup-my.sharepoint.com/personal/tom_chamkrai_hoel_telenor_no/Documents/"
           "Projects/prodreg2026/analyser/a%20b.xlsx")
    assert local_path(url).endswith(r"Projects\prodreg2026\analyser\a b.xlsx")
