"""workbench snapshot logic that needs no Excel."""
from smartview_mcp.workbench.history import _coerced, _constants


def test_constants_skip_formula_cells_and_keep_types():
    f2 = [["Line", "2026"], ["=B2", ""]]          # Formula2 grid of B2:C3 ('2026' typed as text)
    v2 = [["Line", "2026"], [5.0, None]]          # Value2 grid
    assert _constants(f2, v2, 2, 2) == {"B2": "Line", "C2": "2026", "C3": None}


def test_coerced_finds_text_turned_number_and_leftovers():
    consts = {"B2": "Line", "C2": "2026", "D2": 80, "C3": None}
    now = [["Line", 2026.0, 80.0], [7.0, 1.0, None]]   # B2:D3 after the XML import
    assert _coerced(consts, now, 2, 2) == [("C2", "2026"), ("C3", None)]


def test_dependents_outside_a_row_delete_are_listed_once(monkeypatch):
    import types

    from smartview_mcp.workbench import changes, fmt
    sheet = lambda n: types.SimpleNamespace(Name=n)          # noqa: E731
    b = types.SimpleNamespace(wb=types.SimpleNamespace(Worksheets=[sheet("Data"), sheet("Calc")]))
    f = {"Data": [("Data", "D9", "=SUM(B3:B4)"), ("Data", "B3", "=B4"), ("Data", "D10", "=SUM(B2:B5)")],
         "Calc": [("Calc", "A1", "=Data!B3+Data!C3"), ("Calc", "A2", "=Other!B3")]}
    monkeypatch.setattr(fmt.fx, "sheet_formulas", lambda ws, com: f[ws.Name])
    assert changes._dependents_outside(b, sheet("Data"), "3:4", True) == ["Data!D9", "Calc!A1"]


def test_constants_skip_spill_children():
    f2 = [["=SEQUENCE(2)"], [""], [""]]          # B2 spills into B3; B4 is really empty
    v2 = [[1.0], [2.0], [None]]
    assert _constants(f2, v2, 2, 2) == {"B4": None}


def test_format_grid_has_the_range_shape():
    from smartview_mcp.workbench.books import _format_grid
    assert _format_grid("0.0", 2, 3, lambda i, j: 1 / 0) == [["0.0"] * 3] * 2      # uniform: no per-cell reads
    assert _format_grid(None, 2, 2, lambda i, j: f"f{i}{j}") == [["f00", "f01"], ["f10", "f11"]]


def test_changes_filter_matches_a_local_path_to_a_onedrive_url(monkeypatch):
    from smartview_mcp.workbench import history
    url = "https://telenor-my.sharepoint.com/personal/x/Documents/analyser/Q.xlsx"
    monkeypatch.setattr(history, "local_path", lambda s: r"C:\OneDrive\analyser\Q.xlsx" if s == url else s)
    ch = {"workbook": "Q.xlsx", "full_name": url}
    assert history._same_book(r"C:\OneDrive\analyser\Q.xlsx", ch)
    assert history._same_book("q", ch) and history._same_book(url, ch)
    assert not history._same_book(r"C:\Other\Q.xlsx", ch)


def test_a_sheet_with_only_formats_has_no_content():
    from types import SimpleNamespace as NS

    from smartview_mcp.workbench.history import _has_content
    app = NS(WorksheetFunction=NS(CountA=lambda cells: cells["n"]))
    assert not _has_content(app, NS(Cells={"n": 0}, Shapes=NS(Count=0)))   # formatted after a revert, no values
    assert _has_content(app, NS(Cells={"n": 3}, Shapes=NS(Count=0)))
    assert _has_content(app, NS(Cells={"n": 0}, Shapes=NS(Count=1)))      # a chart the user added
