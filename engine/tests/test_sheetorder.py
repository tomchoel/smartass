
import pytest

from smartass.xl.sheetorder import arrange_sheets, classify, plan_moves, sheet_order, simulate

SCRAMBLED = ["Sources", "Q2 FinMax", "Q1 Lines", "Q1 Answer", "Intervjuguide", "Q10 Answer", "Q2 Answer",
             "Q1 FinLite 2", "Q1 FinLite", "_svmeta"]


def vis(names, hidden=("_svmeta",)):
    return {n: n not in hidden for n in names}


def test_full_order_scrambled():
    assert sheet_order(SCRAMBLED, vis(SCRAMBLED)) == [
        "Q1 Answer", "Q2 Answer", "Q10 Answer", "Intervjuguide",
        "Q1 FinLite", "Q1 FinLite 2", "Q1 Lines", "Q2 FinMax", "Sources", "_svmeta"]


def test_numeric_question_sort_not_text():
    names = ["Q10 Answer", "Q9 Answer", "Q2 Answer", "Q1 Answer", "Q11 FinMax", "Q9 FinMax"]
    assert sheet_order(names, set(names)) == ["Q1 Answer", "Q2 Answer", "Q9 Answer", "Q10 Answer", "Q9 FinMax", "Q11 FinMax"]


def test_cubes_then_suffixes_then_lines_aliases():
    names = ["Q1 Lines b", "Q1 FinPlan", "Q1 Lines", "Q1 FinMax w", "Q1 FinLite 3", "Q1 FinMax",
             "Q1 FinLite 2", "Q1 FinLite", "Q1 Lines a", "Q1 Answer"]
    assert sheet_order(names, set(names)) == [
        "Q1 Answer", "Q1 FinMax", "Q1 FinMax w", "Q1 FinLite", "Q1 FinLite 2", "Q1 FinLite 3", "Q1 FinPlan",
        "Q1 Lines", "Q1 Lines b", "Q1 Lines a"]


def test_unknown_sheets_keep_relative_order_in_middle():
    names = ["FY2025 split", "Q1 FinMax", "Q3 Forecast", "Q1 Answer", "Intervjuguide", "Sources SV", "Sources"]
    assert sheet_order(names, set(names)) == [
        "Q1 Answer", "FY2025 split", "Q3 Forecast", "Intervjuguide", "Q1 FinMax", "Sources", "Sources SV"]


def test_hidden_sheets_untouched_at_end():
    names = ["_svmeta", "Q1 FinMax", "~Q1 FinMax", "Q1 Answer", "Sources"]
    order = sheet_order(names, vis(names, hidden=("_svmeta", "~Q1 FinMax")))
    assert order == ["Q1 Answer", "Q1 FinMax", "Sources", "_svmeta", "~Q1 FinMax"]
    moves = plan_moves(names, order, {"_svmeta", "~Q1 FinMax"})
    assert all(m[0] not in ("_svmeta", "~Q1 FinMax") for m in moves)
    assert simulate(names, moves) == order


def test_active_mode_only_reorders_our_sheets_in_their_slots():
    names = ["Forside", "Sources", "Mine tall", "Q2 Answer", "Q1 FinMax", "Q1 Answer", "Notater"]
    order = sheet_order(names, set(names), mode="active")
    assert order == ["Forside", "Q1 Answer", "Mine tall", "Q2 Answer", "Q1 FinMax", "Sources", "Notater"]


def test_active_mode_user_sources_is_not_ours():
    names = ["Sources", "Q1 FinMax", "Sources SV", "Q1 Answer"]
    assert sheet_order(names, set(names), mode="active") == ["Sources", "Q1 Answer", "Q1 FinMax", "Sources SV"]


def test_plan_moves_minimal_and_none_when_ordered():
    order = sheet_order(SCRAMBLED, vis(SCRAMBLED))
    assert plan_moves(order, order, {"_svmeta"}) == []
    moves = plan_moves(SCRAMBLED, order, {"_svmeta"})
    assert simulate(SCRAMBLED, moves) == order
    assert len(moves) <= len(SCRAMBLED) - 3          # several sheets stay put
    one = ["Q1 FinMax", "Q1 Answer", "Sources"]
    assert plan_moves(one, ["Q1 Answer", "Q1 FinMax", "Sources"]) in (
        [("Q1 FinMax", "before", "Sources")], [("Q1 Answer", "before", "Q1 FinMax")])


def test_active_mode_never_moves_user_sheets():
    names = ["Forside", "Sources", "Mine tall", "Q2 Answer", "Q1 FinMax", "Q1 Answer", "Notater", "skjult"]
    v = vis(names, hidden=("skjult",))
    order = sheet_order(names, v, mode="active")
    moves = plan_moves(names, order, {"Forside", "Mine tall", "Notater", "skjult"})
    assert {m[0] for m in moves} <= {"Sources", "Q2 Answer", "Q1 FinMax", "Q1 Answer"}
    assert simulate(names, moves) == order


# ------------------------------------------------------------------ on the fake Excel (batches through the COM executor)
def _fake(names, hidden=(), a1=None):
    import fakexl
    wb = fakexl.Workbook(list(names), styles_ver=False)
    for ws in wb.order:
        ws.Visible = 0 if ws.Name in hidden else -1
    for name, value in (a1 or {}).items():
        wb.Worksheets(name).put(1, 1, value)
    wb.moves = []
    real = fakexl.Worksheet.Move

    def move(ws, *args, **kw):
        assert ws.Visible == -1, "hidden sheets must never be moved"
        wb.moves.append(ws.Name)
        return real(ws, *args, **kw)
    for ws in wb.order:
        ws.Move = lambda *a, ws=ws, **k: move(ws, *a, **k)
    return wb


def _names(wb) -> list[str]:
    return [s.Name for s in wb.order]


def test_arrange_sheets_fake_com_full():
    wb = _fake(SCRAMBLED, hidden=("_svmeta",))
    out = arrange_sheets(wb, "full", activate="Q10 Answer")
    assert _names(wb) == sheet_order(SCRAMBLED, vis(SCRAMBLED))
    assert out["order"] == _names(wb) and wb.moves == out["moved"]
    assert wb.active.Name == "Q10 Answer"
    tabs = {s.Name: s.Tab.Color for s in wb.order}
    assert tabs["Q1 Answer"] == 0x64381F and tabs["Q1 FinLite 2"] == 0xA6A6A6 and tabs["Q1 Lines"] == 0xA6A6A6
    assert tabs["Sources"] is False and tabs["Intervjuguide"] is False
    app = wb.Application
    assert app.ScreenUpdating is True and app.Calculation == -4105 and app.EnableEvents is True   # restored
    again = arrange_sheets(wb, "full")
    assert again["moved"] == [] and wb.active.Name == "Q10 Answer"


def test_arrange_sheets_fake_com_active_keeps_user_sources():
    names = ["Sources", "Q1 FinMax", "Mine tall", "Q1 Answer"]
    wb = _fake(names, a1={"Sources": "Min egen kildeliste"})
    arrange_sheets(wb, "active", activate="Q1 Answer")
    assert _names(wb) == ["Sources", "Q1 Answer", "Mine tall", "Q1 FinMax"]


def test_a_move_excel_refuses_is_raised_by_arrange_sheets():
    import fakexl
    wb = _fake(["Q2 Answer", "Q1 Answer"])

    def refuse(*a, **k):
        raise fakexl.ComError("the move was refused")
    for ws in wb.order:
        ws.Move = refuse
    with pytest.raises(RuntimeError, match="refused"):
        arrange_sheets(wb, "full")


def test_plan_moves_rejects_mismatch():
    with pytest.raises(ValueError):
        plan_moves(["a", "b"], ["a", "c"])


def test_answer_sheets_count_as_questions():
    from smartass.xl.sheets import _Q
    names = ["Q2 FinMax", "Q2 Answer", "Q1 Answer"]
    assert sheet_order(names, set(names)) == ["Q1 Answer", "Q2 Answer", "Q2 FinMax"]
    assert _Q.match("Q3 Answer").group(1) == "3" and not _Q.match("Q3 Svar")


@pytest.mark.parametrize("name", ["Q3 Answer", "Q3 Lines", "Q3 FinMax", "Q3 FinPlan 2"])
def test_our_sheets_are_english(name):
    assert classify(name)[0] != "other"


@pytest.mark.parametrize("name", ["Q3 Svar", "Q3 Linjer", "Kilder", "Kilder SV"])
def test_norwegian_names_are_the_users(name):
    assert classify(name)[0] == "other"
    names = [name, "Q1 Answer", "Sources"]
    assert sheet_order(names, set(names)) == ["Q1 Answer", name, "Sources"]   # stays with the user's sheets
