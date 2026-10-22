from types import SimpleNamespace

import pytest

from smartview_mcp.sheetorder import arrange_sheets, plan_moves, sheet_order, simulate

SCRAMBLED = ["Kilder", "Q2 FinMax", "Q1 Linjer", "Q1 Svar", "Intervjuguide", "Q10 Svar", "Q2 Svar",
             "Q1 FinLite 2", "Q1 FinLite", "_svmeta"]


def vis(names, hidden=("_svmeta",)):
    return {n: n not in hidden for n in names}


def test_full_order_scrambled():
    assert sheet_order(SCRAMBLED, vis(SCRAMBLED)) == [
        "Q1 Svar", "Q2 Svar", "Q10 Svar", "Intervjuguide",
        "Q1 FinLite", "Q1 FinLite 2", "Q1 Linjer", "Q2 FinMax", "Kilder", "_svmeta"]


def test_numeric_question_sort_not_text():
    names = ["Q10 Svar", "Q9 Svar", "Q2 Svar", "Q1 Svar", "Q11 FinMax", "Q9 FinMax"]
    assert sheet_order(names, set(names)) == ["Q1 Svar", "Q2 Svar", "Q9 Svar", "Q10 Svar", "Q9 FinMax", "Q11 FinMax"]


def test_cubes_then_suffixes_then_linjer_aliases():
    names = ["Q1 Linjer b", "Q1 FinPlan", "Q1 Linjer", "Q1 FinMax w", "Q1 FinLite 3", "Q1 FinMax",
             "Q1 FinLite 2", "Q1 FinLite", "Q1 Linjer a", "Q1 Svar"]
    assert sheet_order(names, set(names)) == [
        "Q1 Svar", "Q1 FinMax", "Q1 FinMax w", "Q1 FinLite", "Q1 FinLite 2", "Q1 FinLite 3", "Q1 FinPlan",
        "Q1 Linjer", "Q1 Linjer b", "Q1 Linjer a"]


def test_unknown_sheets_keep_relative_order_in_middle():
    names = ["FY2025 split", "Q1 FinMax", "Q3 Forecast", "Q1 Svar", "Intervjuguide", "Kilder SV", "Kilder"]
    assert sheet_order(names, set(names)) == [
        "Q1 Svar", "FY2025 split", "Q3 Forecast", "Intervjuguide", "Q1 FinMax", "Kilder", "Kilder SV"]


def test_hidden_sheets_untouched_at_end():
    names = ["_svmeta", "Q1 FinMax", "~Q1 FinMax", "Q1 Svar", "Kilder"]
    order = sheet_order(names, vis(names, hidden=("_svmeta", "~Q1 FinMax")))
    assert order == ["Q1 Svar", "Q1 FinMax", "Kilder", "_svmeta", "~Q1 FinMax"]
    moves = plan_moves(names, order, {"_svmeta", "~Q1 FinMax"})
    assert all(m[0] not in ("_svmeta", "~Q1 FinMax") for m in moves)
    assert simulate(names, moves) == order


def test_active_mode_only_reorders_our_sheets_in_their_slots():
    names = ["Forside", "Kilder", "Mine tall", "Q2 Svar", "Q1 FinMax", "Q1 Svar", "Notater"]
    order = sheet_order(names, set(names), mode="active")
    assert order == ["Forside", "Q1 Svar", "Mine tall", "Q2 Svar", "Q1 FinMax", "Kilder", "Notater"]


def test_active_mode_user_kilder_is_not_ours():
    names = ["Kilder", "Q1 FinMax", "Kilder SV", "Q1 Svar"]
    assert sheet_order(names, set(names), mode="active") == ["Kilder", "Q1 Svar", "Q1 FinMax", "Kilder SV"]


def test_plan_moves_minimal_and_none_when_ordered():
    order = sheet_order(SCRAMBLED, vis(SCRAMBLED))
    assert plan_moves(order, order, {"_svmeta"}) == []
    moves = plan_moves(SCRAMBLED, order, {"_svmeta"})
    assert simulate(SCRAMBLED, moves) == order
    assert len(moves) <= len(SCRAMBLED) - 3          # several sheets stay put
    one = ["Q1 FinMax", "Q1 Svar", "Kilder"]
    assert plan_moves(one, ["Q1 Svar", "Q1 FinMax", "Kilder"]) in (
        [("Q1 FinMax", "before", "Kilder")], [("Q1 Svar", "before", "Q1 FinMax")])


def test_active_mode_never_moves_user_sheets():
    names = ["Forside", "Kilder", "Mine tall", "Q2 Svar", "Q1 FinMax", "Q1 Svar", "Notater", "skjult"]
    v = vis(names, hidden=("skjult",))
    order = sheet_order(names, v, mode="active")
    moves = plan_moves(names, order, {"Forside", "Mine tall", "Notater", "skjult"})
    assert {m[0] for m in moves} <= {"Kilder", "Q2 Svar", "Q1 FinMax", "Q1 Svar"}
    assert simulate(names, moves) == order


# ------------------------------------------------------------------ fake COM
class FakeSheets:
    def __init__(self, wb):
        self.wb = wb

    def __call__(self, key):
        return next(s for s in self.wb.order if s.Name == key)

    def __iter__(self):
        return iter(list(self.wb.order))


class FakeSheet:
    def __init__(self, wb, name, visible=-1, a1=None):
        self.wb, self.Name, self.Visible = wb, name, visible
        self.Tab = SimpleNamespace(Color=False)
        self.Range = lambda a: SimpleNamespace(Value=a1)

    def Move(self, *args, **kw):
        assert not kw, "keyword Before=/After= is ignored through late-bound COM: positional only"
        before = args[0] if args else None
        after = args[1] if len(args) > 1 else None
        assert self.Visible == -1, "hidden sheets must never be moved"
        self.wb.order.remove(self)
        k = self.wb.order.index(before if before is not None else after)
        self.wb.order.insert(k if before is not None else k + 1, self)
        self.wb.moves += 1

    def Activate(self):
        self.wb.ActiveSheet = self


class FakeWb:
    def __init__(self, names, hidden=(), a1=None):
        self.order = [FakeSheet(self, n, 0 if n in hidden else -1, (a1 or {}).get(n)) for n in names]
        self.Worksheets = FakeSheets(self)
        self.ActiveSheet = self.order[0]
        self.Application = SimpleNamespace(ScreenUpdating=True, Calculation=-4105)
        self.moves = 0

    @property
    def names(self):
        return [s.Name for s in self.order]


def test_arrange_sheets_fake_com_full():
    wb = FakeWb(SCRAMBLED, hidden=("_svmeta",))
    out = arrange_sheets(wb, "full", activate="Q10 Svar")
    assert wb.names == sheet_order(SCRAMBLED, vis(SCRAMBLED))
    assert out["order"] == wb.names and wb.moves == len(out["moved"])
    assert wb.ActiveSheet.Name == "Q10 Svar"
    tabs = {s.Name: s.Tab.Color for s in wb.order}
    assert tabs["Q1 Svar"] == 0x64381F and tabs["Q1 FinLite 2"] == 0xA6A6A6 and tabs["Q1 Linjer"] == 0xA6A6A6
    assert tabs["Kilder"] is False and tabs["Intervjuguide"] is False
    assert wb.Application.ScreenUpdating is True and wb.Application.Calculation == -4105   # restored
    again = arrange_sheets(wb, "full")
    assert again["moved"] == [] and wb.ActiveSheet.Name == "Q10 Svar"


def test_arrange_sheets_fake_com_active_keeps_user_kilder():
    names = ["Kilder", "Q1 FinMax", "Mine tall", "Q1 Svar"]
    wb = FakeWb(names, a1={"Kilder": "Min egen kildeliste"})
    arrange_sheets(wb, "active", activate="Q1 Svar")
    assert wb.names == ["Kilder", "Q1 Svar", "Mine tall", "Q1 FinMax"]


def test_plan_moves_rejects_mismatch():
    with pytest.raises(ValueError):
        plan_moves(["a", "b"], ["a", "c"])
