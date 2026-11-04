from types import SimpleNamespace

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
    out = arrange_sheets(wb, "full", activate="Q10 Answer")
    assert wb.names == sheet_order(SCRAMBLED, vis(SCRAMBLED))
    assert out["order"] == wb.names and wb.moves == len(out["moved"])
    assert wb.ActiveSheet.Name == "Q10 Answer"
    tabs = {s.Name: s.Tab.Color for s in wb.order}
    assert tabs["Q1 Answer"] == 0x64381F and tabs["Q1 FinLite 2"] == 0xA6A6A6 and tabs["Q1 Lines"] == 0xA6A6A6
    assert tabs["Sources"] is False and tabs["Intervjuguide"] is False
    assert wb.Application.ScreenUpdating is True and wb.Application.Calculation == -4105   # restored
    again = arrange_sheets(wb, "full")
    assert again["moved"] == [] and wb.ActiveSheet.Name == "Q10 Answer"


def test_arrange_sheets_fake_com_active_keeps_user_sources():
    names = ["Sources", "Q1 FinMax", "Mine tall", "Q1 Answer"]
    wb = FakeWb(names, a1={"Sources": "Min egen kildeliste"})
    arrange_sheets(wb, "active", activate="Q1 Answer")
    assert wb.names == ["Sources", "Q1 Answer", "Mine tall", "Q1 FinMax"]


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
