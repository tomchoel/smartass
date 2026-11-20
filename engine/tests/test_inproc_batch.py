"""The batch recorder (xl/inproc/batch.py): Python code that reads like the COM code it replaces records plain data."""
import json
import math

import pytest

from smartass.xl.inproc.batch import MISSING, VERBS, Batch, Node


def ops(b):
    return b.to_dict()["ops"]


def test_a_chain_records_one_operation_per_step_it_needs():
    b = Batch()
    ws = b.workbook("Book1.xlsx").Worksheets("Q3 Data")
    lo = ws.ListObjects.Add(1, ws.Range("A1:C9"), MISSING, 1)
    assert isinstance(lo, Node)
    assert ops(b) == [
        {"k": "call", "on": 0, "name": "Workbooks", "args": ["Book1.xlsx"], "as": 1},
        {"k": "call", "on": 1, "name": "Worksheets", "args": ["Q3 Data"], "as": 2},
        {"k": "get", "on": 2, "name": "ListObjects", "as": 3},
        {"k": "call", "on": 2, "name": "Range", "args": ["A1:C9"], "as": 4},
        {"k": "call", "on": 3, "name": "Add", "args": [1, {"r": 4}, {"missing": 1}, 1], "as": 5},
    ]


def test_a_get_with_no_arguments_is_recorded_only_when_something_is_done_with_it():
    b = Batch()
    ws = b.workbook("B").Worksheets("S")
    used = ws.UsedRange                            # nothing done with it yet
    ws.Calculate                                   # an attribute never called and never used: nothing
    assert [o["name"] for o in ops(b)] == ["Workbooks", "Worksheets"]
    b.ret(used.Address, "addr")
    assert ops(b)[2:] == [{"k": "get", "on": 2, "name": "UsedRange", "as": 3},
                          {"k": "get", "on": 3, "name": "Address", "as": 4},
                          {"k": "ret", "on": 4, "name": "addr"}]


def test_a_get_node_called_later_calls_the_member_with_the_arguments():
    b = Batch()
    wss = b.workbook("B").Worksheets
    b.ret(wss.Count, "n")                          # wss recorded as a get
    wss("S").Name = "T"
    assert ops(b)[4] == {"k": "call", "on": 1, "name": "Worksheets", "args": ["S"], "as": 4}


def test_a_call_with_no_arguments_is_a_call():
    b = Batch()
    b.workbook("B").Worksheets("S").Calculate()
    assert ops(b)[-1] == {"k": "call", "on": 2, "name": "Calculate", "args": [], "as": 3}


def test_a_node_used_twice_is_recorded_once():
    b = Batch()
    ws = b.workbook("B").Worksheets("S")
    font = ws.Range("A1").Font
    font.Bold = True
    font.Size = 9
    assert [o["k"] + " " + o["name"] for o in ops(b)] == [
        "call Workbooks", "call Worksheets", "call Range", "get Font", "set Bold", "set Size"]


def test_assignment_records_a_set():
    b = Batch()
    lo = b.workbook("B").Worksheets("S").ListObjects("T")
    lo.Name = "Data_Q3"
    lo.ShowTableStyleRowStripes = False
    assert ops(b)[-2:] == [{"k": "set", "on": 3, "name": "Name", "value": "Data_Q3"},
                           {"k": "set", "on": 3, "name": "ShowTableStyleRowStripes", "value": False}]


def test_ret_asks_for_a_value_back_under_its_name():
    b = Batch()
    ws = b.workbook("B").Worksheets("S")
    b.ret(ws.Range("A1:B2").Value2, "block")
    b.ret(b.app.Calculation, "calc")
    assert ops(b)[-3:] == [{"k": "ret", "on": 4, "name": "block"},
                           {"k": "get", "on": 0, "name": "Calculation", "as": 5},
                           {"k": "ret", "on": 5, "name": "calc"}]
    with pytest.raises(ValueError, match="block"):
        b.ret(ws.Name, "block")                    # one value per name


def test_a_com_object_asked_back_is_refused_when_the_recorder_knows_it_is_one():
    b = Batch()
    wb = b.workbook("B")
    ws = wb.Worksheets("S")
    ws.Name = "T"                                  # used as an object
    for node in (b.app, wb, ws):
        with pytest.raises(TypeError, match="COM object"):
            b.ret(node, "x")
    rng = ws.Range("A1")
    b.ret(rng, "r")                                # not known yet: the executors refuse it when it runs
    with pytest.raises(TypeError, match="COM object"):
        rng.Value2 = 1                             # ... and the recorder once it is used as one


def test_missing_and_the_value_kinds_are_plain_data():
    b = Batch()
    rng = b.workbook("B").Worksheets("S").Range("A1:C2")
    rng.Value2 = ((None, True, 3), (4.5, 'quote " and \u00e6\u00f8\u00e5', -2))
    b.workbook("B").Worksheets(["S", "T"]).Copy(MISSING, None)
    d = b.to_dict()
    assert d["ops"][3] == {"k": "set", "on": 3, "name": "Value2",
                           "value": {"grid": [[None, True, 3], [4.5, 'quote " and \u00e6\u00f8\u00e5', -2]]}}
    assert d["ops"][4]["args"] == [{"list": ["S", "T"]}]
    assert d["ops"][5]["args"] == [{"missing": 1}, None]
    assert json.loads(json.dumps(d)) == d


def test_a_grid_is_a_tuple_of_rows_and_must_be_rectangular_and_scalar():
    b = Batch()
    rng = b.workbook("B").Worksheets("S").Range("A1:B2")
    rng.Formula2 = [["=1", "=2"], ("x", None)]                 # lists of rows are grids too
    assert ops(b)[-1]["value"] == {"grid": [["=1", "=2"], ["x", None]]}
    with pytest.raises(ValueError, match="rectangular"):
        rng.Value2 = ((1, 2), (3,))
    with pytest.raises(TypeError):
        rng.Value2 = ((1, b.app),)
    with pytest.raises(ValueError):
        rng.Value2 = ()


def test_values_that_cannot_travel_are_refused_when_recorded():
    rng = Batch().workbook("B").Worksheets("S").Range("A1")
    for bad in (math.nan, math.inf, 2 ** 63, {"a": 1}, object(), b"x"):
        with pytest.raises((TypeError, ValueError)):
            rng.Value2 = bad


def test_a_node_used_as_an_argument_is_recorded_first():
    b = Batch()
    wb = b.workbook("B")
    last = wb.Worksheets.Item(3)
    wb.Worksheets("S").Move(MISSING, last)
    assert ops(b)[1:] == [{"k": "get", "on": 1, "name": "Worksheets", "as": 2},
                          {"k": "call", "on": 2, "name": "Item", "args": [3], "as": 3},
                          {"k": "call", "on": 1, "name": "Worksheets", "args": ["S"], "as": 4},
                          {"k": "call", "on": 4, "name": "Move", "args": [{"missing": 1}, {"r": 3}], "as": 5}]


def test_always_records_into_the_finally_list():
    b = Batch()
    src = b.app.Workbooks.Open("C:/x.xlsx", 0, True)
    with b.always():
        src.Close(False)
    b.mark("copied")
    d = b.to_dict()
    assert d["finally"] == [{"k": "call", "on": 2, "name": "Close", "args": [False], "as": 3}]
    assert d["ops"][-1] == {"k": "mark", "name": "copied"}


def test_a_node_made_in_always_cannot_be_used_outside_it():
    b = Batch()
    with b.always():
        ws = b.workbook("B").Worksheets("S")
    with pytest.raises(ValueError, match="always"):
        ws.Name = "x"
    with pytest.raises(RuntimeError):
        with b.always():
            with b.always():
                pass


def test_attempt_marks_an_operation_whose_failure_does_not_stop_the_batch():
    b = Batch()
    ws = b.workbook("B").Worksheets("S")
    lo = b.attempt(ws.ListObjects("Old"))
    with b.attempt():
        lo.Delete()
        ws.Name = "T"
    ws.Calculate()
    tried = [o.get("try", False) for o in ops(b)]
    assert tried == [False, False, True, True, True, False]


def test_each_is_a_verb_and_another_is_refused():
    """The file job's loops are one verb (Task 2b): each(collection, paths); break_links went with the copy (Task 2a)."""
    b = Batch()
    rows = b.verb("each", b.workbook("Book1.xlsx").Worksheets, ["Name", "ListObjects.*.Name"])
    b.ret(rows, "rows")
    assert ops(b)[2:] == [{"k": "verb", "name": "each", "args": [{"r": 2}, {"list": ["Name", "ListObjects.*.Name"]}],
                           "as": 3},
                          {"k": "ret", "on": 3, "name": "rows"}]
    with pytest.raises(ValueError, match="verb"):
        Batch().verb("break_links", Batch().app, "smartass_build_1.xlsx")


def test_the_verbs_arguments_are_checked_when_recorded():
    """Review I1: each takes a non-empty list of names (the executors disagreed on a text for paths); nudge a property
    name, a number, None or a number; exist a non-empty list of texts and True or False. Task 2c review Minor 6: each
    has no prefixes any more (nothing used them once the names were looked up). Task 3: mentions a non-empty list of
    texts to look for and one of the sheets to skip; swap a property name and two texts."""
    assert VERBS == {"each": 2, "nudge": 4, "exist": 3, "mentions": 3, "swap": 4}
    b = Batch()
    wss = b.workbook("Book1.xlsx").Worksheets
    for paths in ("Name", [], ["Name", ""]):
        with pytest.raises(ValueError, match="each"):
            b.verb("each", wss, paths)
    with pytest.raises(ValueError, match="no verb 'each' with 3 arguments"):
        b.verb("each", wss, ["Name"], ["Data_Q"])
    for prop, delta, above in (("", 1, None), ("Top", "1", None), ("Top", True, None), ("Top", 1, "x")):
        with pytest.raises(ValueError, match="nudge"):
            b.verb("nudge", wss, prop, delta, above)
    for keys, refuse in (("Data_Q1", False), ([], False), (["Data_Q1", ""], False), (["Data_Q1"], None)):
        with pytest.raises(ValueError, match="exist"):
            b.verb("exist", wss, keys, refuse)
    for texts, skip in (("old Q1 ", ["Q1 Answer"]), ([], ["Q1 Answer"]), (["old Q1 "], []), (["old Q1 ", ""], ["x"]),
                        (["old Q1 "], "Q1 Answer")):
        with pytest.raises(ValueError, match="mentions"):
            b.verb("mentions", b.workbook("Book1.xlsx"), texts, skip)
    b.verb("each", wss, ["Name"])
    b.verb("nudge", wss, "Top", -10.0, 10.0)
    b.verb("exist", wss, ["Data_Q1"], True)
    b.verb("mentions", b.workbook("Book1.xlsx"), ["old Q1 ", "Data_Q1_old"], ["old Q1 Answer", "Q1 Answer"])
    for prop, expected, value in (("", "=a", "=b"), ("Formula2", None, "=b"), ("Formula2", "=a", 1)):
        with pytest.raises(ValueError, match="swap"):
            b.verb("swap", wss, prop, expected, value)
    b.verb("swap", wss, "Formula2", "=a", "=b")


def test_a_quiet_batch_may_leave_calculation_alone():
    assert "calculation" not in Batch().to_dict() and "calculation" not in Batch(quiet=False, calculation=False).to_dict()
    assert Batch(calculation=False).to_dict()["calculation"] is False


def test_position_is_the_index_the_next_operation_gets():
    b = Batch()
    assert b.position == 0
    ws = b.workbook("Book1.xlsx").Worksheets("S")
    ws.Range("A1").Value2 = 1
    with b.always():
        ws.Range("A2").Value2 = 2
    assert b.position == 4 and b.text(5) == 'Workbooks("Book1.xlsx").Worksheets("S").Range("A2").Value2 = 2'


def test_a_verb_and_a_mark(monkeypatch):
    monkeypatch.setitem(VERBS, "count_links", 2)
    b = Batch()
    wb = b.workbook("Book1.xlsx")
    n = b.verb("count_links", wb, "smartass_build_1.xlsx")
    b.ret(n, "n")
    b.mark("links")
    assert ops(b)[1:] == [{"k": "verb", "name": "count_links", "args": [{"r": 1}, "smartass_build_1.xlsx"], "as": 2},
                          {"k": "ret", "on": 2, "name": "n"},
                          {"k": "mark", "name": "links"}]
    with pytest.raises(ValueError, match="verb"):
        b.verb("no_such_verb", wb)


def test_the_batch_has_a_version_and_its_quiet_flag():
    assert Batch().to_dict() == {"version": 1, "quiet": True, "ops": [], "finally": []}
    assert Batch(quiet=False).to_dict()["quiet"] is False


def test_workbook_is_one_node_per_name():
    b = Batch()
    assert b.workbook("A") is b.workbook("A")
    assert len(ops(b)) == 1


def test_what_a_node_cannot_do_raises():
    b = Batch()
    ws = b.workbook("B").Worksheets("S")
    with pytest.raises(TypeError):
        bool(ws.Visible)
    with pytest.raises(TypeError):
        iter(ws.Rows)
    with pytest.raises(TypeError):
        len(ws.Rows)
    with pytest.raises(TypeError, match="keyword"):
        ws.Range(Cell1="A1")
    with pytest.raises(AttributeError):
        ws._private
    with pytest.raises(TypeError):
        b.app()                                    # the Application is not something to call
    with pytest.raises(TypeError, match="Item"):
        ws.Range("A1:C3")(2, 2)                    # would record ws.Range(2, 2)
    with pytest.raises(TypeError, match="Item"):
        b.workbook("B")("other")                   # would record Workbooks("other")
    with pytest.raises(TypeError):
        ws.Name == "S"
    with pytest.raises(TypeError):
        ws.Name != "S"
    assert len({ws, ws}) == 1                      # still usable as a key


def test_the_text_of_an_operation_reads_like_the_code():
    b = Batch()
    ws = b.workbook("Book1.xlsx").Worksheets("Q3 Data")
    ws.Range("A1:B2").Value2 = ((1, 2), (3, 4))
    ws.ListObjects.Add(1, ws.Range("A1:B2"), MISSING, 1)
    b.ret(ws.Range("C1").Value2, "c1")
    texts = [b.text(i) for i in range(len(ops(b)))]
    assert texts[3] == 'Workbooks("Book1.xlsx").Worksheets("Q3 Data").Range("A1:B2").Value2 = <grid 2x2>'
    assert texts[6] == ('Workbooks("Book1.xlsx").Worksheets("Q3 Data").ListObjects.Add(1, '
                        '<Workbooks("Book1.xlsx").Worksheets("Q3 Data").Range("A1:B2")>, MISSING, 1)')
    assert texts[-1] == 'ret c1 = Workbooks("Book1.xlsx").Worksheets("Q3 Data").Range("C1").Value2'
    long = "x" * 30000
    ws.Range("A1").Value2 = long
    assert len(b.text(len(ops(b)) - 1)) < 200
