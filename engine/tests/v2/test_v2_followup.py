# engine/tests/v2/test_v2_followup.py
import pytest

from smartass.followup import find_question, merge_patch, v2_questions
from smartass.model.validate import GuardError

STORED = {1: {"spec": {"version": 1, "question": "old v1"}, "ids": {}, "owned": None},
          2: {"spec": {"v2": {"version": 2, "question": "a"}}, "ids": {}, "owned": {"rows": 30, "cols": 6}},
          5: {"spec": {"v2": {"version": 2, "question": "b"}}, "ids": {}, "owned": {"rows": 40, "cols": 7}}}


def test_merge_patch_is_rfc7386():
    base = {"plan": {"depth": "quick", "purpose": "x"}, "blocks": [1, 2], "question": "q"}
    out = merge_patch(base, {"plan": {"purpose": "y"}, "blocks": [3], "question": None})
    assert out == {"plan": {"depth": "quick", "purpose": "y"}, "blocks": [3]}
    assert base["plan"]["purpose"] == "x"                        # input untouched


def test_v2_questions_skip_v1():
    assert v2_questions(STORED) == [2, 5]


def test_find_question_explicit_active_latest():
    assert find_question(STORED, q=2, active_sheet="Q5 Answer")[0] == 2
    assert find_question(STORED, q=None, active_sheet="Q2 Answer")[0] == 2
    assert find_question(STORED, q=None, active_sheet="Mine")[0] == 5
    q, raw, owned = find_question(STORED, q=None, active_sheet=None)
    assert (q, raw["question"], owned["rows"]) == (5, "b", 40)
    with pytest.raises(GuardError, match="Q1 is not a v2 answer"):
        find_question(STORED, q=1, active_sheet=None)
    with pytest.raises(GuardError, match="no v2 answer"):
        find_question({}, q=None, active_sheet=None)
