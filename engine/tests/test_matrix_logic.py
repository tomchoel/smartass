from smartview_mcp.auditlib import error_class, merge_before, verdict


def test_error_class_normalises_digits_and_keeps_first_line():
    a = error_class("HostError: drill page never asked for its data within 45s\nnext: ...")
    b = error_class("HostError: drill page never asked for its data within 47s")
    assert a == b == "HostError: drill page never asked for its data within #s"
    assert error_class(" \n ") == ""


def test_verdict_rules():
    assert verdict(ok=True, secs=10, budget=30) == "PASS"
    assert verdict(ok=True, secs=31, budget=30) == "SLOW"
    assert verdict(ok=False, secs=5, budget=30) == "FAIL"
    assert verdict(ok=False, secs=5, budget=30, limit="no FinLite in test env") == "LIMIT"


def test_merge_before_adds_previous_result():
    now = [{"id": "S1.query", "result": "PASS"}]
    before = [{"id": "S1.query", "result": "FAIL"}, {"id": "gone", "result": "PASS"}]
    assert merge_before(now, before)[0]["before"] == "FAIL"
