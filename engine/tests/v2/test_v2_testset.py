import json

from test_v2_pipeline import FakeProvider
from v2util import facts, spec_dict

from smartass.testset import runner


def test_cases_are_valid_specs():
    from smartass.model.spec import AnalysisSpec
    cases = runner.load_cases()
    assert {c["id"] for c in cases} >= {"b01", "b02", "b05", "b06", "b07", "b13", "b15", "b16"}
    for c in cases:
        AnalysisSpec.model_validate(c["spec"])


def test_compare_within_one_nok():
    golden = {"K · EBITDA": {"Actual": 100.0, "Variance %": 0.1}}
    assert runner.compare(golden, {"K · EBITDA": {"Actual": 100.4, "Variance %": 0.1}}) == []
    diffs = runner.compare(golden, {"K · EBITDA": {"Actual": 102.0, "Variance %": 0.1}})
    assert diffs and "Actual" in diffs[0]
    assert runner.compare(golden, {}) == ["K · EBITDA: missing"]


def test_compare_none_equals_none_and_small_fractions_are_exact():
    assert runner.compare({"L": {"A": None}}, {"L": {"A": None}}) == []
    assert runner.compare({"L": {"A": None}}, {"L": {"A": 1.0}})
    assert runner.compare({"L": {"A": 0.10}}, {"L": {"A": 0.11}})


def test_record_approve_run_offline(tmp_path, monkeypatch):
    monkeypatch.setattr(runner, "HOME", tmp_path)
    case = {"id": "t1", "spec": spec_dict()}
    from v2util import resolver
    monkeypatch.setattr(runner, "_resolver_for", lambda closed: resolver())
    runner.record(case, FakeProvider(facts()), lines={})
    assert (tmp_path / "t1" / "candidate.json").exists()
    runner.approve("t1")
    report = runner.run_offline([case])
    assert report == [{"id": "t1", "ok": True, "diffs": []}]
    gold = json.loads((tmp_path / "t1" / "golden.json").read_text(encoding="utf-8"))
    assert "Key figures · EBITDA" in gold["key_values"]


def test_a_case_without_golden_is_reported(tmp_path, monkeypatch):
    monkeypatch.setattr(runner, "HOME", tmp_path)
    assert runner.run_offline([{"id": "t2", "spec": spec_dict()}]) == [
        {"id": "t2", "ok": None, "diffs": ["no golden"]}]
