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


# ---- Plan 2D task 5: a candidate from the replay (offline), to diff against the golden before approving

def _recorded(tmp_path, monkeypatch):
    monkeypatch.setattr(runner, "HOME", tmp_path)
    from v2util import resolver
    monkeypatch.setattr(runner, "_resolver_for", lambda closed: resolver())
    case = {"id": "t3", "spec": spec_dict()}
    runner.record(case, FakeProvider(facts()), lines={})
    runner.approve("t3")
    return case


def test_candidate_offline_writes_the_candidate_from_the_replay(tmp_path, monkeypatch):
    case = _recorded(tmp_path, monkeypatch)
    d = tmp_path / "t3"
    gold = json.loads((d / "golden.json").read_text(encoding="utf-8"))
    (d / "candidate.json").unlink()
    out = runner.candidate_offline(case)
    cand = json.loads((d / "candidate.json").read_text(encoding="utf-8"))
    assert out["diffs"] == [] and out["id"] == "t3"
    assert cand["closed"] == gold["closed"] and cand["recorded_at"] == gold["recorded_at"]
    assert cand["key_values"] == gold["key_values"] and "replayed_at" in cand
    assert json.loads((d / "golden.json").read_text(encoding="utf-8")) == gold      # the golden is untouched


def test_candidate_offline_lists_every_difference_to_the_golden(tmp_path, monkeypatch):
    case = _recorded(tmp_path, monkeypatch)
    g = tmp_path / "t3" / "golden.json"
    gold = json.loads(g.read_text(encoding="utf-8"))
    gold["key_values"]["Key figures · EBITDA"]["Variance %"] = 0.0317
    g.write_text(json.dumps(gold), encoding="utf-8")
    out = runner.candidate_offline(case)
    assert out["diffs"] == ["Key figures · EBITDA / Variance %: golden 0.0317 vs now None"]


def test_candidate_needs_a_recording(tmp_path, monkeypatch):
    import pytest
    monkeypatch.setattr(runner, "HOME", tmp_path)
    with pytest.raises(FileNotFoundError, match="record it first"):
        runner.candidate_offline({"id": "t4", "spec": spec_dict()})


def test_the_cli_candidate_is_offline_only(tmp_path, monkeypatch, capsys):
    import pytest

    from smartass.testset.__main__ import main
    case = _recorded(tmp_path, monkeypatch)
    monkeypatch.setattr(runner, "load_cases", lambda: [case])
    with pytest.raises(SystemExit, match="--offline"):
        main(["candidate", "t3"])
    assert main(["candidate", "--offline", "t3"]) == 0
    assert "t3" in capsys.readouterr().out
