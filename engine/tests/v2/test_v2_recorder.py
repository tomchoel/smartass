import pytest
from test_v2_pipeline import FakeProvider
from v2util import facts, resolver, spec_dict

from smartass.data.recorder import RecordingProvider, ReplayProvider, load_lines, save_lines
from smartass.model.lines import LineSet
from smartass.pipeline import prepare


def test_record_then_replay_gives_the_same_answer(tmp_path):
    rec = RecordingProvider(FakeProvider(facts()))
    live = prepare(spec_dict(), resolver(), rec)
    rec.save(tmp_path / "b02.json")
    replay = prepare(spec_dict(), resolver(), ReplayProvider(tmp_path / "b02.json"))
    assert replay.result.key_values == live.result.key_values


def test_replay_of_an_unrecorded_request_is_an_error(tmp_path):
    RecordingProvider(FakeProvider(facts())).save(tmp_path / "empty.json")
    with pytest.raises(KeyError, match="not recorded"):
        prepare(spec_dict(), resolver(), ReplayProvider(tmp_path / "empty.json"))


def test_replay_of_a_missing_grid_is_an_error_not_a_failed_grid(tmp_path):
    rec = RecordingProvider(FakeProvider(facts()))
    rec.closed_months()
    rec.save(tmp_path / "nogrids.json")
    with pytest.raises(KeyError, match="not recorded"):
        prepare(spec_dict(), resolver(), ReplayProvider(tmp_path / "nogrids.json"))


def test_line_sets_round_trip(tmp_path):
    ls = {"b0": LineSet("b0", [{"AMOUNT": -1.0, "AP_CUSTOMER_NAME": "X"}], ["AMOUNT", "AP_CUSTOMER_NAME"],
                        cube_value=-1.0, cell={"All Accounts": "ac.C60"})}
    save_lines(tmp_path / "l.json", ls)
    back = load_lines(tmp_path / "l.json")
    assert back["b0"].rows == ls["b0"].rows and back["b0"].cube_value == -1.0
