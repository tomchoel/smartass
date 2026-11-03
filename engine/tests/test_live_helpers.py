import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent / "live"))
import helpers  # noqa: E402


@pytest.fixture
def cookies(tmp_path, monkeypatch):
    d = tmp_path / "Network"
    d.mkdir()
    (d / "Cookies").write_text("fresh")
    monkeypatch.setattr(helpers, "_cookie_dirs", lambda: [d])
    monkeypatch.setattr(helpers, "kill_engine_edge", lambda: [])
    monkeypatch.setattr(helpers.time, "sleep", lambda s: None)
    return d


def test_expired_sso_never_overwrites_a_leftover_backup(cookies, monkeypatch):
    (cookies / "Cookies.matrix-bak").write_text("real sign-in")
    monkeypatch.setattr(helpers, "restore_matrix_bak", lambda: [])        # the restore failed (file locked)
    with pytest.raises(helpers.Limit):
        with helpers.expired_sso():
            pass
    assert (cookies / "Cookies.matrix-bak").read_text() == "real sign-in"
    assert (cookies / "Cookies").read_text() == "fresh"


def test_expired_sso_restores_a_leftover_backup_first(cookies):
    (cookies / "Cookies.matrix-bak").write_text("real sign-in")
    with helpers.expired_sso() as moved:
        assert moved == [str(cookies / "Cookies")]
        assert (cookies / "Cookies.matrix-bak").read_text() == "real sign-in"
        assert not (cookies / "Cookies").exists()
    assert (cookies / "Cookies").read_text() == "real sign-in"
    assert not list(cookies.glob("*.matrix-bak"))


def test_answer_ok_fails_on_error_cell_warnings():
    import scenarios
    out = {"checks": {"c": "OK"}, "warnings": [
        "error cell 'Q6 FinMax recon'!B3 is #NoAccess: no read access or member not valid in this cube"]}
    with pytest.raises(AssertionError, match="#NoAccess"):
        scenarios._answer_ok(out, "actual_vs_plan.sy")
    scenarios._answer_ok({"checks": {"c": "OK"}, "warnings": ["could not log Q1 in 'Sources'"]}, "ok")


def test_b16_problems_flag_missing_sheets_and_slow_frame():
    import bench
    good = {"sheets": ["Q3 Answer", "Q3 By project", "Q3 Lines d1", "Q3 Lines d2", "Q3 FinLite cube"],
            "checks": {"Blocks = lines total": "OK", "Total = cube": "OK"},
            "key_values": {"Total (cube)": {"YTD Oct 26": -1.5e8}},
            "timings": {"first frame": 2.6, "open target": 0.64}}           # 1.96 s after the workbook opened
    assert bench.b16_problems(good, 5.2) == []
    bad = {"sheets": ["Q3 Lines d1"], "checks": {"Total = cube": "DIFF 12"}, "key_values": {"Other": {}},
           "timings": {"first frame": 3.7, "open target": 0.6}}
    got = " | ".join(bench.b16_problems(bad, 3.1))
    for must in ("Answer", "By project", "Total = cube", "Total (cube)", "3.1"):
        assert must in got, must
    assert bench.b16_problems({**good, "checks": {}}, 1.0)                  # no checks at all is a problem
    assert bench.b16_problems({**good, "timings": {}}, 1.0)                 # no first frame from the engine


def test_bench_pins_its_recipe_with_a_deviation_only_when_it_departs():
    import bench
    gate = {"recipe": {"name": "cost_base", "params": {"parent": "un.SY.2TEC"}}}
    same = {"recipe": {"name": "cost_base", "params": {"parent": "un.SY.2TEC"}}}
    assert bench.pin_deviation(gate, same) is None
    assert bench.pin_deviation(gate, {"recipe": {"name": "cost_base", "params": {"parent": "un.SY.2BUS"}}}) is None
    assert bench.pin_deviation(gate, {"recipe": {"name": "cost_base", "params": {"top": 15}}}) == \
        "bench pins recipe cost_base"
    assert bench.pin_deviation(gate, {"recipe": {"name": "pnl", "params": {}}}) == "bench pins recipe pnl"
    assert bench.pin_deviation(gate, {"layout": {"sections": []}}) == "bench pins recipe free-form"
    assert bench.pin_deviation({"recipe": None}, {"layout": {"sections": []}}) is None
    before = {100, 200}
    assert helpers.leftover_excel(300, before, 0, "python.exe")           # started in the run, now empty
    assert not helpers.leftover_excel(100, before, 0, "python.exe")       # was there before the run
    assert not helpers.leftover_excel(300, before, 1, "python.exe")       # still has a workbook
    assert not helpers.leftover_excel(300, before, 0, "explorer.exe")     # the user opened it meanwhile
