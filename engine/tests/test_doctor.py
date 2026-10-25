import json
import zipfile

from smartview_mcp import config, doctor


def test_overall_and_exceptions_become_fail(monkeypatch):
    def ok():
        return doctor.Check("a", "OK", "fine")

    def boom():
        raise RuntimeError("kaput")
    monkeypatch.setattr(doctor, "CHECKS", (ok, boom))
    rs = doctor.run()
    assert [c.status for c in rs] == ["OK", "FAIL"] and "kaput" in rs[1].detail and rs[1].fix
    assert doctor.overall(rs) == "FAIL"
    assert doctor.overall([doctor.Check("a", "OK", ""), doctor.Check("b", "WARN", "")]) == "WARN"


def test_config_check(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "CONFIG_ERROR", "config.toml: Expected ']'")
    assert doctor._config().status == "FAIL"
    monkeypatch.setattr(config, "CONFIG_ERROR", None)
    monkeypatch.setattr(config, "USER_CONFIG", tmp_path / "config.toml")
    assert doctor._config().status == "WARN"
    (tmp_path / "config.toml").write_text("", encoding="utf-8")
    assert doctor._config().status == "OK"


def test_output_check_writes_and_cleans_up(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "ANALYSER_DIR", tmp_path / "a")
    assert doctor._output().status == "OK" and list((tmp_path / "a").iterdir()) == []


def test_bundle_has_doctor_versions_logs_and_redacted_config(monkeypatch, tmp_path):
    (tmp_path / "logs").mkdir()
    (tmp_path / "logs" / "server.log").write_text("x sso_token=abcdefghijklmnopqrstu\n", encoding="utf-8")  # allow-secret
    (tmp_path / "config.toml").write_text('[http]\nsecret = "verysecretvalue123456"\n', encoding="utf-8")  # allow-secret
    monkeypatch.setattr(config, "HOME", tmp_path)
    monkeypatch.setattr(config, "USER_CONFIG", tmp_path / "config.toml")
    monkeypatch.setattr(doctor, "CHECKS", (lambda: doctor.Check("a", "OK", "fine"),))
    path = doctor.bundle()
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        assert {"doctor.json", "versions.json", "config.toml", "logs/server.log"} <= names
        assert "verysecret" not in z.read("config.toml").decode()
        assert "abcdefghij" not in z.read("logs/server.log").decode()
        assert json.loads(z.read("doctor.json"))["status"] == "OK"


def test_warehouse_check(monkeypatch, tmp_path):
    import time

    from smartview_mcp import warehouse as wh
    monkeypatch.setattr(config, "WAREHOUSE_DB", tmp_path / "w.duckdb")
    c = doctor._warehouse()
    assert c.status == "OK" and "not filled yet" in c.detail             # a fresh install
    w = wh.Warehouse(tmp_path / "w.duckdb")
    w.put([wh.Row(("FinMax", "Actual", "", "E", "ac.C30", "Sep 26", "Periodic"), 1.0, "http", 1.0, "Sep 26")])
    assert doctor._warehouse().status == "WARN"                          # cells but no fill recorded
    w.set_state(closed_month="Sep 26")
    w.record_fill(wh.Fill("base", "x", time.time(), 1.0, 1, "ok", {}))
    c = doctor._warehouse()
    assert c.status == "OK" and "closed month Sep 26" in c.detail and "(base: ok)" in c.detail
    w.record_fill(wh.Fill("new_month", "x", time.time(), 1.0, 0, "error", {"errors": ["FinMax: provider down"]}))
    c = doctor._warehouse()
    assert c.status == "WARN" and "provider down" in c.detail and "sv_warehouse" in c.fix


def test_warehouse_check_names_missing_cells_of_a_warn_fill(monkeypatch, tmp_path):
    import time

    from smartview_mcp import warehouse as wh
    monkeypatch.setattr(config, "WAREHOUSE_DB", tmp_path / "w.duckdb")
    w = wh.Warehouse(tmp_path / "w.duckdb")
    w.put([wh.Row(("FinMax", "Actual", "", "E", "ac.C30", "Sep 26", "Periodic"), 1.0, "http", 1.0, "Sep 26")])
    warns = ["FinMax E 1: ac.C30 / Aug 26: #Missing", "FinMax E 2: ac.C40S / Aug 26: #Missing",
             "FinMax E 3: unknown members ac.X", "FinMax E 4: ac.Y / Jul 26: #Missing"]
    w.record_fill(wh.Fill("base", "x", time.time(), 1.0, 1, "warn", {"warnings": warns, "problems": []}))
    c = doctor._warehouse()
    assert c.status == "WARN" and "ac.C30 / Aug 26" in c.detail and "ac.C40S / Aug 26" in c.detail
    assert "+1 more" in c.detail
    w.record_fill(wh.Fill("top_up", "x", time.time(), 1.0, 1, "warn", {"problems": ["E: EBITDA differs by 5"]}))
    c = doctor._warehouse()
    assert "EBITDA differs by 5" in c.detail


def test_warehouse_check_shows_paused_jobs_with_their_error(monkeypatch, tmp_path):
    import time

    from smartview_mcp import warehouse as wh
    monkeypatch.setattr(config, "WAREHOUSE_DB", tmp_path / "w.duckdb")
    w = wh.Warehouse(tmp_path / "w.duckdb")
    w.put([wh.Row(("FinMax", "Actual", "", "E", "ac.C30", "Sep 26", "Periodic"), 1.0, "http", 1.0, "Sep 26")])
    w.record_fill(wh.Fill("base", "x", time.time(), 1.0, 1, "ok", {}))
    w.set_state(retry_after_forecast=time.time() + 900, fails_forecast=2, last_error_forecast="FinPlan: no such member",
                retry_after_new_month=0, last_error_new_month="old")
    c = doctor._warehouse()
    assert c.status == "WARN" and "forecast" in c.detail and "FinPlan: no such member" in c.detail
    assert "new_month" not in c.detail                                    # not paused any more


def test_warehouse_check_never_throws_on_a_broken_file(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "WAREHOUSE_DB", tmp_path / "w.duckdb")
    (tmp_path / "w.duckdb").write_bytes(b"not a duckdb file")
    c = doctor._warehouse()
    assert c.status == "WARN" and c.fix
