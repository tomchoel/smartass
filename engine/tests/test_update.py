import json

from smartview_mcp import update


def _release(tmp_path, latest, versions):
    (tmp_path / "latest.json").write_text(json.dumps({"version": latest}), encoding="utf-8")
    for v in versions:
        (tmp_path / v / "installer").mkdir(parents=True)
        (tmp_path / v / "installer" / "install.ps1").write_text("", encoding="utf-8")
    return tmp_path


def test_parse_orders_numerically():
    assert update.parse("0.10.0") > update.parse("0.9.3")


def test_not_configured_and_missing_folder(tmp_path):
    assert update.check("", "0.2.0", "")["status"] == "not configured"
    assert update.check(str(tmp_path / "nope"), "0.2.0", "")["status"] == "error"


def test_latest_newer_is_available(tmp_path):
    r = update.check(str(_release(tmp_path, "0.2.1", ["0.2.1"])), "0.2.0", "")
    assert r["status"] == "ok" and r["update_available"] and r["target"] == "0.2.1"
    assert r["installer"].endswith(r"0.2.1\installer\install.ps1")
    assert not update.check(str(tmp_path), "0.2.1", "")["update_available"]
    assert not update.check(str(tmp_path), "0.3.0", "")["update_available"]   # never downgrade by itself


def test_pin_wins_also_downwards(tmp_path):
    rel = _release(tmp_path, "0.3.0", ["0.2.0", "0.3.0"])
    r = update.check(str(rel), "0.3.0", "0.2.0")
    assert r["update_available"] and r["target"] == "0.2.0"


def test_missing_target_folder_is_error(tmp_path):
    r = update.check(str(_release(tmp_path, "0.2.1", [])), "0.2.0", "")
    assert r["status"] == "error" and "0.2.1" in r["error"]


def test_dev_checkout_never_auto_updates(monkeypatch):
    monkeypatch.setattr(update, "is_dev_checkout", lambda: True)
    assert update.start_if_available() is False
