from smartview_mcp import config


def _defaults(tmp_path):
    d = tmp_path / "defaults.toml"
    d.write_text('[epm]\nbase = "https://org"\n[drill]\nchannels = ["msedge", "chrome"]\n', encoding="utf-8")
    return d


def test_user_file_overrides_defaults_and_env_overrides_both(tmp_path):
    u = tmp_path / "config.toml"
    u.write_text('[epm]\nbase = "https://user"\n', encoding="utf-8")
    s, err = config.load_settings(_defaults(tmp_path), u)
    assert err is None
    assert s["epm"]["base"] == "https://user" and s["drill"]["channels"] == ["msedge", "chrome"]
    assert config.resolve({"SV_EPM_BASE": "https://env"}, s, "SV_EPM_BASE", "epm.base") == "https://env"
    assert config.resolve({}, s, "SV_EPM_BASE", "epm.base") == "https://user"
    assert config.resolve({}, s, None, "nope.x", "dflt") == "dflt"


def test_missing_user_file_is_fine_and_bad_toml_is_reported(tmp_path):
    s, err = config.load_settings(_defaults(tmp_path), tmp_path / "absent.toml")
    assert err is None and s["epm"]["base"] == "https://org"
    bad = tmp_path / "bad.toml"
    bad.write_text("[paths\nx = 1\n", encoding="utf-8")
    s, err = config.load_settings(_defaults(tmp_path), bad)
    assert "bad.toml" in err and s["epm"]["base"] == "https://org"


def test_find_first_prefers_configured_then_candidates(tmp_path):
    exe = tmp_path / "EXCEL.EXE"
    exe.write_text("")
    assert config.find_first(str(exe), None, []) == str(exe)
    assert config.find_first(None, None, [str(tmp_path / "nope.exe"), str(exe)]) == str(exe)
    assert config.find_first(None, None, [str(tmp_path / "nope.exe")]) is None


def test_knowledge_file_falls_back_to_enabled_pack(tmp_path, monkeypatch):
    core, packs = tmp_path / "knowledge", tmp_path / "packs"
    (packs / "p1" / "knowledge").mkdir(parents=True)
    core.mkdir()
    (core / "a.json").write_text("{}")
    (packs / "p1" / "knowledge" / "b.json").write_text("{}")
    monkeypatch.setattr(config, "KNOWLEDGE_DIR", core)
    monkeypatch.setattr(config, "PACKS_DIR", packs)
    monkeypatch.setattr(config, "PACKS", ["p1"])
    assert config.knowledge_file("a.json") == core / "a.json"
    assert config.knowledge_file("b.json") == packs / "p1" / "knowledge" / "b.json"
    monkeypatch.setattr(config, "PACKS", [])
    assert config.knowledge_file("b.json") == core / "b.json"


def test_shipped_defaults_parse_and_hold_org_values():
    s, err = config.load_settings(config.DEFAULTS_FILE, config.DEFAULTS_FILE.with_name("absent.toml"))
    assert err is None
    assert s["epm"]["base"].startswith("https://") and s["drill"]["channels"][0] == "msedge"
    assert s["label"]["id"] and s["http"]["port"] > 1024
