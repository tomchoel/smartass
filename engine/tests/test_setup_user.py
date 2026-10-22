import json
import subprocess

from smartview_mcp import setup_user


def test_user_config_is_written_once_with_a_secret(tmp_path):
    p = tmp_path / "config.toml"
    assert setup_user.write_user_config(p) == "created"
    first = p.read_text(encoding="utf-8")
    assert "[http]" in first and "secret = " in first and "[packs]" in first
    assert setup_user.write_user_config(p) == "kept"
    assert p.read_text(encoding="utf-8") == first


def test_register_mcp_creates_updates_and_keeps_env(tmp_path):
    p = tmp_path / "mcp-config.json"
    p.write_text(json.dumps({"mcpServers": {"other": {"command": "x"},
                                            "smartview": {"command": "old", "env": {"K": "1"}}}}), encoding="utf-8")
    entry = {"type": "local", "command": "py.exe", "args": ["run_server.py"], "env": {}, "tools": ["*"]}
    assert setup_user.register_mcp(p, entry) == "updated"
    d = json.loads(p.read_text(encoding="utf-8"))
    assert d["mcpServers"]["other"] == {"command": "x"}
    assert d["mcpServers"]["smartview"]["command"] == "py.exe" and d["mcpServers"]["smartview"]["env"] == {"K": "1"}
    assert (tmp_path / "mcp-config.json.bak-smartass").exists()
    assert setup_user.register_mcp(p, entry) == "unchanged"
    assert setup_user.register_mcp(tmp_path / "absent.json", entry) == "missing"
    assert setup_user.register_mcp(tmp_path / "new.json", entry, create=True) == "updated"
    assert setup_user.register_mcp(tmp_path / "nodir" / "x.json", entry, create=True) == "missing"


def test_link_skill_replaces_dir_and_old_junction(tmp_path):
    src, dst, backup = tmp_path / "src", tmp_path / "skills" / "smartview", tmp_path / "backup"
    src.mkdir()
    (src / "SKILL.md").write_text("new", encoding="utf-8")
    dst.mkdir(parents=True)
    (dst / "SKILL.md").write_text("old", encoding="utf-8")
    assert setup_user.link_skill(src, dst, backup) == "linked (old copy moved to backup)"
    assert (dst / "SKILL.md").read_text(encoding="utf-8") == "new"
    assert any(backup.iterdir())
    assert setup_user.link_skill(src, dst, backup) == "unchanged"
    other = tmp_path / "other"
    other.mkdir()
    subprocess.run(["cmd", "/c", "rmdir", str(dst)], check=True)
    subprocess.run(["cmd", "/c", "mklink", "/J", str(dst), str(other)], check=True, capture_output=True)
    assert setup_user.link_skill(src, dst, backup) == "linked"
    assert (dst / "SKILL.md").exists()
