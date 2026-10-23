import json
import subprocess
import winreg

from smartview_mcp import setup_user


class FakeReg:
    def __init__(self, value, kind=winreg.REG_EXPAND_SZ):
        self.value, self.kind, self.writes, self.broadcasts = value, kind, [], 0

    def kw(self):
        return {"read": lambda: (self.value, self.kind), "write": self.write, "broadcast": self.bc}

    def write(self, value, kind):
        self.value, self.kind = value, kind
        self.writes.append((value, kind))

    def bc(self):
        self.broadcasts += 1


def test_scripts_dir_is_added_to_user_path_once(tmp_path):
    scripts = tmp_path / "Scripts"
    scripts.mkdir()
    (scripts / "smartass.exe").write_bytes(b"")
    reg = FakeReg(r"%USERPROFILE%\bin;C:\Tools;")
    assert setup_user.path_step(str(scripts / "python.exe"), **reg.kw()).startswith("path: added")
    assert reg.value == rf"%USERPROFILE%\bin;C:\Tools;{scripts}" and reg.kind == winreg.REG_EXPAND_SZ
    assert reg.broadcasts == 1
    upper = FakeReg(str(scripts).upper() + "\\")
    assert setup_user.path_step(str(scripts / "python.exe"), **upper.kw()).startswith("path: unchanged")
    assert setup_user.path_step(str(scripts / "python.exe"), **reg.kw()).startswith("path: unchanged")
    assert len(reg.writes) == 1 and not upper.writes
    empty = FakeReg("", winreg.REG_SZ)
    setup_user.path_step(str(scripts / "python.exe"), **empty.kw())
    assert empty.value == str(scripts) and empty.kind == winreg.REG_SZ


def test_path_step_honours_registry_override(tmp_path, monkeypatch):
    def no_registry(*a):
        raise AssertionError("the real HKCU Path must not be written")
    monkeypatch.setattr(winreg, "SetValueEx", no_registry)
    fake = tmp_path / "path.txt"
    monkeypatch.setenv("SMARTASS_USER_PATH_FILE", str(fake))
    scripts = tmp_path / "Scripts"
    scripts.mkdir()
    (scripts / "smartass.exe").write_text("")
    out = setup_user.path_step(str(scripts / "python.exe"))
    assert str(scripts) in fake.read_text(encoding="utf-8")
    assert "added" in out or "kept" in out
    assert setup_user.path_step(str(scripts / "python.exe")).startswith("path: unchanged")
    assert fake.read_text(encoding="utf-8") == str(scripts)


def test_path_step_skips_without_smartass_exe(tmp_path):
    reg = FakeReg("C:\\Tools")
    assert setup_user.path_step(str(tmp_path / "python.exe"), **reg.kw()).startswith("path: skipped")
    assert not reg.writes

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
                                            "smartass": {"command": "old", "env": {"K": "1"}}}}), encoding="utf-8")
    entry = {"type": "local", "command": "py.exe", "args": ["run_server.py"], "env": {}, "tools": ["*"]}
    assert setup_user.register_mcp(p, entry) == "updated"
    d = json.loads(p.read_text(encoding="utf-8"))
    assert d["mcpServers"]["other"] == {"command": "x"}
    assert d["mcpServers"]["smartass"]["command"] == "py.exe" and d["mcpServers"]["smartass"]["env"] == {"K": "1"}
    assert (tmp_path / "mcp-config.json.bak-smartass").exists()
    assert setup_user.register_mcp(p, entry) == "unchanged"
    assert setup_user.register_mcp(tmp_path / "absent.json", entry) == "missing"
    assert setup_user.register_mcp(tmp_path / "new.json", entry, create=True) == "updated"
    assert setup_user.register_mcp(tmp_path / "nodir" / "x.json", entry, create=True) == "missing"


def test_register_mcp_renames_legacy_smartview_entry(tmp_path):
    p = tmp_path / "mcp-config.json"
    p.write_text(json.dumps({"mcpServers": {"smartview": {"command": "old", "env": {"K": "1"}}}}), encoding="utf-8")
    entry = {"type": "local", "command": "py.exe", "args": ["run_server.py"], "env": {}, "tools": ["*"]}
    assert setup_user.register_mcp(p, entry) == "updated"
    servers = json.loads(p.read_text(encoding="utf-8"))["mcpServers"]
    assert list(servers) == ["smartass"] and servers["smartass"]["env"] == {"K": "1"}
    assert setup_user.register_mcp(p, entry) == "unchanged"


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


def test_make_shortcut_creates_missing_startup_folder(tmp_path):
    lnk = tmp_path / "Microsoft" / "Windows" / "Startup" / "x.lnk"
    setup_user.make_shortcut(lnk, "cmd.exe", "/c exit", str(tmp_path))
    assert lnk.exists()
