import json
import os
import subprocess
import winreg

import pytest

from smartass.ops import setup_user


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

def test_user_config_is_written_once_without_a_plain_secret(tmp_path):
    p = tmp_path / "config.toml"
    assert setup_user.write_user_config(p) == "created"
    first = p.read_text(encoding="utf-8")
    assert "[http]" in first and "\nsecret = " not in first and "[packs]" in first
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


def test_drop_legacy_skill_removes_link_keeps_target(tmp_path):
    src, old, backup = tmp_path / "src", tmp_path / "skills" / "smartview", tmp_path / "backup"
    src.mkdir()
    (src / "SKILL.md").write_text("x", encoding="utf-8")
    old.parent.mkdir()
    subprocess.run(["cmd", "/c", "mklink", "/J", str(old), str(src)], check=True, capture_output=True)
    assert setup_user.drop_legacy_skill(old, backup) == "old link removed"
    assert not old.exists() and (src / "SKILL.md").exists()
    assert setup_user.drop_legacy_skill(old, backup) == "none"
    old.mkdir()
    assert setup_user.drop_legacy_skill(old, backup) == "old copy moved to backup"
    assert not old.exists() and any(backup.iterdir())


def test_make_shortcut_creates_missing_startup_folder(tmp_path):
    lnk = tmp_path / "Microsoft" / "Windows" / "Startup" / "x.lnk"
    setup_user.make_shortcut(lnk, "cmd.exe", "/c exit", str(tmp_path))
    assert lnk.exists()


def test_register_mcp_drops_the_v2_side_by_side_entry(tmp_path):
    p = tmp_path / ".claude.json"
    p.write_text(json.dumps({"mcpServers": {
        "smartass": {"command": "py.exe", "args": ["run_server.py"], "env": {"SV_DRILL_KEEP_WARM": "0"}},
        "smartass2": {"command": "py.exe", "args": ["run_server_v2.py"], "env": {}}}}), encoding="utf-8")
    entry = {"type": "stdio", "command": "py.exe", "args": ["run_server.py"], "env": {}}
    assert setup_user.register_mcp(p, entry) == "updated"
    servers = json.loads(p.read_text(encoding="utf-8"))["mcpServers"]
    assert list(servers) == ["smartass"] and servers["smartass"]["env"] == {"SV_DRILL_KEEP_WARM": "0"}


def _replace_failing(monkeypatch, fails: int):
    """os.replace as setup_user calls it, refused with WinError 5 the first fails times (Windows holds a file it has
    just written for a moment: antivirus, the indexer), then the real one; the pauses recorded, not slept."""
    calls, pauses = [], []
    real = os.replace

    def replace(src, dst):
        calls.append(dst)
        if len(calls) <= fails:
            raise PermissionError(13, "Access is denied", str(dst))
        real(src, dst)
    monkeypatch.setattr(setup_user.os, "replace", replace)
    monkeypatch.setattr(setup_user.time, "sleep", pauses.append)
    return calls, pauses


def test_register_mcp_retries_a_replace_windows_refuses_for_a_moment(tmp_path, monkeypatch):
    """The flaky WinError 5 of test_register_mcp_drops_the_v2_side_by_side_entry: the replace is tried again briefly."""
    p = tmp_path / ".claude.json"
    p.write_text(json.dumps({"mcpServers": {}}), encoding="utf-8")
    calls, pauses = _replace_failing(monkeypatch, fails=2)
    assert setup_user.register_mcp(p, {"type": "stdio", "command": "py.exe", "args": [], "env": {}}) == "updated"
    assert len(calls) == 3 and len(pauses) == 2 and "smartass" in json.loads(p.read_text(encoding="utf-8"))["mcpServers"]


def test_register_mcp_gives_up_on_a_replace_that_stays_refused(tmp_path, monkeypatch):
    p = tmp_path / ".claude.json"
    p.write_text(json.dumps({"mcpServers": {}}), encoding="utf-8")
    calls, _pauses = _replace_failing(monkeypatch, fails=10_000)
    with pytest.raises(PermissionError):
        setup_user.register_mcp(p, {"type": "stdio", "command": "py.exe", "args": [], "env": {}})
    assert len(calls) == setup_user.REPLACE_TRIES and json.loads(p.read_text(encoding="utf-8")) == {"mcpServers": {}}


def test_setup_makes_no_sso_watcher_shortcut_and_removes_an_older_releases(tmp_path, monkeypatch):
    """Task 6: the Smart View SSO watcher is gone. Setup makes only the engine's Startup shortcut, removes the
    watcher's (and the svkit one) where an older release left it, and stops a watcher only when it is ours."""
    monkeypatch.setattr(setup_user, "STARTUP", tmp_path)
    monkeypatch.setattr(setup_user.config, "HOME", tmp_path / "home")
    made = []
    monkeypatch.setattr(setup_user, "make_shortcut", lambda lnk, *a: made.append(lnk.name))
    for name in ("smartass-sso-watcher.lnk", "svkit-sso-watcher.lnk"):
        (tmp_path / name).write_text("x")
    monkeypatch.setattr(setup_user, "_watchers", lambda: [])
    out = setup_user.startup_entries()
    assert made == ["smartass-engine.lnk"] and not list(tmp_path.glob("*sso-watcher.lnk"))
    assert out[0] == "startup: smartass-engine.lnk" and sum("removed" in s for s in out) == 2


def test_only_this_installs_watcher_is_stopped_another_is_left_and_said(monkeypatch):
    killed = []
    out = setup_user.stop_old_watcher(found=lambda: [(11, True), (22, False)], kill=killed.append)
    assert killed == [11]
    assert out == ["startup: stopped the old Smart View SSO watcher (pid 11)",
                   "startup: an SSO watcher of another smartass install (pid 22) is left running"]


def test_a_watcher_is_ours_only_when_its_command_line_names_this_installs_script(monkeypatch, tmp_path):
    import psutil

    class P:
        def __init__(self, pid, name, cmd):
            self.pid, self.info = pid, {"name": name, "cmdline": cmd}
    ours = str(tmp_path / "engine" / "sso_watcher.ps1")
    procs = [P(1, "powershell.exe", ["powershell.exe", "-File", ours.upper()]),
             P(2, "powershell.exe", ["powershell.exe", "-File", r"C:\other\engine\sso_watcher.ps1"]),
             P(3, "powershell.exe", ["powershell.exe", "-File", "x.ps1"]), P(4, "python.exe", ["sso_watcher.ps1"])]
    monkeypatch.setattr(psutil, "process_iter", lambda attrs: procs)
    assert setup_user._watchers(tmp_path / "engine" / "sso_watcher.ps1") == [(1, True), (2, False)]
