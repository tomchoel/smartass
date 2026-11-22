"""What the HTTP API and the Excel task pane of an older release (up to 0.7.0) left on a PC (light-fast phase 4 Task 1):
`smartass setup` removes the Startup entry smartass-engine.lnk and stops a running `smartass serve` whose command line
proves it is this install's; `smartass doctor` lists those, the secret file, a plain [http] secret in config.toml and
the pane's registration; `doctor --fix` removes this install's own files and Startup entries, each said first, and
only REPORTS what lives in Office's own settings (the trusted catalogue, Excel's add-in list). Registry, Startup folder
and process list are fakes here (tests/conftest.py keeps the suite off this PC's)."""
import re
import types
from pathlib import Path

import psutil
import pytest

from smartass.core import config
from smartass.ops import doctor, pane_leftovers, setup_user

VENV = r"C:\work\smartass\engine\.venv\Scripts"
CATALOGUE_URL = r"\\files.example\home\someone\TNClaudeCatalog"
XLL = r"C:\Users\someone\AppData\Local\svkit\addin\dna\TNClaude64.xll"


def _proc(pid, cmd, ppid=1, name=None, running=True, exits=False, readable=True):
    """A process as _process_list gives it: its name in info; the command line and the parent only when asked for,
    and only of a Python or smartass.exe (reading them for every process is what made the real list slow).
    running: False once its pid belongs to another process (psutil's is_running on the listed object); exits: it ends
    within the wait (a launcher whose child was stopped); readable: False when its command line is refused."""
    exe = name or re.split(r"[\\/]", cmd[0])[-1]

    def python_only(value):
        def read():
            assert exe.lower() in ("python.exe", "pythonw.exe", "smartass.exe"), "read for a process that is not ours"
            if not readable:
                raise psutil.AccessDenied(pid)
            return value
        return read

    def wait(timeout=None):
        if exits:
            return 0
        raise psutil.TimeoutExpired(timeout, pid)
    return types.SimpleNamespace(pid=pid, info={"name": exe}, cmdline=python_only(cmd), ppid=python_only(ppid),
                                 is_running=lambda: running, wait=wait)


def _lnk(target: str) -> bytes:
    """A shortcut's bytes as Windows writes them: its target path in UTF-16, at some offset."""
    return b"\x4c\x00\x00\x00" + target.encode("utf-16-le") + b"\x00\x00"


OUR_LNK = _lnk(VENV + r"\pythonw.exe")


def _serve(exe, pid, ppid=1, **kw):
    return _proc(pid, [exe, "-m", "smartass.ops.cli", "serve"], ppid, **kw)


@pytest.fixture
def pc(monkeypatch, tmp_path):
    """A PC of the test's own: home, config, old svkit home, Startup folder, process list and Office's settings."""
    home, startup = tmp_path / "home", tmp_path / "Startup"
    home.mkdir()
    startup.mkdir()
    for name in ("HOME", "SVKIT"):
        monkeypatch.setattr(config, name, home)
    monkeypatch.setattr(config, "USER_CONFIG", home / "config.toml")
    monkeypatch.setattr(config, "LEGACY_HOME", None)
    monkeypatch.setattr(setup_user, "STARTUP", startup)
    monkeypatch.setattr(setup_user, "VENV_SCRIPTS", Path(VENV))
    state = types.SimpleNamespace(home=home, startup=startup, procs=[], killed=[], subkeys={}, options={})
    monkeypatch.setattr(setup_user, "_process_list", lambda: state.procs)
    monkeypatch.setattr(setup_user, "_kill", lambda p: state.killed.append(p.pid))
    monkeypatch.setattr(pane_leftovers, "_hkcu_subkeys", lambda path: state.subkeys)
    monkeypatch.setattr(pane_leftovers, "_hkcu_values", lambda path: state.options)
    return state


# ---------------------------------------------------------------- setup


def test_setup_makes_no_startup_entry_and_no_secret_any_more():
    assert not hasattr(setup_user, "make_shortcut") and not hasattr(setup_user, "secret_step")
    assert "[http]" not in setup_user.CONFIG_TEMPLATE and "secret" not in setup_user.CONFIG_TEMPLATE


def test_setup_removes_the_engines_startup_entry_and_says_so(pc):
    (pc.startup / "smartass-engine.lnk").write_bytes(OUR_LNK)
    (pc.startup / "other-tool.lnk").write_bytes(b"lnk")
    out = setup_user.startup_entries()
    assert not (pc.startup / "smartass-engine.lnk").exists() and (pc.startup / "other-tool.lnk").exists()
    assert out == ["startup: removing smartass-engine.lnk (it started `smartass serve`, the HTTP API, which is gone)"]


def test_another_installs_startup_entry_is_left_and_said(pc):
    """Review 10: the entry is this install's only when its target names this install's venv."""
    lnk = pc.startup / "smartass-engine.lnk"
    lnk.write_bytes(_lnk(r"C:\other\smartass\engine\.venv\Scripts\pythonw.exe"))
    out = setup_user.startup_entries()
    assert lnk.exists() and out == ["startup: smartass-engine.lnk starts another smartass install's `smartass serve` "
                                    "(it does not name this install's .venv): left as it is"]
    c = doctor._leftovers()
    assert "smartass-engine.lnk of another smartass install" in c.detail
    assert doctor.fix_leftovers(say=lambda s: None) == [] and lnk.exists()


def test_setup_with_nothing_left_says_there_is_nothing_to_do(pc):
    assert setup_user.startup_entries() == ["startup: nothing to remove"]


def test_a_serve_is_ours_only_when_its_command_line_names_this_installs_venv(pc):
    pc.procs = [
        _serve(VENV + r"\pythonw.exe", 10),                                          # the venv's launcher ...
        _serve(r"C:\Users\someone\AppData\Roaming\uv\python\cpython-3.12\pythonw.exe", 11, ppid=10),  # its child
        _serve(VENV.upper() + r"\PYTHON.EXE", 12),                                   # a lone one, any case
        _serve(r"C:\other\smartass\engine\.venv\Scripts\pythonw.exe", 20),         # another install's
        _serve(r"C:\Python312\pythonw.exe", 21, ppid=999),                          # names no venv at all
        _serve(VENV + r"2\pythonw.exe", 22),                                         # a look-alike path
        _proc(30, [VENV + r"\pythonw.exe", "-m", "smartass.ops.cli", "doctor"]),    # ours, but not serve
        _proc(31, ["node.exe", "server.mjs"]),
        _proc(32, ["cmd.exe", "/c", "smartass.ops.cli", "serve"]),                  # not a Python
        _proc(13, [VENV + r"\smartass.exe", "serve"]),                                # the console script (review 4)
        _proc(14, [r"C:\uv\python\python.exe", VENV + r"\smartass.exe", "serve"], ppid=13),   # ... its child
        _proc(23, [r"C:\Python312\python.exe", "-m", "smartass.ops.cli", "serve", "--log", VENV + r"\x"]),  # review 3
        _proc(24, [r"C:\other\.venv\Scripts\smartass.exe", "serve"]),             # another install's console script
        _proc(25, [VENV + r"\python.exe", "-m", "smartass.ops.cli", "serve"], readable=False),   # review 5
    ]
    found, unreadable = setup_user._serves()
    pids = [s.proc.pid for s in found]
    assert sorted((s.proc.pid, s.ours) for s in found) == [(10, True), (11, True), (12, True), (13, True), (14, True),
                                                            (20, False), (21, False), (22, False), (23, False),
                                                            (24, False)]
    assert pids.index(11) < pids.index(10) and pids.index(14) < pids.index(13)   # the child before its launcher
    assert unreadable == 1
    assert next(s for s in found if s.proc.pid == 10).children == (11,)


def test_setup_stops_this_installs_serve_saying_so_first_and_leaves_another(pc):
    pc.procs = [_serve(VENV + r"\pythonw.exe", 10), _serve(r"C:\other\.venv\Scripts\pythonw.exe", 20)]
    said = []
    setup_user.stop_old_serve(lambda s: said.append((s, list(pc.killed))))
    assert pc.killed == [10]
    assert said == [("startup: stopping this install's old `smartass serve` (pid 10): the HTTP API is gone", []),
                    ("startup: a `smartass serve` of another smartass install (pid 20) is left running", [10])]


def test_a_serve_that_ended_meanwhile_or_cannot_be_stopped_is_said(pc):
    def kill(p):
        raise (psutil.NoSuchProcess(p.pid) if p.pid == 10 else psutil.AccessDenied(p.pid))
    said = []
    found = [setup_user.Serve(_serve(VENV + r"\pythonw.exe", 10), True, ()),
             setup_user.Serve(_serve(VENV + r"\pythonw.exe", 11), True, ())]
    setup_user.stop_old_serve(said.append, found=lambda: (found, 0), kill=kill)
    assert "(pid 10) had already ended" in said[1]
    assert "(pid 11) could not be stopped (AccessDenied)" in said[3] and "next logoff" in said[3]


def test_a_pid_reused_between_the_listing_and_the_stop_is_never_killed(pc):
    """Review 2: the stop goes to the process object the listing found; once its pid belongs to another process
    (is_running is False on that object) nothing is killed."""
    pc.procs = [_serve(VENV + r"\pythonw.exe", 10, running=False)]
    said = []
    setup_user.stop_old_serve(said.append)
    assert pc.killed == [] and said == ["startup: the old `smartass serve` (pid 10) had already ended"]


@pytest.mark.parametrize("exits, killed", [(True, [11]), (False, [11, 10])])
def test_a_launcher_that_ends_with_its_child_is_not_killed(pc, exits, killed):
    pc.procs = [_serve(VENV + r"\pythonw.exe", 10, exits=exits),
                _serve(r"C:\Users\someone\AppData\Roaming\uv\python\cpython-3.12\pythonw.exe", 11, ppid=10)]
    said = []
    setup_user.stop_old_serve(said.append)
    assert pc.killed == killed
    assert any("(pid 10) ended with it" in s for s in said) is exits


def test_a_python_whose_command_line_cannot_be_read_is_said(pc):
    pc.procs = [_proc(25, [VENV + r"\python.exe", "-m", "smartass.ops.cli", "serve"], readable=False)]
    said = []
    setup_user.stop_old_serve(said.append)
    assert pc.killed == [] and said == ["startup: the command line of 1 Python process(es) could not be read: an old "
                                        "`smartass serve` among them was not looked at"]
    assert "could not be read" in doctor._leftovers().detail


def test_setup_stops_the_serve_and_removes_the_entry_in_one_go(pc):
    (pc.startup / "smartass-engine.lnk").write_bytes(OUR_LNK)
    pc.procs = [_serve(VENV + r"\pythonw.exe", 10)]
    out = setup_user.startup_entries()
    assert pc.killed == [10] and len(out) == 2 and not (pc.startup / "smartass-engine.lnk").exists()


# ---------------------------------------------------------------- the pane's registration (read only)


def test_the_panes_catalogue_is_found_by_its_folder_name_only_and_no_folder_is_read(pc, tmp_path, monkeypatch):
    """A catalogue is the pane's by its Url's last folder (TNClaudeCatalog, any case or separator). No catalogue's
    folder is read: a network share that does not answer would make every doctor run (and fin_status) wait."""
    shared = tmp_path / "shared"
    shared.mkdir()
    (shared / "manifest.xml").write_text("<OfficeApp><Id>6f3c9a52-2b1e-4d7a-9c41-7a8e5d2f1b60</Id></OfficeApp>")
    pc.subkeys = {"{AAAA}": {"Id": "{AAAA}", "Url": CATALOGUE_URL + "\\", "Flags": 1},
                  "{BBBB}": {"Id": "{BBBB}", "Url": str(shared), "Flags": 1},        # the pane's manifest, another name
                  "{CCCC}": {"Id": "{CCCC}", "Url": "https://files.example/x/tnclaudecatalog/", "Flags": 1},
                  "{DDDD}": {"Id": "{DDDD}", "Url": r"\\unreachable.example\share\Catalog", "Flags": 1},
                  "{EEEE}": {"Id": "{EEEE}"}}

    def no_read(*args, **kwargs):
        raise AssertionError("a catalogue's folder was read")
    with monkeypatch.context() as m:
        for name in ("glob", "iterdir", "read_text", "read_bytes", "exists", "is_dir", "stat"):
            m.setattr(Path, name, no_read)
        m.setattr("builtins.open", no_read)
        found = pane_leftovers.catalogues()
    assert found == [CATALOGUE_URL + "\\", "https://files.example/x/tnclaudecatalog/"]
    assert not hasattr(pane_leftovers, "MANIFEST_ID")


def test_the_panes_classic_add_in_is_found_in_excels_open_values(pc):
    pc.options = {"OPEN": '/R "C:\\Program Files\\Other\\other.xll"', "OPEN1": "other.xlam",
                  "OPEN2": f'/R "{XLL}"', "Options": 0, "OPENX": f'"{XLL}"'}
    assert pane_leftovers.excel_addins() == [("OPEN2", XLL)]


PANE_VBS = ("' Starts the TN Claude pane server without a window (for the Startup folder).\n"
            'Set sh = CreateObject("WScript.Shell")\nsh.Run "cmd /c node server.mjs", 0, False\n')   # 0.7.0's script
PANE_CMD = "@echo off\nrem TN Claude task pane - local server (https://localhost:3000).\ncd /d \"%~dp0\"\nnode server.mjs\n"


def test_a_pane_startup_entry_is_found_by_what_it_names(pc):
    (pc.startup / "pane.lnk").write_bytes(b"\x4c\x00" + "C:\\x\\start-claude-pane-hidden.vbs".encode("utf-16-le"))
    (pc.startup / "folder.lnk").write_bytes(b"\x4c" + "C:\\dev\\addin\\tn-claude\\run.cmd".encode("utf-16-le"))
    (pc.startup / "Claude.vbs").write_text(PANE_VBS)
    (pc.startup / "Claude.cmd").write_text(PANE_CMD)
    (pc.startup / "smartass-engine.lnk").write_bytes("pythonw.exe -m smartass.ops.cli serve".encode("utf-16-le"))
    (pc.startup / "other.lnk").write_bytes(b"\x00" + "C:\\tools\\sync.exe".encode("utf-16-le"))
    (pc.startup / "my-node-app.cmd").write_text("cd /d C:\\work\\my-app\nnode server.mjs\n")   # another Node project
    found = [p.name for p in pane_leftovers.startup_entries(pc.startup)]
    assert found == ["Claude.cmd", "Claude.vbs", "folder.lnk", "pane.lnk"]


def test_nothing_writes_under_offices_settings():
    """Ruling of 22.11: the trusted catalogue and Excel's add-in list are Office's own state. The doctor reads them;
    no code of the engine can write or delete there (the user removes them in Excel)."""
    src = Path(pane_leftovers.__file__).read_text(encoding="utf-8") + Path(doctor.__file__).read_text(encoding="utf-8")
    assert not any(w in src for w in ("DeleteKey", "DeleteValue", "SetValue", "CreateKey", "KEY_SET_VALUE",
                                      "KEY_WRITE", "KEY_ALL_ACCESS"))


REAL = {"values": pane_leftovers._hkcu_values, "subkeys": pane_leftovers._hkcu_subkeys,
        "processes": setup_user._process_list}                 # taken at import, before conftest's fakes are set


def test_the_real_readers_read_without_changing_anything(monkeypatch):
    """The readers conftest fakes in every other test, run once on neutral ground: HKCU values and subkeys of keys
    that are not smartass's (read only), a missing key, and this process in the process list."""
    import os
    monkeypatch.setattr(pane_leftovers, "_hkcu_values", REAL["values"])
    env = REAL["values"]("Environment")
    assert isinstance(env, dict) and all(isinstance(k, str) for k in env)
    assert REAL["values"](r"Software\smartass-test\no such key") == {}
    assert REAL["subkeys"](r"Software\smartass-test\no such key") == {}
    panel = REAL["subkeys"]("Control Panel")
    assert panel and all(isinstance(v, dict) for v in panel.values())
    me = [p for p in REAL["processes"]() if p.pid == os.getpid()]
    assert len(me) == 1 and me[0].info["name"].lower().startswith("python") and me[0].ppid() == os.getppid()
    assert any("pytest" in c or "py.test" in c for c in me[0].cmdline())


# ---------------------------------------------------------------- doctor


def test_the_http_api_check_is_gone():
    assert "_http" not in [f.__name__ for f in doctor.CHECKS]


def test_a_pc_with_nothing_left_is_ok(pc):
    c = doctor._leftovers()
    assert c.status == "OK" and doctor.fix_leftovers(say=lambda s: None) == []


def test_doctor_lists_everything_the_http_api_and_the_pane_left(pc):
    (pc.startup / "smartass-engine.lnk").write_bytes(OUR_LNK)
    (pc.startup / "Claude.vbs").write_text(PANE_VBS)
    (pc.home / "secret.bin").write_bytes(b"DPAPI")
    secret = "plain-secret-" + "x" * 12
    (pc.home / "config.toml").write_text(f'[paths]\nanalyser = "C:/a"\n[http]\nport = 1\nsecret = "{secret}"\n')
    pc.procs = [_serve(VENV + r"\pythonw.exe", 10), _serve(r"C:\other\.venv\Scripts\pythonw.exe", 20)]
    pc.subkeys = {"{AAAA}": {"Url": CATALOGUE_URL}}
    pc.options = {"OPEN2": f'/R "{XLL}"'}
    c = doctor._leftovers()
    assert c.status == "WARN" and secret not in c.detail and secret not in c.fix
    for said in ("smartass-engine.lnk", "pid 10", "pid 20", "another smartass install", str(pc.home / "secret.bin"),
                 "plain [http] secret", "Claude.vbs", CATALOGUE_URL, XLL, "OPEN2"):
        assert said in c.detail, said
    assert "smartass doctor --fix" in c.fix and "Trusted Add-in Catalogs" in c.fix and "Manage: Excel Add-ins" in c.fix
    assert "renumbers" in c.fix and str(Path(XLL).parent) in c.fix


def test_office_settings_alone_are_reported_with_the_manual_steps_and_no_fix(pc):
    pc.subkeys = {"{AAAA}": {"Url": CATALOGUE_URL}}
    c = doctor._leftovers()
    assert c.status == "WARN" and CATALOGUE_URL in c.detail and "--fix" not in c.fix


def test_doctor_fix_removes_this_installs_files_and_entries_each_said_first(pc):
    lnk, vbs, sec = pc.startup / "smartass-engine.lnk", pc.startup / "Claude.vbs", pc.home / "secret.bin"
    lnk.write_bytes(OUR_LNK)
    vbs.write_text(PANE_VBS)
    sec.write_bytes(b"DPAPI")
    secret = "plain-secret-" + "y" * 12
    cfg = pc.home / "config.toml"
    cfg.write_text(f'[paths]\nanalyser = "C:/a"\n\n[http]\nport = 1\nsecret = "{secret}"\n\n[packs]\nenabled = []\n')
    pc.procs = [_serve(VENV + r"\pythonw.exe", 10)]
    pc.subkeys = {"{AAAA}": {"Url": CATALOGUE_URL}}
    pc.options = {"OPEN2": f'/R "{XLL}"'}
    said = []

    def say(s):
        said.append(s)
        assert pc.killed == [] or "stopping" not in s           # the stop is said before it is done
    removed = doctor.fix_leftovers(say=say)
    assert sorted(removed) == sorted([lnk, vbs, sec]) and not lnk.exists() and not vbs.exists() and not sec.exists()
    assert pc.killed == [10]
    assert cfg.read_text() == '[paths]\nanalyser = "C:/a"\n\n[packs]\nenabled = []\n'       # the section whole (Task 2)
    text = "\n".join(said)
    assert secret not in text
    for what in (f"removing {vbs}", f"removing the HTTP API's secret file {sec}", "removing the plain [http] secret",
                 "stopping this install's old `smartass serve` (pid 10)", "removing smartass-engine.lnk"):
        assert what in text, what
    assert "Trusted Add-in Catalogs" in text and "TNClaude" in text            # left for the user, and said
    left = doctor._leftovers()
    assert CATALOGUE_URL in left.detail and "secret" not in left.detail and "smartass-engine" not in left.detail


def test_a_plain_secret_is_read_from_the_file_so_the_check_is_clean_right_after_the_fix(pc):
    cfg = pc.home / "config.toml"
    cfg.write_text('[http]\nsecret = "abcdefghijklmnop"\n')
    assert doctor._leftovers().status == "WARN"
    doctor.fix_leftovers(say=lambda s: None)
    assert doctor._leftovers().status == "OK" and cfg.read_text() == ""


def test_the_fix_takes_the_secret_only_from_the_http_section(pc):
    cfg = pc.home / "config.toml"
    cfg.write_text('[other]\nsecret = "keep-me"\n\n[http]\nsecret = "drop-me"\n')
    doctor.fix_leftovers(say=lambda s: None)
    assert cfg.read_text() == '[other]\nsecret = "keep-me"\n\n'


def test_a_secret_the_fix_cannot_take_out_is_said_to_be_left(pc):
    """A secret written as a dotted key outside the [http] section is a plain secret to TOML, but not a line of that
    section: the fix does not claim it went."""
    cfg, secret = pc.home / "config.toml", "plain-secret-" + "z" * 12
    cfg.write_text(f'http.secret = "{secret}"\n')
    said = []
    doctor.fix_leftovers(say=said.append)
    assert cfg.read_text() == f'http.secret = "{secret}"\n' and doctor._leftovers().status == "WARN"
    assert any("delete it by hand" in s for s in said) and secret not in " ".join(said)


def test_a_broken_config_is_not_touched_by_the_fix(pc):
    cfg = pc.home / "config.toml"
    cfg.write_text('[http]\nsecret = "abc\n')                                # not TOML: the config check says so
    assert doctor._leftovers().status == "OK"
    doctor.fix_leftovers(say=lambda s: None)
    assert cfg.read_text() == '[http]\nsecret = "abc\n'


# ---------------------------------------------------------------- fin_status stays cheap (review 6)


def test_fin_status_makes_no_process_listing_and_reads_no_startup_file(pc, monkeypatch):
    """The leftovers check (a process scan, every Startup file's bytes) runs in `smartass doctor` only: fin_status's
    doctor summary leaves it out."""
    import asyncio

    from smartass import server

    def never(*args, **kwargs):
        pytest.fail("fin_status looked for an older release's leftovers")
    monkeypatch.setattr(setup_user, "_process_list", never)
    monkeypatch.setattr(pane_leftovers, "startup_entries", never)
    assert doctor._leftovers not in doctor.CHECKS
    monkeypatch.setattr(doctor, "CHECKS", tuple(f for f in doctor.CHECKS if f is not doctor._epm))   # no network
    (pc.home / "secret.bin").write_bytes(b"DPAPI")                     # a leftover the CLI's doctor would list
    out = asyncio.run(server.fin_status())
    assert "leftovers" not in [p["name"] for p in out["doctor"]["problems"]]


def test_the_doctor_command_and_the_bundle_include_the_leftovers(pc, monkeypatch, capsys):
    from smartass.ops import cli
    (pc.home / "secret.bin").write_bytes(b"DPAPI")
    monkeypatch.setattr(doctor, "CHECKS", ())
    assert cli.main(["doctor"]) == 0
    assert "leftovers" in capsys.readouterr().out
    assert [c.name for c in doctor.run(leftovers=True)] == ["leftovers"] and doctor.run() == []
