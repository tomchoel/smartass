"""Per-user setup (`smartass setup`, the second of the README's three install commands). Idempotent: every step reports
created/updated/unchanged. Never needs admin rights. What an older release started at logon (the Smart View SSO
watcher, the HTTP API `smartass serve`) is removed and stopped here when it is provably this install's."""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import NamedTuple

from ..core import config, retry

SERVER_NAME = "smartass"
LEGACY_SERVER_NAMES = ("smartview", "smartass2")   # smartview renamed 2026-10-23, smartass2 (v2 beside v1) retired 2026-11-04
USER = Path(os.environ.get("USERPROFILE", str(Path.home())))
STARTUP = Path(os.environ.get("APPDATA", str(USER / "AppData" / "Roaming"))) / \
    "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"

CONFIG_TEMPLATE = """# smartass user settings. Defaults live in engine/defaults.toml - only override what you need.

[paths]
# analyser = 'C:\\Users\\you\\OneDrive - Telenor\\analyser'   # where answer workbooks go
# excel_exe = 'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE'

[packs]
enabled = []            # e.g. ["produktregnskap"]
"""

def write_user_config(path: Path = config.USER_CONFIG) -> str:
    if path.exists():
        return "kept"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(CONFIG_TEMPLATE, encoding="utf-8")
    return "created"


def mcp_entries(python: Path, relay: Path) -> dict:
    return {
        "copilot": {"type": "local", "command": str(python), "args": [str(relay)], "env": {}, "tools": ["*"]},
        "claude": {"type": "stdio", "command": str(python), "args": [str(relay)], "env": {}},
    }


def register_mcp(path: Path, entry: dict, create: bool = False) -> str:
    """Set mcpServers.smartass in an agent config (a legacy `smartview` entry is replaced). Other servers and keys
    are kept, an existing `env` is kept, the first original is backed up once, and the file is replaced atomically."""
    if not path.exists():
        if not (create and path.parent.exists()):
            return "missing"
        path.write_text("{}", encoding="utf-8")
    data = json.loads(path.read_text(encoding="utf-8") or "{}")
    servers = data.setdefault("mcpServers", {})
    legacy = [n for n in LEGACY_SERVER_NAMES if n in servers]
    old = servers.get(SERVER_NAME) or next((servers[n] for n in legacy if servers[n]), {})
    new = {**entry, "env": old.get("env") or entry.get("env", {})}
    if servers.get(SERVER_NAME) == new and not legacy:
        return "unchanged"
    bak = path.with_name(path.name + ".bak-smartass")
    if not bak.exists():
        shutil.copy2(path, bak)
    for n in legacy:
        del servers[n]
    servers[SERVER_NAME] = new
    tmp = path.with_name(path.name + ".tmp-smartass")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    _replace(tmp, path)
    return "updated"


REPLACE_TRIES, REPLACE_PAUSE_S = 10, 0.1    # Windows can hold a file it has just written for a moment (WinError 5)


def _replace(tmp: Path, path: Path) -> None:
    """os.replace, tried again briefly while Windows refuses it (PermissionError: antivirus or the indexer holds the
    file just written); after REPLACE_TRIES the error is raised and the temp file removed (path is unchanged)."""
    try:
        retry.tried_again(lambda: os.replace(tmp, path), REPLACE_TRIES, REPLACE_PAUSE_S)
    except PermissionError:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def _is_junction(p: Path) -> bool:
    try:
        return p.is_junction() or p.is_symlink()
    except OSError:
        return False


def link_skill(src: Path, dst: Path, backup_root: Path) -> str:
    """dst becomes a junction to src. A real folder is moved to backup_root (NOT next to the skills, or agents
    would load it as a second skill)."""
    if not src.is_dir():
        raise FileNotFoundError(f"skill source not found: {src}")
    if _is_junction(dst) and os.path.realpath(dst) == os.path.realpath(src):
        return "unchanged"
    note = ""
    if _is_junction(dst):
        os.rmdir(dst)  # removes only the link, never the target's content
    elif dst.exists():
        backup_root.mkdir(parents=True, exist_ok=True)
        shutil.move(str(dst), str(backup_root / f"{dst.name}-{time.strftime('%Y%m%d-%H%M%S')}"))
        note = " (old copy moved to backup)"
    dst.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(["cmd", "/c", "mklink", "/J", str(dst), str(src)], check=True, capture_output=True)
    return "linked" + note


def drop_legacy_skill(dst: Path, backup_root: Path) -> str:
    """Removes the skill under its old name ('smartview'), so agents do not load it twice. A link is just removed;
    a real folder is moved to backup_root."""
    if _is_junction(dst):
        os.rmdir(dst)  # removes only the link, never the target's content
        return "old link removed"
    if dst.exists():
        backup_root.mkdir(parents=True, exist_ok=True)
        shutil.move(str(dst), str(backup_root / f"{dst.name}-{time.strftime('%Y%m%d-%H%M%S')}"))
        return "old copy moved to backup"
    return "none"


LEGACY_STARTUP = ("smartass-sso-watcher.lnk", "svkit-sso-watcher.lnk")   # the SSO watcher's, removed (Task 6)
WATCHER_PS1 = config.ENGINE_DIR / "sso_watcher.ps1"      # an older release's Smart View SSO watcher (deleted)


ENGINE_STARTUP = "smartass-engine.lnk"      # an older release's (up to 0.7.0): `smartass serve` at logon (removed)
VENV_SCRIPTS = config.ENGINE_DIR / ".venv" / "Scripts"   # this install's interpreters


def startup_entries() -> list[str]:
    """smartass starts nothing at logon any more (the agent's MCP client starts the server): what an older release
    put in the Startup folder is removed, and what it started is stopped when it is provably this install's."""
    out = []
    for name in LEGACY_STARTUP:
        legacy = STARTUP / name
        if legacy.exists():
            legacy.unlink()
            out.append(f"startup: removed {name} (the Smart View SSO watcher is gone)")
    drop_engine_startup(out.append)
    stop_old_serve(out.append)
    out += stop_old_watcher()
    return out or ["startup: nothing to remove"]


def _names(data: bytes, text: str) -> bool:
    """True when a file's bytes carry text (any case): a shortcut keeps its target as UTF-16 at any offset, and
    sometimes as ANSI too."""
    texts = [data.decode("latin-1")]
    for start in (0, 1):
        part = data[start:]
        texts.append(part[:len(part) // 2 * 2].decode("utf-16-le", errors="ignore"))
    return text.lower() in " ".join(texts).lower()


def engine_startup() -> tuple[Path, bool] | None:
    """(the Startup entry smartass-engine.lnk, ours) when there is one: ours when its target names THIS install's venv
    Scripts folder (every release made it start that folder's pythonw.exe); None when there is none."""
    lnk = STARTUP / ENGINE_STARTUP
    if not lnk.exists():
        return None
    try:
        data = lnk.read_bytes()
    except OSError:
        return lnk, False
    return lnk, _names(data, str(VENV_SCRIPTS).rstrip("\\") + "\\")


def drop_engine_startup(say) -> list[Path]:
    """The Startup entry smartass-engine.lnk an older release made (it started `smartass serve`), removed after it is
    said when it is this install's (engine_startup); another install's is left, and said. [the file] when it went."""
    found = engine_startup()
    if found is None:
        return []
    lnk, ours = found
    if not ours:
        say(f"startup: {ENGINE_STARTUP} starts another smartass install's `smartass serve` (it does not name this "
            "install's .venv): left as it is")
        return []
    say(f"startup: removing {ENGINE_STARTUP} (it started `smartass serve`, the HTTP API, which is gone)")
    try:
        lnk.unlink()
    except OSError as e:
        say(f"startup: could not remove {lnk} ({e.strerror or e}); delete it by hand")
        return []
    return [lnk]


SERVE_NAMES = ("python.exe", "pythonw.exe", "smartass.exe")   # what runs a `serve`: a Python, or the console script
LAUNCHER_WAIT_S = 2.0          # how long a launcher whose child was stopped may take to end with it


def _process_list():
    """The running processes with their name only: a command line, and a serve's parent, are read one by one, only
    for SERVE_NAMES (reading every process's parent took over three minutes on the owner's PC, 525 processes, 22.11).
    The objects are kept: psutil's Process carries its creation time, so a stop through it never hits a pid that a
    new process has taken since."""
    import psutil
    return psutil.process_iter(["name"])


def _is_serve(cmd: list[str]) -> bool:
    """`python -m smartass.ops.cli serve`, or the console script `smartass(.exe) serve` (any folder)."""
    low = [c.lower() for c in cmd]
    return any(b == "serve" and (a == "smartass.ops.cli" or re.split(r"[\\/]", a)[-1] in ("smartass", "smartass.exe"))
               for a, b in zip(low, low[1:]))


class Serve(NamedTuple):
    proc: object                  # the psutil.Process the listing found (its identity includes the creation time)
    ours: bool
    children: tuple[int, ...]     # the pids of the serves it launched


def _serves(scripts: Path | None = None) -> tuple[list[Serve], int]:
    """(each running `serve` - a Python or smartass.exe whose command line runs it - a child before its parent; how
    many Pythons' command lines could not be read). Ours when the command line's FIRST element, the executable, is in
    THIS install's venv (VENV_SCRIPTS, case-insensitive, as a whole folder), or when its parent is such a one: the
    venv's python.exe and smartass.exe are launchers that run the base interpreter as their child, whose command line
    names that interpreter instead. Any other (another checkout's, or one whose parent is gone) is not."""
    import psutil
    want = str(scripts or VENV_SCRIPTS).replace("/", "\\").rstrip("\\").lower() + "\\"
    found, unreadable = [], 0
    for p in _process_list():
        try:
            if (p.info["name"] or "").lower() not in SERVE_NAMES:
                continue
            cmd = [str(c) for c in (p.cmdline() or [])]
            if _is_serve(cmd):
                found.append((p, p.ppid(), cmd[0].replace("/", "\\").lower().startswith(want)))
        except psutil.NoSuchProcess:
            continue
        except psutil.Error:
            unreadable += 1
    ours = {p.pid for p, _, named in found if named}
    ours |= {p.pid for p, ppid, _ in found if ppid in ours}
    pids = {p.pid for p, _, _ in found}
    found.sort(key=lambda f: 0 if f[1] in pids else 1)              # a child first: its launcher ends with it
    return [Serve(p, p.pid in ours, tuple(c.pid for c, ppid, _ in found if ppid == p.pid)) for p, _, _ in found], \
        unreadable


def stop_old_serve(say, found=None, kill=None) -> None:
    """The HTTP API (`smartass serve`) an older release started at logon, stopped when its command line proves it is
    this install's (_serves), each stop said before it is done; another install's is left running, and said so. The
    stop goes to the process object the listing found, and only while it still runs (a pid a new process has taken
    is never stopped); a launcher whose child was stopped is given LAUNCHER_WAIT_S to end with it first. Pythons
    whose command line could not be read are said."""
    import psutil
    found, kill = found or _serves, kill or _kill
    serves, unreadable = found()
    stopped: set[int] = set()
    for s in serves:
        p = s.proc
        if not s.ours:
            say(f"startup: a `smartass serve` of another smartass install (pid {p.pid}) is left running")
            continue
        if stopped.intersection(s.children):
            try:
                p.wait(timeout=LAUNCHER_WAIT_S)
                say(f"startup: the old `smartass serve`'s launcher (pid {p.pid}) ended with it")
                continue
            except psutil.TimeoutExpired:
                pass                                # still running: stopped below, through the same object
            except psutil.Error:
                say(f"startup: the old `smartass serve`'s launcher (pid {p.pid}) ended with it")
                continue
        if not p.is_running():                      # ended, or its pid belongs to a new process now
            say(f"startup: the old `smartass serve` (pid {p.pid}) had already ended")
            continue
        say(f"startup: stopping this install's old `smartass serve` (pid {p.pid}): the HTTP API is gone")
        try:
            kill(p)
            stopped.add(p.pid)
        except psutil.NoSuchProcess:              # its launcher ended with its child
            say(f"startup: the old `smartass serve` (pid {p.pid}) had already ended")
        except psutil.Error as e:
            say(f"startup: the old `smartass serve` (pid {p.pid}) could not be stopped ({type(e).__name__}); it ends at "
                "the next logoff")
    if unreadable:
        say(f"startup: the command line of {unreadable} Python process(es) could not be read: an old `smartass serve` "
            "among them was not looked at")


def _watchers(script: Path = WATCHER_PS1) -> list[tuple[int, bool]]:
    """(pid, ours) of each running PowerShell whose command line names an sso_watcher.ps1: ours only when it names
    THIS install's script by its full path (case-insensitive) - another checkout's watcher on this PC is not."""
    import psutil
    want = str(script).lower()
    out = []
    for p in psutil.process_iter(["name", "cmdline"]):
        try:
            if (p.info["name"] or "").lower() in ("powershell.exe", "pwsh.exe"):
                cmd = " ".join(p.info["cmdline"] or []).lower()
                if "sso_watcher.ps1" in cmd:
                    out.append((p.pid, want in cmd))
        except psutil.Error:
            continue
    return out


def _kill(p) -> None:
    """Stop a process through the psutil.Process its listing found (psutil refuses one whose pid a new process has
    taken: NoSuchProcess); a bare pid (the SSO watcher's listing) is looked up now."""
    import psutil
    (p if isinstance(p, psutil.Process) else psutil.Process(p)).kill()


def stop_old_watcher(found=None, kill=None) -> list[str]:
    """The SSO watcher an older release started at logon, stopped when its command line proves it is ours (it names
    this install's script path); any other is left running, and said so."""
    import psutil
    found, kill = found or _watchers, kill or _kill
    out = []
    for pid, ours in found():
        if not ours:
            out.append(f"startup: an SSO watcher of another smartass install (pid {pid}) is left running")
            continue
        try:
            kill(pid)
            out.append(f"startup: stopped the old Smart View SSO watcher (pid {pid})")
        except psutil.Error as e:                 # gone meanwhile, or not ours to end: setup goes on
            out.append(f"startup: the old Smart View SSO watcher (pid {pid}) could not be stopped ({type(e).__name__});"
                       " it ends at the next logoff")
    return out


REG_EXPAND_SZ = 2


def _path_override() -> Path | None:
    """SMARTASS_USER_PATH_FILE: a file stands in for HKCU\\Environment Path (fresh-profile tests)."""
    f = os.environ.get("SMARTASS_USER_PATH_FILE")
    return Path(f) if f else None


def _read_user_path() -> tuple[str, int]:
    fake = _path_override()
    if fake:
        try:
            return fake.read_text(encoding="utf-8"), REG_EXPAND_SZ
        except FileNotFoundError:
            return "", REG_EXPAND_SZ
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as k:
        try:
            value, kind = winreg.QueryValueEx(k, "Path")
            return value or "", kind
        except FileNotFoundError:
            return "", winreg.REG_EXPAND_SZ


def _write_user_path(value: str, kind: int) -> None:
    fake = _path_override()
    if fake:
        fake.parent.mkdir(parents=True, exist_ok=True)
        fake.write_text(value, encoding="utf-8")
        return
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_SET_VALUE) as k:
        winreg.SetValueEx(k, "Path", 0, kind, value)


def _broadcast_env_change() -> None:
    """Tell Explorer (and so new terminals) that the environment changed. Best effort."""
    if _path_override():
        return
    try:
        import ctypes
        from ctypes import wintypes
        result = wintypes.DWORD()
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001A, 0, "Environment", 0x0002, 5000,
                                                 ctypes.byref(result))   # HWND_BROADCAST, WM_SETTINGCHANGE
    except Exception:  # noqa: BLE001 - only a notification
        pass


def _path_key(entry: str) -> str:
    return os.path.expandvars(entry.strip().strip('"')).rstrip("\\/").lower()


def add_to_user_path(folder: Path, read=_read_user_path, write=_write_user_path,
                     broadcast=_broadcast_env_change) -> str:
    """Append folder to HKCU\\Environment Path once (case-insensitive), keeping the value type (REG_EXPAND_SZ)."""
    value, kind = read()
    if _path_key(str(folder)) in {_path_key(e) for e in value.split(";") if e.strip()}:
        return "unchanged"
    write(f"{value.rstrip(';')};{folder}" if value.strip() else str(folder), kind)
    broadcast()
    return "added"


def path_step(executable: str | None = None, **reg) -> str:
    """Put the engine venv's Scripts folder (where smartass.exe lives) on the user PATH."""
    scripts = Path(executable or sys.executable).parent
    if not (scripts / "smartass.exe").exists():
        return f"path: skipped - no smartass.exe next to {scripts}"
    try:
        r = add_to_user_path(scripts, **reg)
    except OSError as e:
        return f"path: failed ({e}) - run {scripts / 'smartass.exe'} by its full path"
    return f"path: {r} ({scripts}; open a new terminal to use `smartass`)"


def run() -> list[str]:
    config.HOME.mkdir(parents=True, exist_ok=True)
    out = [f"config: {write_user_config()} ({config.USER_CONFIG})"]
    e = mcp_entries(config.ENGINE_DIR / ".venv" / "Scripts" / "python.exe", config.ENGINE_DIR / "run_server.py")
    out.append(f"copilot mcp: {register_mcp(USER / '.copilot' / 'mcp-config.json', e['copilot'], create=True)}")
    out.append(f"claude mcp: {register_mcp(USER / '.claude.json', e['claude'])}")
    skill_src = config.REPO_DIR / "skills" / "smartass"
    backup = config.HOME / "backup"
    for agent in (".claude", ".agents"):
        skills = USER / agent / "skills"
        out.append(f"skill ({agent[1:]}): {link_skill(skill_src, skills / 'smartass', backup)}"
                   f"; legacy smartview: {drop_legacy_skill(skills / 'smartview', backup)}")
    out += startup_entries()
    out.append(path_step())
    return out
