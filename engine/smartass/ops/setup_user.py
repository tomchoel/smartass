"""Per-user setup, run by installer\\install.ps1 (`smartass setup`). Idempotent: every step reports
created/updated/unchanged. Never needs admin rights."""
from __future__ import annotations

import json
import os
import re
import secrets
import shutil
import subprocess
import sys
import time
import tomllib
from pathlib import Path

from ..core import config

SERVER_NAME = "smartass"
LEGACY_SERVER_NAMES = ("smartview",)   # renamed 2026-10-23; register_mcp moves these entries to SERVER_NAME
USER = Path(os.environ.get("USERPROFILE", str(Path.home())))
STARTUP = Path(os.environ.get("APPDATA", str(USER / "AppData" / "Roaming"))) / \
    "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"

CONFIG_TEMPLATE = """# smartass user settings. Defaults live in engine/defaults.toml - only override what you need.

[paths]
# analyser = 'C:\\Users\\you\\OneDrive - Telenor\\analyser'   # where answer workbooks go
# excel_exe = 'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE'

[packs]
enabled = []            # e.g. ["produktregnskap"]

[release]
# version = "0.2.0"     # pin a version (also downwards); empty = latest

[http]
# port = 47811          # the API secret is in secret.bin next to this file (DPAPI: only this Windows user)
"""

_SECTION = re.compile(r"(?m)^[ \t]*\[([^\]]+)\][ \t]*(?:#.*)?$")
_PLAIN_SECRET = re.compile(r"(?m)^[ \t]*secret[ \t]*=.*(?:\r?\n|$)")


def _drop_http_secret(text: str) -> str:
    """config.toml without the `secret = ...` line of its [http] section (other sections untouched)."""
    heads = list(_SECTION.finditer(text))
    for i, h in enumerate(heads):
        if h.group(1).strip() == "http":
            end = heads[i + 1].start() if i + 1 < len(heads) else len(text)
            return text[:h.end()] + _PLAIN_SECRET.sub("", text[h.end():end], count=1) + text[end:]
    return text


def write_user_config(path: Path = config.USER_CONFIG) -> str:
    if path.exists():
        return "kept"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(CONFIG_TEMPLATE, encoding="utf-8")
    return "created"


def _plain_secret(cfg: Path) -> str | None:
    try:
        return config.get(tomllib.loads(cfg.read_text(encoding="utf-8")), "http.secret")
    except (OSError, tomllib.TOMLDecodeError):
        return None


def secret_step(cfg: Path = config.USER_CONFIG, secret_file: Path = config.SECRET_FILE) -> str:
    """The HTTP API secret, DPAPI-protected in secret.bin. A plain-text http.secret in config.toml is moved
    there once and removed from the file; without either a new one is generated. An existing readable
    secret.bin is kept."""
    plain = _plain_secret(cfg)
    if config.read_secret(secret_file) and not plain:
        return "secret: unchanged"
    how = "migrated from config.toml" if plain else "created"
    if not config.read_secret(secret_file):
        secret_file.parent.mkdir(parents=True, exist_ok=True)
        tmp = secret_file.with_name(secret_file.name + ".tmp")
        tmp.write_bytes(config.protect_secret(plain or secrets.token_urlsafe(32)))
        os.replace(tmp, secret_file)
    if plain:
        text = cfg.read_text(encoding="utf-8")
        tmp = cfg.with_name(cfg.name + ".tmp")
        tmp.write_text(_drop_http_secret(text), encoding="utf-8")
        os.replace(tmp, cfg)
    return f"secret: {how} (secret.bin, DPAPI)"


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
    os.replace(tmp, path)
    return "updated"


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


def make_shortcut(lnk: Path, target: str, args: str, workdir: str) -> None:
    import win32com.client
    lnk.parent.mkdir(parents=True, exist_ok=True)
    sc = win32com.client.Dispatch("WScript.Shell").CreateShortcut(str(lnk))
    sc.TargetPath, sc.Arguments, sc.WorkingDirectory, sc.WindowStyle = target, args, workdir, 7
    sc.Save()


def startup_entries() -> list[str]:
    venv = config.ENGINE_DIR / ".venv" / "Scripts"
    logs = config.HOME / "logs"
    logs.mkdir(parents=True, exist_ok=True)
    make_shortcut(STARTUP / "smartass-engine.lnk", str(venv / "pythonw.exe"), "-m smartass.ops.cli serve",
                  str(config.ENGINE_DIR))
    make_shortcut(STARTUP / "smartass-sso-watcher.lnk", "powershell.exe",
                  f'-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "{config.ENGINE_DIR / "sso_watcher.ps1"}"'
                  f' -LogDir "{logs}"', str(config.ENGINE_DIR))
    out = ["startup: smartass-engine.lnk, smartass-sso-watcher.lnk"]
    legacy = STARTUP / "svkit-sso-watcher.lnk"
    if legacy.exists():
        legacy.unlink()
        out.append("startup: removed legacy svkit-sso-watcher.lnk")
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


def copy_atlas() -> str:
    if not config.RELEASE_FOLDER:
        return "atlas: no release folder configured"
    src = Path(config.RELEASE_FOLDER) / "atlas" / "atlas.db"
    if not src.exists():
        return f"atlas: none in {src.parent}"
    dst = config.ATLAS_DB
    if dst.exists() and dst.stat().st_mtime >= src.stat().st_mtime:
        return "atlas: local copy is current"
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_suffix(".db.tmp")
    shutil.copy2(src, tmp)
    os.replace(tmp, dst)
    return f"atlas: copied from {src}"


def run() -> list[str]:
    config.HOME.mkdir(parents=True, exist_ok=True)
    out = [f"config: {write_user_config()} ({config.USER_CONFIG})", secret_step()]
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
    out.append(copy_atlas())
    return out
