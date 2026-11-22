r"""The fresh-home proof (light-fast phase 4 Task 3): a new PC's install comes up clean. In a temporary folder of its
own this script makes a fresh home (SMARTASS_HOME), a user profile (USERPROFILE / APPDATA / LOCALAPPDATA: the agents'
MCP files, the skills folder, the Startup folder, the old svkit home) and a stand-in for the user PATH, runs `smartass
setup`'s steps there, then `smartass doctor` twice: as setup leaves the PC (oauth must say what config.toml needs,
then `smartass login`; the README must name that step before login), and after the README's [oauth] step with
invented values (https://idcs.example ...). The second doctor must show every check OK except the sign-in and the
atlas: oauth WARN, atlas FAIL and knowledge WARN (it waits for the atlas), each fix naming exactly one command, the
same one. Prints the checks and exits 0 only then.

Each stage runs in a fresh Python (config is read at import, as by a user's separate commands) with these fakes: the
process list is empty and a kill is a tripwire (no running `serve` or watcher is looked at or stopped); the registry
readers of Smart View and Office answer "not set"; doctor's EPM check is a stub (it needs the network). A registry
WRITE (winreg SetValue*/CreateKey*/DeleteKey*/DeleteValue), a COM object, a network connection and any child process
but `mklink /J` of the skill links inside the temp folder are tripwires. Before anything runs, every place a stage can
write (home, config, atlas, knowledge, the sign-ins, the analyser and output folders, the user profile, the Startup
folder, the PATH stand-in, the old svkit home) is checked to be inside the temp folder; the stage refuses otherwise.
The real profile's MCP files, home config, Startup folder and user PATH are read (stat, a listing, the value) before
and after: the script fails if any changed. No EPM, no Excel, no network.

Usage (from the repo root):
    engine\.venv\Scripts\python.exe scripts\check_fresh_home.py [--keep]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import tomllib
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
ENGINE = REPO / "engine"
LOGIN = "smartass login"
STARTUP = ("AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
OAUTH_NOTE = """
[oauth]          # the README's step before `smartass login`, with invented values (this check reaches no EPM)
idcs = "https://idcs.example"
client_id = "fresh-home-check"

[oauth.scopes]
"{pod}" = "urn:opc:serviceInstanceID=fresh-home-check"
"""


# ---------------------------------------------------------------- one stage (a fresh Python per command)


class Tripwire(BaseException):
    """A fake stopped something this check must never do (a BaseException: a broad except cannot hide it)."""


TRIPPED: list[str] = []
CALLS: dict[str, int] = {}


def _trip(what: str):
    TRIPPED.append(what)
    raise Tripwire(what)


def _counted(name: str, value):
    def fake(*args, **kwargs):
        CALLS[name] = CALLS.get(name, 0) + 1
        return value
    return fake


def _inside(root: Path, p) -> bool:
    try:
        Path(p).resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _tripwires(root: Path) -> None:
    import socket
    import winreg

    import win32com.client
    for name in ("SetValue", "SetValueEx", "CreateKey", "CreateKeyEx", "DeleteKey", "DeleteKeyEx", "DeleteValue"):
        setattr(winreg, name, lambda *a, _n=name, **k: _trip(f"registry write winreg.{_n}"))
    for name in ("Dispatch", "DispatchEx", "GetObject", "GetActiveObject"):
        setattr(win32com.client, name, lambda *a, _n=name, **k: _trip(f"COM win32com.client.{_n}"))
    real_connect, real_lookup, real_run = socket.socket.connect, socket.getaddrinfo, subprocess.run
    local = ("127.0.0.1", "::1", "localhost")

    def connect(self, address):
        if isinstance(address, tuple) and address[0] not in local:
            _trip(f"network connect {address[0]}")
        return real_connect(self, address)

    def lookup(host, *args, **kwargs):
        return real_lookup(host, *args, **kwargs) if host in local else _trip(f"name lookup {host}")

    def run(args, *a, **k):
        if (isinstance(args, list) and [str(x).lower() for x in args[:4]] == ["cmd", "/c", "mklink", "/j"]
                and _inside(root, args[4])):
            CALLS["mklink /J inside the temp folder"] = CALLS.get("mklink /J inside the temp folder", 0) + 1
            return real_run(args, *a, **k)
        _trip(f"child process {args!r}")
    socket.socket.connect, socket.getaddrinfo, subprocess.run = connect, lookup, run


def _fakes() -> None:
    from smartass.ops import doctor, pane_leftovers, setup_user
    from smartass.xl import dialogs
    setup_user._process_list = _counted("process list (empty)", [])
    setup_user._kill = lambda *process: _trip("a process kill")
    pane_leftovers._hkcu_values = _counted("Office registry values (none)", {})
    pane_leftovers._hkcu_subkeys = _counted("Office registry subkeys (none)", {})
    dialogs.read_hkcu = _counted("Smart View registry values (not set)", None)

    def epm():
        CALLS["epm check (a stub)"] = CALLS.get("epm check (a stub)", 0) + 1
        return doctor.Check("epm", "OK", "not checked here: a stub (this check reaches no network)")

    def smartview_startup():                        # its reader is bound at import: given the fake here
        return doctor._smartview_startup(dialogs.read_hkcu)
    stubs = {"_epm": epm, "_smartview_startup": smartview_startup}
    doctor.CHECKS = tuple(stubs.get(f.__name__, f) for f in doctor.CHECKS)
    for name, f in stubs.items():
        f.__name__ = name


def _outside(root: Path) -> list[str]:
    """The places a stage can write that are not inside root."""
    from smartass.core import config
    from smartass.ops import setup_user
    places = {"home": config.HOME, "config": config.USER_CONFIG, "atlas": config.ATLAS_DB,
              "knowledge": config.KNOWLEDGE_DB, "sign-ins": config.OAUTH_DIR, "analyser": config.ANALYSER_DIR,
              "output": config.OUT_DIR, "user profile": setup_user.USER, "Startup": setup_user.STARTUP,
              "old svkit home": config.LEGACY_HOME, "user PATH stand-in": os.environ.get("SMARTASS_USER_PATH_FILE")}
    return [f"{k}: {v}" for k, v in places.items() if not v or not _inside(root, v)]


def stage(name: str, root: Path) -> int:
    sys.path.insert(0, str(ENGINE))
    _tripwires(root)
    if bad := _outside(root):
        print("REFUSED - outside the temp folder: " + "; ".join(bad))
        return 2
    _fakes()
    from dataclasses import asdict

    from smartass.epm import oauth
    from smartass.ops import doctor, setup_user
    out: dict = {}
    try:
        if name == "setup":
            out["lines"] = setup_user.run()
        else:
            out["checks"] = [asdict(c) for c in doctor.run(leftovers=True)]
            out["not_set_up"] = oauth.NOT_SET_UP
    except Tripwire:
        pass
    out["tripped"], out["calls"] = TRIPPED, CALLS
    print("@@" + json.dumps(out))
    return 0


# ---------------------------------------------------------------- the driver


def _env(root: Path) -> dict:
    env = {k: v for k, v in os.environ.items() if not k.upper().startswith(("SV_", "SMARTASS_", "SVKIT"))}
    user = root / "user"
    env.update({"SMARTASS_HOME": str(root / "home"), "USERPROFILE": str(user),
                "APPDATA": str(user / "AppData" / "Roaming"), "LOCALAPPDATA": str(user / "AppData" / "Local"),
                "SMARTASS_USER_PATH_FILE": str(root / "user_path.txt"), "SV_ANALYSER_DIR": str(root / "analyser"),
                "SV_OUT_DIR": str(root / "output"), "OneDrive": str(root / "onedrive"),
                "OneDriveCommercial": str(root / "onedrive"), "PYTHONIOENCODING": "utf-8"})
    return env


def _run(name: str, root: Path) -> dict:
    r = subprocess.run([sys.executable, str(Path(__file__).resolve()), "--stage", name, "--root", str(root)],
                       env=_env(root), capture_output=True, text=True, encoding="utf-8", timeout=300)
    line = next((x for x in r.stdout.splitlines() if x.startswith("@@")), None)
    if r.returncode or line is None:
        raise SystemExit(f"stage {name} failed (exit {r.returncode}):\n{r.stdout}\n{r.stderr}")
    return json.loads(line[2:])


def _commands(text: str) -> set[str]:
    return set(re.findall(r"`(smartass [a-z][a-z -]*)`", text))


def _real_places() -> dict[str, tuple]:
    """The real profile's places a setup or doctor would change, read (never written) to show they were not."""
    import winreg
    user = Path(os.environ.get("USERPROFILE", str(Path.home())))
    startup = Path(os.environ.get("APPDATA") or user / "AppData" / "Roaming").joinpath(*STARTUP[2:])
    home = Path(os.environ.get("SMARTASS_HOME") or Path(os.environ["LOCALAPPDATA"]) / "smartass")
    out: dict[str, tuple] = {}
    for p in (user / ".claude.json", user / ".copilot" / "mcp-config.json", home / "config.toml"):
        out[str(p)] = (p.stat().st_mtime_ns, p.stat().st_size) if p.exists() else ("missing",)
    out[str(startup)] = tuple(sorted(x.name for x in startup.iterdir())) if startup.is_dir() else ("missing",)
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as k:
            out[r"HKCU\Environment Path"] = (winreg.QueryValueEx(k, "Path")[0],)
    except OSError:
        out[r"HKCU\Environment Path"] = ("missing",)
    return out


def judge_before_oauth(doc: dict) -> list[str]:
    """As setup leaves a PC: oauth says what config.toml needs, then `smartass login`; nothing else is wrong."""
    bad = []
    for c in doc["checks"]:
        if c["name"] == "oauth":
            if c["status"] != "FAIL" or c["fix"] != doc["not_set_up"] or _commands(c["fix"]) != {LOGIN}:
                bad.append(f"oauth before the [oauth] step: {c}")
        elif c["name"] not in ("atlas", "knowledge") and c["status"] != "OK":
            bad.append(f"{c['name']} before the [oauth] step: {c}")
    return bad


def judge(doc: dict) -> list[str]:
    """After the [oauth] step: every check OK but the sign-in's and the atlas's, each naming one command."""
    bad, by = [], {c["name"]: c for c in doc["checks"]}
    atlas = by.get("atlas", {})
    one = _commands(atlas.get("fix", ""))
    if atlas.get("status") != "FAIL" or len(one) != 1 or not atlas["detail"].startswith("no atlas on this PC yet"):
        bad.append(f"atlas: {atlas}")
    for c in doc["checks"]:
        if c["name"] == "oauth":
            if c["status"] != "WARN" or _commands(c["fix"]) != {LOGIN}:
                bad.append(f"oauth: {c}")
        elif c["name"] == "knowledge":
            if c["status"] != "WARN" or "waits for the atlas" not in c["detail"] or _commands(c["fix"]) != one:
                bad.append(f"knowledge: {c}")
        elif c["name"] != "atlas" and c["status"] != "OK":
            bad.append(f"{c['name']}: {c}")
    return bad


def readme_names_the_oauth_step_before_login() -> bool:
    """The README's install says that [oauth] goes into config.toml before `smartass login` (a fresh PC starts with
    doctor's set-up text)."""
    readme = (REPO / "README.md").read_text(encoding="utf-8")
    install = " ".join(readme.split("## Install", 1)[1].split("\n## ", 1)[0].split())
    return re.search(r"\[oauth\]` in that config\.toml .{0,200}? before `smartass login`", install) is not None


def _text(checks: list[dict]) -> str:
    lines = []
    for c in checks:
        lines.append(f"  {c['status']:<4}  {c['name']:<17} {c['detail']}")
        if c["fix"] and c["status"] != "OK":
            lines.append(f"        fix: {c['fix']}")
    return "\n".join(lines)


def _unlink_skills(root: Path) -> bool:
    """The skill links setup made (junctions to the repo's skills folder) removed as links, never followed; True when
    none is left."""
    for skills in root.glob("user/*/skills"):
        for p in skills.iterdir():
            if p.is_junction() or p.is_symlink():
                os.rmdir(p)
    return not any(p.is_junction() or p.is_symlink() for p in root.glob("user/*/skills/*"))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--keep", action="store_true", help="keep the temp folder (its path is printed)")
    ap.add_argument("--stage", help=argparse.SUPPRESS)
    ap.add_argument("--root", help=argparse.SUPPRESS)
    a = ap.parse_args(argv)
    if a.stage:
        return stage(a.stage, Path(a.root))
    root = Path(tempfile.mkdtemp(prefix="smartass_fresh_home_"))
    user = root / "user"
    (user / ".copilot").mkdir(parents=True)                   # the agents' MCP files, as on a PC that has them
    (user / ".claude.json").write_text("{}", encoding="utf-8")
    user.joinpath(*STARTUP).mkdir(parents=True)
    (user / "AppData" / "Local").mkdir(parents=True)
    before, problems = _real_places(), []
    try:
        print(f"fresh home in {root}")
        setup = _run("setup", root)
        print("smartass setup:\n" + "\n".join(f"  {line}" for line in setup.get("lines", [])))
        first = _run("doctor", root)
        print("smartass doctor, as setup leaves the PC:\n" + _text(first["checks"]))
        problems += judge_before_oauth(first)
        if not readme_names_the_oauth_step_before_login():
            problems.append("the README's install does not name the [oauth] step before `smartass login`")
        pod = tomllib.loads((ENGINE / "defaults.toml").read_text(encoding="utf-8"))["epm"]["base"]
        with (root / "home" / "config.toml").open("a", encoding="utf-8") as f:
            f.write(OAUTH_NOTE.format(pod=pod))
        print("the README's [oauth] step, with invented values")
        doc = _run("doctor", root)
        print("smartass doctor:\n" + _text(doc["checks"]))
        problems += judge(doc)
        claude = json.loads((user / ".claude.json").read_text(encoding="utf-8"))
        if "smartass" not in claude.get("mcpServers", {}):
            problems.append("setup wrote no Claude Code MCP entry in the temp profile")
        if not (user / ".copilot" / "mcp-config.json").exists():
            problems.append("setup wrote no Copilot MCP entry in the temp profile")
        if any(user.joinpath(*STARTUP).iterdir()):
            problems.append("setup wrote a Startup entry")
        if (root / "home" / "atlas.db").exists():
            problems.append("an atlas.db was made in the fresh home")
        problems += [f"tripwire: {t}" for s in (setup, first, doc) for t in s["tripped"]]
        calls: dict[str, int] = {}
        for s in (setup, first, doc):
            for k, v in s["calls"].items():
                calls[k] = calls.get(k, 0) + v
        print("stubbed or faked (the output proves no more than this): "
              + "; ".join(f"{k} x{v}" for k, v in sorted(calls.items())))
        print("not stubbed: excel (whether EXCEL.EXE exists on this PC), config, output (a probe file in the temp "
              "folder), oauth (the temp home's config and sign-ins), atlas and knowledge (the temp home), leftovers "
              "(the temp home and Startup folder)")
    finally:
        after = _real_places()
        changed = [k for k in before if before[k] != after.get(k)]
        problems += [f"changed outside the temp folder: {k}" for k in changed]
        print(f"untouched outside the temp folder: {len(before) - len(changed)} of {len(before)} places read before "
              "and after (the real MCP files, home config, Startup folder and user PATH)")
        if a.keep:
            print(f"kept: {root}")
        elif _unlink_skills(root):
            shutil.rmtree(root, ignore_errors=True)
        else:
            print(f"left (a link could not be removed): {root}")
    for p in problems:
        print(f"PROBLEM: {p}")
    print("OK: a fresh home comes up with only the sign-in and the atlas to do" if not problems else
          f"FAIL: {len(problems)} problem(s)")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main())
