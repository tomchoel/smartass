"""smartass doctor: cheap health checks (starts neither Excel nor a browser). Each check gives OK / WARN / FAIL,
a detail and the fix. Used by the CLI and fin_status. `smartass doctor --fix` removes what an older release left
behind (leftovers, fix_leftovers) - each thing said before it goes; what lives in Office's own settings is only
reported, with the steps to remove it in Excel."""
from __future__ import annotations

import contextlib
import json
import os
import platform
import re
import sys
import time
import tomllib
import urllib.error
import urllib.request
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path

from .. import __version__
from ..core import config, logsetup
from ..xl import dialogs

BUNDLE_FIX = "Send the file from `smartass doctor --bundle` to the maintainer."


@dataclass
class Check:
    name: str
    status: str            # OK | WARN | FAIL
    detail: str
    fix: str = ""


def _config() -> Check:
    if config.CONFIG_ERROR:
        return Check("config", "FAIL", config.CONFIG_ERROR,
                     "Fix the TOML syntax, or delete config.toml and re-run installer\\install.ps1.")
    if not config.USER_CONFIG.exists():
        return Check("config", "WARN", f"no {config.USER_CONFIG}", "Run installer\\install.ps1.")
    return Check("config", "OK", str(config.USER_CONFIG))


def _excel() -> Check:
    if config.EXCEL_EXE and os.path.exists(config.EXCEL_EXE):
        return Check("excel", "OK", config.EXCEL_EXE)
    return Check("excel", "FAIL", f"not found: {config.EXCEL_EXE}",
                 "Install desktop Excel (Microsoft 365), or set paths.excel_exe in config.toml.")


def _epm(timeout: float = 5) -> Check:
    try:
        urllib.request.urlopen(config.EPM_BASE, timeout=timeout)
    except urllib.error.HTTPError:
        pass                                            # any HTTP answer means the server is reachable
    except (urllib.error.URLError, OSError) as e:
        return Check("epm", "FAIL", f"{config.EPM_BASE} unreachable: {e}",
                     "Connect to the Telenor network (or VPN) and check the proxy.")
    return Check("epm", "OK", f"{config.EPM_BASE} reachable (the login itself is checked on the first query)")


def _oauth() -> Check:
    """The [oauth] set-up and the stored sign-in of each pod, from config and the token files only (no network). Grids
    and drill lines go over the user's token: without the set-up no tool call reaches EPM; the home pod (epm.base)
    without a scope is a FAIL too (every answer reads the closed month and the data-source list there, with no Excel
    fallback); a pod without a stored sign-in answers its first call with the sign-in page and code."""
    from ..epm import oauth
    if not oauth.configured():
        why = ("[oauth] idcs is not an https URL" if config.OAUTH_IDCS_REFUSED else
               "config.toml has no [oauth] idcs, client_id and scopes")
        return Check("oauth", "FAIL", f"not set up: {why}", oauth.NOT_SET_UP)
    host = lambda pod: pod.removeprefix("https://").split(".")[0]   # noqa: E731
    st = oauth.status()
    parts = [f"{host(s['pod'])}: " + ("signed in" if s["state"] == "signed_in" else
                                      "sign-in ended" if s["state"] == "ended" else "not signed in") for s in st]
    home = oauth.origin(config.EPM_BASE or "")
    no_scope = bool(home) and home not in oauth.pods()
    if no_scope:
        parts.append(f"{host(home)}: no scope (epm.base: no answers)")
        return Check("oauth", "FAIL", "; ".join(parts),
                     f"Add the scope of {home} under [oauth.scopes] in config.toml (the identity domain's administrator "
                     "reads it where that environment was added to the OAuth client as a resource), then run "
                     "`smartass login`.")
    if any(s["state"] != "signed_in" for s in st):
        return Check("oauth", "WARN", "; ".join(parts),
                     "Run `smartass login` (one approval in the browser), or approve the code a tool reply shows.")
    return Check("oauth", "OK", "; ".join(parts))


def _atlas(max_age_days: float = 35) -> Check:
    from ..knowledge import atlas
    from .atlas_build import no_atlas_step
    p = config.ATLAS_DB
    if not atlas.present(p):
        step = no_atlas_step()
        return Check("atlas", "FAIL", f"no atlas on this PC yet ({p}{'' if p.exists() else ': missing'})",
                     f"{step[0].upper()}{step[1:]}.")
    age = (time.time() - p.stat().st_mtime) / 86400
    trust = atlas.trust(p)
    if trust is not None and trust[0] < trust[1]:             # an atlas built before 0.7.0 (or rebuilt by one)
        what = "no dimension has" if not trust[0] else f"{trust[1] - trust[0]} of {trust[1]} dimensions have no"
        return Check("atlas", "WARN", f"{p} ({age:.0f} days old); {what} EPM's real member names",
                     "Rebuild once with `smartass atlas build` - until then the strict member check over Planning REST "
                     "is off.")
    status = "OK" if age <= max_age_days else "WARN"
    return Check("atlas", status, f"{p} ({age:.0f} days old)",
                 "" if status == "OK" else "Run `smartass atlas build`.")


def _knowledge() -> Check:
    """The semantic layer (knowledge/semantic) against the current atlas, with the build checks of `smartass
    knowledge build`: FAIL on a build error (a metric member not in the atlas, an alias clash); WARN when knowledge.db
    is missing, older than the atlas or the semantic files, or when drafts await approval."""
    import os

    from ..knowledge.build import BuildReport, _atlas, semantic_hash, check, summary
    from ..knowledge.semantic import SemanticError, load_all
    try:
        sem = load_all(config.SEMANTIC_DIR)
    except (OSError, SemanticError) as e:
        return Check("knowledge", "FAIL", f"{type(e).__name__}: {e}",
                     "Fix the semantic file named in the error (knowledge\\semantic), or update smartass.")
    detail = (f"{len(sem.metrics)} metrics, {len(sem.facts)} facts, {len(sem.concepts)} concepts, "
              f"{len(sem.rules)} rules, {len(sem.entities)} entities, {len(sem.kpis)} kpis, "
              f"{len(sem.vocab)} vocab, {len(sem.notes)} notes")
    from ..knowledge import atlas
    from .atlas_build import no_atlas_step
    if not atlas.present(config.ATLAS_DB):
        step = no_atlas_step()
        return Check("knowledge", "WARN", f"{detail}; members not checked: waits for the atlas",
                     f"{step[0].upper()}{step[1:]}; the knowledge is checked against it after.")
    report = BuildReport(out=config.KNOWLEDGE_DB)
    with contextlib.closing(_atlas(config.ATLAS_DB)) as db:
        check(sem.metrics, sem.facts, db, report, sem.concepts, sem)
    if report.errors:
        return Check("knowledge", "FAIL", f"{len(report.errors)} build error(s): " + "; ".join(report.errors[:3]),
                     "Fix the semantic files (a reorg or rename?), then run `smartass knowledge build`.")
    rebuild = "Run `smartass knowledge build` (or `smartass knowledge refresh` after a month-end)."
    k = config.KNOWLEDGE_DB
    if not k.exists():
        return Check("knowledge", "WARN", f"{detail}; knowledge.db not built yet (built on first use)", rebuild)
    problems = []
    if os.path.getmtime(k) < os.path.getmtime(config.ATLAS_DB):
        problems.append("knowledge.db is older than the atlas")
    if summary(k).get("semantic_hash") != semantic_hash(config.SEMANTIC_DIR):
        problems.append("knowledge.db is older than the semantic files")
    if report.drafts:
        problems.append(f"{report.drafts} draft(s) await approval")
        rebuild = "Review the drafts, then `smartass knowledge approve <id|all>`."
    if problems:
        return Check("knowledge", "WARN", f"{detail}; " + "; ".join(problems), rebuild)
    return Check("knowledge", "OK", f"{detail} match the atlas ({k})")


def _output() -> Check:
    d = config.ANALYSER_DIR
    try:
        d.mkdir(parents=True, exist_ok=True)
        t = d / f".smartass_write_test_{os.getpid()}"
        t.write_text("ok", encoding="utf-8")
        t.unlink()
    except OSError as e:
        return Check("output", "FAIL", f"cannot write {d}: {e}",
                     "Set paths.analyser in config.toml to a folder you can write.")
    return Check("output", "OK", str(d))


def _smartview_startup(read=dialogs.read_hkcu) -> Check:
    """Smart View settings of the user's Excel (smartass uses no Smart View, but it stays installed there; found
    2026-11-03): the extension update check at Office start opens Smart View's Login ~20 s after Excel opens, and a
    job is refused while that Login is open in an Excel (xl/dialogs.py: COM into it blocks); diagnostics capture slows
    Smart View and keeps sign-in cookies in plain text."""
    problems, fixes = [], []
    if str(read(*dialogs.STARTUP_CHECK)) == "1":
        problems.append("Smart View checks for extension updates at every Excel start (its Login opens ~20 s after "
                        "Excel opens, and smartass does not touch that Excel until it is signed in or closed)")
        fixes.append("Smart View > Options > Extensions: untick 'Check for updates when Microsoft Office starts'")
    if str(read(r"Software\Hyperion Solutions\HyperionSmartView\Options", "Profile")) == "1":
        problems.append("Smart View diagnostics capture is on (slower; sign-in cookies stored in plain text)")
        fixes.append("Smart View > Options > Advanced: turn diagnostics off, then delete "
                     "%APPDATA%\\Oracle\\SmartView\\Diagnostics")
    if problems:
        return Check("smartview_startup", "WARN", "; ".join(problems), "; ".join(fixes) + " (close Excel first).")
    return Check("smartview_startup", "OK", "no login at Excel start-up; diagnostics capture off")


# Smart View's options an older release (up to 0.6.0) changed while its hidden Excel ran, and restored at its shutdown
# from this file; the file left behind means a run crashed with them changed.
SV_OPTIONS_FILE = "smartview_options_original.json"
SV_OPTIONS = {"104": "Log Message Display (HSV_LOGMESSAGE_DISPLAY: which messages Smart View shows)",
              "109": "Display Home panel on startup (HSV_DISPLAY_HOMEPANEL_ONSTARTUP)",
              "110": "Show comment dialog on refresh (HSV_SHOW_COMMENTDIALOG_ON_REFRESH)"}   # Oracle's option ids


def _options_files() -> list[Path]:
    """The left-behind copies of Smart View's options: the home's, and the old svkit home's (read only)."""
    out = [config.SVKIT / SV_OPTIONS_FILE]
    if config.LEGACY_HOME:
        out.append(config.LEGACY_HOME / SV_OPTIONS_FILE)
    return [p for p in out if p.exists()]


def _bridge_files() -> list[Path]:
    """The VBA bridge add-ins an older release built (svbridge_<hash>.xlam): nothing uses them any more."""
    return sorted(config.SVKIT.glob("svbridge*.xlam"))


def _recorded(p: Path) -> str:
    try:
        rec = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return "unreadable"
    return "; ".join(f"{SV_OPTIONS.get(str(k), f'option {k}')} = {v!r}" for k, v in rec.items()) or "empty"


OLD_SECRET_BIN = "secret.bin"     # the HTTP API's secret (DPAPI) an older release (up to 0.7.0) kept in the home
OLD_UPDATE_ATTEMPT = "update_attempt.json"   # the updater's last try at login, an older release's (up to 0.7.0)
# The sections of config.toml only an older release read (up to 0.7.0): what is gone with them.
UNUSED_SECTIONS = {"release": "the updater", "http": "the HTTP API"}
_SECTION = re.compile(r"(?m)^[ \t]*\[([^\]]+)\][ \t]*(?:#.*)?$")


def _secret_file() -> Path | None:
    p = config.HOME / OLD_SECRET_BIN
    return p if p.exists() else None


def _update_attempt_file() -> Path | None:
    p = config.HOME / OLD_UPDATE_ATTEMPT
    return p if p.exists() else None


def _user_settings() -> dict:
    """The user's config.toml, read from the file now ({} when it is missing or not TOML: the config check says so)."""
    try:
        return tomllib.loads(config.USER_CONFIG.read_text(encoding="utf-8"))
    except (OSError, tomllib.TOMLDecodeError):
        return {}


def _plain_http_secret() -> bool:
    """True when the user's config.toml, read from the file now, still holds a plain [http] secret (an older release
    moved it to secret.bin; nothing reads it any more)."""
    return config.get(_user_settings(), "http.secret") is not None


def _unused_settings() -> list[str]:
    """The settings of UNUSED_SECTIONS the user's config.toml, read from the file now, still sets (release.folder,
    http.port, ...), by name; http.secret is said on its own (_plain_http_secret). An empty section is none."""
    data = _user_settings()
    return [f"{s}.{k}" for s in UNUSED_SECTIONS if isinstance(data.get(s), dict) for k in data[s]
            if f"{s}.{k}" != "http.secret"]


def _drop_sections(text: str, names) -> str:
    """config.toml without the named sections: each header with everything up to the next header (other sections
    untouched)."""
    heads = list(_SECTION.finditer(text))
    for i in reversed(range(len(heads))):
        if heads[i].group(1).strip() in names:
            end = heads[i + 1].start() if i + 1 < len(heads) else len(text)
            text = text[:heads[i].start()] + text[end:]
    return text


CATALOGUE_STEP = ("in Excel: File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs: remove "
                  "{url}, tick 'Next time Office starts, clear all previously-started web add-ins cache', OK, "
                  "restart Excel")
XLL_STEP = ("in Excel: File > Options > Add-ins > Manage: Excel Add-ins > Go: untick TNClaude ({name} loads {xll}; use "
            "the dialog - do not edit the OPEN values in the registry, Excel renumbers them), restart Excel, then "
            "delete {folder} by hand")


def _office_items() -> list[tuple[str, str]]:
    """(what, the manual step) for the pane's registration in Office's own settings: never changed by smartass."""
    from . import pane_leftovers as pane
    out = [(f"Office trusts the Excel task pane's add-in catalogue {url} (the pane is gone)",
            CATALOGUE_STEP.format(url=url)) for url in pane.catalogues()]
    out += [(f"Excel loads the task pane's classic add-in {xll} at start ({name})",
             XLL_STEP.format(name=name, xll=xll, folder=Path(xll).parent)) for name, xll in pane.excel_addins()]
    return out


def _http_api_items() -> list[str]:
    """What the HTTP API and the task pane of an older release left that `doctor --fix` removes or stops."""
    from . import pane_leftovers as pane
    from . import setup_user
    out = []
    if found := setup_user.engine_startup():
        out.append(f"the Startup entry {setup_user.ENGINE_STARTUP} starts `smartass serve` (the HTTP API, gone)"
                   if found[1] else f"the Startup entry {setup_user.ENGINE_STARTUP} of another smartass install: left "
                   "as it is - remove it there")
    serves, unreadable = setup_user._serves()
    for s in serves:
        out.append(f"an older `smartass serve` of this install runs (pid {s.proc.pid})" if s.ours else
                   f"a `smartass serve` of another smartass install runs (pid {s.proc.pid}): left as it is - stop it "
                   "there")
    if unreadable:
        out.append(f"the command line of {unreadable} Python process(es) could not be read (an older `smartass serve` "
                   "among them would not be found)")
    out += [f"the Excel task pane's Startup entry {p} (the pane is gone)"
            for p in pane.startup_entries(setup_user.STARTUP)]
    if p := _secret_file():
        out.append(f"the HTTP API's secret file {p} (unused)")
    return out


def _config_items() -> list[str]:
    """What the user's config.toml and the home still hold for the HTTP API and the updater (gone), never a value."""
    out = []
    if _plain_http_secret():
        out.append(f"{config.USER_CONFIG} still holds a plain [http] secret (unused; the value is not shown)")
    if names := _unused_settings():
        out.append(f"{config.USER_CONFIG} sets {', '.join(names)}: no longer used ({_gone(names)} gone)")
    if p := _update_attempt_file():
        out.append(f"the updater's record {p} (unused: the updater is gone)")
    return out


def _gone(names: list[str]) -> str:
    """'the updater is' / 'the updater and the HTTP API are', for the sections the names are in."""
    parts = [what for s, what in UNUSED_SECTIONS.items() if any(n.startswith(s + ".") for n in names)]
    return " and ".join(parts) + (" are" if len(parts) > 1 else " is")


def _leftovers() -> Check:
    """What an older release left: in the home, from 0.6.0 and before (its hidden Excel and VBA bridge), a copy of
    Smart View's options that a crashed run could not put back (nothing can restore them now: set them by hand) and
    the bridge add-ins it built; from 0.7.0 and before, the HTTP API (its Startup entry, a running `serve`, its secret),
    the updater (its record, the release settings in config.toml) and the Excel task pane (its Startup entry; its registration in Office's settings, which the user removes)."""
    opts, bridges, office = _options_files(), _bridge_files(), _office_items()
    http = _http_api_items() + _config_items()
    if not opts and not bridges and not http and not office:
        return Check("leftovers", "OK", "nothing an older release left behind")
    said, fixes = [], []
    for p in opts:
        said.append(f"an older smartass crashed with Smart View's options changed; the originals recorded in {p}: "
                    f"{_recorded(p)}")
        fixes.append("set those back by hand in Excel: Smart View ribbon > Options (close the other Excels first)")
    if bridges:
        said.append(f"{len(bridges)} VBA bridge add-in file(s) of an older release in {config.SVKIT} (unused)")
    said += http
    if opts or bridges or http:
        fixes.append("then run `smartass doctor --fix` to remove the left-behind files and entries")
    for what, step in office:
        said.append(what)
        fixes.append(step)
    return Check("leftovers", "WARN", "; ".join(said), "; ".join(fixes))


def fix_leftovers(say=print) -> list[Path]:
    """`smartass doctor --fix`: the left-behind files removed, each said first (the options file after its recorded
    values; the old svkit home's copy is never changed); a file in use (a bridge still loaded in an Excel) stays and
    is said so. The HTTP API's Startup entry, secret file and plain secret go, this install's `serve` is stopped
    (setup_user), the updater's record and the settings only an older release read go, the task pane's Startup entry
    goes; its registration in Office's settings is only said, with the
    steps in Excel. Returns the files removed."""
    from . import pane_leftovers as pane
    from . import setup_user
    removed = []
    for p in _options_files():
        say(f"Smart View options an older smartass left changed (set them by hand in Smart View > Options): "
            f"{_recorded(p)}")
        if config.LEGACY_HOME and p.parent == config.LEGACY_HOME:
            say(f"left as it is (the old svkit home is never changed): {p}")
            continue
        say(f"removing {p}")
        removed += _remove(p, say)
    for p in _bridge_files():
        say(f"removing the unused VBA bridge add-in {p}")
        removed += _remove(p, say)
    setup_user.stop_old_serve(say)
    removed += setup_user.drop_engine_startup(say)
    for p in pane.startup_entries(setup_user.STARTUP):
        say(f"removing {p} (the Excel task pane's Startup entry; the pane is gone)")
        removed += _remove(p, say)
    if p := _secret_file():
        say(f"removing the HTTP API's secret file {p} (the HTTP API is gone)")
        removed += _remove(p, say)
    if p := _update_attempt_file():
        say(f"removing the updater's record {p} (the updater is gone)")
        removed += _remove(p, say)
    changed = _fix_config(say)
    for what, step in _office_items():
        say(f"left for you (Office's own settings, never changed by smartass): {what} - {step}")
    if not removed and not changed:
        say("nothing removed")
    return removed


def _fix_config(say) -> bool:
    """The sections only an older release read ([release]: the updater's; [http]: the HTTP API's port and plain
    secret) taken out of config.toml whole (written whole, then put in place), each said first; a value is never said.
    What is still set after it (a dotted key, an inline table) is said to be left."""
    secret, names = _plain_http_secret(), _unused_settings()
    if not secret and not names:
        return False
    cfg = config.USER_CONFIG
    if secret:
        say(f"removing the plain [http] secret from {cfg} (the HTTP API is gone; the value is not shown)")
    if names:
        say(f"removing {', '.join(names)} from {cfg}: no longer used ({_gone(names)} gone)")
    try:
        tmp = cfg.with_name(cfg.name + ".tmp")
        tmp.write_text(_drop_sections(cfg.read_text(encoding="utf-8"), UNUSED_SECTIONS), encoding="utf-8")
        os.replace(tmp, cfg)
    except OSError as e:
        say(f"could not change {cfg} ({e.strerror or e}); delete its {' and '.join(f'[{s}]' for s in UNUSED_SECTIONS)} "
            "sections by hand")
        return False
    left = _unused_settings()
    if _plain_http_secret():                        # written another way (a dotted key, an inline table)
        say(f"the plain [http] secret is still in {cfg}: delete it by hand")
    if left:
        say(f"{', '.join(left)} still set in {cfg}: delete it by hand")
    return not (left or _plain_http_secret())


def _remove(p: Path, say) -> list[Path]:
    try:
        p.unlink()
        return [p]
    except OSError as e:
        say(f"could not remove {p} ({e.strerror or e}); it may be in use (an add-in loaded in an Excel: close that "
            "Excel) - then run `smartass doctor --fix` again")
        return []


CHECKS = (_config, _excel, _smartview_startup, _epm, _oauth, _atlas, _knowledge, _output, _leftovers)


def run() -> list[Check]:
    out = []
    for f in CHECKS:
        try:
            out.append(f())
        except Exception as e:   # noqa: BLE001 - a check never breaks the doctor
            out.append(Check(getattr(f, "__name__", "check").lstrip("_"), "FAIL", f"{type(e).__name__}: {e}",
                             BUNDLE_FIX))
    return out


def overall(results: list[Check]) -> str:
    statuses = {c.status for c in results}
    return "FAIL" if "FAIL" in statuses else ("WARN" if "WARN" in statuses else "OK")


def as_dict(results: list[Check]) -> dict:
    return {"status": overall(results), "version": __version__, "checks": [asdict(c) for c in results]}


def to_text(results: list[Check]) -> str:
    lines = [f"smartass {__version__}: {overall(results)}"]
    for c in results:
        lines.append(f"  {c.status:<4}  {c.name:<14} {c.detail}")
        if c.fix and c.status != "OK":
            lines.append(f"        fix: {c.fix}")
    return "\n".join(lines)


def _versions() -> dict:
    from importlib.metadata import PackageNotFoundError, version
    pk = {}
    for name in ("mcp", "pydantic", "pywin32", "psutil", "openpyxl"):
        try:
            pk[name] = version(name)
        except PackageNotFoundError:
            pk[name] = None
    return {"smartass": __version__, "python": sys.version, "os": platform.platform(), "packages": pk}


def _tail(p: Path, max_bytes: int = 2_000_000) -> str:
    with p.open("rb") as f:
        f.seek(max(0, p.stat().st_size - max_bytes))
        return f.read().decode("utf-8", errors="replace")


def bundle(dest_dir: Path | None = None) -> Path:
    """Zip doctor result, versions, redacted config.toml and the tail of every log, for the maintainer."""
    dest_dir = Path(dest_dir or config.HOME / "bundles")
    dest_dir.mkdir(parents=True, exist_ok=True)
    path = dest_dir / f"smartass-doctor-{time.strftime('%Y%m%d-%H%M%S')}.zip"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("doctor.json", json.dumps(as_dict(run()), indent=2))
        z.writestr("versions.json", json.dumps(_versions(), indent=2))
        if config.USER_CONFIG.exists():
            z.writestr("config.toml", logsetup.redact(config.USER_CONFIG.read_text(encoding="utf-8")))
        for f in sorted((config.HOME / "logs").glob("*.log*")):
            z.writestr(f"logs/{f.name}", logsetup.redact(_tail(f)))
    return path
