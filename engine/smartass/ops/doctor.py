"""smartass doctor: cheap health checks (starts neither Excel nor a browser). Each check gives OK / WARN / FAIL,
a detail and the fix. Used by the CLI and fin_status (fin_status leaves out the leftovers check: run(leftovers=)).
`smartass doctor --fix` removes what an older release left behind (leftovers, fix_leftovers) - each thing said
before it goes; what lives in Office's own settings is only reported, with the steps to remove it in Excel."""
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
                     "Fix the TOML syntax, or delete config.toml and run `smartass setup`.")
    if not config.USER_CONFIG.exists():
        return Check("config", "WARN", f"no {config.USER_CONFIG}", "Run `smartass setup`.")
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
    p = config.ATLAS_DB
    if not atlas.present(p):
        step = atlas.no_atlas_step()
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
    if not atlas.present(config.ATLAS_DB):
        step = atlas.no_atlas_step()
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


# Smart View's stored options (HKCU, a string value holding XML: <option key="messageDisplay" value="2"/> ...): read
# with winreg (xl/dialogs.read_hkcu), never through COM. Oracle's option ids of SV_OPTIONS as its keys there.
SV_STORED = (r"Software\Hyperion Solutions\HyperionSmartView\Options", "CAOptionsXML")
SV_STORED_KEYS = {"104": "messageDisplay", "109": "bDisplayHomePanelOnStartup", "110": "bCommentDlgEnable"}


def _sv_stored() -> dict[str, str] | None:
    """Smart View's stored options {key: value}; None when they cannot be read (no value, not XML)."""
    from xml.etree import ElementTree
    try:
        xml = dialogs.read_hkcu(*SV_STORED)          # looked up at the call: the suite's fake reader
        root = ElementTree.fromstring(xml) if isinstance(xml, str) and xml.strip() else None
    except Exception:   # noqa: BLE001 - unreadable options prove nothing
        return None
    if root is None:
        return None
    return {str(o.get("key")): str(o.get("value")) for o in root.iter()
            if str(o.tag).rsplit("}", 1)[-1] == "option" and o.get("key") is not None}


def _number(value) -> int | None:
    """An option's value as Smart View stores it (true -> 1, "0" -> 0, 2 -> 2); None when it is not one."""
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    text = str(value).strip().lower()
    if text in ("true", "false"):
        return int(text == "true")
    try:
        return int(text)
    except ValueError:
        return None


def _differs(record: Path) -> bool:
    """True only when it is proven that Smart View's stored options differ from the originals the record holds: the
    record and the stored options both readable and a recorded option there with another value. Anything that cannot
    be read or compared is no difference."""
    try:
        rec = json.loads(record.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return False
    stored = _sv_stored()
    if not isinstance(rec, dict) or not stored:
        return False
    for oid, value in rec.items():
        key = SV_STORED_KEYS.get(str(oid))
        was, now = _number(value), _number(stored[key]) if key in stored else None
        if was is not None and now is not None and was != now:
            return True
    return False


def _options_files() -> list[Path]:
    """The left-behind copies of Smart View's options: the home's (only a crashed run leaves it), and the old svkit
    home's (read only) only when Smart View's stored options differ from it - an older engine always kept its record
    there, so the file alone proves nothing (the false alarm on the owner's PC, 22.11)."""
    out = [p for p in [config.SVKIT / SV_OPTIONS_FILE] if p.exists()]
    if config.LEGACY_HOME and (old := config.LEGACY_HOME / SV_OPTIONS_FILE).exists() and _differs(old):
        out.append(old)
    return out


def _legacy(p: Path) -> bool:
    return bool(config.LEGACY_HOME) and p.parent == config.LEGACY_HOME


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
# A table header line of config.toml: [name] or [[name]], with either line end.
_HEADER = re.compile(r"(?m)^[ \t]*(\[\[?)([^\]\r\n]+)\]\]?[ \t]*(?:#[^\r\n]*)?\r?$")


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


def _real_headers(text: str) -> list[re.Match]:
    """The table headers of config.toml's text ([name] and [[name]]): a header-looking line counts only where the text
    before it is TOML on its own (one inside a multi-line string is not a header)."""
    out = []
    for m in _HEADER.finditer(text):
        try:
            tomllib.loads(text[:m.start()])
        except tomllib.TOMLDecodeError:
            continue
        out.append(m)
    return out


def _section_end(text: str, heads: list[re.Match], i: int) -> int:
    """Where header i's section ends: at the next header, before the comment lines just above it (they belong to it)."""
    end = heads[i + 1].start() if i + 1 < len(heads) else len(text)
    while i + 1 < len(heads):
        line = text.rfind("\n", heads[i].end(), end - 1) + 1
        if line <= heads[i].end() or not text[line:end].lstrip(" \t").startswith("#"):
            break
        end = line
    return end


def _drop_sections(raw: bytes, names) -> bytes | None:
    """config.toml's bytes without the named tables (each header with its lines up to the next header; comment lines
    just above that one stay with it); every other byte as it was - line ends, no newline at the end. None when in
    doubt: not UTF-8 TOML, or the result is not exactly the old settings without those tables."""
    try:
        text = raw.decode("utf-8")
        old = tomllib.loads(text)
    except (UnicodeDecodeError, tomllib.TOMLDecodeError):
        return None
    heads = _real_headers(text)
    out = text
    for i in reversed(range(len(heads))):
        if heads[i].group(2).strip() in names:
            out = out[:heads[i].start()] + out[_section_end(text, heads, i):]
    try:
        new = tomllib.loads(out)
    except tomllib.TOMLDecodeError:
        return None
    return out.encode("utf-8") if new == {k: v for k, v in old.items() if k not in names} else None


def _by_hand(raw: bytes, names) -> str:
    """The lines of config.toml that set the named tables, for the user to delete by hand ('lines 4-6 ([http])')."""
    text = raw.decode("utf-8", errors="replace")
    heads = _real_headers(text)
    line = lambda pos: text.count("\n", 0, pos) + 1   # noqa: E731
    out = [f"lines {line(h.start())}-{line(max(h.start(), _section_end(text, heads, i) - 1))} "
           f"({h.group(1)}{h.group(2).strip()}{h.group(1).replace('[', ']')})"   # the header, never a comment on it
           for i, h in enumerate(heads) if h.group(2).strip().split(".")[0].strip() in names]
    keys = re.compile(r"^[ \t]*(" + "|".join(map(re.escape, names)) + r")[ \t]*[.=]", re.M)
    first = heads[0].start() if heads else len(text)
    out += [f"line {line(m.start())}" for m in keys.finditer(text, 0, first)]
    return ", ".join(out) or "the lines that set " + " and ".join(f"[{n}]" for n in names)


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


def _http_api_items(serves: list) -> list[str]:
    """What the HTTP API and the task pane of an older release left that `doctor --fix` removes or stops (serves: the
    running `serve`s, setup_user._serves())."""
    from . import pane_leftovers as pane
    from . import setup_user
    out = []
    if found := setup_user.engine_startup():
        out.append(f"the Startup entry {setup_user.ENGINE_STARTUP} starts `smartass serve` (the HTTP API, gone)"
                   if found[1] else f"the Startup entry {setup_user.ENGINE_STARTUP} of another smartass install: left "
                   "as it is - remove it there")
    for s in serves:
        out.append(f"an older `smartass serve` of this install runs (pid {s.proc.pid})" if s.ours else
                   f"a `smartass serve` of another smartass install runs (pid {s.proc.pid}): left as it is - stop it "
                   "there")
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
    Smart View's options that a crashed run could not put back (nothing can restore them now: set them by hand) - the
    old svkit home's copy only when Smart View's stored options differ from it - and the bridge add-ins it built; from
    0.7.0 and before, the HTTP API (its Startup entry, a running `serve`, its secret), the updater (its record, the
    release settings in config.toml) and the Excel task pane (its Startup entry; its registration in Office's settings,
    which the user removes)."""
    from . import setup_user
    opts, bridges, office = _options_files(), _bridge_files(), _office_items()
    serves, unreadable = setup_user._serves()
    http = _http_api_items(serves) + _config_items()
    # Pythons whose command line could not be read are said, never a finding of their own: a managed PC's service
    # Python stays unreadable, and nothing here could clear it.
    unread = (f"the command line of {unreadable} Python process(es) could not be read: an older `smartass serve` "
              "among them would not be found" if unreadable else "")
    if not opts and not bridges and not http and not office:
        return Check("leftovers", "OK", "nothing an older release left behind" + (f" ({unread})" if unread else ""))
    said, fixes = [], []
    for p in opts:
        if _legacy(p):
            said.append(f"Smart View's options differ from the originals an older smartass recorded in {p}: "
                        f"{_recorded(p)}")
            fixes.append("set those back by hand in Excel: Smart View ribbon > Options (close the other Excels first), "
                         f"or, when the values you have now are the ones you want, delete {p} by hand")
            continue
        said.append(f"an older smartass crashed with Smart View's options changed; the originals recorded in {p}: "
                    f"{_recorded(p)}")
        fixes.append("set those back by hand in Excel: Smart View ribbon > Options (close the other Excels first)")
    if bridges:
        said.append(f"{len(bridges)} VBA bridge add-in file(s) of an older release in {config.SVKIT} (unused)")
    said += http
    if [p for p in opts if not _legacy(p)] or bridges or http:
        fixes.append("then run `smartass doctor --fix` to remove the left-behind files and entries")
    for what, step in office:
        said.append(what)
        fixes.append(step)
    return Check("leftovers", "WARN", "; ".join(said + ([unread] if unread else [])), "; ".join(fixes))


def pending_leftovers() -> bool:
    """True when `smartass doctor` would list a leftover that is a file, a setting in config.toml or an entry in
    Office's settings (what `smartass setup` leaves to `doctor --fix` or to the user) - cheap: no process listing."""
    from . import pane_leftovers as pane
    from . import setup_user
    return bool(_options_files() or _bridge_files() or _secret_file() or _config_items()
                or pane.startup_entries(setup_user.STARTUP) or _office_items())


def fixable_leftovers() -> bool:
    """True when `smartass doctor --fix` can clear one of pending_leftovers(): a file or a setting in config.toml.
    Office's settings (the pane's registration) and a record in the old svkit home are only the user's to change."""
    from . import pane_leftovers as pane
    from . import setup_user
    return bool([p for p in _options_files() if not _legacy(p)] or _bridge_files() or _secret_file()
                or _config_items() or pane.startup_entries(setup_user.STARTUP))


def fix_leftovers(say=print) -> list[Path]:
    """`smartass doctor --fix`: the left-behind files removed, each said first (the options file after its recorded
    values; the old svkit home's copy is never changed); a file in use (a bridge still loaded in an Excel) stays and
    is said so. The HTTP API's Startup entry, secret file and plain secret go, this install's `serve` is stopped
    (setup_user), the updater's record and the settings only an older release read go, the task pane's Startup entry
    goes; its registration in Office's settings is only said, with the steps in Excel. Returns the files removed."""
    from . import pane_leftovers as pane
    from . import setup_user
    removed = []
    for p in _options_files():
        if _legacy(p):
            say(f"Smart View's options differ from the originals an older smartass recorded (set them by hand in Smart "
                f"View > Options): {_recorded(p)}")
            say(f"left as it is (the old svkit home is never changed): {p}")
            continue
        say(f"Smart View options an older smartass left changed (set them by hand in Smart View > Options): "
            f"{_recorded(p)}")
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
    secret) taken out of config.toml whole, each said first; a value is never said. A copy of the file goes next to it
    first (config.toml.bak-<date>, never over an earlier one), and it is said. The new file is written only when it is
    exactly the old settings without those two tables (_drop_sections); otherwise the file is left as it is and the
    lines to delete by hand are said. What is still set after it (a dotted key, an inline table) is said to be left."""
    secret, names = _plain_http_secret(), _unused_settings()
    if not secret and not names:
        return False
    cfg = config.USER_CONFIG
    try:
        raw = cfg.read_bytes()
    except OSError as e:
        say(f"could not read {cfg} ({e.strerror or e}); delete its {_sections()} sections by hand")
        return False
    new = _drop_sections(raw, UNUSED_SECTIONS)
    if new is None:
        what = ", ".join((["a plain [http] secret"] if secret else []) + names)
        say(f"{cfg} is left as it is ({what}: no longer used; taking out its {_sections()} sections would change "
            f"more than them) - delete it by hand: {_by_hand(raw, UNUSED_SECTIONS)}")
        return False
    if secret:
        say(f"removing the plain [http] secret from {cfg} (the HTTP API is gone; the value is not shown)")
    if names:
        say(f"removing {', '.join(names)} from {cfg}: no longer used ({_gone(names)} gone)")
    if new != raw:
        tmp = cfg.with_name(cfg.name + ".tmp")
        try:
            bak = _backup(cfg, raw, say)
            tmp.write_bytes(new)
            os.replace(tmp, cfg)
        except OSError as e:
            try:
                tmp.unlink(missing_ok=True)         # no half-done copy left next to it
            except OSError:
                pass
            say(f"could not change {cfg} ({e.strerror or e}); delete its {_sections()} sections by hand")
            return False
        say(f"{cfg} changed; the copy from before is {bak}")
        if secret:
            say(f"the backup {bak} still holds the plain [http] secret - delete it when you no longer need it")
    left = _unused_settings()
    if _plain_http_secret():                        # written another way (a dotted key, an inline table)
        say(f"the plain [http] secret is still in {cfg}: delete it by hand")
    if left:
        say(f"{', '.join(left)} still set in {cfg}: delete it by hand")
    return new != raw and not (left or _plain_http_secret())


def _sections() -> str:
    return " and ".join(f"[{s}]" for s in UNUSED_SECTIONS)


def _backup(cfg: Path, raw: bytes, say) -> Path:
    """A copy of config.toml next to it, config.toml.bak-<date> (-2, -3 ... when that name is taken: an earlier backup
    is never overwritten), said before it is written; a copy of the same day that already holds these bytes (a retry
    after a write that failed) is used as it is. OSError when it cannot be written (then nothing is changed)."""
    stamp = time.strftime("%Y-%m-%d")
    for n in range(1, 1000):
        bak = cfg.with_name(f"{cfg.name}.bak-{stamp}" + (f"-{n}" if n > 1 else ""))
        if bak.exists():
            try:
                if bak.read_bytes() == raw:
                    say(f"{cfg} is already backed up in {bak}")
                    return bak
            except OSError:
                pass
            continue
        say(f"backing up {cfg} to {bak} before it is changed")
        with open(bak, "xb") as f:                 # x: never over a file that appeared meanwhile
            f.write(raw)
        return bak
    raise FileExistsError(f"no free backup name for {cfg}")


def _remove(p: Path, say) -> list[Path]:
    try:
        p.unlink()
        return [p]
    except OSError as e:
        say(f"could not remove {p} ({e.strerror or e}); it may be in use (an add-in loaded in an Excel: close that "
            "Excel) - then run `smartass doctor --fix` again")
        return []


# What fin_status runs on every call (each cheap); `smartass doctor` and the bundle add LEFTOVERS: a process scan and
# a read of every Startup file (an older release's leftovers), paid once by the command, never by a tool call.
CHECKS = (_config, _excel, _smartview_startup, _epm, _oauth, _atlas, _knowledge, _output)
LEFTOVERS = (_leftovers,)


def run(leftovers: bool = False) -> list[Check]:
    out = []
    for f in CHECKS + (LEFTOVERS if leftovers else ()):
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
        z.writestr("doctor.json", json.dumps(as_dict(run(leftovers=True)), indent=2))
        z.writestr("versions.json", json.dumps(_versions(), indent=2))
        if config.USER_CONFIG.exists():
            z.writestr("config.toml", logsetup.redact(config.USER_CONFIG.read_text(encoding="utf-8")))
        for f in sorted((config.HOME / "logs").glob("*.log*")):
            z.writestr(f"logs/{f.name}", logsetup.redact(_tail(f)))
    return path
