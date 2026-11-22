"""The Excel task pane of an older release (addin/tn-claude, up to 0.7.0) as it may still be registered on a PC, found
for `smartass doctor`. The pane was registered by hand (0.7.0's addin/README.md): a web add-in whose manifest sat
in a shared-folder catalogue that Office trusts (a subkey of HKCU CATALOGUES_KEY whose Url ends in TNClaudeCatalog),
and a classic Excel-DNA add-in TNClaude64.xll in the old svkit home that Excel loads from an OPENn value of its Options
(the Add-ins dialog wrote it); its server's start script was meant for the Startup folder. Both registry places are
Office's own settings (ruling of 22.11): this module only READS them - nothing in smartass writes under
HKCU\\Software\\Microsoft\\Office - and the doctor says how to remove them in Excel. A Startup entry of the pane is a
file of this user's, removed by `smartass doctor --fix`."""
from __future__ import annotations

import re
from pathlib import Path

CATALOGUE = "TNClaudeCatalog"                               # the shared folder the manifest was put in
XLL = "TNClaude64.xll"                                      # the classic add-in Excel loads
OFFICE = r"Software\Microsoft\Office\16.0"
CATALOGUES_KEY = OFFICE + r"\WEF\TrustedCatalogs"          # one subkey per trusted catalogue: Id, Url, Flags
EXCEL_OPTIONS_KEY = OFFICE + r"\Excel\Options"             # OPEN, OPEN1, ...: the add-ins Excel loads at start
# What only the pane's start scripts and a shortcut to them carry (the node command line they run is any Node app's):
# their names, the folder, and the title in their first lines ("TN Claude pane", "TN Claude task pane").
STARTUP_MARKS = ("start-claude-pane", "tn-claude", "tn claude pane", "tn claude task pane")
STARTUP_KINDS = (".lnk", ".vbs", ".cmd", ".bat", ".js", ".ps1", ".url")


def _hkcu_values(path: str) -> dict[str, object]:
    """The values of an HKCU key (read only); {} when it is missing or cannot be read."""
    import winreg
    out: dict[str, object] = {}
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as k:
            for i in range(winreg.QueryInfoKey(k)[1]):
                name, value, _ = winreg.EnumValue(k, i)
                out[name] = value
    except OSError:
        return out
    return out


def _hkcu_subkeys(path: str) -> dict[str, dict[str, object]]:
    """{subkey name: its values} of an HKCU key (read only); {} when it is missing or cannot be read."""
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as k:
            names = [winreg.EnumKey(k, i) for i in range(winreg.QueryInfoKey(k)[0])]
    except OSError:
        return {}
    return {n: _hkcu_values(f"{path}\\{n}") for n in names}


def catalogues() -> list[str]:
    """The Url of every trusted add-in catalogue that is the pane's: its last folder is named TNClaudeCatalog. Matched
    by the Url alone - no catalogue's folder is read (a share that does not answer would make every doctor run wait)."""
    out = []
    for values in _hkcu_subkeys(CATALOGUES_KEY).values():
        url = str(values.get("Url") or "")
        last = re.split(r"[\\/]", url.rstrip("\\/"))[-1]
        if url and last.lower() == CATALOGUE.lower():
            out.append(url)
    return out


def _addin_path(value: str) -> str:
    """The file of an OPENn value ('/R "C:\\x\\a.xll"' or 'a.xlam')."""
    quoted = re.search(r'"([^"]+)"', value)
    return quoted.group(1) if quoted else value.split()[-1]


def excel_addins() -> list[tuple[str, str]]:
    """(value name, the add-in's file) of each OPENn value of Excel's Options that loads the pane's TNClaude64.xll."""
    return [(name, _addin_path(str(v))) for name, v in sorted(_hkcu_values(EXCEL_OPTIONS_KEY).items())
            if re.fullmatch(r"OPEN\d*", name, re.IGNORECASE) and XLL.lower() in str(v).lower()]


def _names_pane(data: bytes) -> bool:
    """True when a file's bytes carry one of STARTUP_MARKS (a .lnk keeps its target as UTF-16, at any offset; a
    script as plain text)."""
    texts = [data.decode("latin-1")]
    for start in (0, 1):
        part = data[start:]
        texts.append(part[:len(part) // 2 * 2].decode("utf-16-le", errors="ignore"))
    text = " ".join(texts).lower()
    return any(m in text for m in STARTUP_MARKS)


def startup_entries(folder: Path) -> list[Path]:
    """The pane's entries in the Startup folder: a shortcut or script there that names its start scripts or folder, or
    carries their title."""
    try:
        files = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in STARTUP_KINDS)
    except OSError:
        return []
    out = []
    for p in files:
        try:
            if _names_pane(p.read_bytes()):
                out.append(p)
        except OSError:
            continue
    return out
