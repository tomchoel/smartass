"""The engine has no Excel of its own (light-fast phase 2 Task 6): no module, import or name of the hidden Excel, its
VBA bridge, the Smart View automation or the sign-in dialog handling is left, and no code path can create an Excel
process except launching the USER's Excel on a workbook (xl/target._launch_excel, called by _open_path and
_spawn_excel). Imports are checked in a fresh interpreter (this suite's conftest has imported much already); names and
call sites by reading every engine source file."""
import ast
import importlib.util
import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

ENGINE = Path(__file__).resolve().parents[1]
GONE = ["smartass.xl.host", "smartass.xl.hidden", "smartass.xl.login", "smartass.epm.ops", "smartass.epm.sso"]
# Names of the deleted parts: none may appear in the engine's code, comments or texts - except where a leftover of
# an older release is cleaned up (each such use named here, with why).
NAMES = ["ExcelHost", "HiddenExcel", "svbridge", "sso_watcher", "HypConnect", "ensure_global_watcher",
         "GLOBAL_OPTIONS", "BAS_FILES", "modBridge", "smartview_body", "wait_login_cleared", "SmartViewBoxes",
         "on_host", "_attach_bridge", "ensure_bridge_file", "smartview_version"]
BRIDGE_MACRO = re.compile(r"\bB_[A-Z][A-Za-z]+\b")          # the bridge's macros: B_Ping, B_Connect, B_SetOption ...
LEFTOVER_USES = {("smartass/ops/doctor.py", "svbridge"): "doctor --fix removes an older release's bridge files",
                 ("smartass/ops/setup_user.py", "sso_watcher"): "setup stops an older release's watcher"}
WALK = r"""
import importlib, json, pkgutil, sys
mods = ["smartass.server"]
import smartass
mods += [m.name for m in pkgutil.walk_packages(smartass.__path__, "smartass.")]
for m in mods:
    importlib.import_module(m)
print(json.dumps({"walked": len(mods), "loaded": sorted(m for m in sys.modules if m.startswith("smartass."))}))
"""


def _sources():
    files = [*(ENGINE / "smartass").rglob("*.py"), *ENGINE.glob("*.py")]
    return sorted(p for p in files if "__pycache__" not in p.parts)


def _rel(p: Path) -> str:
    return p.relative_to(ENGINE).as_posix()


def test_the_deleted_modules_are_gone_and_every_engine_module_imports_without_them():
    out = subprocess.run([sys.executable, "-c", WALK], capture_output=True, text=True, timeout=300,
                         stdin=subprocess.DEVNULL, cwd=ENGINE)
    assert out.returncode == 0, out.stderr[-3000:]
    got = json.loads(out.stdout.strip().splitlines()[-1])
    assert got["walked"] > 100 and [m for m in GONE if m in got["loaded"]] == []
    assert [m for m in GONE if importlib.util.find_spec(m) is not None] == []
    assert not (ENGINE / "vba").exists() and not (ENGINE / "sso_watcher.ps1").exists()


def test_no_name_of_the_deleted_parts_is_left_in_the_engine():
    found = []
    for p in _sources():
        text = p.read_text(encoding="utf-8")
        for name in NAMES:
            if name in text and (_rel(p), name) not in LEFTOVER_USES:
                found.append(f"{_rel(p)}: {name}")
        found += [f"{_rel(p)}: {m}" for m in sorted(set(BRIDGE_MACRO.findall(text)))]
    assert found == []


PROCESS_STARTS = {"Popen", "run", "call", "check_call", "check_output", "startfile", "ShellExecute", "ShellExecuteEx",
                  "CreateProcess", "spawnv", "spawnl", "system"}


def _calls(tree):
    for node in ast.walk(tree):
        if isinstance(node, ast.Call):
            f = node.func
            yield node, (f.attr if isinstance(f, ast.Attribute) else f.id if isinstance(f, ast.Name) else "")


def test_only_the_users_excel_is_ever_started_and_only_on_a_workbook():
    """No COM creation of Excel (DispatchEx, EnsureDispatch, GetObject, CoCreateInstance, Dispatch of an Excel progid);
    a process start names Excel (config.EXCEL_EXE or 'excel' in a literal) only in target._launch_excel, which only
    _open_path (a workbook file) and _spawn_excel (the TN template, a new workbook) call."""
    creating, excel_exe, launch_callers = [], [], set()
    for p in _sources():
        tree = ast.parse(p.read_text(encoding="utf-8"))
        parents = {c: n for n in ast.walk(tree) for c in ast.iter_child_nodes(n)}

        def enclosing(node):
            while node in parents:
                node = parents[node]
                if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    return node.name
            return "<module>"
        for node, name in _calls(tree):
            literal = [a.value for a in node.args if isinstance(a, ast.Constant) and isinstance(a.value, str)]
            if name in ("DispatchEx", "EnsureDispatch", "GetObject", "GetActiveObject", "CoCreateInstance",
                        "CoCreateInstanceEx") or (name == "Dispatch" and any("excel" in s.lower() for s in literal)):
                creating.append(f"{_rel(p)}:{node.lineno} {name}")
            names_excel = any((isinstance(n, ast.Attribute) and n.attr == "EXCEL_EXE") or
                              (isinstance(n, ast.Constant) and isinstance(n.value, str) and "excel" in n.value.lower())
                              for n in ast.walk(node))
            if name in PROCESS_STARTS and names_excel:
                excel_exe.append(f"{_rel(p)}:{enclosing(node)}")
            if name == "_launch_excel":
                launch_callers.add(f"{_rel(p)}:{enclosing(node)}")
    assert creating == []
    assert excel_exe == ["smartass/xl/target.py:_launch_excel"]
    assert launch_callers == {"smartass/xl/target.py:_open_path", "smartass/xl/target.py:_spawn_excel"}


# Every process start (Task 6b): os.startfile / ShellExecute* start whatever a file's type opens - a workbook path
# starts Excel without naming it - and a subprocess or os call can take Excel's path from a variable. So each start is
# allowed by its (file, function), with why; anything else, whatever it names, fails the guard.
START_NAMES = {"startfile", "ShellExecute", "ShellExecuteW", "ShellExecuteEx", "ShellExecuteExW", "CreateProcess",
               "CreateProcessW", "CreateProcessAsUser", "Popen"}   # Popen on any receiver: subprocess under an alias
SUBPROCESS_STARTS = {"Popen", "run", "call", "check_call", "check_output", "getoutput", "getstatusoutput"}
OS_STARTS = {"startfile", "system", "popen", "spawnl", "spawnle", "spawnlp", "spawnv", "spawnve", "spawnvp", "execv",
             "execve", "execl", "execlp", "execvp", "posix_spawn"}
ALLOWED_STARTS = {
    "smartass/xl/target.py:_launch_excel": "the user's Excel on a workbook (EXCEL.EXE /x <file>)",
    "run_server.py:start_child": "the relay's server child (this Python)",
    "smartass/ops/login.py:_copy": "clip.exe: the sign-in code to the clipboard",
    "smartass/ops/setup_user.py:link_skill": "cmd mklink /J: the skill's junction",
    "smartass/ops/update.py:apply": "the release's installer (smartass update)",
    "smartass/ops/update.py:start_if_available": "the release's installer, detached (serve at login)",
    "smartass/xl/inproc/build.py:_compile": "csc.exe: the in-process add-in's compiler",
}


def process_starts(source: str, rel: str) -> list[str]:
    """'rel:function receiver.name' of every call in source that starts a process: os.startfile / ShellExecute* /
    CreateProcess* on any receiver or imported bare, subprocess.<start> and os.<start> (the receiver resolved: a
    bare .run is asyncio's, a doctor's or a COM object's, not a process), Popen / startfile imported bare."""
    tree = ast.parse(source)
    parents = {c: n for n in ast.walk(tree) for c in ast.iter_child_nodes(n)}

    def enclosing(node):
        while node in parents:
            node = parents[node]
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                return node.name
        return "<module>"
    out = []
    for node, name in _calls(tree):
        f = node.func
        recv = f.value.id if isinstance(f, ast.Attribute) and isinstance(f.value, ast.Name) else ""
        if (name in START_NAMES or (recv == "subprocess" and name in SUBPROCESS_STARTS)
                or (recv == "os" and name in OS_STARTS)):
            out.append(f"{rel}:{enclosing(node)} {recv + '.' if recv else ''}{name}")
    return out


@pytest.mark.parametrize("source, found", [
    ("import os\ndef open_book(book):\n    os.startfile(book)\n", ["x.py:open_book os.startfile"]),
    ("import subprocess\ndef go(exe, p):\n    subprocess.Popen([exe, p])\n", ["x.py:go subprocess.Popen"]),
    ("from subprocess import Popen\ndef go(exe, p):\n    Popen([exe, p])\n", ["x.py:go Popen"]),
    ("import subprocess as sp\ndef go(exe, p):\n    sp.Popen([exe, p])\n", ["x.py:go sp.Popen"]),
    ("import win32api\ndef go(p):\n    win32api.ShellExecute(0, 'open', p, None, None, 1)\n",
     ["x.py:go win32api.ShellExecute"]),
    ("from win32com.shell import shell\ndef go(p):\n    shell.ShellExecuteEx(lpFile=p)\n",
     ["x.py:go shell.ShellExecuteEx"]),
    ("import os\ndef go(p):\n    os.system(f'start {p}')\n", ["x.py:go os.system"]),
    ("import asyncio\ndef go(c, doctor, app):\n    asyncio.run(c)\n    doctor.run()\n    app.Run('M')\n", []),
], ids=["startfile", "popen-variable", "popen-imported", "popen-alias", "shellexecute", "shellexecuteex", "os-system",
        "not-a-start"])
def test_the_launch_guard_sees_a_process_start_that_does_not_name_excel(source, found):
    assert process_starts(source, "x.py") == found


def test_every_process_start_is_one_of_the_allowed_and_only_launch_excel_starts_excel():
    starts = [s for p in _sources() for s in process_starts(p.read_text(encoding="utf-8"), _rel(p))]
    assert sorted({s.split(" ")[0] for s in starts}) == sorted(ALLOWED_STARTS)
    assert all(reason for reason in ALLOWED_STARTS.values())


# The documents a user or an agent reads for how smartass works now (light-fast phase 2 Task 7). The CHANGELOG, the
# audits, the plans, the specs and docs/archive are history and keep the old names.
DOCS = ["README.md", "docs/DESIGN.md", "docs/troubleshooting.md", "docs/release.md", "installer/install.ps1",
        "engine/defaults.toml", "skills/smartass/SKILL.md", "skills/smartass/reference/*.md"]
DOC_NAMES = [*NAMES, "the hidden Excel", "engine's own hidden Excel", "VBA bridge", "vba_access", "VBA project",
             "smartview_dll", "HsAddin", "SSO watcher", "pull_adhoc", "adhoc.py", "--excel", "B_Ping", "SV_LOGIN_GRACE",
             "SV_CALL_TIMEOUT", "WebView2", "xl/host", "epm/ops", "Smart View options", "Global options"]
DOC_LEFTOVERS = {("docs/troubleshooting.md", "svbridge"): "doctor --fix removes an older release's bridge files",
                 ("docs/troubleshooting.md", "VBA bridge"): "the doctor's leftovers text, quoted",
                 ("skills/smartass/reference/operations.md", "svbridge"): "the files an older release left"}


def test_the_docs_a_user_or_an_agent_reads_name_no_deleted_part():
    repo = ENGINE.parent
    found = []
    for pattern in DOCS:
        for p in sorted(repo.glob(pattern)):
            rel, text = p.relative_to(repo).as_posix(), p.read_text(encoding="utf-8")
            found += [f"{rel}: {n}" for n in DOC_NAMES if n in text and (rel, n) not in DOC_LEFTOVERS]
    assert found == []


def test_the_errors_live_in_one_neutral_module():
    from smartass.core import errors
    from smartass.xl import userjobs
    assert (userjobs.HostError, userjobs.HostTimeout, userjobs.InputError) == (errors.HostError, errors.HostTimeout,
                                                                               errors.InputError)
