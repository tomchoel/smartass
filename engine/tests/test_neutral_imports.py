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
# allowed by its (file, function), with why; anything else, whatever it names, fails the guard. The allowlist is by
# function: a second start inside an allowed function passes too (each allowed function is short; read it).
START_NAMES = {"startfile", "ShellExecute", "ShellExecuteW", "ShellExecuteEx", "ShellExecuteExW", "CreateProcess",
               "CreateProcessW", "CreateProcessAsUser", "Popen"}   # Popen on any receiver: subprocess under an alias
SUBPROCESS_STARTS = {"Popen", "run", "call", "check_call", "check_output", "getoutput", "getstatusoutput"}
OS_STARTS = {"startfile", "system", "popen", "spawnl", "spawnle", "spawnlp", "spawnv", "spawnve", "spawnvp", "execv",
             "execve", "execl", "execlp", "execvp", "posix_spawn"}
# The modules whose functions start a process (webbrowser.open: os.startfile underneath on Windows), by module; a
# module or a function of one is found under its own name, an alias (import subprocess as sp) or imported bare (from
# subprocess import run). A start function handed on as a value (login.run's open_page=webbrowser.open) counts where
# it is named; Popen only when called (as a name it is a type: an annotation).
START_MODULES = {"subprocess": SUBPROCESS_STARTS, "os": OS_STARTS,
                 "asyncio": {"create_subprocess_exec", "create_subprocess_shell"},
                 "webbrowser": {"open", "open_new", "open_new_tab"}}
ALLOWED_STARTS = {
    "smartass/xl/target.py:_launch_excel": "the user's Excel on a workbook (EXCEL.EXE /x <file>)",
    "run_server.py:start_child": "the relay's server child (this Python)",
    "smartass/ops/login.py:_copy": "clip.exe: the sign-in code to the clipboard",
    "smartass/ops/login.py:run": "the browser on the identity domain's sign-in page (webbrowser.open, as open_page)",
    "smartass/ops/setup_user.py:link_skill": "cmd mklink /J: the skill's junction",
    "smartass/xl/inproc/build.py:_compile": "csc.exe: the in-process add-in's compiler",
}


def process_starts(source: str, rel: str) -> list[str]:
    """'rel:function receiver.name' of every call in source that starts a process: os.startfile / ShellExecute* /
    CreateProcess* on any receiver or imported bare, a start function of START_MODULES on its module under any alias
    or imported bare under any name (the receiver resolved: a bare .run is asyncio's, a doctor's or a COM object's,
    not a process), Popen / startfile imported bare; and such a start function named without a call (handed on)."""
    tree = ast.parse(source)
    parents = {c: n for n in ast.walk(tree) for c in ast.iter_child_nodes(n)}
    modules = {m: m for m in START_MODULES}                 # local name -> module
    bare: set[str] = set()                                  # local names of start functions imported bare
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            modules.update({a.asname or a.name: a.name for a in node.names if a.name in START_MODULES})
        elif isinstance(node, ast.ImportFrom) and node.module in START_MODULES:
            bare.update(a.asname or a.name for a in node.names if a.name in START_MODULES[node.module])

    def start_of(f) -> str:
        """'receiver.name' or 'name' when f names a start function of START_MODULES, else ''."""
        if isinstance(f, ast.Attribute) and isinstance(f.value, ast.Name) and f.value.id in modules:
            return f"{f.value.id}.{f.attr}" if f.attr in START_MODULES[modules[f.value.id]] else ""
        return f.id if isinstance(f, ast.Name) and f.id in bare else ""

    def enclosing(node):
        while node in parents:
            node = parents[node]
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                return node.name
        return "<module>"
    out, called = [], set()
    for node, name in _calls(tree):
        f = node.func
        called.add(f)
        recv = f.value.id if isinstance(f, ast.Attribute) and isinstance(f.value, ast.Name) else ""
        if name in START_NAMES or start_of(f):
            out.append(f"{rel}:{enclosing(node)} {recv + '.' if recv else ''}{name}")
    for node in ast.walk(tree):
        if (isinstance(node, (ast.Attribute, ast.Name)) and isinstance(node.ctx, ast.Load) and node not in called
                and (named := start_of(node)) and not named.endswith("Popen")):
            out.append(f"{rel}:{enclosing(node)} {named}")
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
    ("import subprocess as sp\ndef go(p):\n    sp.run(p)\n", ["x.py:go sp.run"]),
    ("from subprocess import run\ndef go(p):\n    run(p)\n", ["x.py:go run"]),
    ("from os import system\ndef go(p):\n    system(p)\n", ["x.py:go system"]),
    ("import asyncio\nasync def go(p):\n    await asyncio.create_subprocess_exec(p)\n",
     ["x.py:go asyncio.create_subprocess_exec"]),
    ("import asyncio as aio\nasync def go(p):\n    await aio.create_subprocess_shell(p)\n",
     ["x.py:go aio.create_subprocess_shell"]),
    ("from asyncio import create_subprocess_exec\nasync def go(p):\n    await create_subprocess_exec(p)\n",
     ["x.py:go create_subprocess_exec"]),
    ("import webbrowser\ndef go(book):\n    webbrowser.open(book)\n", ["x.py:go webbrowser.open"]),
    ("from webbrowser import open as browse\ndef go(book):\n    browse(book)\n", ["x.py:go browse"]),
    ("import webbrowser\ndef go(page, open_page=webbrowser.open):\n    open_page(page)\n",
     ["x.py:go webbrowser.open"]),
    ("import subprocess\nfrom asyncio import run\ndef go(c, proc: subprocess.Popen) -> subprocess.Popen:\n"
     "    run(c)\n    return proc\n", []),
], ids=["startfile", "popen-variable", "popen-imported", "popen-alias", "shellexecute", "shellexecuteex", "os-system",
        "not-a-start", "run-alias", "run-imported", "system-imported", "asyncio-exec", "asyncio-alias-shell",
        "asyncio-exec-imported", "webbrowser-open", "webbrowser-open-imported", "webbrowser-open-handed-on",
        "not-a-start-asyncio-run-and-annotations"])
def test_the_launch_guard_sees_a_process_start_that_does_not_name_excel(source, found):
    assert process_starts(source, "x.py") == found


def test_every_process_start_is_one_of_the_allowed_and_only_launch_excel_starts_excel():
    starts = [s for p in _sources() for s in process_starts(p.read_text(encoding="utf-8"), _rel(p))]
    assert sorted({s.split(" ")[0] for s in starts}) == sorted(ALLOWED_STARTS)
    assert all(reason for reason in ALLOWED_STARTS.values())


# The documents a user or an agent reads for how smartass works now (light-fast phase 2 Task 7). The CHANGELOG, the
# audits, the plans, the specs and docs/archive are history and keep the old names.
DOCS = ["README.md", "docs/DESIGN.md", "docs/troubleshooting.md", "docs/release.md", "engine/defaults.toml",
        "skills/smartass/SKILL.md", "skills/smartass/reference/*.md"]
DOC_NAMES = [*NAMES, "hidden Excel", "VBA bridge", "vba_access", "VBA project", "smartview_dll", "HsAddin",
             "SSO watcher", "pull_adhoc", "adhoc.py", "--excel", "B_Ping", "SV_LOGIN_GRACE", "SV_CALL_TIMEOUT",
             "WebView2", "xl/host", "epm/ops", "Smart View options", "Global options"]
DOC_LEFTOVERS = {("docs/troubleshooting.md", "svbridge"): "doctor --fix removes an older release's bridge files",
                 ("docs/troubleshooting.md", "VBA bridge"): "the doctor's leftovers text, quoted",
                 ("skills/smartass/reference/operations.md", "svbridge"): "the files an older release left"}
# A name a document may use a known number of times (history, or a script's own Excel): one use more or less is a
# finding, so a new line that says the engine has a hidden Excel is judged, not let through with the file.
DOC_TIMES = {
    ("docs/DESIGN.md", "hidden Excel"): (4, "check_insert ran in the engine's hidden Excel of 0.6.0 and now in one of "
                                           "its own; two measurements made in a script's own hidden Excel"),
    ("docs/troubleshooting.md", "hidden Excel"): (2, "what a release before 0.7.0 ran; check_fastcom's own Excel"),
    ("docs/troubleshooting.md", "SSO watcher"): (2, "setup's line when it stops an older release's watcher, quoted"),
}


def doc_findings(repo: Path, docs: list[str], names=None, leftovers=None, times=None) -> list[str]:
    """'file: name' for every deleted part a listed document names (whitespace folded, so a name wrapped over two
    lines is found), unless DOC_LEFTOVERS allows it or DOC_TIMES allows exactly that many uses; a pattern that matches
    no file is a finding too (a renamed document would leave the guard nothing to read). Another guard passes its own
    names, leftovers and times."""
    names = DOC_NAMES if names is None else names
    leftovers = DOC_LEFTOVERS if leftovers is None else leftovers
    counts = DOC_TIMES if times is None else times
    found = []
    for pattern in docs:
        paths = sorted(repo.glob(pattern))
        if not paths:
            found.append(f"{pattern}: no such file")
        for p in paths:
            rel, text = p.relative_to(repo).as_posix(), " ".join(p.read_text(encoding="utf-8").split())
            for n in names:
                allowed, used = counts.get((rel, n), (0, ""))[0], text.count(n)
                if used != allowed and (rel, n) not in leftovers:
                    found.append(f"{rel}: {n} {used} time(s), {allowed} allowed")
    return found


def test_the_docs_a_user_or_an_agent_reads_name_no_deleted_part():
    assert doc_findings(ENGINE.parent, DOCS) == []
    assert all(why for _, why in DOC_TIMES.values()) and all(DOC_LEFTOVERS.values())


def test_the_docs_guard_fails_on_a_listed_file_that_is_missing(tmp_path):
    """A document renamed or moved would leave the guard nothing to read: a pattern that matches no file is a finding."""
    (tmp_path / "docs").mkdir()
    (tmp_path / "docs" / "here.md").write_text("plain text", encoding="utf-8")
    assert doc_findings(tmp_path, ["docs/here.md", "docs/gone.md", "skills/*/reference/*.md"]) == [
        "docs/gone.md: no such file", "skills/*/reference/*.md: no such file"]


def test_the_docs_guard_finds_a_name_wrapped_over_two_lines(tmp_path):
    (tmp_path / "d.md").write_text("smartass starts a hidden\n  Excel for the grids.", encoding="utf-8")
    assert doc_findings(tmp_path, ["d.md"]) == ["d.md: hidden Excel 1 time(s), 0 allowed"]


def test_the_errors_live_in_one_neutral_module():
    from smartass.core import errors
    from smartass.xl import userjobs
    assert (userjobs.HostError, userjobs.HostTimeout, userjobs.InputError) == (errors.HostError, errors.HostTimeout,
                                                                               errors.InputError)
