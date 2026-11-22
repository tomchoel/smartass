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


def test_the_errors_live_in_one_neutral_module():
    from smartass.core import errors
    from smartass.xl import userjobs
    assert (userjobs.HostError, userjobs.HostTimeout, userjobs.InputError) == (errors.HostError, errors.HostTimeout,
                                                                               errors.InputError)
