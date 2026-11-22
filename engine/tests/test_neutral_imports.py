"""Nothing on a tool call's path imports the engine's hidden Excel (light-fast phase 2 Task 5): the server and every
EPM route, data and knowledge module load without xl/host.py, its hidden half (xl/hidden.py), Smart View's dialogs and
login handling, the SSO watcher or the hidden Excel's operations (epm/ops.py) - the errors come from core/errors.py,
which imports no COM. Checked in a fresh interpreter: this suite's conftest has imported xl/host.py already."""
import json
import subprocess
import sys

HIDDEN = ["smartass.xl.host", "smartass.xl.hidden", "smartass.epm.ops"]
FORBIDDEN = HIDDEN + ["smartass.xl.dialogs", "smartass.xl.login", "smartass.epm.sso"]
WALK = r"""
import importlib, json, pkgutil, sys
mods = ["smartass.server"]
for pkg in ("smartass.epm", "smartass.data", "smartass.knowledge"):
    p = importlib.import_module(pkg)
    mods += [m.name for m in pkgutil.walk_packages(p.__path__, pkg + ".")]
mods = [m for m in mods if m not in ("smartass.epm.ops", "smartass.epm.sso")]     # the hidden half itself
for m in mods:
    importlib.import_module(m)
print(json.dumps({"walked": len(mods), "loaded": sorted(m for m in sys.modules if m.startswith("smartass."))}))
"""
USER_PATH = r"""
import importlib, json, sys
for m in sys.argv[1:]:
    importlib.import_module(m)
print(json.dumps({"loaded": sorted(m for m in sys.modules if m.startswith("smartass."))}))
"""


def _run(code: str, *args: str) -> dict:
    out = subprocess.run([sys.executable, "-c", code, *args], capture_output=True, text=True, timeout=300,
                         stdin=subprocess.DEVNULL)
    assert out.returncode == 0, out.stderr[-3000:]
    return json.loads(out.stdout.strip().splitlines()[-1])


def test_the_server_and_every_epm_data_and_knowledge_module_import_nothing_of_the_hidden_excel():
    got = _run(WALK)
    assert got["walked"] > 20
    assert [m for m in FORBIDDEN if m in got["loaded"]] == []


def test_the_user_excel_jobs_import_nothing_of_the_hidden_excel():
    got = _run(USER_PATH, "smartass.xl.jobs", "smartass.xl.target", "smartass.xl.instances", "smartass.exceljob",
               "smartass.workbench", "smartass.xl.workbench.fmt", "smartass.xl.workbench.history",
               "smartass.core.serverkit")
    assert [m for m in HIDDEN if m in got["loaded"]] == []


def test_the_com_thread_module_holds_no_hidden_excel():
    from smartass.xl import hidden, host, hostproc
    assert not hasattr(hostproc, "HiddenExcel") and host.HiddenExcel is hidden.HiddenExcel


def test_the_errors_are_the_same_classes_wherever_they_are_imported_from():
    from smartass.core import errors
    from smartass.xl import host
    assert (host.HostError, host.HostTimeout, host.InputError) == (errors.HostError, errors.HostTimeout,
                                                                   errors.InputError)
