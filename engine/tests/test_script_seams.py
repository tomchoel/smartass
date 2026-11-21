"""The scripts in scripts/ patch engine functions to watch or replace them (check_followups wraps exceljob._look,
check_insert swaps exceljob._save, ...). Neither pytest nor ruff runs those scripts, so a signature the engine changes
would only show when a live script dies (Task 2c review Important 1: pipeline._look(t, name) became _look(job)).
This reads every script's source (it does not import or run them) and checks each patched name still exists and
that the replacement takes the arguments the engine's own function takes."""
from __future__ import annotations

import ast
import importlib
import inspect
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[2] / "scripts"


def _modules(tree: ast.AST) -> dict[str, str]:
    """alias -> module path, from every 'from smartass[...] import x [as y]' in the script."""
    out = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module and node.module.split(".")[0] == "smartass":
            for a in node.names:
                out[a.asname or a.name] = f"{node.module}.{a.name}"
    return out


def _shape(fn: ast.Lambda | ast.FunctionDef) -> tuple[int, int | None]:
    """(required, at most) positional arguments of a lambda or def (None: *args)."""
    args = fn.args
    n = len(args.posonlyargs) + len(args.args)
    return n - len(args.defaults), None if args.vararg else n


def _patches(path: Path):
    """(line, module path, name, the replacement's shape or None) of every patch of an engine module's attribute:
    'mod.name = value' and '(mod, "name"): value' in a dict."""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    mods, defs = _modules(tree), {n.name: n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef)}

    def shape(value):
        if isinstance(value, ast.Lambda):
            return _shape(value)
        if isinstance(value, ast.Name) and value.id in defs:
            return _shape(defs[value.id])
        return None                                         # a value, a class, an outer name: no arity to check
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            for tgt in node.targets:
                if isinstance(tgt, ast.Attribute) and isinstance(tgt.value, ast.Name) and tgt.value.id in mods:
                    yield node.lineno, mods[tgt.value.id], tgt.attr, shape(node.value)
        elif isinstance(node, ast.Dict):
            for k, v in zip(node.keys, node.values):
                if (isinstance(k, ast.Tuple) and len(k.elts) == 2 and isinstance(k.elts[0], ast.Name)
                        and k.elts[0].id in mods and isinstance(k.elts[1], ast.Constant)):
                    yield k.lineno, mods[k.elts[0].id], k.elts[1].value, shape(v)


def _module(path: str):
    """The module a dotted path names (smartass.xl.movein), or its attribute (smartass.xl.inproc.runner)."""
    try:
        return importlib.import_module(path)
    except ImportError:
        parent, _, name = path.rpartition(".")
        return getattr(importlib.import_module(parent), name)


CASES = [(p.name, *patch) for p in sorted(SCRIPTS.glob("*.py")) for patch in _patches(p)]


def test_the_scripts_patch_engine_functions():
    """The check below has something to check (check_followups, check_insert and audit_testset patch the Excel job,
    exceljob)."""
    assert {name for _, _, mod, name, _ in CASES if mod.endswith("exceljob")} >= {"_look", "_save", "_target"}


@pytest.mark.parametrize("script, line, mod, name, shape", CASES,
                         ids=[f"{c[0]}:{c[1]} {c[3]}" for c in CASES])
def test_a_patched_engine_function_is_patched_in_the_module_that_calls_it(script, line, mod, name, shape):
    """A function patched on a module that neither defines nor calls it is not replaced where the engine calls it
    (the Excel job moved out of pipeline.py into exceljob.py, phase 3c Task 3): a script patches the module that
    defines the function, or one that imports it and calls it by that name (server.redraw_job)."""
    module = _module(mod)
    real = getattr(module, name, None)
    if inspect.isfunction(real) and real.__module__ != module.__name__:
        calls = [ln for ln in inspect.getsource(module).splitlines() if f"{name}(" in ln and "import" not in ln]
        assert calls, f"{script}:{line} patches {mod}.{name}, defined in {real.__module__} and not called there"


@pytest.mark.parametrize("script, line, mod, name, shape", CASES,
                         ids=[f"{c[0]}:{c[1]} {c[3]}" for c in CASES])
def test_a_patched_engine_function_still_exists_and_takes_the_same_arguments(script, line, mod, name, shape):
    real = getattr(_module(mod), name, None)
    assert real is not None or hasattr(_module(mod), name), f"{script}:{line} patches {mod}.{name}: no such name"
    if shape is None or not inspect.isfunction(real):
        return
    params = [p for p in inspect.signature(real).parameters.values()
              if p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD)]
    required = sum(p.default is p.empty for p in params)
    low, high = shape
    assert low <= len(params) and (high is None or high >= required), (
        f"{script}:{line} replaces {mod}.{name}{inspect.signature(real)} with a function of {low}"
        f"{'' if high == low else f'..{high}'} positional arguments")
