"""One semantic layer (design §5): catalog.toml says what lines, KPIs and scenarios mean in the cubes;
capabilities.toml says which words mean which measure, which grains can carry which measure, and which currencies
exist. The scope gate (core/scope.py) reads both through this module."""
from __future__ import annotations

import functools
import re
import tomllib
from dataclasses import dataclass

from .. import catalog, config


@dataclass(frozen=True)
class Measure:
    key: str
    synonyms: tuple[str, ...]
    concept: str | None = None
    kpi: str | None = None
    account: str | None = None
    accounts: tuple[str, ...] = ()   # account members summed (FinLite/FinMax names), e.g. consultants
    finplan: str | None = None       # the FinPlan account for a total-level budget, when one maps
    label: str | None = None
    vague: bool = False
    options: tuple[str, ...] = ()
    note: str | None = None
    matched: str = ""                # the words of the question that matched (set by measures())


@dataclass(frozen=True)
class Grain:
    key: str
    synonyms: tuple[str, ...]
    feasible: tuple[str, ...]
    proxy: str | None = None
    note: str | None = None
    matched: str = ""


@functools.cache
def _caps() -> dict:
    return tomllib.loads(config.knowledge_file("capabilities.toml").read_text(encoding="utf-8"))


def _tuple(d: dict, *keys: str) -> dict:
    return {k: tuple(v) if k in keys else v for k, v in d.items()}


@functools.cache
def all_measures() -> tuple[Measure, ...]:
    return tuple(Measure(**_tuple(m, "synonyms", "options", "accounts")) for m in _caps().get("measure", []))


@functools.cache
def all_grains() -> tuple[Grain, ...]:
    return tuple(Grain(**_tuple(g, "synonyms", "feasible")) for g in _caps().get("grain", []))


def _find(items, text: str) -> list:
    """Items whose synonyms occur in text as whole words: longest match first, no overlaps, in order of appearance."""
    spans = []
    for it in items:
        for s in it.synonyms:
            for m in re.finditer(r"(?<![\w])" + re.escape(s) + r"(?![\w])", text, re.I):
                spans.append((m.start(), m.end(), it, m.group(0)))
    kept: list = []
    for a, b, it, word in sorted(spans, key=lambda x: (-(x[1] - x[0]), x[0])):
        if not any(a < b2 and a2 < b for a2, b2, *_ in kept):
            kept.append((a, b, it, word))
    out, seen = [], set()
    for _, _, it, word in sorted(kept, key=lambda x: x[0]):
        if it.key not in seen:
            seen.add(it.key)
            out.append(type(it)(**{**it.__dict__, "matched": word}))
    return out


def measures(text: str) -> list[Measure]:
    """The measures a question names. A vague word ('costs') is dropped when a specific measure is also named
    ('spend on consultants' = consultants)."""
    found = _find(all_measures(), text)
    specific = [m for m in found if not m.vague]
    return specific or found


def grains(text: str) -> list[Grain]:
    return _find(all_grains(), text)


def measure(key: str) -> Measure | None:
    return next((m for m in all_measures() if m.key == key), None)


def feasible(measure_key: str, grain: Grain) -> tuple[bool, str | None]:
    if "*" in grain.feasible or measure_key in grain.feasible:
        return True, None
    return False, grain.proxy


def currency(text: str) -> tuple[str, bool]:
    """(currency code named in the question or NOK, whether the cubes report it)."""
    cur = _caps().get("currency", {})
    words = cur.get("words", {})
    for m in re.finditer(r"[A-Za-zæøå]+", text):
        code = words.get(m.group(0).lower())
        if code:
            return code, code in cur.get("reporting", ["NOK"])
    return "NOK", True


def currency_proxy() -> str:
    return _caps().get("currency", {}).get("proxy", "")


def describe(m: Measure) -> str:
    """How the answer states a measure's definition: label, catalog line/KPI, note."""
    cat = catalog.load()
    if m.kpi:
        k = next(x for x in cat.kpis if x.key == m.kpi)
        return m.note or k.label
    if m.concept:
        line = cat.line_by_concept("FinMax", m.concept)
        base = f"{line.label} ({line.account})" if line else m.concept
        return f"{base}; {m.note}" if m.note else base
    base = m.label or m.key
    if m.accounts:
        base = f"{base} ({' + '.join(m.accounts)})"
    return f"{base}; {m.note}" if m.note else base
