"""Bridges: an EBITDA walk by P&L line (the leaf lines of the catalogue's identity chain) and by sub-unit (the top N
children by |variance| plus 'Other units' = the total variance less the top N, so that walk closes exactly)."""
from __future__ import annotations

import re
from dataclasses import dataclass, replace

from .facts import Fact, combine, missing, total
from .reconcile import TOL
from .variance import Variance


@dataclass(frozen=True)
class Step:
    key: str
    label: str
    fact: Fact


@dataclass(frozen=True)
class Bridge:
    start: Fact                       # the comparison total
    steps: tuple[Step, ...]
    end: Fact                         # the current total

    def residual(self) -> float | None:
        vals = [self.start.value, self.end.value, *(s.fact.value for s in self.steps)]
        if any(v is None for v in vals):
            return None
        return self.end.value - self.start.value - sum(s.fact.value for s in self.steps)

    @property
    def unexplained(self) -> float:
        """The walk's residual step when it is more than NOK 1 (a problem to report), else 0."""
        step = next((x for x in self.steps if x.key == "unexplained"), None)
        return step.fact.value if step is not None and step.fact.value is not None else 0.0


def _terms(expr: str) -> list[tuple[int, str]]:
    return [(-1 if s == "-" else 1, k) for s, k in re.findall(r"([+-]?)\s*([a-z_]\w*)", expr)]


def components(checks, lines, total_key: str) -> list[tuple[int, str]]:
    """The leaf lines (with sign) that add up to total_key, from identity checks of the form 'a+b-c' (c = a + b),
    in catalogue line order. FinMax ebitda -> rev, cogs, pers, oth, c73."""
    defs = {}
    for c in checks:
        terms = _terms(c.expr)
        neg = [k for s, k in terms if s < 0]
        if len(neg) == 1:
            defs[neg[0]] = [(s, k) for s, k in terms if s > 0]
    if total_key not in defs:
        raise ValueError(f"no identity check defines '{total_key}'")

    def expand(key: str, sign: int, depth: int) -> list[tuple[int, str]]:
        if key not in defs or depth > len(defs):
            return [(sign, key)]
        return [x for s, k in defs[key] for x in expand(k, sign * s, depth + 1)]
    order = {x.key: i for i, x in enumerate(lines)}
    return sorted(expand(total_key, 1, 0), key=lambda t: order.get(t[1], len(order)))


def by_line(start: Fact, end: Fact, variances: dict[str, Variance], leaves: list[tuple[int, str]],
            labels: dict[str, str], id: str = "bridge") -> Bridge:
    """start -> end through the variance of each leaf line, plus a 'Rounding / unexplained' step for any residual,
    so the walk always closes ('Unexplained' above NOK 1, else 'Rounding'; see Bridge.unexplained)."""
    steps = []
    for sign, key in leaves:
        v = variances.get(key)
        f = v.abs if v is not None else missing(f"bridge.{key}", "nok", f"variance of {key}", "no variance")
        if sign < 0:
            f = combine(f.id + ".neg", f.unit, f"-({f.definition})", lambda x: -x, f)
        steps.append(Step(key, labels.get(key, key), f))
    b = Bridge(start, tuple(steps), end)
    r = b.residual()
    if r is not None and abs(r) >= 0.005:                      # what the identities do not explain: shown, never hidden
        big = abs(r) > TOL
        gap = Fact(f"{id}.{'unexplained' if big else 'rounding'}", r, start.unit,
                   "end - start less the steps shown (" + ("unexplained: more than NOK 1" if big else "rounding") + ")",
                   tuple(dict.fromkeys(k for x in (start, end) for k in x.inputs)))
        key, label = ("unexplained", "Unexplained") if big else ("rounding", "Rounding")
        b = Bridge(start, b.steps + (Step(key, label, gap),), end)
    return b


def by_unit(start: Fact, end: Fact, units: list[tuple[str, str, Fact, Fact]], top_n: int = 5,
            id: str = "units") -> Bridge:
    """units: (member, label, current, comparison). The top_n units by |variance| (known variances only), then
    'Other units' = (end - start) - sum(top)."""
    vs = [(name, label, combine(f"{id}.{name}.var", "nok", f"{cur.id} - {cmp.id}", lambda a, b: a - b, cur, cmp))
          for name, label, cur, cmp in units]
    known = sorted((x for x in vs if x[2].value is not None), key=lambda x: -abs(x[2].value))
    top = known[:top_n]
    change = combine(f"{id}.total.var", "nok", f"{end.id} - {start.id}", lambda a, b: a - b, end, start)
    shown = total(f"{id}.top", [f for _, _, f in top], "nok", f"sum of the top {len(top)} units")
    other = combine(f"{id}.other", "nok", "total variance less the units shown", lambda a, b: a - b, change, shown)
    unknown = [lab for _, lab, f in vs if f.value is None]
    if unknown:
        other = replace(other, definition=f"{other.definition}; includes units with unknown variance: "
                                          + ", ".join(unknown))
    steps = tuple(Step(n, lab, f) for n, lab, f in top)
    if unknown or other.value is None or abs(other.value) >= 0.005:   # a zero 'other' is left out - unless units are unknown
        steps += (Step("other", "Other units and eliminations", other),)
    return Bridge(start, steps, end)
