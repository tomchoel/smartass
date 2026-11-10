"""Behaviour cards derived from data (spec 2026-11-10 §3; plan phase B). Pure functions: values in, card fields out.

Smart View gives no member properties on this provider (probe 10.11), so the cards are read from the numbers:
- operator of a child towards its parent: parent == sum of children -> '+'; exactly one child that must be left out
  -> '~'; exactly one that must be subtracted -> '-'; a parent that is 0 while its children carry data does not
  consolidate -> every child '~'. Anything else (several children could explain the gap, no child can) is unknown.
- time balance: FinPlan stores levels as levels (YearTotal = December); FinMax books every account as monthly
  movements (Yr = sum of the months), so a level shows as a January that carries the whole level followed by small
  changes.
- account type: the branch the member sits in (over every hierarchy it is shared into), and inside the P&L the
  stored sign (revenue +, costs -)."""
from __future__ import annotations

from collections.abc import Iterable, Mapping

Op = str | None
TOL_ABS, TOL_REL, TOL_SMALL, TOL_FLOOR = 1.0, 1e-6, 1e-4, 1e-6


def tol(scale: float) -> float:
    """The tolerance at a magnitude: 1 (NOK rounding) for money, relative 1e-6 for billions, but never more than
    1e-4 of the value itself, so ratios, ARPU and FTE (0.53, 739, 2 782) are not 'equal' within a whole unit
    (live 10.11: EBITDA % 0.534 vs December 0.483 passed as equal with a flat 1.0)."""
    return max(TOL_FLOOR, TOL_REL * scale, min(TOL_ABS, TOL_SMALL * scale))


def close(a: float, b: float) -> bool:
    return abs(a - b) <= tol(max(abs(a), abs(b)))


def same(a: float, b: float) -> bool:
    """Equal but for float noise: FinPlan computes its year totals from the same doubles (live 10.11 the mean of
    'No of FTE Total Average' matched to 1e-12, while EBITDA % - a ratio of the year totals - was within 1e-4 of the
    mean of its months, so 'close' would call a ratio an average)."""
    return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))


def _one(kids: list[str], val: dict[str, float | None], parent: str,
         blocked: set[str] = frozenset()) -> dict[str, Op]:
    """Operators of one parent's children from one slice; {} when nothing can be said. A member in `blocked`
    (#NoAccess / #Invalid in this slice) has no known value: a blocked parent says nothing; with a blocked child only
    'parent == sum of the visible children' still gives '+' for those (FinMax: 3 of 4 entities are #NoAccess to a
    normal user, live 10.11 - read as 0 they made 'non-consolidating' parents and '~' children that were not)."""
    if parent in blocked:
        return {}
    pv = val.get(parent)
    visible = [c for c in kids if c not in blocked]
    cv = {c: val.get(c) for c in visible}
    if pv is None and all(x is None for x in cv.values()):
        return {}
    pv = pv or 0.0
    s = sum(x or 0.0 for x in cv.values())
    with_data = [c for c, x in cv.items() if x]
    if close(pv, s):
        return {c: "+" for c in with_data}
    if len(visible) < len(kids):          # a hidden child may carry the gap: no '~' / '-' from this slice
        return {}
    if with_data and abs(pv) <= tol(max(abs(cv[c]) for c in with_data)):     # a zero parent, at the children's scale
        return {c: "~" for c in with_data}
    diff = pv - s
    hits = [(c, "~") for c in with_data if close(diff, -cv[c])] + \
           [(c, "-") for c in with_data if close(diff, -2 * cv[c])]
    if len(hits) != 1:                    # no child, or several, explain the gap: do not guess
        return {}
    c, op = hits[0]
    return {**{k: "+" for k in with_data}, c: op}


def operators(kids: dict[str, list[str]], *slices: dict[str, float | None],
              prior: dict[tuple[str, str], Op] | None = None,
              blocked: Iterable[set[str]] | None = None) -> dict[tuple[str, str], Op]:
    """{(parent, child): '+' | '~' | '-' | None} over every parent in `kids`. Later slices fill what earlier ones
    left unknown (a child without data in one slice), never overwrite a known operator. `blocked`: per slice, the
    members whose cell was #NoAccess / #Invalid."""
    out: dict[tuple[str, str], Op] = {(p, c): None for p, cs in kids.items() for c in cs}
    if prior:
        out.update({k: v for k, v in prior.items() if v is not None})
    blocks = list(blocked or [])
    for i, val in enumerate(slices):
        bl = blocks[i] if i < len(blocks) else set()
        for p, cs in kids.items():
            for c, op in _one(cs, val, p, bl).items():
                if out.get((p, c)) is None:
                    out[(p, c)] = op
    return out


STOCK_DOMINANCE = 3.0      # FinMax: |January| > 3 x every later movement
STOCK_ACTIVE_MONTHS = 6    # ... and at least 6 of the 11 later months move (else a January-only flow looks the same)


def time_balance(year: float | None, months: list[float | None], movements: bool = False) -> str | None:
    """flow | balance_last | average | None from the year total and the 12 months of one year.

    movements=False (FinPlan): the year equals exactly one of the last month (balance_last), the sum (flow) or the
    mean of the months with data (average: live 10.11 'No of FTE Total Average', 'Fixed Broadband FWA Access ARPU
    Recurring'); none or several (a flat series), or fewer than two months with data -> None. A ratio whose year is
    computed from the year totals (EBITDA %) matches none -> None.
    movements=True (FinMax, every account booked as monthly movements, Yr = sum): a level is booked whole in January
    and followed by small changes, so January dominates, Yr / January is near 1 and most later months move ->
    balance_last. A January that dominates while the later months are (nearly) empty is a one-off booking or an
    unchanged level: undecidable -> None. Otherwise flow. P&L accounts are flows whatever the numbers say: see
    settle_time_balance."""
    ms = [m for m in months if m is not None]
    if year is None or not ms:
        return None
    total = sum(ms)
    if movements:
        jan = months[0] or 0.0
        later = [abs(m or 0.0) for m in months[1:]]
        dominant = abs(jan) > TOL_ABS and abs(jan) > STOCK_DOMINANCE * max(later, default=0.0)
        if dominant and 0.5 <= year / jan <= 2.0:
            return "balance_last" if sum(1 for m in later if m > 0) >= STOCK_ACTIVE_MONTHS else None
        return "flow" if close(year, total) else None
    if len(ms) < 2:
        return None
    last = next(m for m in reversed(months) if m is not None)
    hits = [k for k, v in (("balance_last", last), ("flow", total), ("average", total / len(ms))) if same(year, v)]
    return hits[0] if len(hits) == 1 else None


def settle_time_balance(derived: str | None, account_type: str | None) -> str | None:
    """Revenue and expense accounts are flows: a P&L account booked mostly in January (live 10.11: FinMax ac.31204
    3.8 MNOK in January, a few hundred NOK later) looks like a level in the numbers, and a January-only booking is
    undecidable from them."""
    return "flow" if account_type in ("revenue", "expense") else derived


# Branch roots per cube (FinMax/FinLite 'ac.*'; FinPlan plain names - no name is in both cubes). The nearest root
# decides (account_type).
STAT_ROOTS = {"ac.STA", "ac.Statistics", "Statistics"}
BALANCE_ROOTS = {"ac.C1_C2", "Balance"}
REVENUE_ROOTS = {"ac.C30"}
EXPENSE_ROOTS = {"ac.C40S", "ac.C70", "ac.C50S", "ac.C60", "ac.C73", "ac.C76S", "ac.C75"}
PNL_ROOTS = {"ac.C3_C8", "P&L"}
_ROOTS = ((STAT_ROOTS, "statistic"), (BALANCE_ROOTS, "balance"), (REVENUE_ROOTS, "revenue"), (EXPENSE_ROOTS, "expense"),
          (PNL_ROOTS, "pnl"))


def account_type(ancestry: Mapping[str, int] | Iterable[str], value: float | None) -> str | None:
    """revenue | expense | statistic | balance | None from the member's ancestry over every hierarchy it sits in
    ({name: distance}, or a chain where the index is the distance; the member itself at 0). The NEAREST branch root
    decides (ties: statistic, balance, revenue, expense, P&L): the whole FinMax P&L is also shared into equity
    (ac.C89T under ac.C203 ... ac.C1_C2, live atlas 10.11), so 'any root above' would make every P&L account a
    balance. Inside the P&L outside the known roots the stored sign decides (revenue +, costs -). Outside these
    branches (ratios, calculations, export trees) None."""
    kind = branch(ancestry)
    if kind != "pnl":
        return kind
    if not value:
        return None
    return "revenue" if value > 0 else "expense"


def branch(ancestry: Mapping[str, int] | Iterable[str]) -> str | None:
    """statistic | balance | revenue | expense | pnl | None: the kind of the nearest branch root (see account_type)."""
    dist = dict(ancestry) if isinstance(ancestry, Mapping) else {n: i for i, n in reversed(list(enumerate(ancestry)))}
    best: tuple[int, int, str] | None = None
    for rank, (roots, kind) in enumerate(_ROOTS):
        for r in roots & dist.keys():
            if best is None or (dist[r], rank) < best[:2]:
                best = (dist[r], rank, kind)
    return best[2] if best else None


def account_types(kids: dict[str, list[str]], ancestries: dict[str, Mapping[str, int]],
                  values: dict[str, float | None]) -> dict[str, str | None]:
    """account_type for every member; a P&L total typed only by its sign (no root of its own) whose descendants are
    both revenue and expense is a mixed total (EBITDA, EBIT: positive, yet no revenue) -> None."""
    out = {n: account_type(a, values.get(n)) for n, a in ancestries.items()}
    below: dict[str, set[str]] = {}

    def kinds(n: str, path: frozenset = frozenset()) -> set[str]:
        if n not in below:
            got: set[str] = set()
            for c in kids.get(n, []):
                if c not in path:
                    got |= {out.get(c)} | kinds(c, path | {n})
            below[n] = got
        return below[n]
    for n, a in ancestries.items():
        if n in kids and branch(a) == "pnl" and {"revenue", "expense"} <= kinds(n):
            out[n] = None
    return out
