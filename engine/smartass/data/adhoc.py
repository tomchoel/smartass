"""The ad hoc fallback: grids the HTTP route could not fetch are pulled as Smart View ad hoc grids in the user's
Excel (v1 excel/adhoc.pull, which needs Smart View and the bridge), parsed by the same facts_from_parsed and the
temporary sheet removed. Runs as one user-Excel job; a grid that fails here is a warning, never a crash."""
from __future__ import annotations

from .provider import FetchResult, _grids, facts_from_parsed


def _target(h, target):
    from ..xl import target as xltarget
    return xltarget.resolve_target(h, target, allow_archive=False, title="Ad hoc fetch")


def _next_q(t):
    from ..xl import sheets
    return sheets.next_question(t)


def _pull(h, t, q, conn, pov, rows, cols):
    from ..xl import pull_adhoc as v1adhoc
    return v1adhoc.pull(h, t, q, conn, pov, rows, cols)


def _drop(t, name):
    from ..xl import sheets
    sheets.drop_parked(t, {name: name})      # drop_parked deletes every key of the mapping


def adhoc_fetch(h, reqs, target: str) -> dict:
    t = _target(h, target)
    q = _next_q(t)
    out: dict = {}
    for req in reqs:
        res = FetchResult(req, route="adhoc")
        for pov, rows, cols, labels in _grids(req):
            try:
                p = _pull(h, t, q, req.cube, pov, rows, cols)
            except Exception as e:   # noqa: BLE001 - one grid failing must not stop the others
                res.warnings.append(f"{req.cube}: ad hoc fallback failed - {e}")
                continue
            try:
                facts, warns = facts_from_parsed(req, p.parsed, labels)
                res.facts += facts
                res.warnings += warns
            finally:
                _drop(t, p.sheet)
        out[req] = res
    return out
