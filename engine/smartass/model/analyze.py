# engine/smartass/model/analyze.py
"""The model's two entry points. plan_requests: guards, then what to fetch. compute: guards, blocks -> trees ->
values, the reconciliation guard, findings, material lines and key values. The data layer (Plan 1B) fetches between
the two; the render layer writes the same trees as formulas."""
from __future__ import annotations

from dataclasses import dataclass, field, replace

from .blocks import BUILDERS
from .blocks.common import BlockOut, evaluate
from .blocks.concept import concept_findings
from .context import Ctx
from .expr import Cells, Param, Ratio
from .findings import bridge_findings, key_values, lines_findings, material
from .frame import Frame
from .periods import Closed
from .requests import GridRequest, requests
from .resolver import Resolver
from .spec import AnalysisSpec
from .validate import Finding, guard_result, guard_spec, sign_findings


@dataclass
class Result:
    outs: list[BlockOut]
    cells: Cells
    findings: list[Finding]
    material: list[dict]
    key_values: dict[str, dict[str, float | None]]
    floors: dict[str, float] = field(default_factory=dict)    # metric -> variance % floor used (sheet: Thresholds)


FLOOR_KEY = "floor."


def apply_floors(outs: list[BlockOut], r: Resolver) -> dict[str, float]:
    """Every variance % of a line with a metric gets the metric's floor: below it the % is empty (|base| too small
    for a percentage to say anything). Returns the floors used, per metric, for the sheet's Thresholds cells."""
    used: dict[str, float] = {}
    for o in outs:
        for ln in o.lines:
            if not ln.metric:
                continue
            for j, e in enumerate(ln.cells):
                if o.cols[j] != "var_pct" or not isinstance(e, Ratio) or e.floor is not None:
                    continue
                f = r.pct_floor(ln.metric)
                if f is None:
                    continue
                used[ln.metric] = f
                ln.cells[j] = replace(e, floor=Param(FLOOR_KEY + ln.metric, f))
    return used


def plan_requests(spec: AnalysisSpec, r: Resolver, closed: Closed) -> list[GridRequest]:
    guard_spec(spec, r, closed)
    return requests(spec, r, closed)


def _unique(names: list[str]) -> list[str]:
    taken = set(names)          # every original name is reserved, so a suffix never lands on a later real name
    seen: set[str] = set()
    out = []
    for n in names:
        if n in seen:
            k = 2
            while f"{n} ({k})" in taken | seen:
                k += 1
            n = f"{n} ({k})"
        seen.add(n)
        out.append(n)
    return out


def _uniquify(outs: list[BlockOut]) -> None:
    """Block titles and, within a block, line labels become unique so key_values cannot overwrite itself."""
    for o, t in zip(outs, _unique([o.title for o in outs])):
        o.title = t
        for ln, lab in zip(o.lines, _unique([ln.label for ln in o.lines])):
            ln.label = lab


def compute(spec: AnalysisSpec, r: Resolver, closed: Closed, frame: Frame, kv_limit: int = 40) -> Result:
    soft = guard_spec(spec, r, closed)
    ctx = Ctx(spec, r, closed, frame)
    outs = [BUILDERS[b.type](b, ctx, f"b{i}") for i, b in enumerate(spec.blocks) if b.type in BUILDERS]
    _uniquify(outs)
    floors = apply_floors(outs, r)
    cells = evaluate(outs, frame)
    guard_result(outs, cells)
    return Result(outs, cells, soft + sign_findings(ctx) + bridge_findings(outs, cells, r) + lines_findings(outs, cells)
                  + concept_findings(outs, cells), material(outs, cells, r),
                  key_values(outs, cells, kv_limit), floors)
