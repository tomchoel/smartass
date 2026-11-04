# engine/smartass/model/analyze.py
"""The model's two entry points. plan_requests: guards, then what to fetch. compute: guards, blocks -> trees ->
values, the reconciliation guard, findings, material lines and key values. The data layer (Plan 1B) fetches between
the two; the render layer writes the same trees as formulas."""
from __future__ import annotations

from dataclasses import dataclass

from .blocks import BUILDERS
from .blocks.common import BlockOut, evaluate
from .context import Ctx
from .expr import Cells
from .findings import bridge_findings, key_values, material
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


def plan_requests(spec: AnalysisSpec, r: Resolver, closed: Closed) -> list[GridRequest]:
    guard_spec(spec, r, closed)
    return requests(spec, r, closed)


def _unique(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out = []
    for n in names:
        seen[n] = seen.get(n, 0) + 1
        out.append(n if seen[n] == 1 else f"{n} ({seen[n]})")
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
    cells = evaluate(outs, frame)
    guard_result(outs, cells)
    return Result(outs, cells, soft + sign_findings(ctx) + bridge_findings(outs, cells, r), material(outs, cells, r),
                  key_values(outs, cells, kv_limit))
