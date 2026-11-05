# engine/smartass/model/analyze.py
"""The model's two entry points. plan_requests: guards, then what to fetch. compute: guards, blocks -> trees ->
values, the reconciliation guard, findings, material lines and key values. The data layer (Plan 1B) fetches between
the two; the render layer writes the same trees as formulas."""
from __future__ import annotations

from dataclasses import dataclass, field, fields, replace

from .blocks import BUILDERS
from .blocks.chart import implicit_chart
from .blocks.common import MAT_ABS, MAT_PCT, MAT_SMALL, BlockOut, evaluate
from .blocks.concept import concept_findings
from .context import Ctx
from .expr import Cells, Expr, Param, Ratio
from .findings import bridge_findings, key_values, lines_findings, material
from .frame import Frame
from .periods import Closed
from .requests import GridRequest, requests
from .resolver import Resolver
from .spec import AnalysisSpec, ChartBlock, has_chart
from .validate import Finding, guard_result, guard_spec, sign_findings


@dataclass
class Result:
    outs: list[BlockOut]
    cells: Cells
    findings: list[Finding]
    material: list[dict]
    key_values: dict[str, dict[str, float | None]]
    params: dict[str, tuple[str, float]] = field(default_factory=dict)   # Param key -> (label, value): the
    # labelled threshold cells the formulas reference (sheet: Thresholds)
    entity_label: str = ""      # the first query's entity name (the answer's subtitle)
    entity_labels: dict[str, str] = field(default_factory=dict)   # query id -> entity name (the data header)

    @property
    def floors(self) -> dict[str, float]:
        """metric -> variance % floor used (the floor.<metric> params)."""
        return {k[len(FLOOR_KEY):]: v for k, (_, v) in self.params.items() if k.startswith(FLOOR_KEY)}


FLOOR_KEY = "floor."
PARAM_LABELS = {FLOOR_KEY: "Variance % floor: {}", MAT_ABS: "Materiality: {} (NOK)", MAT_PCT: "Materiality %: {}",
                MAT_SMALL: "Materiality with %: {} (NOK)"}


def param_label(key: str) -> str:
    for prefix, label in PARAM_LABELS.items():
        if key.startswith(prefix):
            return label.format(key[len(prefix):])
    return key


def _params_of(e, out: dict[str, tuple[str, float]]) -> None:
    if isinstance(e, Param):
        out.setdefault(e.key, (param_label(e.key), e.value))
        return
    if isinstance(e, (tuple, list)):
        for x in e:
            _params_of(x, out)
        return
    if isinstance(e, Expr):
        for f in fields(e):
            _params_of(getattr(e, f.name), out)


def collect_params(outs: list[BlockOut]) -> dict[str, tuple[str, float]]:
    """Every Param the cells reference, in sheet order: the variance % floors first (as before), then the
    materiality thresholds of the Material columns."""
    found: dict[str, tuple[str, float]] = {}
    for o in outs:
        for ln in o.lines:
            for e in ln.cells:
                _params_of(e, found)
    floors = {k: v for k, v in found.items() if k.startswith(FLOOR_KEY)}
    return {**floors, **{k: v for k, v in found.items() if k not in floors}}


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
    outs = [o for o in outs if o.type != "chart"]     # a chart out has no lines (and its title is a message)
    for o, t in zip(outs, _unique([o.title for o in outs])):
        o.title = t
        for ln, lab in zip(o.lines, _unique([ln.label for ln in o.lines])):
            ln.label = lab


def compute(spec: AnalysisSpec, r: Resolver, closed: Closed, frame: Frame, kv_limit: int = 40) -> Result:
    soft = guard_spec(spec, r, closed)
    ctx = Ctx(spec, r, closed, frame)
    # chart blocks are built after the rest (keys stay b<i>): a chart title never renames a table (_uniquify), and
    # own chart data never takes a material line or a key_values slot from the table it repeats
    order = sorted(range(len(spec.blocks)), key=lambda i: isinstance(spec.blocks[i], ChartBlock))
    outs = [BUILDERS[spec.blocks[i].type](spec.blocks[i], ctx, f"b{i}") for i in order
            if spec.blocks[i].type in BUILDERS]
    _uniquify(outs)
    if not has_chart(spec):          # every answer has a chart; the stored spec stays as given
        extra = implicit_chart(outs)
        if extra is not None:
            outs.append(extra)
    apply_floors(outs, r)
    cells = evaluate(outs, frame)
    guard_result(outs, cells)
    return Result(outs, cells, soft + sign_findings(ctx) + bridge_findings(outs, cells, r) + lines_findings(outs, cells)
                  + concept_findings(outs, cells), material(outs, cells, r),
                  key_values(outs, cells, kv_limit), collect_params(outs), _entity_label(spec, r),
                  {q.id: _label(r, q) for q in spec.queries})


def _entity_label(spec: AnalysisSpec, r: Resolver) -> str:
    if not spec.queries:
        return ""
    return _label(r, spec.queries[0])


def _label(r: Resolver, q) -> str:
    try:
        return r.label(q.entity, q.cube)
    except Exception:   # noqa: BLE001 - a label is cosmetic
        return q.entity
