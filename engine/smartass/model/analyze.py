# engine/smartass/model/analyze.py
"""The model's two entry points. plan_requests: guards, then what to fetch. compute: guards, blocks -> trees ->
values, the reconciliation guard, findings, material lines and key values. The data layer (Plan 1B) fetches between
the two; the render layer writes the same trees as formulas."""
from __future__ import annotations

from dataclasses import dataclass, field, fields, replace

from .blocks import BUILDERS
from .blocks.chart import implicit_chart
from .blocks.common import MAT_ABS, MAT_PCT, MAT_SMALL, BlockOut, base_of, ck, evaluate
from .blocks.concept import concept_findings
from .context import Ctx
from .expr import Cells, Expr, Material, Num, Param, Ratio, Ref, Sub, Var
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
    units: dict[str, str] = field(default_factory=dict)   # metric token -> unit class (model/units.py): charts
    # never mix units, Key points and the subtitle say NOK only for amounts

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
    if missing_bases(outs, cells):     # a base missing on every line: its variance is missing too (re-evaluated)
        cells = evaluate(outs, frame)
    guard_result(outs, cells)
    return Result(outs, cells, soft + sign_findings(ctx) + bridge_findings(outs, cells, r) + lines_findings(outs, cells)
                  + concept_findings(outs, cells), material(outs, cells, r),
                  key_values(outs, cells, kv_limit), collect_params(outs), _entity_label(spec, r),
                  {q.id: _label(r, q) for q in spec.queries}, metric_units(spec, outs, r))


def metric_units(spec: AnalysisSpec, outs: list[BlockOut], r: Resolver) -> dict[str, str]:
    """The unit class of every metric a block shows (its columns' metrics and its lines' metrics), in the cube of
    the block's query."""
    from .units import unit_class
    cube = {q.id: q.cube for q in spec.queries}
    out: dict[str, str] = {}
    for o in outs:
        c = cube.get(o.query or "", "FinMax")
        for m in [*o.col_metrics, *(ln.metric for ln in o.lines if ln.metric)]:
            if m not in out:
                try:
                    out[m] = unit_class(r, m, c)
                except Exception:   # noqa: BLE001 - a unit is cosmetic
                    out[m] = "other"
    return out


def _entity_label(spec: AnalysisSpec, r: Resolver) -> str:
    if not spec.queries:
        return ""
    return _label(r, spec.queries[0])


def _label(r: Resolver, q) -> str:
    try:
        return r.label(q.entity, q.cube)
    except Exception:   # noqa: BLE001 - a label is cosmetic
        return q.entity


def missing_bases(outs: list[BlockOut], cells: Cells) -> bool:
    """A variance column whose base column is missing (None) on every line of its block - no PY for a product
    launched this year, a FinPlan answer with only budget data - gets Var cells: missing, not the whole current
    value. A base missing on only some lines (one cost centre without last year) keeps current - 0, as before.
    Returns True when a cell changed (the caller evaluates again)."""
    changed = False
    for o in outs:
        for j, c in enumerate(o.cols):
            if c != "var":
                continue
            try:
                bi = base_of(o.cols, j)
            except StopIteration:
                continue
            lines = [ln for ln in o.lines if ln.kind != "check" and bi < len(ln.cells) and ln.cells[bi] is not None]
            if not lines or any(cells.get(ck(o.key, ln.key, bi)) is not None for ln in lines):
                continue
            for ln in o.lines:
                e = ln.cells[j] if j < len(ln.cells) else None
                if isinstance(e, Sub) and not isinstance(e, Var):
                    ln.cells[j] = Var(e.a, e.b)
                    changed = True
                    _blank_safe(ln, ck(o.key, ln.key, j))
    return changed


def _blank_safe(ln, key: str) -> None:
    """The line's Material marker reads the (now possibly "") variance cell through N(), never ABS("")."""
    def wrap(x):
        return Num(x) if isinstance(x, Ref) and x.key == key else x
    for i, c in enumerate(ln.cells):
        if isinstance(c, Material):
            ln.cells[i] = replace(c, var=wrap(c.var), more=tuple((wrap(v), p) for v, p in c.more))
