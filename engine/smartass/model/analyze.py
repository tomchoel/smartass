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
from .expr import Cells, Expr, Material, Num, Param, Ratio, Ref, Sub, Sum, Var
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
    if drop_empty_bases(outs, cells, r):   # a base missing on every line: its columns go (no empty ΔPY columns)
        cells = evaluate(outs, frame)
    if drop_leading_months(outs, cells):   # a trend starts at its first month with data (a product's launch)
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


DROP_TYPES = ("table", "concept", "kpi_tiles", "trend", "ranking")


def _remap(e, f):
    """The expression with every Ref key passed through f (frozen dataclasses, nested tuples)."""
    if isinstance(e, Ref):
        return Ref(f(e.key))
    if isinstance(e, tuple):
        return tuple(_remap(x, f) for x in e)
    if isinstance(e, Expr) and hasattr(e, "__dataclass_fields__"):
        ch = {fl.name: _remap(getattr(e, fl.name), f) for fl in fields(e)
              if isinstance(getattr(e, fl.name), (Expr, tuple))}
        return replace(e, **ch) if ch else e
    return e


def _prune(e, dead: set[str]):
    """The expression without Refs to dead cells: a Sum keeps its other items (none left: None)."""
    if isinstance(e, Ref):
        return None if e.key in dead else e
    if isinstance(e, Sum):
        items = tuple(x for x in (_prune(i, dead) for i in e.items) if x is not None)
        return Sum(items) if items else None
    return e


def drop_leading_months(outs: list[BlockOut], cells: Cells) -> bool:
    """A trend whose first months are empty on every column (a product launched in Nov 25 asked from Jan 25) starts
    at its first month with data: the empty rows go, the Total and cumulative sums lose their refs to them, and the
    block's period label follows. A trend without any data keeps its rows. Returns True when a block changed."""
    from .periods import parse, period_label
    changed = False
    for o in outs:
        if o.type != "trend":
            continue
        months = [ln for ln in o.lines if ln.kind == "line"]
        lead = []
        for ln in months:
            if any(cells.get(ck(o.key, ln.key, j)) is not None for j, e in enumerate(ln.cells) if e is not None):
                break
            lead.append(ln)
        if not lead or len(lead) == len(months):
            continue
        gone = {ln.key for ln in lead}
        dead = {ck(o.key, k, j) for k in gone for j in range(len(o.cols))}
        o.lines = [ln for ln in o.lines if ln.key not in gone]
        for ln in o.lines:
            ln.cells = [_prune(e, dead) if e is not None else None for e in ln.cells]
        try:
            o.period = period_label([parse(ln.label) for ln in o.lines if ln.kind == "line"])
        except Exception:   # noqa: BLE001 - a label is cosmetic
            pass
        changed = True
    return changed


def _refs_col(e, prefix: str, cols: set[int]) -> bool:
    """True when the expression holds a Ref to one of the block's columns in cols."""
    if isinstance(e, Ref):
        head, _, j = e.key.rpartition(".")
        return e.key.startswith(prefix) and j.isdigit() and int(j) in cols
    if isinstance(e, tuple):
        return any(_refs_col(x, prefix, cols) for x in e)
    if isinstance(e, Expr) and hasattr(e, "__dataclass_fields__"):
        return any(_refs_col(getattr(e, fl.name), prefix, cols) for fl in fields(e))
    return False


def drop_empty_bases(outs: list[BlockOut], cells: Cells, r: Resolver) -> bool:
    """A base column (PY/BU/FC) missing on every line of a table-like block - no PY for a product launched this
    year - is dropped with its variance and variance % columns (an empty ΔPY column says nothing; the key points say
    'no PY (new)' from o.no_base). The Material column is rebuilt over the remaining variances (none left: no
    Material). Ref keys of the block are re-indexed. A base missing on only some lines stays (current - 0).
    Returns True when a block changed (the caller evaluates again)."""
    from .blocks.common import BASES, MATERIAL, add_material
    changed = False
    for o in outs:
        if o.type not in DROP_TYPES or not o.lines:
            continue
        empty = []
        for bi, c in enumerate(o.cols):
            # a trend's AC column empty on every month (a FinPlan KPI with only Target) goes too: the key points
            # say 'no AC'; a table keeps its value column (its variances read it)
            if c not in BASES and not (o.type == "trend" and c == "value"):
                continue
            lines = [ln for ln in o.lines if ln.kind != "check" and bi < len(ln.cells) and ln.cells[bi] is not None]
            if lines and all(cells.get(ck(o.key, ln.key, bi)) is None for ln in lines):
                empty.append(bi)
        if not empty:
            continue
        drop = set(empty)
        for j, c in enumerate(o.cols):
            if c in ("var", "var_pct"):
                try:
                    if base_of(o.cols, j) in drop:
                        drop.add(j)
                except StopIteration:
                    pass
        had_material = MATERIAL in o.cols
        if had_material:
            drop.add(o.cols.index(MATERIAL))
        prefix = f"{o.key}."
        grown = True
        while grown:                   # a column that reads a dropped column goes too (concept helper columns)
            grown = False
            for j in range(len(o.cols)):
                if j in drop:
                    continue
                if any(_refs_col(ln.cells[j], prefix, drop) for ln in o.lines if j < len(ln.cells)):
                    drop.add(j)
                    grown = True
        keep = [j for j in range(len(o.cols)) if j not in drop]
        new = {old: i for i, old in enumerate(keep)}

        def f(key, new=new, prefix=prefix):
            if not key.startswith(prefix):
                return key
            head, _, j = key.rpartition(".")
            return f"{head}.{new[int(j)]}" if j.isdigit() and int(j) in new else key
        if len(keep) == 0 or all(o.cols[j] == "gap" for j in keep):
            continue                   # nothing would be left: keep the block as it is
        for code in dict.fromkeys(o.cols[j] for j in empty if o.cols[j] in BASES):
            if code not in o.no_base:
                o.no_base.append(code)
        o.dropped += [c for c in dict.fromkeys(o.cols[j] for j in empty) if c not in o.dropped]
        o.cols = [o.cols[j] for j in keep]
        o.headers = [o.headers[j] for j in keep]
        o.keys = [o.keys[j] for j in keep]
        if o.col_metrics:
            o.col_metrics = [o.col_metrics[j] for j in keep if j < len(o.col_metrics)]
        for ln in o.lines:
            ln.cells = [_remap(ln.cells[j], f) if j < len(ln.cells) else None for j in keep]
        if had_material and "var" in o.cols:
            add_material(o, r)
        changed = True
    return changed


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
