# engine/smartass/render/keypoints.py
"""The Key points (sub-spec 5 §1.4): one live formula per point that builds the number sentence from cells of the
main table's last column group, plus the cause row from the commentary drivers, then the agent's free points (static
text), at most 5 in total. Plugs into layout(key_points=...); the number of points never depends on addr. Pure.
The sentence uses the base the line is material against (vs PY/BU/FC); a stock metric (FTE) says FTE, no MNOK."""
from __future__ import annotations

import textwrap
from collections.abc import Callable, Iterable

from ..model.blocks.common import SCEN, BlockOut, Line, base_of, ck, group_start
from ..model.spec import CommentaryBlock
from .layout import HUNDRED, NOTE_WIDTH, UNIT_MNOK, _last_period

MAX_POINTS = 5
MAX_MATERIAL = 4
BULLET = "•  "
INDENT = "   "
NO_CAUSE = INDENT + "Cause not yet explained"


def _esc(s: str) -> str:
    """Text inside an Excel string literal: a double quote is doubled."""
    return s.replace('"', '""')


def _columns(o: BlockOut, base: str | None = None) -> tuple[int | None, int | None, int | None]:
    """(value, var, var_pct against the same base) column indexes in the last column group: the var against `base`
    (the base the line is material against), else the first var."""
    last = group_start(o.cols, len(o.cols))
    vi = next((j for j in range(last, len(o.cols)) if o.cols[j] == "value"), None)
    vars_ = [j for j in range(last, len(o.cols)) if o.cols[j] == "var"]
    di = next((j for j in vars_ if o.cols[base_of(o.cols, j)] == base), vars_[0] if vars_ else None)
    pi = None
    if di is not None:
        b = o.cols[base_of(o.cols, di)]
        pi = next((j for j in range(last, len(o.cols))
                   if o.cols[j] == "var_pct" and o.cols[base_of(o.cols, j)] == b), None)
    return vi, di, pi


def _cell(ln: Line, j: int | None) -> bool:
    return j is not None and j < len(ln.cells) and ln.cells[j] is not None


def _formula(o: BlockOut, ln: Line, addr: Callable[[str], str], stock: set[str],
             base: str | None = None) -> str | None:
    vi, di, pi = _columns(o, base)
    if not _cell(ln, vi):
        return None
    mnok = not (ln.metric and ln.metric in stock)
    unit = addr(UNIT_MNOK) if mnok else ""

    def num(j: int) -> str:
        a = addr(ck(o.key, ln.key, j))
        return f"FIXED({a}/{unit},1)" if mnok else f"FIXED({a},1)"

    u = " MNOK" if mnok else " FTE"
    head = " ".join(x for x in (ln.label, _last_period(o)) if x)
    f = f'="{BULLET}{_esc(head)} "&{num(vi)}'
    if not _cell(ln, di):
        return f + f'&"{u}"'
    code = SCEN.get(o.cols[base_of(o.cols, di)], "")
    f += f'&"{u}; "&{num(di)}&"{u} vs {code}'
    if not _cell(ln, pi):
        return f + '"'
    p, h = addr(ck(o.key, ln.key, pi)), addr(HUNDRED)
    return f + f' ("&IF({p}="","n/a",FIXED({p}*{h},1)&" %")&")"'


def _cause(ln: Line, drivers: Iterable, material: bool) -> str | None:
    names = {ln.label.strip().lower()} | ({ln.metric.lower()} if ln.metric else set())
    d = next((d for d in drivers if d.line.strip().lower() in names), None)
    if d is None:
        return NO_CAUSE if material else None
    if d.basis == "data":
        return f"{INDENT}{d.text} (data: {d.evidence})"
    return f"{INDENT}{d.text} (hypothesis — not verified in data)"


def _chosen(main: BlockOut, res) -> list[tuple[Line, bool, str | None]]:
    """(line, material, base) for the material lines of the main table by |variance| descending (one point per
    line, at most 4, against the base of its largest variance); else the grand line, else the first line (first
    base)."""
    by_key = {f"{main.key}.{ln.key}": ln for ln in main.lines}
    seen: dict[str, tuple[float, str | None]] = {}
    for m in getattr(res, "material", None) or []:
        k = m.get("key")
        v = abs(m.get("variance") or 0.0)
        if k in by_key and by_key[k].kind != "check" and (k not in seen or v > seen[k][0]):
            seen[k] = (v, m.get("base"))
    if seen:
        keys = sorted(seen, key=lambda k: -seen[k][0])[:MAX_MATERIAL]
        return [(by_key[k], True, seen[k][1]) for k in keys]
    lines = [ln for ln in main.lines if ln.kind != "check"]
    top = next((ln for ln in lines if ln.kind == "grand"), lines[0] if lines else None)
    return [(top, False, None)] if top is not None else []


def key_points(main: BlockOut | None, res, drivers, addr: Callable[[str], str], *, stock: Iterable[str] = (),
               free: Iterable[str] = ()) -> list[tuple[str, str | None]]:
    """(formula or text, cause or None) per point: the engine's live points (none without a main table), then the
    free points ('•  ' + text, so a point starting with '=' stays text), at most 5 in total."""
    stock = set(stock)
    out: list[tuple[str, str | None]] = []
    for ln, material, base in (_chosen(main, res) if main is not None else []):
        f = _formula(main, ln, addr, stock, base)
        if f is not None:
            out.append((f, _cause(ln, drivers, material)))
    out += [(BULLET + t, None) for t in free]
    return out[:MAX_POINTS]


def point_rows(points: list[tuple[str, str | None]]) -> list[str]:
    """The sheet rows of the points: a formula is one row; a static point and a cause wrap like notes (NOTE_WIDTH,
    at word boundaries) with indented continuation rows, so no row is cut off and none starts with '='. Depends only
    on the texts, never on addr (the layout's two passes reserve and fill the same rows)."""
    rows: list[str] = []
    for point, cause in points:
        for t in (point, cause):
            if t is None:
                continue
            if t.startswith("="):
                rows.append(t)
            else:
                rows += textwrap.wrap(t, NOTE_WIDTH, subsequent_indent=INDENT) or [t]
    return rows


def free_points(spec) -> tuple[list[str], list[str]]:
    """The agent's free key points (every commentary block, in order) and a warning per block with more than 5:
    its first 5 are kept (only a new question refuses them: validate.guard_key_points)."""
    pts: list[str] = []
    warnings: list[str] = []
    for b in spec.blocks:
        if isinstance(b, CommentaryBlock):
            if len(b.key_points) > MAX_POINTS:
                warnings.append(f"commentary '{b.title}': {len(b.key_points)} key points, the first {MAX_POINTS} "
                                "were kept; fix: keep at most 5")
            pts += b.key_points[:MAX_POINTS]
    return pts, warnings
