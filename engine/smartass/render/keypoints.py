# engine/smartass/render/keypoints.py
"""The Key points (sub-spec 5 §1.4): one live formula per point that builds the number sentence from cells of the
main table's last column group, plus the cause row from the commentary drivers, then the agent's free points (static
text), at most 5 in total. Plugs into layout(key_points=...); the number of points never depends on addr. Pure.
The sentence uses the base the line is material against (vs PY/BU/FC); a stock metric (FTE) says FTE, no MNOK; an
amount under 1 m says NOK in whole kroner (layout.unit_of: never '-0,0 MNOK'). A ledger lines main block (a
lines-only answer) gets its top 3 groups with their share and the lines total with the gap to the cube. A trend main
block gets one point per metric (its last month, the total for a NOK flow, against the base); a concept its total and
its largest items. A base missing on every line says 'no PY (new)', never the whole value as the variance. A line
without a driver gets no cause row."""
from __future__ import annotations

import textwrap
from collections.abc import Callable, Iterable

from ..model.blocks.common import SCEN, BlockOut, Line, base_of, ck, group_start
from ..model.spec import CommentaryBlock
from .layout import HUNDRED, NOTE_WIDTH, UNIT_MNOK, _last_period, sheet_label, unit_of

MAX_POINTS = 5
MAX_MATERIAL = 4
TOP_LINES = 3                    # a lines main block: its top 3 groups
BULLET = "•  "
INDENT = "   "
MAX_TREND = 4                    # a trend main block: one point per metric


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


SMALL_MNOK = 50_000              # under this an MNOK amount gets two decimals (never '0,0 MNOK')


def _amount(a: str, unit: str, addr: Callable[[str], str], value=None) -> str:
    """An amount cell as text in its unit: MNOK one decimal (two below 50 000), NOK whole kroner, FTE one
    decimal."""
    if unit == "MNOK":
        small = isinstance(value, (int, float)) and 0 < abs(value) < SMALL_MNOK
        return _minus(f"FIXED({a}/{addr(UNIT_MNOK)},{2 if small else 1})")
    return _minus(f"FIXED({a},0)" if unit == "NOK" else f"FIXED({a},1)")


def _minus(expr: str) -> str:
    """A number as text with the typographic minus (−, U+2212): text negatives are −1 000, never a hyphen."""
    return f'SUBSTITUTE({expr},"-","−")'


def shared_unit(values) -> str:
    """One unit for every amount of a set of key points (a controller reads one unit per section): MNOK when any
    amount is at least 1 m (or none is known), else NOK."""
    nums = [abs(v) for v in values if isinstance(v, (int, float)) and not isinstance(v, bool)]
    return unit_of(max(nums)) if nums else "MNOK"


def _formula(o: BlockOut, ln: Line, addr: Callable[[str], str], stock: set[str],
             base: str | None = None, cells=None, unit: str | None = None) -> str | None:
    vi, di, pi = _columns(o, base)
    if not _cell(ln, vi):
        return None
    stocked = bool(ln.metric and ln.metric in stock)
    unit = "FTE" if stocked else unit or unit_of((cells or {}).get(ck(o.key, ln.key, vi)))

    def num(j: int) -> str:
        return _amount(addr(ck(o.key, ln.key, j)), unit, addr, _value(cells or {}, ck(o.key, ln.key, j)))

    u = " " + unit
    head = " ".join(x for x in (sheet_label(ln), _last_period(o)) if x)
    f = f'="{BULLET}{_esc(head)} "&{num(vi)}'
    if not _cell(ln, di):
        gone = [SCEN.get(c, "") for c in o.no_base]          # the base was dropped as empty (a new product)
        return f + f'&"{u}' + (f'; {_no_base(gone[0])}"' if gone else '"')
    code = SCEN.get(o.cols[base_of(o.cols, di)], "")
    if cells and _missing(cells, ck(o.key, ln.key, di)):       # no base at all (a product launched this year)
        return f + f'&"{u}; {_no_base(code)}"'
    f += f'&"{u}; "&{num(di)}&"{u} vs {code}'
    if _cell(ln, pi):
        p, h = addr(ck(o.key, ln.key, pi)), addr(HUNDRED)
        f += f' ("&IF({p}="","n/a",{_minus(f"FIXED({p}*{h},1)")}&" %")&")'
    last = group_start(o.cols, len(o.cols))
    for dj in range(last, len(o.cols)):             # the other bases of the group: 'vs BU; ... vs FC' in one point
        if dj == di or o.cols[dj] != "var" or not _cell(ln, dj):
            continue
        if cells and _missing(cells, ck(o.key, ln.key, dj)):
            continue
        f += f'; "&{num(dj)}&"{u} vs {SCEN.get(o.cols[base_of(o.cols, dj)], "")}'
    return f + '"'


def _missing(cells, key: str) -> bool:
    try:
        return cells.get(key) is None
    except KeyError:
        return True


def _no_base(code: str) -> str:
    return "no PY (new)" if code == "PY" else f"no {code}"


def _cause(ln: Line, drivers: Iterable, material: bool = False) -> str | None:
    """The driver sentence of the line (or its metric), None without one: a line without a driver gets no cause row
    (the critic asks for drivers; the sheet shows no placeholder)."""
    names = {ln.label.strip().lower(), sheet_label(ln).strip().lower()} | ({ln.metric.lower()} if ln.metric else set())
    d = next((d for d in drivers if d.line.strip().lower() in names), None)
    if d is None:
        return None
    if d.basis == "data":
        return f"{INDENT}{d.text} (data: {d.evidence})"
    return f"{INDENT}{d.text} (hypothesis — not verified in data)"


def _chosen(main: BlockOut, res) -> list[tuple[Line, bool, str | None]]:
    """(line, material, base) for the material lines of the main table by |variance| descending (one point per
    line, at most 4, against the base of its largest variance); else the grand line, else the first line (first
    base)."""
    lines = main.lines
    if main.type == "drivers":          # the total and the leaf drivers; a middle level only sums its children
        from ..model.blocks.drivers import leaf_keys
        keep = leaf_keys(main) | {ln.key for ln in main.lines if ln.level == 0}
        lines = [ln for ln in main.lines if ln.key in keep]
    by_key = {f"{main.key}.{ln.key}": ln for ln in lines}
    seen: dict[str, tuple[float, str | None]] = {}
    for m in getattr(res, "material", None) or []:
        k = m.get("key")
        v = abs(m.get("variance") or 0.0)
        if k in by_key and by_key[k].kind != "check" and (k not in seen or v > seen[k][0]):
            seen[k] = (v, m.get("base"))
    if seen:
        keys = sorted(seen, key=lambda k: -seen[k][0])[:MAX_MATERIAL]
        if main.type in ("ranking", "drivers"):   # the parent first, then its parts (never between them)
            keys.sort(key=lambda k: not (by_key[k].kind == "grand" or (main.type == "drivers"
                                                                      and by_key[k].level == 0)))
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
    cells = getattr(res, "cells", None) or {}
    if main is not None and main.type == "lines":
        out += [(f, None) for f in _lines_points(main, addr, cells)]
    elif main is not None and main.type == "trend":
        out += _trend_points(main, res, drivers, addr, cells)
    elif main is not None:
        chosen = _concept_chosen(main, res, cells) if main.type == "concept" else _chosen(main, res)
        unit = _points_unit(main, chosen, cells)
        for ln, material, base in chosen:
            f = _formula(main, ln, addr, stock, base, cells, unit)
            if f is not None:
                out.append((f, _cause(ln, drivers, material)))
    out = _once(out)
    out += [(BULLET + t, None) for t in free]
    return out[:MAX_POINTS]


def _points_unit(o: BlockOut, chosen, cells) -> str:
    """The one amount unit of a table's or concept's points: by the largest value or variance they show."""
    vals = []
    for ln, _, base in chosen:
        vi, di, _ = _columns(o, base)
        vals += [_value(cells, ck(o.key, ln.key, j)) for j in (vi, di) if j is not None]
    return shared_unit(vals)


def _once(points: list[tuple[str, str | None]]) -> list[tuple[str, str | None]]:
    """'no PY (new)' (or 'no BU') is said once, on the first point that has it: every line of a new product has no
    PY, and saying it four times is noise."""
    said: set[str] = set()
    out = []
    for f, cause in points:
        for code in ("PY", "BU", "FC"):
            t = f"; {_no_base(code)}"
            if t in f:
                if code in said:
                    f = f.replace(t, "")
                said.add(code)
        out.append((f, cause))
    return out


def _concept_chosen(main: BlockOut, res, cells) -> list[tuple[Line, bool, str | None]]:
    """A concept: its total first, then its material items by |variance|, else its largest items by |value| (the
    items of the total), at most 4 points."""
    by_key = {ln.key: ln for ln in main.lines}
    total = by_key.get("total")
    out = [(total, False, None)] if total is not None else []
    steps = set(main.steps) - {"total"}           # memo items are outside the total: never a key point
    picked = [(ln, m, b) for ln, m, b in _chosen(main, res) if m and ln.key in steps]
    if not picked:
        vi = next((j for j, c in enumerate(main.cols) if c == "value"), None)
        items = [by_key[k] for k in main.steps if k in by_key and k != "total"] or [
            ln for ln in main.lines if ln.kind == "line" and ln.key.startswith("i_")]
        val = {ln.key: abs(_value(cells, ck(main.key, ln.key, vi)) or 0.0) for ln in items} if vi is not None else {}
        items = sorted((ln for ln in items if val.get(ln.key)), key=lambda ln: -val[ln.key])
        picked = [(ln, False, None) for ln in items]
    return (out + picked)[:MAX_MATERIAL]


def _value(cells, key: str) -> float | None:
    try:
        v = cells.get(key)
    except KeyError:
        return None
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _unit_text(unit: str, value, nok: str | None = None) -> tuple[str, str]:
    """(unit word, FIXED decimals) for a trend metric by its unit class: NOK amounts in MNOK or NOK (nok: the one
    unit of all the points), counts as whole numbers without a unit word, ARPU in whole NOK, FTE and % with one
    decimal."""
    if unit == "nok":
        u = nok or unit_of(value)
        return u, "mnok" if u == "MNOK" else "0"
    return {"count": ("", "0"), "per_unit": ("NOK", "0"), "fte": ("FTE", "1"), "pct": ("%", "1")}.get(unit, ("", "1"))


def _fixed(a: str, how: str, addr: Callable[[str], str]) -> str:
    return _minus(f"FIXED({a}/{addr(UNIT_MNOK)},1)" if how == "mnok" else f"FIXED({a},{how})")


def _trend_points(o: BlockOut, res, drivers, addr: Callable[[str], str], cells) -> list[tuple[str, str | None]]:
    """One point per metric of a trend main block (at most 4): '<metric> <last month>: <value>', vs the base in
    that month when the base has data (else 'no PY (new)'), and ' · total <period>: <sum>' for a NOK flow. The
    series shown is AC, or the first series with data (a budget-only answer: 'BU', and 'no AC')."""
    units = getattr(res, "units", None) or {}
    months = [ln for ln in o.lines if ln.kind == "line"]
    total = next((ln for ln in o.lines if ln.key == "total"), None)
    out: list[tuple[str, str | None]] = []
    nok = shared_unit(_value(cells, ck(o.key, ln.key, j)) for ln in [*months, *([total] if total else [])]
                      for j, m in enumerate(o.col_metrics) if units.get(m, "nok") == "nok")
    for m in list(dict.fromkeys(o.col_metrics))[:MAX_TREND]:
        js = [j for j, x in enumerate(o.col_metrics) if x == m]
        has = {j: [ln for ln in months if _value(cells, ck(o.key, ln.key, j)) is not None] for j in js}
        shown = next((j for j in js if o.cols[j] == "value" and has[j]), None)
        if shown is None:
            shown = next((j for j in js if has[j]), None)
        if shown is None:
            continue
        last = has[shown][-1]
        a = addr(ck(o.key, last.key, shown))
        unit = units.get(m, "nok")
        word, how = _unit_text(unit, _value(cells, ck(o.key, last.key, shown)), nok)
        tail = f" {word}" if word else ""
        name = o.headers[shown].rsplit(" ", 1)[0]
        scen = SCEN.get(o.cols[shown], "")
        head = f"{name} {last.label}" if o.cols[shown] == "value" else f"{name} {scen} {last.label}"
        f = f'="{BULLET}{_esc(head)}: "&{_fixed(a, how, addr)}&"{tail}'
        bases = [j for j in js if j != shown and o.cols[j] in ("ly", "budget", "forecast")]
        for b in bases[:1]:
            code = SCEN.get(o.cols[b], "")
            if not has[b]:
                f += f"; {_no_base(code)}"
            elif _value(cells, ck(o.key, last.key, b)) is not None:
                f += f' vs {code} "&{_fixed(addr(ck(o.key, last.key, b)), how, addr)}&"{tail}'
        if not bases and o.no_base:                          # the base was dropped as empty (a new product)
            f += "; " + _no_base(SCEN.get(o.no_base[0], ""))
        if o.cols[shown] != "value":
            f += "; no AC"
        if (total is not None and unit == "nok" and shown < len(total.cells) and total.cells[shown] is not None
                and o.period):
            t = ck(o.key, "total", shown)
            tw, thow = _unit_text(unit, _value(cells, t), nok)
            f += f' · total {_esc(o.period)}: "&{_fixed(addr(t), thow, addr)}&" {tw}'
        out.append((f + '"', _cause(Line(m, m, [], metric=m, display=name), drivers)))
    return out


LINES_CONTROLS = ("total", "cube", "gap", "other", "blank")    # not groups: totals, controls and the remainder


def _lines_points(o: BlockOut, addr: Callable[[str], str], cells) -> list[str]:
    """A ledger lines main block: '<group>: <amount> (<share> %)' for its top 3 groups (ranked by size), then
    'Ledger lines total <amount>; gap to cube <gap>'. Live formulas; one unit for the amounts (shared_unit), the gap in its own."""
    one = shared_unit(_value(cells, ck(o.key, ln.key, 0)) for ln in o.lines if ln.cells and ln.cells[0] is not None)

    def amount(key: str) -> str:
        return _amount(addr(ck(o.key, key, 0)), one, addr, _value(cells, ck(o.key, key, 0)))

    def unit(key: str) -> str:
        return " " + one
    out = []
    groups = [ln for ln in o.lines if ln.key not in LINES_CONTROLS and ln.kind == "line"][:TOP_LINES]
    for ln in groups:
        f = f'="{BULLET}{_esc(sheet_label(ln))}: "&{amount(ln.key)}&"{unit(ln.key)}'
        if _cell(ln, 1):
            p, h = addr(ck(o.key, ln.key, 1)), addr(HUNDRED)
            f += f' ("&IF({p}="","n/a",{_minus(f"FIXED({p}*{h},1)")}&" %")&")"'
        else:
            f += '"'
        out.append(f)
    keys = {ln.key for ln in o.lines}
    if "total" in keys:
        f = f'="{BULLET}Ledger lines total "&{amount("total")}&"{unit("total")}'
        if "gap" in keys:
            g = unit_of(cells.get(ck(o.key, "gap", 0)))           # a control gap: in its own unit (1 NOK, not 0,0 MNOK)
            f += f'; gap to cube "&{_amount(addr(ck(o.key, "gap", 0)), g, addr)}&" {g}'
        out.append(f + '"')
    return out


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
