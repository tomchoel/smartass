"""Split a drill cell into child cells when it holds too many lines (children come from the atlas).
Order: period (year -> quarters -> months), then entity, then account. The caller recurses on pieces that are
still too big, so each call goes down one level of one dimension. Also: the cells EPM refused as too wide (kept on
disk) and the one-retrieve read of a split's child values."""
from __future__ import annotations

import logging
import re
from typing import Callable

from . import config

log = logging.getLogger("smartview")

ORDER = ("All Time Periods", "All Entities", "All Accounts")
_MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
_PERIOD = re.compile(r"^(?:([A-Za-z]{3})|Q([1-4]))\s+(\d{2})$")


def _period_key(name: str) -> tuple[int, int] | None:
    """'Sep 26' -> (26, 9), 'Q3 26' -> (26, 7); None if not a month or quarter."""
    m = _PERIOD.match(str(name).strip())
    if not m or (m.group(1) and m.group(1).lower() not in _MONTHS):
        return None
    sub = _MONTHS.index(m.group(1).lower()) + 1 if m.group(1) else 3 * int(m.group(2)) - 2
    return int(m.group(3)), sub


def chronological(names: list[str]) -> list[str]:
    """Periods in calendar order when every name is a month or quarter (the FinLite outline lists them newest
    first: Q3 26 -> Sep, Aug, Jul); anything else keeps its order."""
    keys = [_period_key(n) for n in names]
    return [n for _, n in sorted(zip(keys, names))] if names and None not in keys else list(names)


def plan_split(cell: dict[str, str], children: Callable[[str, str], list[str]],
               order: tuple[str, ...] = ORDER) -> list[dict[str, str]] | None:
    """Child cells along the first dimension in `order` whose member has children; None if all are leaves."""
    for dim in order:
        member = cell.get(dim)
        kids = children(dim, member) if member else []
        if kids:
            if dim == "All Time Periods":
                kids = chronological(kids)
            return [{**cell, dim: k} for k in kids]
    return None


def atlas_children(connection: str) -> Callable[[str, str], list[str]]:
    """children(dim, member) from the local atlas (instant)."""
    from . import atlas
    db = atlas.connect()

    def kids(dim: str, member: str) -> list[str]:
        return [m["name"] for m in atlas.children(db, connection, member) if m.get("dimension", dim) == dim]
    return kids


def too_wide_key(connection: str, cell: dict) -> tuple:
    return (connection, tuple(sorted(cell.items())))


_too_wide: set | None = None    # cells EPM refused as too wide (EPMFDM-140373): split them right away (kept on disk)


def known_too_wide() -> set:
    global _too_wide
    if _too_wide is None:
        import json
        try:
            raw = json.loads((config.SVKIT / "drill_too_wide.json").read_text(encoding="utf-8"))
            _too_wide = {(c, tuple(tuple(kv) for kv in cell)) for c, cell in raw}
        except Exception:   # noqa: BLE001 - none yet or unreadable: learn again
            _too_wide = set()
    return _too_wide


def remember_too_wide(key: tuple) -> None:
    import json
    known_too_wide().add(key)
    try:
        config.SVKIT.mkdir(parents=True, exist_ok=True)
        (config.SVKIT / "drill_too_wide.json").write_text(json.dumps(sorted(_too_wide)), encoding="utf-8")
    except Exception as e:   # noqa: BLE001 - only an optimisation
        log.info("too-wide memo not saved: %s", e)


def child_values_on(h, connection: str, parts: list[dict]) -> list | None:
    """Host job (a drill split): the cells differ in one dimension - retrieve them as one column (#Missing -> None)."""
    from . import ops
    from .grid import Axis
    first = parts[0]
    dim = next((k for k in first if any(p.get(k) != first[k] for p in parts)), None)
    if dim is None:
        return None
    col = "All Time Periods" if dim != "All Time Periods" else "All Accounts"
    pov = {**ops.DRILL_DEFAULTS, **{k: v for k, v in first.items() if k not in (dim, col)}}
    lay, p, _ = ops.query(h, connection, pov, [Axis(dim, [pc[dim] for pc in parts])], [Axis(col, [first[col]])],
                          False, False)
    if p is not None and p.pov_problems:              # a member was dropped: rows no longer line up
        return None
    ws = h.wb.Worksheets(ops.ensure_grid_sheet(h, connection))
    c = chr(ord("A") + lay.nr)
    out = []
    for i in range(len(parts)):
        v = ws.Range(f"${c}${2 + lay.nc + i}").Value
        out.append(v if isinstance(v, float) else None)
    return out
