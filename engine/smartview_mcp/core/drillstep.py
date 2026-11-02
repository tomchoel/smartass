"""The async drill (ledger lines behind a FinLite cell) and its Excel half, shared by sv_drill and the recipes that
drill (lines_ranking). drill() runs on the event loop (Smart View link via the host, then the hidden browser);
write() runs inside a host job."""
from __future__ import annotations

import logging
import time
from contextlib import contextmanager
from dataclasses import dataclass, field

from mcp.server.mcpserver.exceptions import ToolError

from .. import config, ops

log = logging.getLogger("smartview")
SOURCE = "Kilde: Data Integration drill-through (FinLite)."
EMPTY_PROVIDER = "Empty response from provider"
_drill_lock = None          # asyncio.Lock: one browser fetch at a time, so the hang budget starts when it runs


@dataclass
class DrillResult:
    cell: dict
    rows: list[dict]
    columns: list[str]
    cell_value: float | None
    reconciliation: dict
    profile: dict | None
    source: str
    secs: float
    pieces: list[dict] = field(default_factory=list)      # per-piece reconciliation (only when split)
    text_repairs: dict | None = None
    split_check: float | None = None
    warnings: list[str] = field(default_factory=list)
    stages: dict = field(default_factory=dict)


def browser():
    from ..drill_browser import shared_browser
    return shared_browser()


async def _call(fn, what: str = "call"):
    from ..server import _call as call
    return await call(fn, what=what)


@contextmanager
def timed(stages: dict, name: str):
    """Stage time recorded in `stages` (stage() is thread-local, and the drill awaits on the event loop)."""
    t0 = time.monotonic()
    try:
        yield
    finally:
        stages[name] = round(stages.get(name, 0.0) + time.monotonic() - t0, 2)


def say(text: str) -> None:
    """A progress note to the client of the running tool call (event loop side). Never raises."""
    import asyncio
    log.info("note: %s", text)
    try:
        from ..server import _ctx, progress_secs
        ctx = _ctx.get()
        if ctx is not None:
            asyncio.get_running_loop().create_task(ctx.report_progress(progress_secs(), None, text))
    except Exception:   # noqa: BLE001 - a progress note must never break the drill
        pass


async def drill(cell: dict, *, say=say, connection: str = "FinLite", max_lines: int = config.DRILL_MAX_LINES,
                split: str = "auto") -> DrillResult:
    """Drill a cell to its ledger lines: link (Smart View), lines (hidden browser), normalise, reconcile, profile.
    One retry (with a note) when Smart View answers 'Empty response from provider'."""
    import asyncio

    from .. import drill_lines as dl
    t0, stages = time.monotonic(), {}
    cell, warnings = ops.drill_cell_defaults(cell)
    with timed(stages, "drill"):
        try:
            pieces, parent_cube = await _drill_cell(connection, cell, max_lines, split)
        except Exception as e:   # noqa: BLE001 - only the empty provider response is retried
            if EMPTY_PROVIDER.lower() not in str(e).lower():
                raise
            say(f"Smart View gave an empty response for the drill link of {' × '.join(cell.values())}; retrying once")
            pieces, parent_cube = await _drill_cell(connection, cell, max_lines, split)
    view = next((p["view"] for p in pieces if p.get("view")), None)
    if view is None:                                # spec 6: no drill report -> cube value, 0 lines, the reason
        total = parent_cube if parent_cube is not None else sum(p["cube_value"] or 0.0 for p in pieces)
        reasons = sorted({p.get("note") or "the drill page returned no view" for p in pieces})
        return DrillResult(cell, [], [], total, dl.reconcile(total, []), None, SOURCE,
                           round(time.monotonic() - t0, 1),
                           [{"cell": p["cell"], "cube_value": p["cube_value"], "note": p.get("note")} for p in pieces]
                           if len(pieces) > 1 else [],
                           warnings=warnings + ["No drill-through lines: " + "; ".join(reasons)], stages=stages)
    with timed(stages, "normalise + profile"):
        L = await asyncio.to_thread(_lines_of, pieces, view, parent_cube)
    rec, stats = L["rec"], L["stats"]
    if stats["unfixed"]:
        warnings.append("Text broken in EPM (letters lost at load): " + "; ".join(stats["unfixed"][:5])
                        + " - add corrections to %LOCALAPPDATA%\\smartass\\text_fixes.json")
    if rec["status"] == "gap":
        warnings.append(_gap_warning(rec))
    if L["split_check"]:
        warnings.append(f"The split pieces' cube values differ from the cell by {L['split_check']:,.2f} NOK (the "
                        "atlas children may be capped, out of date or shared): the lines are reconciled to the cell "
                        "itself.")
    return DrillResult(cell, L["rows"], L["cols"], rec.get("cube_value"), rec, L["profile"], SOURCE,
                       round(time.monotonic() - t0, 1), L["piece_recs"] if len(pieces) > 1 else [], stats,
                       L["split_check"], warnings, stages)


def write(t, q: int, alias: str, result: DrillResult, title: str, question: str = "", log: bool = True) -> dict:
    """The Excel half (inside a host job): 'Q<q> Lines[ alias]' (table Lines_Q<q>[_alias]) + a Sources row
    (log=False: the caller logs its own). 0 lines empties an existing table, so Excel matches the drill."""
    from .. import drill_xl
    from ..excel import sheets
    from ..grid import Layout
    if result.rows:
        info = drill_xl.write_lines(t, q, result.rows, result.columns, title=title, source=result.source, alias=alias)
    else:
        info = drill_xl.clear_lines(t, q, title=title, source=result.source, alias=alias)
    if info and log:
        rec = result.reconciliation

        class _Src:                                 # log_source wants a pull: the drilled cell is its POV
            sheet, connection = info["sheet"], "FinLite"
            layout = Layout([(d, m) for d, m in result.cell.items()], [], [])
        try:
            sheets.log_source(t, q, question or title, _Src,
                              notes=f"drill {alias}: {rec.get('lines', 0)} lines, gap {rec.get('gap')}")
        except Exception:   # noqa: BLE001 - the source log must never undo a finished drill
            pass
    return info


async def _drill_fetch(url: str, max_lines: int) -> dict:
    """browser().fetch, one at a time in this server, with a hang guard: past the budget only OUR headless Chrome
    is killed (the next drill starts a fresh one) and the tool fails readably instead of waiting forever."""
    import asyncio
    import math

    from ..drill_browser import DrillBusy
    global _drill_lock
    if _drill_lock is None:
        _drill_lock = asyncio.Lock()
    budget = 300 + 10 * math.ceil(max_lines / config.DRILL_PAGE)      # login + capture (+ one browser restart),
                                                                       # then ~10 s per page
    async with _drill_lock:
        tick = asyncio.create_task(_progress("drill (hidden browser: sign-in, drill page, lines)"))
        try:
            return await asyncio.wait_for(asyncio.wrap_future(browser().fetch(url, max_lines=max_lines)), budget)
        except asyncio.TimeoutError:
            await asyncio.to_thread(browser().kill)
            raise ToolError(f"The drill page did not finish within {budget} s; the hidden browser was restarted. "
                            "Try again, or narrow the cell.") from None
        except DrillBusy as e:
            raise ToolError(str(e)) from None
        finally:
            tick.cancel()


async def _progress(what: str) -> None:
    """'still working: <what> (n s)' to the client every timing.PROGRESS_AFTER_S (tools that declare ctx)."""
    import asyncio

    from .. import timing
    from ..server import _ctx, progress_secs
    ctx = _ctx.get()
    if ctx is None:
        return
    while True:
        await asyncio.sleep(timing.PROGRESS_AFTER_S)
        try:
            await ctx.report_progress(progress_secs(), None, f"still working: {what} ({progress_secs():.0f}s)")
        except Exception:   # noqa: BLE001 - a progress note must never break the drill
            return


_too_wide: set = set()      # cells EPM refused as too wide (EPMFDM-140373) in this server: split them right away


class _KnownTooWide(Exception):
    def __str__(self):
        return "EPMFDM-140373 earlier in this session"


def _too_wide_key(connection: str, cell: dict) -> tuple:
    return (connection, tuple(sorted(cell.items())))


async def _drill_cell(connection: str, cell: dict, max_lines: int, split: str) -> tuple[list[dict], float | None]:
    """Drill a cell (split into child cells when too big). Returns (pieces, the cell's own cube value)."""
    state = {"left": max_lines, "cube": None}
    pieces = await _drill_piece(connection, cell, max_lines, split, state)
    return pieces, state["cube"]


async def _drill_piece(connection: str, cell: dict, max_lines: int, split: str, state: dict,
                       depth: int = 0) -> list[dict]:
    """Drill one cell; if it holds more than DRILL_SPLIT_AT lines (and split='auto'), drill its children instead.
    state['left'] is the line budget still free (max_lines minus the lines already fetched): a cell bigger than
    that fails with its size BEFORE any split (spec 4). Returns pieces: [{cell, cube_value, items, view, secs}]."""
    from .. import drill_split
    from ..drill_browser import CellTooWide, TooManyLines
    link = await _call(lambda h: ops.drill_link(h, connection, cell), what="drill link")
    if depth == 0:
        state["cube"] = link["cube_value"]
    elif link["cube_value"] is None:                # an empty (#Missing) child of a split: nothing to drill
        return [{"cell": cell, "cube_value": link["cube_value"], "items": [], "view": None, "secs": 0,
                 "note": "empty cell (skipped)"}]
    if not link["url"]:
        return [{"cell": cell, "cube_value": link["cube_value"], "items": [], "view": None, "secs": 0,
                 "note": "this cell has no drill-through report: it is not a FinLite base-level intersection "
                         "loaded through Data Integration (use a FinLite account and a base entity/period)"}]
    can_split = split == "auto" and depth < 4
    left = state["left"]
    limit = min(config.DRILL_SPLIT_AT, left) if can_split else left
    try:
        if _too_wide_key(connection, cell) in _too_wide and can_split:
            raise _KnownTooWide()                   # seen before: skip the ~20 s round trip to EPM's refusal
        res = await _drill_fetch(link["url"], limit)
    except (CellTooWide, _KnownTooWide) as e:
        # too many dimensions open at once: narrow by entity (then account) - a period split would not help
        parts = drill_split.plan_split(cell, drill_split.atlas_children(connection),
                                       order=("All Entities", "All Accounts")) if can_split else None
        if not parts:
            raise ToolError(f"EPM refused the drill on {cell} as too wide ({e}). Choose a lower entity or "
                            "account.") from None
        _too_wide.add(_too_wide_key(connection, cell))
        out = []
        for pc in parts:
            out += await _drill_piece(connection, pc, max_lines, split, state, depth + 1)
        return out
    except TooManyLines as e:
        if e.total > left:
            raise ToolError(f"{cell} holds {e.total} lines, more than the {left} still allowed "
                            f"(max_lines={max_lines}). Narrow the cell or raise max_lines.") from None
        parts = drill_split.plan_split(cell, drill_split.atlas_children(connection)) if can_split else None
        if parts:                                   # too big: drill each child cell instead (recursively)
            out = []
            for pc in parts:
                out += await _drill_piece(connection, pc, max_lines, split, state, depth + 1)
            return out
        res = await _drill_fetch(link["url"], left)                       # leaf: page it all
    state["left"] -= len(res["items"])
    return [{"cell": cell, "cube_value": link["cube_value"], **res}]


def _lines_of(pieces: list[dict], view: dict, parent_cube: float | None = None) -> dict:
    """CPU work off the event loop: normalise each piece ONCE, reconcile each piece, and the whole against the
    PARENT cell's own cube value; split_check = sum of the piece cubes - parent (None without a split)."""
    from .. import drill_lines as dl
    cols = dl.columns_from_view(view)
    fixes = dl.load_text_fixes(config.TEXT_FIXES, config.USER_TEXT_FIXES)
    rows, piece_recs = [], []
    stats = {"fixed_mojibake": 0, "fixed_by_list": 0, "unfixed": set()}
    for p in pieces:
        pr, st = dl.normalize(p["items"], cols, fixes)
        rows += pr
        stats["fixed_mojibake"] += st["fixed_mojibake"]
        stats["fixed_by_list"] += st["fixed_by_list"]
        stats["unfixed"].update(st["unfixed"])
        piece_recs.append({"cell": p["cell"], **dl.reconcile(p["cube_value"], pr),
                           **({"note": p["note"]} if p.get("note") else {})})
    stats["unfixed"] = sorted(stats["unfixed"])
    piece_cubes = round(sum(p["cube_value"] or 0.0 for p in pieces), 2)
    parent = piece_cubes if parent_cube is None else parent_cube
    rec = dl.reconcile(parent, rows)
    split_check = round(piece_cubes - round(float(parent), 2), 2) if len(pieces) > 1 and parent is not None else None
    return {"cols": [c[0] for c in cols], "rows": rows, "stats": stats, "piece_recs": piece_recs, "rec": rec,
            "profile": dl.profile(rows), "split_check": split_check}


def _gap_warning(rec: dict) -> str:
    share = f" ({abs(rec['gap'] / rec['cube_value']) * 100:.1f} % of the cell)" if rec.get("cube_value") else ""
    return (f"Lines differ from the cube by {rec['gap']:,.2f} NOK{share} (data not loaded via Data Integration, "
            "or postings after the load).")
