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
from . import drilltime

log = logging.getLogger("smartview")
SOURCE = "Kilde: Data Integration drill-through (FinLite)."
EMPTY_PROVIDER = "Empty response from provider"
_ticks: list = []           # drill pages in flight (event loop only)
_tick = None                # their one 'still working' progress task


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
    One retry (with a note) when Smart View answers 'Empty response from provider'. Always live: nothing is cached."""
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
            say(f"Smart View gave an empty response for the drill link of {' × '.join(cell.values())}; "
                "retrying once")
            pieces, parent_cube = await _drill_cell(connection, cell, max_lines, split)
    source = SOURCE
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
    return DrillResult(cell, L["rows"], L["cols"], rec.get("cube_value"), rec, L["profile"], source,
                       round(time.monotonic() - t0, 1), L["piece_recs"] if len(pieces) > 1 else [], stats,
                       L["split_check"], warnings, stages)


def combine(parts: list[DrillResult]) -> DrillResult:
    """The drills of one alias (one cell per account) as one result: rows concatenated, columns in first-seen order."""
    if len(parts) == 1:
        return parts[0]
    cols = list(dict.fromkeys(c for p in parts for c in p.columns))
    rows = [r for p in parts for r in p.rows]
    vals = [p.cell_value for p in parts]
    gaps = [p.reconciliation.get("gap") for p in parts]
    cell = {**parts[0].cell, "All Accounts": ", ".join(str(p.cell.get("All Accounts")) for p in parts)}
    notes = list(dict.fromkeys(p.source.removeprefix(SOURCE) for p in parts if p.source != SOURCE))
    return DrillResult(cell, rows, cols, None if None in vals else sum(vals),
                       {"lines": len(rows), "gap": None if None in gaps else sum(gaps)}, None,
                       SOURCE + "".join(notes), round(sum(p.secs for p in parts), 1),
                       warnings=[w for p in parts for w in p.warnings],
                       stages={"drill": round(sum(p.stages.get("drill", 0.0) for p in parts), 2)})


async def drill_all(plan: list[dict], say=say) -> tuple[dict[str, DrillResult], list[str], str | None]:
    """A recipe's drills ([{alias, year, cells}], d1 first) -> ({alias: result}, warnings, why d1 failed or None).
    Every drill runs at once (Excel links queue on the host; the pages load in parallel tabs). A failed first drill
    drops the answer to the cube (no last year without this year); a later failure is a warning."""
    import asyncio

    async def one(d: dict):
        t0 = time.monotonic()
        res = combine(await _all(drill(cell, say=say) for cell in d["cells"]))
        say(f"Drilled {len(res.rows):,} lines for 20{d['year']:02d} ({time.monotonic() - t0:.0f} s)")
        return res
    tasks = [asyncio.ensure_future(one(d)) for d in plan]
    out, warnings = {}, []
    try:
        for i, (d, t) in enumerate(zip(plan, tasks)):
            try:
                res = await t
            except Exception as e:   # noqa: BLE001 - the answer falls back to the cube (D21)
                why = str(e)[:300]
                if i == 0:
                    return {}, [], why
                warnings.append(f"The drill of 20{d['year']:02d} ({d['alias']}) failed: {why}")
                continue
            out[d["alias"]] = res
            warnings += res.warnings
    finally:
        for t in tasks:
            t.cancel()
        await asyncio.gather(*tasks, return_exceptions=True)
    return out, warnings, None


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
    """One drill page through the tab pump (several at once, drill_tabs), with a hang guard: past the budget every
    running drill is stopped, only OUR headless browser is killed (the next drill starts a fresh one) and the tool
    fails readably instead of waiting forever."""
    import asyncio
    import math

    from ..drill_browser import DrillBusy
    from ..drill_tabs import shared_pump
    budget = 300 + 10 * math.ceil(max_lines / config.DRILL_PAGE)      # login + capture (+ one browser restart),
                                                                       # then ~10 s per page
    global _tick
    if _tick is None or _tick.done():               # one 'still working' note for all the drills running at once
        _tick = asyncio.create_task(_progress("drill (hidden browser: sign-in, drill pages, lines)"))
    _ticks.append(1)
    try:
        with drilltime.span("browser page"):
            return await asyncio.wait_for(asyncio.wrap_future(shared_pump().fetch(url, max_lines)), budget)
    except asyncio.TimeoutError:
        shared_pump().cancel_all(f"a drill page did not finish within {budget} s")
        await asyncio.to_thread(browser().kill)
        raise ToolError(f"The drill page did not finish within {budget} s; the hidden browser was restarted. "
                        "Try again, or narrow the cell.") from None
    except DrillBusy as e:
        raise ToolError(str(e)) from None
    finally:
        _ticks.pop()


async def _all(coros) -> list:
    """Run coroutines at once, results in order; the first failure cancels the rest and is raised as it is."""
    import asyncio
    tasks = []
    try:
        async with asyncio.TaskGroup() as tg:
            tasks = [tg.create_task(c) for c in coros]
    except BaseExceptionGroup as eg:   # noqa: F821 - builtin since Python 3.11
        raise eg.exceptions[0] from None
    return [t.result() for t in tasks]


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
        if not _ticks:                              # every drill page has finished
            return
        try:
            await ctx.report_progress(progress_secs(), None, f"still working: {what} ({progress_secs():.0f}s)")
        except Exception:   # noqa: BLE001 - a progress note must never break the drill
            return


class _KnownTooWide(Exception):
    def __str__(self):
        return "EPMFDM-140373 earlier in this session"


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
    link = await _call(lambda h: drilltime.timed("excel drill link", ops.drill_link, h, connection, cell),
                       what="drill link")
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
        if drill_split.too_wide_key(connection, cell) in drill_split.known_too_wide() and can_split:
            raise _KnownTooWide()                   # seen before: skip the ~20 s round trip to EPM's refusal
        res = await _drill_fetch(link["url"], limit)
    except (CellTooWide, _KnownTooWide) as e:
        # too many dimensions open at once: narrow by entity (then account) - a period split would not help
        parts = drill_split.plan_split(cell, drill_split.atlas_children(connection),
                                       order=("All Entities", "All Accounts")) if can_split else None
        if not parts:
            raise ToolError(f"EPM refused the drill on {cell} as too wide ({e}). Choose a lower entity or "
                            "account.") from None
        drill_split.remember_too_wide(drill_split.too_wide_key(connection, cell))
        return await _children(connection, parts, max_lines, split, state, depth)
    except TooManyLines as e:
        if e.total > left:
            raise ToolError(f"{cell} holds {e.total} lines, more than the {left} still allowed "
                            f"(max_lines={max_lines}). Narrow the cell or raise max_lines.") from None
        parts = drill_split.plan_split(cell, drill_split.atlas_children(connection)) if can_split else None
        if parts:                                   # too big: drill each child cell instead (recursively)
            return await _children(connection, parts, max_lines, split, state, depth)
        res = await _drill_fetch(link["url"], left)                       # leaf: page it all
    state["left"] -= len(res["items"])
    return [{"cell": cell, "cube_value": link["cube_value"], **res}]


async def _children(connection: str, parts: list[dict], max_lines: int, split: str, state: dict,
                    depth: int) -> list[dict]:
    """The split children drilled at once (their pages load in parallel tabs), pieces in the plan's order. Their
    values come first in ONE retrieve, so an empty (#Missing) child costs no drill link (most children of a wide
    entity are empty). Each starts with the budget free at that moment, so the total is checked again at the end."""
    vals = await _child_values(connection, parts)

    async def one(i: int, pc: dict) -> list[dict]:
        if vals is not None and vals[i] is None:
            return [{"cell": pc, "cube_value": None, "items": [], "view": None, "secs": 0,
                     "note": "empty cell (skipped)"}]
        return await _drill_piece(connection, pc, max_lines, split, state, depth + 1)
    done = await _all(one(i, pc) for i, pc in enumerate(parts))
    out = [p for ps in done for p in ps]
    if state["left"] < 0:
        n = sum(len(p.get("items") or []) for p in out)
        raise ToolError(f"The split pieces hold {n}+ lines, more than allowed (max_lines={max_lines}). Narrow the "
                        "cell or raise max_lines.")
    return out


async def _child_values(connection: str, parts: list[dict]) -> list | None:
    """The split children's cube values in one host job, or None (then every child is linked as before)."""
    from .. import drill_split
    try:
        return await _call(lambda h: drill_split.child_values_on(h, connection, parts), what="split values")
    except Exception as e:   # noqa: BLE001 - only an optimisation
        log.info("split values not read (%s); linking every child", e)
        return None


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
