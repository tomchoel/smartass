"""The async drill (ledger lines behind a FinLite cell) and its Excel half, shared by fin_drill and the recipes that
drill (lines_ranking). drill() runs on the event loop (the drill link over HTTP, then the lines over the user's
token: drill_token); write() runs inside a host job."""
from __future__ import annotations

import contextvars
import itertools
import logging
import math
import time
from contextlib import contextmanager
from dataclasses import dataclass, field

from mcp.server.mcpserver.exceptions import ToolError

from ..core import config, timing
from . import ops
from . import drilltime, oauth
from .drill_token import _REQUEST_S
from .drillcore import CellTooWide, NoDrillData, TooManyLines

log = logging.getLogger("smartview")
SOURCE = "Kilde: Data Integration drill-through (FinLite)."
EMPTY_PROVIDER = "Empty response from provider"
# A drill piece's time budget (piece_budget). Live, a 5 000-line getData page took about 11 s (audit (b): 7 852 lines
# in two pages, 22.8-23.1 s with the key and the view): PAGE_S allows a third more. One renewal on a slow EPM is the
# lost request, the drain and the key request, each at most one request's timeout.
PIECE_FLOOR_S = 300.0            # the key, EPM's drill query, the view and a few pages, before the renewal allowance
PAGE_S = 15.0                    # per getData page of config.DRILL_PAGE lines
RENEWAL_S = 3 * _REQUEST_S       # 360 s
OLD_CAP_S = 900.0                # the old page route's cap (300 + 10 s a page at 300 000 lines)


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


async def _call(fn, what: str = "call"):
    from ..core.serverkit import _call as call
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
        from ..core.serverkit import _ctx, progress_secs
        ctx = _ctx.get()
        if ctx is not None:
            asyncio.get_running_loop().create_task(ctx.report_progress(progress_secs(), None, text))
    except Exception:   # noqa: BLE001 - a progress note must never break the drill
        pass


class _Notes:
    """A drill's sign of life: 'still working: drill (<lines so far> of <their total> lines so far, <n> pieces done)
    (<s>s)', at most one note per timing.PROGRESS_AFTER_S and the first after that long. Counts only: never a figure
    from EPM. One per drill, shared by its pieces (_NOTES)."""

    def __init__(self, say, clock=time.monotonic):
        self.say, self.clock = say, clock
        self.t0 = self.last = clock()
        self.lines: dict[int, tuple[int, int | None]] = {}     # piece -> (lines so far, its total or None)
        self.done = 0
        self._ids = itertools.count()

    def new(self) -> int:
        return next(self._ids)

    def page(self, piece, got: int, total: int | None) -> None:
        self.lines[piece] = (got, total)
        self._tell()

    def forget(self, piece) -> None:
        """A piece whose lines are not used (a built query that did not tie): its count goes."""
        self.lines.pop(piece, None)

    def piece_done(self) -> None:
        self.done += 1
        self._tell()

    def _tell(self) -> None:
        now = self.clock()
        if now - self.last < timing.PROGRESS_AFTER_S:
            return
        self.last = now
        got = sum(g for g, _ in self.lines.values())
        total = sum(g if t is None else t for g, t in self.lines.values())
        self.say(f"still working: drill ({got:,} of {total:,} lines so far, {self.done} piece"
                 f"{'' if self.done == 1 else 's'} done) ({now - self.t0:.0f}s)")


_NOTES: contextvars.ContextVar[_Notes | None] = contextvars.ContextVar("smartass_drill_notes", default=None)


async def drill(cell: dict, *, say=say, connection: str = "FinLite", max_lines: int = config.DRILL_MAX_LINES,
                split: str = "auto") -> DrillResult:
    """Drill a cell to its ledger lines: link (Smart View), lines (token route), normalise, reconcile, profile.
    One retry (with a note) when Smart View answers 'Empty response from provider'. Always live: nothing is cached."""
    import asyncio

    from . import drill_lines as dl
    t0, stages = time.monotonic(), {}
    cell, warnings = ops.drill_cell_defaults(cell)
    notes = _NOTES.set(_Notes(say))                 # the pieces' tasks inherit it
    try:
        with timed(stages, "drill"):
            try:
                pieces, parent_cube = await _drill_cell(connection, cell, max_lines, split)
            except Exception as e:   # noqa: BLE001 - only the empty provider response is retried
                if EMPTY_PROVIDER.lower() not in str(e).lower():
                    raise
                say(f"Smart View gave an empty response for the drill link of {' × '.join(cell.values())}; "
                    "retrying once")
                pieces, parent_cube = await _drill_cell(connection, cell, max_lines, split)
    finally:
        _NOTES.reset(notes)
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


def _lines_source():
    """The object that fetches drill lines: the token route (drill_token), with fetch(url, max_lines) and
    lines(query, url, max_lines). Looked up on every call, so a test's fake (and conftest's guard) binds."""
    from . import drill_token
    return drill_token.shared()


def _need(pages: int) -> float:
    return max(PIECE_FLOOR_S, pages * PAGE_S) + RENEWAL_S


def piece_cap() -> float:
    """The most a piece may take: the old 900 s, or what config.DRILL_MAX_LINES needs if that is more (300 000 lines:
    max(300 s, 60 pages * 15 s) + 360 s = 1 260 s), so a piece within the line limit is never cut short by the cap."""
    return max(OLD_CAP_S, _need(math.ceil(config.DRILL_MAX_LINES / config.DRILL_PAGE)))


def piece_budget(total: int | None) -> float:
    """Seconds one drill piece may take, counted from its start: the larger of PIECE_FLOOR_S and its pages times
    PAGE_S (no pages until its first page tells the total: None), plus one renewal (RENEWAL_S: a lost session on the
    drill query, the view or any page), at most piece_cap()."""
    pages = 0 if total is None else max(1, math.ceil(total / config.DRILL_PAGE))
    return min(_need(pages), piece_cap())


class PieceTimeout(Exception):
    """A drill piece ran past its budget (piece_budget); _drill_piece names the cell."""

    def __init__(self, secs: float, got: int, total: int | None):
        super().__init__(f"the drill did not finish within {secs:.0f} s")
        self.secs, self.got, self.total = secs, got, total


async def _drill_fetch(url: str, max_lines: int) -> dict:
    """Lines of one drill piece over the token route, within piece_budget (it grows once the first page tells the
    total). No data at this cell -> an empty piece with EPM's note; past the budget PieceTimeout; every other error
    is raised as it is."""
    import asyncio

    loop = asyncio.get_running_loop()
    t0, seen = loop.time(), {"got": 0, "total": None, "budget": piece_budget(None)}
    budget = asyncio.timeout_at(t0 + seen["budget"])
    notes = _NOTES.get()
    piece = notes.new() if notes is not None else None

    def progress(got: int, total: int | None) -> None:     # on the event loop, after each page
        seen.update(got=got, total=total)
        seen["budget"] = max(seen["budget"], piece_budget(total if total is not None else got + config.DRILL_PAGE))
        budget.reschedule(t0 + seen["budget"])        # never earlier: a deadline already past would fire at once
        if notes is not None:
            notes.page(piece, got, total)
    try:
        async with budget:
            with drilltime.span("token drill"):
                return await _lines_source().fetch(url, max_lines, progress)
    except NoDrillData as e:
        return {"items": [], "total": 0, "view": None, "secs": 0, "note": str(e)}
    except TimeoutError:
        if not budget.expired():
            raise                                       # a request's own timeout inside the fetch: not the budget
        raise PieceTimeout(seen["budget"], seen["got"], seen["total"]) from None


async def _fetch_piece(cell: dict, url: str, max_lines: int) -> dict:
    """_drill_fetch for one cell; past its budget a readable ToolError that names the cell and what to do."""
    try:
        res = await _drill_fetch(url, max_lines)
    except PieceTimeout as e:
        of = f" of {e.total:,}" if e.total is not None else ""
        raise ToolError(f"The drill of {' × '.join(cell.values())} did not finish within {e.secs:.0f} s ({e.got:,}"
                        f"{of} lines fetched). Narrow the cell (a lower entity or account, or one month) and drill "
                        "again.") from None
    notes = _NOTES.get()
    if notes is not None:
        notes.piece_done()
    return res


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


class _KnownTooWide(Exception):
    def __str__(self):
        return "EPMFDM-140373 earlier in this session"


def _http_link(connection: str, cell: dict) -> dict:
    from . import svhttp
    return svhttp.shared().drill_link(connection, cell)


def _http_off() -> bool:
    import os
    return os.environ.get("SV_TEST_NO_HTTP_DRILL") == "1"


HTTP_OFF = ("the drill link needs the Smart View provider's HTTP route, which is switched off here "
            "(SV_TEST_NO_HTTP_DRILL=1); there is no other route for it")


async def _link(connection: str, cell: dict) -> dict:
    """The drill link over the provider's HTTP route (GetDrillThroughReports). It has no second route: Planning REST
    has no drill-report request, and the Excel route is gone - a failure is a clear error."""
    import asyncio

    from ..xl.host import InputError
    if _http_off():
        raise ToolError(HTTP_OFF)
    try:
        with drilltime.span("http drill link"):
            return await asyncio.to_thread(_http_link, connection, cell)
    except (InputError, oauth.NotSignedIn):              # a missing sign-in is answered with the code
        raise
    except Exception as e:   # noqa: BLE001 - SvHttpError or anything else
        raise ToolError(f"the drill link of {connection} cell {' / '.join(map(str, cell.values()))} could not be "
                        f"fetched over the Smart View provider route ({e}); Planning REST has no drill-report "
                        "request, so there is no other route - try again in a minute") from None


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
    from . import drill_split
    link = await _link(connection, cell)
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
        res = await _fetch_piece(cell, link["url"], limit)
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
        res = await _fetch_piece(cell, link["url"], left)                 # leaf: page it all
    state["left"] -= len(res["items"])
    return [{"cell": cell, "cube_value": link["cube_value"], "drill_url": link["url"], **res}]


async def _children(connection: str, parts: list[dict], max_lines: int, split: str, state: dict,
                    depth: int) -> list[dict]:
    """The split children drilled at once, pieces in the plan's order. Their values come first in ONE retrieve, so an
    empty (#Missing) child costs no drill link (most children of a wide entity are empty). Each starts with the budget
    free at that moment, so the total is checked again at the end. With config.DRILL_DERIVE the siblings' queries are
    built from one of EPM's (_children_derived)."""
    vals = await _child_values(connection, parts)
    if vals is not None and config.DRILL_DERIVE:
        done = await _children_derived(connection, parts, vals, max_lines, split, state, depth)
    else:
        async def one(i: int, pc: dict) -> list[dict]:
            if vals is not None and vals[i] is None:
                return [_skipped(pc)]
            return await _drill_piece(connection, pc, max_lines, split, state, depth + 1)
        done = await _all(one(i, pc) for i, pc in enumerate(parts))
    out = [p for ps in done for p in ps]
    if state["left"] < 0:
        n = sum(len(p.get("items") or []) for p in out)
        raise ToolError(f"The split pieces hold {n}+ lines, more than allowed (max_lines={max_lines}). Narrow the "
                        "cell or raise max_lines.")
    return out


def _skipped(cell: dict) -> dict:
    return {"cell": cell, "cube_value": None, "items": [], "view": None, "secs": 0, "note": "empty cell (skipped)"}


async def _children_derived(connection: str, parts: list[dict], vals: list, max_lines: int, split: str,
                            state: dict, depth: int) -> list[list[dict]]:
    """The split children with ONE drill query asked of EPM (the template: the live child with the fewest members,
    up to 3 tries) and the others' queries built from it (drill_derive). A built piece counts only when its lines
    tie to the child's cube value; anything else - no template, a guard, too many lines, no tie - drills the child
    the usual way. Pieces in the plan's order."""
    import asyncio

    from . import drill_derive
    dim = next((k for k in parts[0] if any(p.get(k) != parts[0][k] for p in parts)), None)
    out: list = [[_skipped(pc)] if vals[i] is None else None for i, pc in enumerate(parts)]
    live = [i for i in range(len(parts)) if vals[i] is not None]
    mem: dict = {}
    if dim in drill_derive.COLUMNS and len(live) > 1:
        mem = dict(zip(live, await asyncio.to_thread(
            lambda: [drill_derive.members(connection, parts[i][dim]) for i in live])))
    template = None
    for i in sorted((i for i in live if mem.get(i) and vals[i]), key=lambda i: len(mem[i]))[:3]:
        out[i] = await _drill_piece(connection, parts[i], max_lines, split, state, depth + 1)
        p = out[i][0] if len(out[i]) == 1 else {}
        if p.get("query") is not None and p.get("drill_url"):
            template = (mem[i], p["query"], p["drill_url"])
            break

    async def one(i: int) -> list[dict]:
        if template is not None and mem.get(i):
            body = drill_derive.derive(template[1].body, dim, template[0], mem[i])
            piece = await _derived_piece(parts[i], vals[i], template, body, state) if body else None
            if piece is not None:
                return [piece]
        return await _drill_piece(connection, parts[i], max_lines, split, state, depth + 1)
    rest = [i for i in live if out[i] is None]
    for i, ps in zip(rest, await _all(one(i) for i in rest)):
        out[i] = ps
    return out


async def _derived_piece(cell: dict, cube, template: tuple, body: str, state: dict) -> dict | None:
    """Lines of a sibling from its built query, or None (then EPM is asked): too many lines for one piece, any
    drill error, or lines that do not tie to the cube value."""
    import asyncio
    import dataclasses

    q = dataclasses.replace(template[1], body=body)
    notes = _NOTES.get()
    piece = notes.new() if notes is not None else None
    progress = (lambda got, total: notes.page(piece, got, total)) if notes is not None else None
    used = False
    try:
        try:
            with drilltime.span("derived drill"):
                res = await asyncio.wait_for(_lines_source().lines(
                    q, template[2], min(config.DRILL_SPLIT_AT, max(state["left"], 0)), progress=progress), 300)
        except Exception as e:   # noqa: BLE001 - TooManyLines, a drill error, a timeout: EPM is asked
            log.info("derived drill of %s not used (%s: %s); asking EPM", cell, type(e).__name__, e)
            return None
        total = 0.0
        for it in res["items"]:
            try:
                total += float(it.get("AMOUNTX") or 0)
            except (TypeError, ValueError):
                return None
        if not res["items"] or cube is None or abs(round(total, 2) - round(float(cube), 2)) >= 0.005:
            # no lines proves nothing (an atlas gap ties to a zero cube): only EPM's own query may say "none"
            log.info("derived drill of %s does not tie to the cube; asking EPM", cell)
            return None
        used = True
        state["left"] -= len(res["items"])
        return {"cell": cell, "cube_value": cube, "drill_url": template[2], **res, "derived": True}
    finally:
        if notes is not None and used:
            notes.piece_done()
        elif notes is not None:                     # lines not used are not counted: EPM is asked for them
            notes.forget(piece)


async def _child_values(connection: str, parts: list[dict]) -> list | None:
    """The split children's cube values as one column grid - over the provider route, else over Planning REST - or
    None (then every child is linked as before: the values are only an optimisation)."""
    import asyncio

    from . import svhttp
    if _http_off():
        return None
    for over_rest in (False, True):
        try:
            return await asyncio.to_thread(lambda: svhttp.shared().child_values(connection, parts, rest=over_rest))
        except oauth.NotSignedIn:                       # answered with the sign-in code
            raise
        except Exception as e:   # noqa: BLE001 - only an optimisation
            log.info("split values not read over %s (%s)", "Planning REST" if over_rest else "HTTP", e)
    return None


def _lines_of(pieces: list[dict], view: dict, parent_cube: float | None = None) -> dict:
    """CPU work off the event loop: normalise each piece ONCE, reconcile each piece, and the whole against the
    PARENT cell's own cube value; split_check = sum of the piece cubes - parent (None without a split)."""
    from . import drill_lines as dl
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
