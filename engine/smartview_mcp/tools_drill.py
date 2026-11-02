"""sv_drill: ledger lines behind a FinLite cell (plus the drill helpers the answers use). Registered on the shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

import re
import time
from contextlib import contextmanager
from typing import Literal

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError

from . import config, ops, qlog
from .excel import adhoc, jobs, sheets, target as xltarget
from .core import reply
from .grid import Axis
from .host import HostError, rewrap
from .server import WRITE, _call, _call_user, _err, _is_our_analysis_file, tool, with_next_step
from .timing import Timer, stage

# ---------------------------------------------------------------- sv_drill: ledger lines behind a FinLite cell
_drill_lock = None          # asyncio.Lock: one browser fetch at a time, so the hang budget starts when it runs
_ALIAS = re.compile(r"^[A-Za-z]\w*$")


def browser():
    from .drill_browser import shared_browser
    return shared_browser()


@contextmanager
def _timed(tm, name: str):
    """Stage time recorded on THIS call's timer (stage() is thread-local, and sv_drill awaits on the event loop)."""
    t0 = time.monotonic()
    try:
        yield
    finally:
        tm.stages[name] = round(tm.stages.get(name, 0.0) + time.monotonic() - t0, 2)


async def _drill_fetch(url: str, max_lines: int) -> dict:
    """browser().fetch, one at a time in this server, with a hang guard: past the budget only OUR headless Chrome
    is killed (the next drill starts a fresh one) and the tool fails readably instead of waiting forever."""
    import asyncio
    import math
    from .drill_browser import DrillBusy
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

    from . import timing
    from .server import _ctx, progress_secs
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
    from . import drill_split
    from .drill_browser import CellTooWide, TooManyLines
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


def _grid_key(connection: str, pov: dict, account: str, periods: list[str]) -> str:
    """Identity of a drill's cube grid: same key = refresh the existing sheet in place."""
    import json
    return json.dumps({"c": connection, "pov": sorted(pov.items()), "a": account, "p": list(periods)})


def _prev_drill(workbook: str, q: int, alias: str) -> dict | None:
    """Metadata of the earlier drill of this question/alias (the cache is saved last, so it is still the old one)."""
    from . import drill_store
    try:
        db = drill_store.connect()
        try:
            did = drill_store.find(db, workbook, q).get(alias)
            return drill_store.meta(db, did) if did else None
        finally:
            db.close()
    except Exception:   # noqa: BLE001 - no earlier drill known: build fresh
        return None


def _restore_lines(t, q: int, alias: str, prev: dict | None, source: str) -> str:
    """A re-drill failed while rewriting the reused table: put the previous lines back from drill.db; if that fails
    too, mark the sheet INCOMPLETE in red. Returns a note for the error message."""
    from . import drill_store, drill_xl
    try:
        if prev is None:
            raise LookupError("no earlier drill in the cache")
        db = drill_store.connect()
        try:
            _, old_rows = drill_store.load(db, prev["id"])
        finally:
            db.close()
        title = (f"Drill-through lines – {' × '.join(prev['cell'].values())} ({prev['connection']}) "
                 f"– forrige uttrekk {prev['created']}")
        if old_rows:
            drill_xl.write_lines(t, q, old_rows, prev["columns"], title=title, source=source, alias=alias)
        else:
            drill_xl.clear_lines(t, q, title=title, source=source, alias=alias)
        return "the previous lines were restored"
    except Exception as e:   # noqa: BLE001
        try:
            drill_xl.mark_incomplete(t, q, alias)
        except Exception:   # noqa: BLE001
            pass
        return f"the previous lines could not be restored ({e}); the sheet is marked INCOMPLETE"


def _cache_drill(**kw) -> str:
    from . import drill_store
    db = drill_store.connect()
    try:
        drill_store.prune(db)
        return drill_store.save(db, **kw)
    finally:
        db.close()


@tool(annotations=WRITE)
async def sv_drill(question: str, cell: dict[str, str], connection: Literal["FinLite"] = "FinLite", id: str = "d",
                   split: Literal["auto", "none"] = "auto", max_lines: int = config.DRILL_MAX_LINES,
                   target: str = "session", q: int | None = None, allow_archive: bool = False,
                   ctx: Context | None = None) -> dict:
    """Drill a FinLite cell down to its ledger lines (supplier, invoice, voucher, JE source, line text, product)
    through a hidden headless browser (Data Integration drill-through; nothing opens on screen).
    cell: {'All Accounts': 'ac.C40S', 'All Entities': 'un.SY.2WHS', 'All Time Periods': 'Q3 26', ...} - parents
    allowed; other dimensions default to Actual / NOK / Periodic / top member. Cells over 20 000 lines are split
    automatically (months, then cost centres). Writes a live FinLite grid of the cell + 'Q<n> Lines' (Excel table
    Lines_Q<n>) to the session workbook (target 'none' = cache only) and returns the reconciliation (cube vs
    lines), a profile (top suppliers/accounts/months/journals, signs_mixed) and text repairs. Then build the report
    with sv_answer using {drill: 'd', where: {...}} cells. id: alias of this drill
    in the question (letters/digits/_; default 'd'). Re-drilling the same q and id refreshes its table in place."""
    if not _ALIAS.match(id or ""):
        raise ToolError(f"id '{id}' must start with a letter and contain only letters, digits and _ (e.g. 'd', 'b2').")
    from .drill_browser import redact
    try:
        return await _sv_drill(question, cell, connection, id, split, max_lines, target, q, allow_archive)
    except Exception as e:
        msg = redact(str(e) if isinstance(e, ToolError) else f"{type(e).__name__}: {e}")   # never an sso_token
        qlog.write({"tool": "sv_drill", "question": question, "cell": cell, "error": msg})
        raise ToolError(with_next_step(msg, e)) from None


async def _sv_drill(question, cell, connection, id, split, max_lines, target, q, allow_archive) -> dict:
    import asyncio
    from . import drill_lines as dl
    from . import drill_xl
    tm = Timer()
    cell, cell_warns = ops.drill_cell_defaults(cell)
    with _timed(tm, "drill"):
        pieces, parent_cube = await _drill_cell(connection, cell, max_lines, split)
    view = next((p["view"] for p in pieces if p.get("view")), None)
    if view is None:                                # spec 6: no drill report -> cube value, 0 lines, the reason
        total_cube = parent_cube if parent_cube is not None else sum(p["cube_value"] or 0.0 for p in pieces)
        reasons = sorted({p.get("note") or "the drill page returned no view" for p in pieces})
        out = {"question": question, "cell": cell, "reconciliation": dl.reconcile(total_cube, []),
               "pieces": [{"cell": p["cell"], "cube_value": p["cube_value"], "note": p.get("note")} for p in pieces]
               if len(pieces) > 1 else [], "warnings": cell_warns + ["No drill-through lines: " + "; ".join(reasons)],
               "timings": tm.summary()}
        qlog.write({"tool": "sv_drill", "question": question, "cell": cell, "lines": 0, "note": reasons})
        return out
    with _timed(tm, "normalise + profile"):
        L = await asyncio.to_thread(_lines_of, pieces, view, parent_cube)
    cols, rows, stats, piece_recs, rec = L["cols"], L["rows"], L["stats"], L["piece_recs"], L["rec"]
    out = {"question": question, "cell": cell, "reconciliation": rec, "pieces": piece_recs if len(pieces) > 1 else [],
           "profile": L["profile"], "text_repairs": stats, "columns": cols}
    if L["split_check"] is not None:
        out["split_check"] = L["split_check"]
    warnings = list(cell_warns)
    if stats["unfixed"]:
        warnings.append("Text broken in EPM (letters lost at load): " + "; ".join(stats["unfixed"][:5])
                        + " - add corrections to %LOCALAPPDATA%\\smartass\\text_fixes.json")
    if rec["status"] == "gap":
        warnings.append(_gap_warning(rec))
    if L["split_check"]:
        warnings.append(f"The split pieces' cube values differ from the cell by {L['split_check']:,.2f} NOK (the "
                        "atlas children may be capped, out of date or shared): the lines are reconciled to the cell "
                        "itself.")

    if target != "none":
        title = f"Drill-through lines – {' × '.join(cell.values())} ({connection})"
        source = "Kilde: Data Integration drill-through (FinLite)."

        def job(h):
            jt = Timer.start()                      # the job runs on the host thread (thread-local timer)
            with stage("open target"):
                t = xltarget.resolve_target(h, target, allow_archive, title=question[:50])
                qn = q or sheets.next_question(t)
            pov = {**ops.DRILL_DEFAULTS, **{d: m for d, m in cell.items()
                                           if d not in ("All Accounts", "All Time Periods")}}
            periods = list(dict.fromkeys([cell["All Time Periods"]] + [pc["cell"]["All Time Periods"]
                                                                       for pc in piece_recs]))
            key = _grid_key(connection, pov, cell["All Accounts"], periods)
            prev = _prev_drill(t.path, qn, id)                 # the earlier drill of this q/alias, if any
            prev_grid = ((prev or {}).get("rec") or {}).get("grid") or {}
            sess, notes, parked, info = reply.load_session(t, qn), [], {}, None
            had_table = drill_xl.table_columns(t.wb, drill_xl.table_name(qn, id)) is not None
            lines_started = False
            try:
                with stage("cube grid"):
                    old = prev_grid.get("sheet")
                    if old and prev_grid.get("key") == key and old in sess:
                        p = adhoc.refresh_pull(t, sess[old])           # same grid: refresh in place
                        notes.append(f"grid {old} refreshed in place")
                    else:
                        parked = sheets.park_sheets(t, [old]) if old else {}
                        for o in parked.values():
                            reply.session(t.path, qn).pop(o, None)
                        p = adhoc.pull(h, t, qn, connection, pov,
                                           [Axis("All Accounts", [cell["All Accounts"]])],
                                           [Axis("All Time Periods", periods)], title=f"Q{qn} {connection}")
                        if parked:
                            notes.append(f"the cell changed: grid {p.sheet} rebuilt - re-run sv_answer for Q{qn} "
                                         "if its formulas used the old grid")
                reply.session(t.path, qn)[p.sheet] = p
                lines_started = True
                with stage("lines sheet"):
                    if rows:
                        info = drill_xl.write_lines(t, qn, rows, cols, title=title, source=source, alias=id)
                    else:                           # 0 lines: empty an existing table so Excel matches drill.db
                        info = drill_xl.clear_lines(t, qn, title=title, source=source, alias=id)
            except Exception as e:
                removed = sheets.rollback(t)
                for name in removed:
                    reply.session(t.path, qn).pop(name, None)
                if parked:
                    sheets.restore_parked(t, parked)
                    reply.forget(t.path, qn)     # reload the restored grid from the workbook
                extra = ""
                if had_table and lines_started:     # the reused table may be half written: put the old lines back
                    extra = "; " + _restore_lines(t, qn, id, prev, source)
                raise rewrap(e, f"{e} (removed the partial sheets: {', '.join(removed) or 'none'}{extra})") from None
            if parked:
                with stage("drop old grid"):
                    sheets.drop_parked(t, parked)
            with stage("log Sources"):
                try:
                    sheets.log_source(t, qn, question, p, notes=f"drill: {rec['lines']} lines, gap {rec['gap']}"
                                         + (f"; sheet {info['sheet']}" if info else "; no lines"))
                except Exception:   # noqa: BLE001 - the source log must never undo a finished drill
                    pass
            t.status(None)
            t.created = t.created or _is_our_analysis_file(t.path)     # the session's own workbook: ours to save
            reply.tidy(t, info["sheet"] if info else p.sheet)
            if t.created:
                with stage("save"):
                    jobs.com(t.wb.Save)
            if info and info.get("columns_changed"):
                notes.append(f"columns changed – re-run sv_answer for Q{qn}")
            elif info and info.get("reused") and rows:
                notes.append(f"re-drill: table {info['table']} refreshed in place (answer formulas stay valid)")
            if info and len(info["columns"]) < len(cols) and rows:
                notes.append(f"{len(rows)} lines > {config.DRILL_XL_FULL}: only the key columns are in Excel "
                             "(excel_columns); where-filters on other columns are not possible in the answer.")
            if not rows:
                notes.append(f"0 lines: table {info['table']} emptied." if info else
                             "0 lines: no lines table was written.")
            return {"workbook": t.path, "q": qn, "lines_sheet": info["sheet"] if info else None,
                    "table": info["table"] if info else None, "excel_columns": info["columns"] if info else [],
                    "cube_pull": {"sheet": p.sheet, "row_keys": [list(k) for k in p.layout.row_tuples],
                                  "col_keys": [list(k) for k in p.layout.col_tuples]},
                    "notes": notes, "_grid": {"sheet": p.sheet, "key": key}, "_excel_stages": dict(jt.stages)}
        try:
            res = await _call_user(job, what="sv_drill")
        except (HostError, ToolError) as e:
            raise _err(e)
        for k, v in res.pop("_excel_stages", {}).items():
            tm.stages[f"excel.{k}"] = round(v, 2)
        grid = res.pop("_grid", None)
        out.update(res)
    else:
        grid = None
    with _timed(tm, "cache"):
        out["drill_id"] = await asyncio.to_thread(
            _cache_drill, connection=connection, cell=cell, rec={**rec, "grid": grid, "excel_columns": out.get("excel_columns")}, columns=cols, rows=rows,
            workbook=out.get("workbook"), q=out.get("q"), alias=id)
    out["warnings"] = warnings
    out["timings"] = tm.summary()
    qlog.write({"tool": "sv_drill", "question": question, "cell": cell, "lines": rec["lines"], "gap": rec["gap"],
                "workbook": out.get("workbook"), "q": out.get("q"), "timings": out["timings"]})
    return out
