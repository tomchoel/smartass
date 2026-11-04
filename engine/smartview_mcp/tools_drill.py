"""sv_drill: ledger lines behind a FinLite cell (the drill itself is core.drillstep, shared with the recipes).
Registered on the shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

import re
from typing import Literal

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError

from smartass.core import config
from . import ops, qlog
from .excel import adhoc, jobs, sheets, target as xltarget
from .core import drillstep, reply
from .grid import Axis
from .host import HostError, rewrap
from .server import WRITE, _call_user, _err, _is_our_analysis_file, tool, with_next_step
from smartass.core.timing import Timer, stage

# ---------------------------------------------------------------- sv_drill: ledger lines behind a FinLite cell
_ALIAS = re.compile(r"^[A-Za-z]\w*$")
NEXT = "To answer a question, use sv_gate → sv_answer (recipe lines_ranking drills and builds the ranking)."


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
    from . import drill_xl
    tm = Timer()
    r = await drillstep.drill(cell, connection=connection, max_lines=max_lines, split=split)
    tm.stages.update(r.stages)
    cell, rec, rows, cols, piece_recs, warnings = r.cell, r.reconciliation, r.rows, r.columns, r.pieces, r.warnings
    out = {"question": question, "cell": cell, "reconciliation": rec, "pieces": piece_recs}
    if r.profile is None:                           # spec 6: no drill report -> cube value, 0 lines, the reason
        out.update(warnings=warnings, timings=tm.summary(), next=NEXT)
        qlog.write({"tool": "sv_drill", "question": question, "cell": cell, "lines": 0, "note": warnings[-1:]})
        return out
    out.update(profile=r.profile, text_repairs=r.text_repairs, columns=cols)
    if r.split_check is not None:
        out["split_check"] = r.split_check

    if target != "none":
        title = f"Drill-through lines – {' × '.join(cell.values())} ({connection})"
        source = drillstep.SOURCE

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
                    info = drillstep.write(t, qn, id, r, title, question, log=False)   # 0 lines: empties the table
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
    with drillstep.timed(tm.stages, "cache"):
        out["drill_id"] = await asyncio.to_thread(
            _cache_drill, connection=connection, cell=cell, rec={**rec, "grid": grid, "excel_columns": out.get("excel_columns")}, columns=cols, rows=rows,
            workbook=out.get("workbook"), q=out.get("q"), alias=id)
    out["warnings"] = warnings
    out["timings"] = tm.summary()
    out["next"] = NEXT
    qlog.write({"tool": "sv_drill", "question": question, "cell": cell, "lines": rec["lines"], "gap": rec["gap"],
                "workbook": out.get("workbook"), "q": out.get("q"), "timings": out["timings"]})
    return out
