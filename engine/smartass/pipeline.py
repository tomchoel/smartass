# engine/smartass/pipeline.py
"""fin_analyze in three steps. prepare(): validate, guard, fetch every grid request in parallel (HTTP, no Excel),
compute the model. excel_job(): runs in the user's Excel (v1 host job) - target, Q<n>, Data + Answer sheets,
styling, readback, spec saved in _svmeta, sheet order, save our own file. reply(): the result dict the agent
quotes from (key_values only - never numbers it computed itself)."""
from __future__ import annotations

import contextlib
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field

from .model.analyze import Result, compute, plan_requests
from .model.frame import Frame
from .model.requests import GridRequest
from .model.spec import AnalysisSpec, LinesBlock
from .model.validate import GuardError

MAX_PARALLEL = 6


class RestoreError(RuntimeError):
    """A follow-up failed AND its restore was incomplete: the workbook may be damaged, nothing more should be drawn."""


@dataclass
class Prepared:
    spec: AnalysisSpec
    raw: dict
    closed: dict
    requests: list[GridRequest]
    frame: Frame
    result: Result
    warnings: list[str] = field(default_factory=list)
    routes: dict[str, int] = field(default_factory=dict)
    fetch_seconds: float = 0.0
    aggregation: dict[str, str] = field(default_factory=dict)


def prepare(raw: dict, resolver, provider, closed: dict | None = None, lines: dict | None = None) -> Prepared:
    spec = AnalysisSpec.model_validate(raw)
    lines = lines or {}
    for i, b in enumerate(spec.blocks):
        if not isinstance(b, LinesBlock):
            continue
        ls = lines.get(f"b{i}")
        if ls is None:
            raise GuardError(f"block {i} (lines) has no ledger lines", "run it through fin_analyze or fin_drill, "
                             "which drill the cell first")
        if not ls.rows:
            why = f" ({'; '.join(ls.warnings)})" if ls.warnings else ""
            raise GuardError(f"no ledger lines behind {' / '.join(b.cell.values())}{why}",
                             "check the cell (account, entity, period) or the cube value; FinLite holds the lines")
    if closed is None:
        closed = provider.closed_months()
    canonical = getattr(resolver, "canonical", None)
    if canonical is not None:                   # atlas spelling: requests and SUMIFS then use one name
        for q in spec.queries:
            if resolver.exists(q.entity, q.cube):
                q.entity = canonical(q.entity, q.cube)
    reqs = plan_requests(spec, resolver, closed)
    with ThreadPoolExecutor(max_workers=MAX_PARALLEL) as ex:
        got = list(ex.map(provider.fetch, reqs))
    frame = Frame([f for g in got for f in g.facts])
    frame.lines = dict(lines)
    warnings = [w for g in got for w in g.warnings]
    routes: dict[str, int] = {}
    for g in got:
        routes[g.route] = routes.get(g.route, 0) + 1
    result = compute(spec, resolver, closed, frame)
    aggregation = {m: resolver.aggregation(m) for q in spec.queries for m in q.metrics}
    return Prepared(spec, raw, closed, reqs, frame, result, warnings, routes, sum(g.seconds for g in got),
                    aggregation)


def run_prepare(raw: dict, provider, lines: dict | None = None) -> Prepared:
    """prepare() for a worker thread: everything that touches sqlite (the atlas connection) is opened, used and closed
    here, in the calling thread. Closed months are fetched once. Nothing after prepare needs the resolver: the
    labels are already in the BlockOut lines."""
    import contextlib

    from smartview_mcp import atlas

    from .knowledge.atlas_resolver import AtlasResolver
    from .knowledge.interim import catalog_metrics
    closed = provider.closed_months()
    with contextlib.closing(atlas.connect()) as db:
        resolver = AtlasResolver(db, catalog_metrics(), forecast_year=closed["FinPlan"][0] % 100)
        return prepare(raw, resolver, provider, closed=closed, lines=lines)


def reply(pre: Prepared, q: int, workbook: str, sheets: list[str], mismatches: list[str], lint: list[str]) -> dict:
    res = pre.result
    return {
        "q": q,
        "workbook": workbook,
        "sheets": sheets,
        "key_values": res.key_values,
        "material": res.material,
        "findings": [{"severity": f.severity, "text": f.text, "where": f.where} for f in res.findings],
        "checks": {"sheet_equals_model": "OK" if not mismatches else f"DIFF in {len(mismatches)} cells"},
        "mismatches": mismatches[:20],
        "warnings": pre.warnings + lint,
        "routes": pre.routes,
        "closed": {c: f"{m[1]:02d}/{m[0]}" for c, m in pre.closed.items()},
        "fetch_seconds": round(pre.fetch_seconds, 2),
        "saved_user_workbook": False,
    }
# ---------------------------------------------------------------- the Excel job (runs on a v1 user-Excel job thread)
from .render.charts import write_charts  # noqa: E402
from .render.excel import (col_number, read_block, read_data, rewrite_data, rewrite_grid,  # noqa: E402
                           write_answer, write_data)
from .render.layout import constant_columns, layout  # noqa: E402
from .render.readback import compare  # noqa: E402

DATA_TAB = 0xA6A6A6      # grey tab for data sheets (BGR), as v1 sheetorder.DATA_TAB


def _target(h, target: str, title: str):
    from smartview_mcp.excel import target as xltarget
    return xltarget.resolve_target(h, target, allow_archive=False, title=title)


def _next_q(t) -> int:
    from smartview_mcp.excel import sheets
    return sheets.next_question(t)


def _add_sheet(t, name: str, q: int, color=None):
    from smartview_mcp.excel import sheets
    return sheets.add_sheet(t, name, after_q=q, color=color)


def _calculate(t) -> None:
    from smartview_mcp.excel import sheets
    from smartview_mcp.excel.jobs import com
    com(lambda: t.app.Calculate())
    sheets.wait_calculated(t.app)


def _save_spec(t, q: int, raw: dict, owned: dict | None = None) -> None:
    from smartview_mcp.excel import jobs, sheets
    if not sheets._sheet_exists(t, sheets.META_SHEET):      # v1 save_spec returns silently without it; same creation as save_meta
        ws = jobs.com(lambda: t.wb.Worksheets.Add(None, t.wb.Worksheets(t.wb.Worksheets.Count)))
        jobs.com(lambda: sheets._name_hidden(ws, sheets.META_SHEET))
    sheets.save_spec(t, q, {"v2": raw}, ids={}, owned=owned)


def _arrange(t, sheet: str) -> str | None:
    """Sheet order like v1 tidy: our own files fully, the user's workbook only our sheets among themselves; land on
    the Answer sheet. Cosmetic: a failure is returned as a warning text, never raised (no rollback)."""
    from smartview_mcp import sheetorder
    try:
        mode = "full" if _is_ours(t) else "active"
        sheetorder.arrange_sheets(t.wb, mode, activate=sheet)
    except Exception as e:   # noqa: BLE001
        return f"could not arrange the sheets ({e}); the answer is built"
    return None


def _is_ours(t) -> bool:
    """The engine's own analysis file (created by us, or sitting in the analyser folder - a follow-up targets the
    stored path), never the user's workbook."""
    from smartview_mcp import server as v1
    return bool(getattr(t, "created", False)) or v1._is_our_analysis_file(str(t.path))


@contextlib.contextmanager
def autosave_paused(t):
    """AutoSave (OneDrive/SharePoint files) off while v2 writes, so a half-written answer is never stored.
    Our own analysis file: restored afterwards (we save it ourselves). The user's workbook: left off, with a note
    (AutoSave would otherwise save our edits into their file; v1 does the same)."""
    from smartview_mcp.excel.jobs import com
    notes: list[str] = []
    try:
        prev = bool(com(lambda: t.wb.AutoSaveOn))
    except Exception:   # noqa: BLE001 - a local file or an Excel without AutoSave
        prev = False
    if prev:
        com(lambda: setattr(t.wb, "AutoSaveOn", False))
    try:
        yield notes
    finally:
        if prev and _is_ours(t):
            with contextlib.suppress(Exception):
                com(lambda: setattr(t.wb, "AutoSaveOn", True))
        elif prev:
            notes.append(f"AutoSave was switched off in {t.wb.Name} so the answer is not saved for you; "
                         "save yourself or switch it back on.")


def _save(t) -> None:
    from smartview_mcp.excel.jobs import com
    if _is_ours(t):                                     # only the engine's own analysis file, never the user's
        com(lambda: t.wb.Save())


def _rollback(t) -> None:
    from smartview_mcp.excel import sheets
    sheets.rollback(t)


def _backup(t, name: str) -> str:
    from smartview_mcp.core import owned
    return owned.backup(t, name)


def _restore(t, name: str, bname: str) -> None:
    """The backup becomes the sheet again. v1's restore first deletes the (broken) sheet `name`, which fails when it
    is already gone (the Data sheet is deleted before it is rebuilt): then only un-hide and rename the backup."""
    from smartview_mcp.core import owned
    from smartview_mcp.excel import sheets
    from smartview_mcp.excel.jobs import com
    if sheets._sheet_exists(t, name):
        owned.restore(t, name, bname)
        return
    b = com(lambda: t.wb.Worksheets(bname))
    com(lambda: setattr(b, "Visible", -1))                          # xlSheetVisible
    com(lambda: setattr(b, "Name", name))


def _delete_sheet(t, name: str) -> None:
    from smartview_mcp.excel import sheets
    sheets.drop_parked(t, {name: name})                 # alerts off while deleting, restored after


def _drop_backup(t, bname: str) -> None:
    _delete_sheet(t, bname)


def _read_data(ws, q: int):
    return read_data(ws, q)


def _sheet(t, name: str):
    from smartview_mcp.excel.jobs import com
    return com(lambda: t.wb.Worksheets(name))


def _make_room(ws, old: dict | None, new: dict) -> None:
    """Insert rows/columns when the larger new answer would cover the user's own cells (v1 owned.make_room)."""
    from smartview_mcp.core import owned as owned_mod
    owned_mod.make_room(ws, old, new)


def _clear(ws, owned: dict | None) -> None:
    from smartview_mcp.core import owned as owned_mod
    owned_mod.clear(ws, owned)                          # the drawn range and our charts (everything when none stored)


def _session_workbook() -> str | None:
    from smartview_mcp import session
    return session.workbook()


def read_stored(h, target: str = "session") -> tuple[dict, str | None, str]:
    """(stored specs by q, active sheet name, absolute workbook path): a short user-Excel job. The redraw uses the
    returned path as its target, so a workbook switch while the data is fetched cannot redirect it. A follow-up never
    creates a workbook: target 'new', or 'session' before the session has one, is refused before Excel is touched."""
    from smartview_mcp.excel import sheets
    from smartview_mcp.excel.jobs import com
    if target == "new":
        raise GuardError("a follow-up or refresh works on an existing answer, not a new workbook",
                         "use target 'session' (default) or 'active', or ask a new question with fin_analyze(spec)")
    if target == "session" and not _session_workbook():
        raise GuardError("no session workbook yet - ask a new question first",
                         "run fin_analyze(spec) for a new question; follow-ups need an existing answer")
    t = _target(h, target, "Follow-up")
    return sheets._stored_specs(t), com(lambda: t.wb.ActiveSheet.Name), str(t.path)


def _write_lines(t, q: int, ls) -> str:
    """One line set -> table Lines_Q<q>_<alias> on 'Q<q> Lines <alias>' (v1 drill_xl.write_lines, which reuses the
    table in place) and the cube value in C2/D2 where CubeRef reads it."""
    from smartview_mcp import drill_xl
    from smartview_mcp.excel.jobs import com

    from .model.lines import CUBE_CELL
    cell = " / ".join(ls.cell.values()) if ls.cell else ls.alias
    res = drill_xl.write_lines(t, q, ls.rows, ls.columns, title=f"Ledger lines behind {cell}",
                               source="Source: Oracle EPM Cloud (FinLite drill-through) via Smart View.",
                               alias=ls.alias)
    want = f"Q{q} Lines {ls.alias}"
    if res["sheet"] != want:                    # CubeRef addresses the sheet by this name
        raise GuardError(f"the lines sheet came out as {res['sheet']!r}, not {want!r}",
                         "rename or delete the sheet with that name, then ask again")
    ws = _sheet(t, want)
    com(lambda: setattr(ws.Range("C2"), "Value", "Cube value"))
    com(lambda: setattr(ws.Range(CUBE_CELL.replace("$", "")), "Value", ls.cube_value))
    return want


def _draw(t, q: int, pre: Prepared, ws_answer, ws_data, redraw: bool = False,
          old_owned: dict | None = None) -> dict:
    """Data sheet, layout, Answer sheet, charts, readback, spec + owned range, sheet order. Shared by excel_job and
    redraw_job; the caller owns rollback and save. A redraw rewrites the existing Data table in place."""
    (rewrite_data if redraw else write_data)(ws_data, pre.frame, q)
    cubes = sorted({x.cube for x in pre.spec.queries}) or ["FinLite"]     # a lines-only spec has no query
    line_sheets = [f"Q{q} Lines {a}" for a in pre.frame.lines]
    for ls in pre.frame.lines.values():         # before the answer, so its formulas resolve
        _write_lines(t, q, ls)
    closed = ", ".join(f"{c} {m[1]:02d}/{m[0]}" for c, m in sorted(pre.closed.items()) if c in cubes)
    plan = layout(q, pre.spec, pre.result, stock=_stock(pre), subtitle=f"NOK · {' + '.join(cubes)} · closed {closed}",
                  const=constant_columns(pre.frame))
    if redraw:                                  # a larger answer pushes the user's cells aside, never overwrites them
        _make_room(ws_answer, old_owned, {"rows": plan.last_row, "cols": col_number(plan.last_col)})
    lint = write_answer(ws_answer, plan, pre.warnings)
    lint = list(lint or []) + write_charts(ws_answer, plan)
    _calculate(t)
    mismatches = compare(plan.addr, pre.result.cells, read_block(ws_answer, plan), plan.fractions)
    _save_spec(t, q, pre.raw, {"rows": plan.last_row, "cols": col_number(plan.last_col)})
    arranged = _arrange(t, f"Q{q} Answer")
    return reply(pre, q, str(t.path), [f"Q{q} Answer", f"Q{q} Data", *line_sheets], mismatches,
                 lint + ([arranged] if arranged else []))


def _with_notes(out: dict, notes: list[str]) -> dict:
    if notes:
        out["warnings"] = list(out.get("warnings", [])) + notes
    return out


def excel_job(h, pre: Prepared, target: str = "session") -> dict:
    t = _target(h, target, pre.spec.question[:60])
    with autosave_paused(t) as notes:
        try:
            q = _next_q(t)
            ws_answer = _add_sheet(t, f"Q{q} Answer", q)
            ws_data = _add_sheet(t, f"Q{q} Data", q, DATA_TAB)
            out = _draw(t, q, pre, ws_answer, ws_data)
            _save(t)
        except Exception:
            _rollback(t)
            raise
    return _with_notes(out, notes)


def redraw_job(h, pre: Prepared, q: int, owned: dict | None, target: str = "session") -> dict:
    """A follow-up or refresh: back up Q<q> Answer (only), clear what smartass drew there, rewrite the existing
    Data table in place (the Data sheet is never copied, deleted or renamed: the table name and every formula that
    points at it survive). On failure: rollback, the Answer backup becomes the sheet again, the old data is written
    back into the table - each step on its own, so one failing step never hides the error or skips the others."""
    t = _target(h, target, pre.spec.question[:60])
    with autosave_paused(t) as notes:
        a_name, d_name = f"Q{q} Answer", f"Q{q} Data"
        ws_data = _sheet(t, d_name)
        old_data = _read_data(ws_data, q)
        b_answer = _backup(t, a_name)
        try:
            ws_answer = _sheet(t, a_name)
            _clear(ws_answer, owned)
            out = _draw(t, q, pre, ws_answer, ws_data, redraw=True, old_owned=owned)
        except Exception as e:
            problems = []
            steps = [("rollback", lambda: _rollback(t)), (f"restore {a_name}", lambda: _restore(t, a_name, b_answer))]
            if old_data is not None:
                steps.append((f"rewrite {d_name}", lambda: rewrite_grid(_sheet(t, d_name), old_data, q)))
            for what, step in steps:
                try:
                    step()
                except Exception as e2:   # noqa: BLE001 - collect, never hide the original error
                    problems.append(f"{what}: {e2}")
            if problems:
                raise RestoreError(f"{e}; the restore was incomplete - " + "; ".join(problems)) from e
            raise
        _drop_backup(t, b_answer)
        _save(t)
        out["followup"] = True
        if not owned:
            out["warnings"] = list(out.get("warnings", [])) + [
                "old answer without a stored range: the whole Answer sheet was redrawn"]
    return _with_notes(out, notes)


def _stock(pre: Prepared) -> set[str]:
    """Stock metrics (FTE): shown with one decimal (tall1), never summed over months."""
    return {m for m, agg in pre.aggregation.items() if agg == "last"}
