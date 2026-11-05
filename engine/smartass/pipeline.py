# engine/smartass/pipeline.py
"""fin_analyze in three steps. prepare(): validate, guard, fetch every grid request in parallel (HTTP, no Excel),
compute the model. excel_job(): runs in the user's Excel (v1 host job) - target, Q<n>, Data + Answer sheets,
styling, readback, spec saved in _svmeta, sheet order, save our own file. reply(): the result dict the agent
quotes from (key_values only - never numbers it computed itself)."""
from __future__ import annotations

import contextlib
import re
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime
from functools import partial

from .model.analyze import Result, compute, plan_requests
from .model.blocks.concept import concept_caveats
from .model.chain import complete_chain, restore_chain
from .model.frame import Frame
from .model.periods import parse
from .model.requests import GridRequest
from .model.spec import AnalysisSpec, LinesBlock
from .model.validate import GuardError
from .epm.svhttp import SvHttpError

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
    failed: list[GridRequest] = field(default_factory=list)          # grids the HTTP route could not fetch
    failed_errors: list[str] = field(default_factory=list)
    fetched: dict = field(default_factory=dict)       # GridRequest -> FetchResult the HTTP route fetched (reused by the fallback round)
    caveats: list[dict] = field(default_factory=list)  # {id, text, severity} of the facts the answer touched


def prepare(raw: dict, resolver, provider, closed: dict | None = None, lines: dict | None = None,
            complete: bool = False) -> Prepared:
    """complete (new questions and follow-up patches): the spec's chained tables get the whole P&L chain first
    (model/chain.py; a query a patch resent without the chain gets it back); the completed spec is the one returned
    as raw, so it is stored and refresh, patch and redraw see the chain."""
    if complete:
        raw = restore_chain(raw)
    spec = AnalysisSpec.model_validate(raw)
    if complete:
        spec = complete_chain(spec, resolver)
        raw = spec.model_dump()
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
            q.products = [canonical(x, q.cube) if resolver.exists(x, q.cube) else x for x in q.products]
    reqs = plan_requests(spec, resolver, closed)
    def fetch_one(req):
        try:
            return provider.fetch(req), None
        except SvHttpError as e:        # only the HTTP route's failure; a wrong member (InputError) is wrong on every route
            return None, str(e)

    with ThreadPoolExecutor(max_workers=MAX_PARALLEL) as ex:
        fetched = list(ex.map(fetch_one, reqs))
    failed = [r for r, (g, _) in zip(reqs, fetched) if g is None]
    failed_errors = [err for g, err in fetched if g is None]
    got = [g for g, _ in fetched if g is not None]
    frame = Frame([f for g in got for f in g.facts])
    frame.lines = dict(lines)
    warnings = [w for g in got for w in g.warnings]
    if failed:
        warnings.append(f"{len(failed)} grid(s) could not be fetched over HTTP: {'; '.join(failed_errors[:3])}")
    routes: dict[str, int] = {}
    for g in got:
        routes[g.route] = routes.get(g.route, 0) + 1
    result = compute(spec, resolver, closed, frame)
    aggregation = {m: resolver.aggregation(m) for q in spec.queries for m in q.metrics}
    return Prepared(spec, raw, closed, reqs, frame, result, warnings, routes, sum(g.seconds for g in got),
                    aggregation, failed, failed_errors, {g.request: g for g in got},
                    caveats(spec, reqs, resolver) + concept_caveats(spec, resolver))


_YEAR = re.compile(r"(?:Yr|FY) ?(\d{2})")


def _cell_months(period: str) -> list:
    """A lines cell's period as months: 'Sep 26' -> [Sep 26], 'Yr 26' -> the year; anything else -> none."""
    with contextlib.suppress(ValueError):
        return [parse(period)]
    hit = _YEAR.fullmatch(period.strip())
    return [(2000 + int(hit.group(1)), m) for m in range(1, 13)] if hit else []


def caveats(spec: AnalysisSpec, reqs: list[GridRequest], resolver) -> list[dict]:
    """The resolver's caveats for what the answer touched: the entities and accounts of every grid request and lines
    cell (per cube; FinLite shares the FinMax members, so it counts for both), the metrics of the queries and the
    months fetched. A resolver without caveats (a test double) gives none."""
    fn = getattr(resolver, "caveats", None)
    if fn is None:
        return []
    touched: dict[str, set[str]] = {}
    months: set = set()

    def touch(cube: str, names) -> None:
        for c in ({cube, "FinMax"} if cube == "FinLite" else {cube}):
            touched.setdefault(c, set()).update(names)
    for r in reqs:
        touch(r.cube, (*r.entities, *r.accounts, *r.products))
        for per in r.periods:
            with contextlib.suppress(ValueError):
                months.add(parse(per))
    for b in spec.blocks:
        if isinstance(b, LinesBlock):
            touch("FinLite", [v for k, v in b.cell.items() if k in ("All Accounts", "All Entities")])
            months.update(_cell_months(b.cell.get("All Time Periods", "")))
    metric = getattr(resolver, "metric", None)
    ids = {m.id for q in spec.queries for name in q.metrics if metric and (m := metric(name)) is not None}
    return fn(touched, ids, months)


def run_prepare(raw: dict, provider, lines: dict | None = None, complete: bool = False) -> Prepared:
    """prepare() for a worker thread: everything that touches sqlite (knowledge.db and the atlas) is opened, used and
    closed here, in the calling thread. knowledge.db is built first when it is missing (said in the warnings). Closed
    months are fetched once. Nothing after prepare needs the resolver: the labels are already in the BlockOut lines
    and the caveats in Prepared."""
    from .core import config
    from .knowledge.refresh import note_closed
    from .knowledge.resolver import open_resolver
    closed = provider.closed_months()
    with contextlib.suppress(Exception):         # fin_status's 'knowledge refresh due'; never stops an answer
        note_closed(config.HOME, closed["FinMax"])
    notes: list[str] = []
    with contextlib.closing(open_resolver(closed["FinPlan"][0] % 100, notes)) as resolver:
        pre = prepare(raw, resolver, provider, closed=closed, lines=lines, complete=complete)
    pre.warnings[:0] = notes
    return pre


def reply(pre: Prepared, q: int, workbook: str, sheets: list[str], mismatches: list[str], lint: list[str]) -> dict:
    res = pre.result
    checks = {"sheet_equals_model": sheet_status(mismatches)}
    out = {
        "q": q,
        "workbook": workbook,
        "sheets": sheets,
        "key_values": res.key_values,
        "material": res.material,
        "findings": [{"severity": f.severity, "text": f.text, "where": f.where} for f in res.findings],
        "checks": checks,
        "mismatches": mismatches[:20],
        "warnings": pre.warnings + lint,
        "caveats": list(pre.caveats),
        "routes": pre.routes,
        "closed": {c: f"{m[1]:02d}/{m[0]}" for c, m in pre.closed.items()},
        "fetch_seconds": round(pre.fetch_seconds, 2),
        "saved_user_workbook": False,
    }
    if pre.spec.plan.depth in ("standard", "deep"):         # the critic pass (sub-spec 2 §5) reads this
        from .model.critic import brief
        out["critic_brief"] = brief(pre.spec, res, list(pre.caveats), checks, stock=_stock(pre))
    return out
# ---------------------------------------------------------------- the Excel job (runs on a v1 user-Excel job thread)
from .render import datasheet  # noqa: E402
from .render.charts import write_charts  # noqa: E402
from .render.datasheet import header_rows  # noqa: E402
from .render.excel import (VALUE_FORMAT, answer_footer, col_number, read_block, read_data, rewrite_data,  # noqa: E402
                           rewrite_grid, sheet_status, write_answer, write_data, write_status)
from .render.layout import answer_subtitle, constant_columns, layout  # noqa: E402
from .render.keypoints import free_points, key_points  # noqa: E402
from .render.readback import compare  # noqa: E402

DATA_TAB = 0xA6A6A6      # grey tab for data sheets (BGR), as v1 sheetorder.DATA_TAB


def _target(h, target: str, title: str):
    from .xl import target as xltarget
    return xltarget.resolve_target(h, target, allow_archive=False, title=title)


def _next_q(t) -> int:
    from .xl import sheets
    return sheets.next_question(t)


def _add_sheet(t, name: str, q: int, color=None):
    from .xl import sheets
    return sheets.add_sheet(t, name, after_q=q, color=color)


def _calculate(t) -> None:
    from .xl import sheets
    from .xl.jobs import com
    com(lambda: t.app.Calculate())
    sheets.wait_calculated(t.app)


def _save_spec(t, q: int, raw: dict, owned: dict | None = None) -> None:
    from .xl import jobs, sheets
    if not sheets._sheet_exists(t, sheets.META_SHEET):      # v1 save_spec returns silently without it; same creation as save_meta
        ws = jobs.com(lambda: t.wb.Worksheets.Add(None, t.wb.Worksheets(t.wb.Worksheets.Count)))
        jobs.com(lambda: sheets._name_hidden(ws, sheets.META_SHEET))
    sheets.save_spec(t, q, {"v2": raw}, ids={}, owned=owned)


def _arrange(t, sheet: str) -> str | None:
    """Sheet order like v1 tidy: our own files fully, the user's workbook only our sheets among themselves; land on
    the Answer sheet. Cosmetic: a failure is returned as a warning text, never raised (no rollback)."""
    from .xl import sheetorder
    try:
        mode = "full" if _is_ours(t) else "active"
        sheetorder.arrange_sheets(t.wb, mode, activate=sheet)
    except Exception as e:   # noqa: BLE001
        return f"could not arrange the sheets ({e}); the answer is built"
    return None


def _is_ours(t) -> bool:
    """The engine's own analysis file (created by us, or sitting in the analyser folder - a follow-up targets the
    stored path), never the user's workbook."""
    from .core import serverkit as kit
    return bool(getattr(t, "created", False)) or kit._is_our_analysis_file(str(t.path))


@contextlib.contextmanager
def autosave_paused(t):
    """AutoSave (OneDrive/SharePoint files) off while v2 writes, so a half-written answer is never stored.
    Our own analysis file: restored afterwards (we save it ourselves). The user's workbook: left off, with a note
    (AutoSave would otherwise save our edits into their file; v1 does the same)."""
    from .xl.jobs import com
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
    from .xl.jobs import com
    if _is_ours(t):                                     # only the engine's own analysis file, never the user's
        com(lambda: t.wb.Save())


def _rollback(t) -> None:
    from .xl import sheets
    sheets.rollback(t)


def _backup(t, name: str) -> str:
    from .xl import owned
    return owned.backup(t, name)


def _restore(t, name: str, bname: str) -> None:
    """The backup becomes the sheet again. v1's restore first deletes the (broken) sheet `name`, which fails when it
    is already gone (the Data sheet is deleted before it is rebuilt): then only un-hide and rename the backup."""
    from .xl import owned
    from .xl import sheets
    from .xl.jobs import com
    if sheets._sheet_exists(t, name):
        owned.restore(t, name, bname)
        return
    b = com(lambda: t.wb.Worksheets(bname))
    com(lambda: setattr(b, "Visible", -1))                          # xlSheetVisible
    com(lambda: setattr(b, "Name", name))


def _delete_sheet(t, name: str) -> None:
    from .xl import sheets
    sheets.drop_parked(t, {name: name})                 # alerts off while deleting, restored after


def _drop_backup(t, bname: str) -> None:
    _delete_sheet(t, bname)


def _read_data(ws, q: int):
    return read_data(ws, q)


def _sheet(t, name: str):
    from .xl.jobs import com
    return com(lambda: t.wb.Worksheets(name))


def _make_room(ws, old: dict | None, new: dict) -> None:
    """Insert rows/columns when the larger new answer would cover the user's own cells (v1 owned.make_room)."""
    from .xl import owned as owned_mod
    owned_mod.make_room(ws, old, new)


def _clear(ws, owned: dict | None) -> None:
    from .xl import owned as owned_mod
    owned_mod.clear(ws, owned)                          # the drawn range and our charts (everything when none stored)


def _session_workbook() -> str | None:
    from .core import session
    return session.workbook()


def read_stored(h, target: str = "session") -> tuple[dict, str | None, str]:
    """(stored specs by q, active sheet name, absolute workbook path): a short user-Excel job. The redraw uses the
    returned path as its target, so a workbook switch while the data is fetched cannot redirect it. A follow-up never
    creates a workbook: target 'new', or 'session' before the session has one, is refused before Excel is touched."""
    from .xl import sheets
    from .xl.jobs import com
    if target == "new":
        raise GuardError("a follow-up or refresh works on an existing answer, not a new workbook",
                         "use target 'session' (default) or 'active', or ask a new question with fin_analyze(spec)")
    if target == "session" and not _session_workbook():
        raise GuardError("no session workbook yet - ask a new question first",
                         "run fin_analyze(spec) for a new question; follow-ups need an existing answer")
    t = _target(h, target, "Follow-up")
    return sheets._stored_specs(t), com(lambda: t.wb.ActiveSheet.Name), str(t.path)


def _write_lines(t, q: int, ls, fetched: str = "") -> str:
    """One line set -> table Lines_Q<q>_<alias> on 'Q<q> Lines <alias>' (v1 drill_xl.write_lines, which reuses the
    table in place), the cube value in C2/D2 where CubeRef reads it, and the provenance + tie-out in F1:K3
    (datasheet.lines_header; fetched: the draw's one timestamp text)."""
    from .xl import drill_xl
    from .xl.jobs import com

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
    _lines_header(ws, res["table"], ls.cell, fetched)
    return want


def _lines_header(ws, table: str, cell: dict, fetched: str) -> None:
    """F1:K3 as one rectangle (text with a leading apostrophe: never a date or a formula), labels grey, values ink."""
    from .xl import styling
    from .xl.jobs import com
    g = datasheet.lines_header(table, cell, fetched)
    grid = tuple(tuple("'" + x if isinstance(x, str) and x and not x.startswith("=") else x for x in r) for r in g)
    rng = ws.Range(datasheet.LINES_HEADER)

    def look():
        rng.NumberFormat = "General"
        rng.Formula = grid
        rng.Font.Name = styling.tn.FONT
        rng.Font.Size = 9
        rng.Font.Color = styling._rgb(styling.tn.INK)
        for a in ("F1:F3", "H3", "J3"):
            ws.Range(a).Font.Color = styling._rgb(styling.tn.GREY)
        ws.Range("G3").NumberFormat = VALUE_FORMAT
    com(look)


def _draw(t, q: int, pre: Prepared, ws_answer, ws_data, redraw: bool = False,
          old_owned: dict | None = None) -> dict:
    """Data sheet, layout, Answer sheet, charts, readback, spec + owned range, sheet order. Shared by excel_job and
    redraw_job; the caller owns rollback and save. A redraw rewrites the existing Data table in place. One timestamp
    per draw: the data header's Fetched and the answer's print footer never disagree."""
    fetched = datetime.now().astimezone()
    user = datasheet.windows_user()
    header = header_rows(q, pre.spec, pre, fetched, user)
    (rewrite_data if redraw else write_data)(ws_data, pre.frame, q, header)   # its cells: Task 11 (Sources link)
    line_sheets = [f"Q{q} Lines {a}" for a in pre.frame.lines]
    for ls in pre.frame.lines.values():         # before the answer, so its formulas resolve
        _write_lines(t, q, ls, datasheet.fetched_text(fetched, user))
    free, too_many = free_points(pre.spec)
    pre.warnings += [w for w in too_many if w not in pre.warnings]
    plan = layout(q, pre.spec, pre.result, stock=_stock(pre),
                  subtitle=answer_subtitle(pre.spec, pre.result, pre.closed),
                  const=constant_columns(pre.frame), caveats=pre.caveats,
                  key_points=partial(key_points, stock=_stock(pre), free=free))
    plan.footer = answer_footer(fetched)
    if redraw:                                  # a larger answer pushes the user's cells aside, never overwrites them
        _make_room(ws_answer, old_owned, {"rows": plan.last_row, "cols": col_number(plan.last_col)})
    lint = write_answer(ws_answer, plan, pre.warnings)
    lint = list(lint or []) + write_charts(ws_answer, plan)
    _calculate(t)
    mismatches = compare(plan.addr, pre.result.cells, read_block(ws_answer, plan), plan.fractions)
    write_status(ws_answer, plan, mismatches)   # the 'Sheet = model' cell: OK or DIFF in <n> cells
    _save_spec(t, q, pre.raw, {"rows": plan.last_row, "cols": col_number(plan.last_col)})
    arranged = _arrange(t, f"Q{q} Answer")
    return reply(pre, q, str(t.path), [f"Q{q} Answer", f"Q{q} Data", *line_sheets], mismatches,
                 lint + ([arranged] if arranged else []))


def remember(pre: Prepared, out: dict) -> None:
    """The run into answer memory (answers.db). Never stops an answer: a failure becomes a warning."""
    from . import answers
    spec = pre.spec
    cells = [b.cell for b in spec.blocks if isinstance(b, LinesBlock)]
    try:
        answers.store({
            "workbook": out["workbook"], "q": out["q"], "question": spec.question, "purpose": spec.plan.purpose,
            "spec": pre.raw, "key_values": out.get("key_values") or {},
            "metrics": sorted({m for q in spec.queries for m in q.metrics}),
            "entities": list(dict.fromkeys([q.entity for q in spec.queries]
                                           + [c["All Entities"] for c in cells if "All Entities" in c])),
            "periods": list(dict.fromkeys([q.periods for q in spec.queries]
                                          + [c["All Time Periods"] for c in cells])),
            "critic_brief": out.get("critic_brief")})
    except Exception as e:   # noqa: BLE001
        out["warnings"] = list(out.get("warnings", [])) + [f"answer memory not updated ({e})"]


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
    remember(pre, out)
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
    remember(pre, out)
    return _with_notes(out, notes)


def _stock(pre: Prepared) -> set[str]:
    """Stock metrics (FTE): shown with one decimal (tall1), never summed over months."""
    return {m for m, agg in pre.aggregation.items() if agg == "last"}
