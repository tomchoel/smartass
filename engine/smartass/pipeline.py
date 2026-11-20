# engine/smartass/pipeline.py
"""fin_analyze in three steps. prepare(): validate, guard, fetch every grid request in parallel (HTTP, no Excel),
compute the model. excel_job(): runs in the user's Excel (v1 host job) - target, Q<n>, the answer's sheets
built as an .xlsx without Excel and moved into the workbook in one copy, finished there, read back, Sources, spec
saved in _svmeta, sheet order, save our own file; redraw_job() replaces a question's sheets the same way. reply(): the
result dict the agent quotes from (key_values only - never numbers it computed itself)."""
from __future__ import annotations

import contextlib
import logging
import re
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime
from functools import partial

from .core import timing
from .core.timing import Timer, stage
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
log = logging.getLogger("smartview.pipeline")


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
    out = fn(touched, ids, months)
    cards = getattr(resolver, "card_caveats", None)
    if cards is not None:                   # a level read as a flow is the commonest wrong answer: say it
        have = {c["id"] for c in out}
        out = out + [c for c in cards(touched) if c["id"] not in have]
    return out


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


def reply(pre: Prepared, q: int, workbook: str, sheets: list[str], mismatches: list[str], lint: list[str],
          layout: list[str] | None = None, charts: list[str] | None = None) -> dict:
    """layout: the layout check's findings (render/check.py), each with its fix; charts: the titles drawn."""
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
        "layout": list(layout or []),
        "charts": list(charts or []),
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
from .render import datasheet, xlsx  # noqa: E402
from .render.datasheet import header_rows  # noqa: E402
from .render.excel import answer_footer, data_grid, read_block, sheet_status, write_status  # noqa: E402
from .render.layout import answer_subtitle, constant_columns, layout  # noqa: E402
from .render.keypoints import free_points, key_points  # noqa: E402
from .render.readback import compare  # noqa: E402


def _target(h, target: str, title: str):
    from .xl import target as xltarget
    return xltarget.resolve_target(h, target, allow_archive=False, title=title)


def _save_spec(t, q: int, raw: dict) -> bool:
    """The spec in _svmeta (the hidden sheet made first when there is none, as save_meta makes it): True when it
    was made. No drawn range is stored: a follow-up replaces the question's sheets whole."""
    from .xl import sheets
    return sheets.save_spec(t, q, {"v2": raw}, ids={}, create=True)


def _arrange(t, sheet: str, book, active: str | None = None) -> str | None:
    """Sheet order like v1 tidy: our own files fully, the user's workbook only our sheets among themselves; land on
    the Answer sheet. book (the job's movein.Book, inside quiet): the order and tabs from it, no read; active: the
    sheet that will be active (quiet lands there), else the Book's. Cosmetic: a failure is returned as a warning
    text, never raised (no rollback)."""
    from .xl import sheetorder
    try:
        mode = "full" if _is_ours(t) else "active"
        rows = [(s.name, s.visible == -1, s.tab) for s in book.sheets]
        out = sheetorder.arrange_sheets(t.wb, mode, activate=sheet, known=(rows, active or book.active))
        book.order(out["order"])
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
    Our own analysis file: restored afterwards (we save it ourselves: switched back on, AutoSave did not save the
    last answer within 30 s live, 19.11.2026). The user's workbook: left off, with a note (AutoSave would otherwise
    save our edits into their file; v1 does the same)."""
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


LINES_SOURCE = "Source: Oracle EPM Cloud, FinLite drill-through (Smart View). AMOUNT: FinMax sign, costs negative."


def _lines_grid(table: str, cell: dict, fetched: str) -> list[list]:
    """datasheet.lines_header ready for Formula2: text with a leading apostrophe (never a date or a formula)."""
    g = datasheet.lines_header(table, cell, fetched)
    return [["'" + x if isinstance(x, str) and x and not x.startswith("=") else x for x in r] for r in g]


def _write_sources(t, q: int, pre: Prepared, ws_data, header, fetched: str) -> str | None:
    """The pull logged on the Sources sheet (xl/sources.py), with each query's OK/DIFF read from the calculated
    Data header and each drill's tie-out from its lines sheet, and the header's Sources cells linked to the rows
    (every draw: the header rewrite cleared them). Provenance, not the answer: a failure (or a query / drill without
    its row) is returned as a warning text, never raised (no rollback); None = every pull is logged."""
    from .xl import sources
    from .xl.jobs import com

    def read():
        return ({qid: ws_data.Range(cell).Value for qid, cell in header.status.items()},
                {a: t.wb.Worksheets(f"Q{q} Lines {a}").Range(datasheet.LINES_STATUS).Value for a in pre.frame.lines})
    try:
        notes: list[str] = []
        got, ties = com(read) if header.status or pre.frame.lines else ({}, {})       # one call for every status
        status = {qid: v if v in ("OK", "DIFF") else "" for qid, v in got.items()}
        tie = {a: v if isinstance(v, str) else "" for a, v in ties.items()}
        rows = sources.write_sources(t, q, pre, ws_data, header, fetched, status, tie, notes)
    except Exception as e:   # noqa: BLE001
        return f"could not log the pull on the Sources sheet ({e}); the answer is built"
    missing = [x for x in [*(qq.id for qq in pre.spec.queries), *pre.frame.lines] if x not in rows]
    if missing:
        notes.append(f"the Sources sheet has no row for {', '.join(missing)}; the answer is built")
    return "; ".join(notes) or None


SOURCES_PROBLEMS = ("could not log", "has no row")      # _write_sources warnings that mean rows are missing


def _check(plan, ws_answer, ws_data, sources_ok: bool, header) -> list[str]:
    """The fast layout check (render/check.py): findings for the reply, never raised."""
    from .render.check import check
    return check(plan, ws_answer, ws_data, sources_ok, header.status)


def _answer_plan(q: int, pre: Prepared, fetched: datetime):
    """The Answer's layout with its print footer (one timestamp per draw); more than 5 free Key points are a
    warning."""
    free, too_many = free_points(pre.spec)
    pre.warnings += [w for w in too_many if w not in pre.warnings]
    plan = layout(q, pre.spec, pre.result, stock=_stock(pre),
                  subtitle=answer_subtitle(pre.spec, pre.result, pre.closed),
                  const=constant_columns(pre.frame), caveats=pre.caveats,
                  key_points=partial(key_points, stock=_stock(pre), free=free))
    plan.footer = answer_footer(fetched)
    return plan


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


def with_timings(job):
    """job(h) -> (its result, its Timer summary). The job runs on the host's COM thread, which does not share the
    tool call's run log (core/runlog.py), so its stage timings travel back with the result. The thread's earlier
    timer (normally none) is put back afterwards, so jobs outside this wrapper stay silent as before."""
    def timed(h):
        before = Timer.current()
        t = Timer.start()
        try:
            return job(h), t.summary()
        finally:
            timing._local.timer = before
    return timed


def _book(t):
    """The workbook's sheets, read once for the whole file job (movein.Book)."""
    from .xl import movein
    return movein.read_book(t)


def _next_q_in(book) -> int:
    from .xl import sheets
    return sheets.next_question_in(book.names())


def _place(book, q: int) -> tuple[str | None, bool]:
    """(the sheet the new answer's sheets go after, first): where sheets.add_sheet would put them (None: first)."""
    from .xl import sheets
    after = sheets.place_after_in(book.names(), q)
    return after, after is None


def _answer_sheets(t, q: int) -> tuple:
    from .xl.jobs import com
    return com(lambda: (t.wb.Worksheets(f"Q{q} Answer"), t.wb.Worksheets(f"Q{q} Data")))


def _lines_sheet(q: int, ls, fetched: str) -> xlsx.LinesSheet:
    """One line set as the built file's 'Q<q> Lines <alias>' sheet: the title, source, cube value and F1:K3 block
    the COM writer gave it before 0.3.0."""
    from .xl.drill_xl import table_name
    table = table_name(q, ls.alias)
    cell = " / ".join(ls.cell.values()) if ls.cell else ls.alias
    return xlsx.LinesSheet(f"Q{q} Lines {ls.alias}", table, f"Ledger lines behind {cell}", LINES_SOURCE,
                           list(ls.columns), ls.rows, ls.cube_value, _lines_grid(table, ls.cell, fetched))


def _temp_path():
    """A new name under config.HOME / tmp (never OneDrive): move_in breaks the copy's links by the file's name."""
    from .core import config
    folder = config.HOME / "tmp"
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{uuid.uuid4().hex}.xlsx"


def _drop_file(path) -> None:
    """The temporary file deleted (also a partial one); a file Excel still holds is logged, never raised."""
    try:
        path.unlink(missing_ok=True)
    except OSError as e:
        log.warning("could not delete the temporary file %s: %s", path, e)


def _height_runs(heights: dict[int, float]) -> list[tuple[float, int, int]]:
    """{row: points} as (points, first row, last row) runs of consecutive rows with the same height."""
    out: list[list] = []
    for r in sorted(heights):
        if out and out[-1][0] == heights[r] and out[-1][2] == r - 1:
            out[-1][2] = r
        else:
            out.append([heights[r], r, r])
    return [(h, a, b) for h, a, b in out]


def _set_breaks(ws, rows: list[int], keep: list[int]) -> None:
    """The sheet's manual page breaks set to rows; when Excel refuses one of them after the reset, the file's breaks
    (keep) are set again before the refusal is raised (a refusal of those too leaves the breaks Excel took)."""
    def put(breaks):
        ws.ResetAllPageBreaks()
        for r in breaks:
            ws.HPageBreaks.Add(ws.Range(f"A{r}"))
    try:
        put(rows)
    except Exception:
        put(keep)
        raise


def _fit_text_rows(ws, plan, block: tuple) -> None:
    """The Key point rows to the height of their calculated text (look.text_heights over the read-back block; one
    call per run of rows with the same new height), then the page breaks again when those heights move them. A
    printer that refuses the breaks keeps the file's (_set_breaks): a break never fails an answer, it is logged."""
    from .render import look
    from .xl import jobs
    lk = look.answer_look(plan)
    heights = look.text_heights(plan, lk, block)
    for h, a, b in _height_runs(heights):
        jobs.com(lambda a=a, b=b, h=h: setattr(ws.Range(f"{a}:{b}"), "RowHeight", h))
    breaks = look.page_breaks(plan, lk, heights)
    if breaks == lk.breaks:
        return
    try:
        jobs.com(lambda: _set_breaks(ws, breaks, lk.breaks))
    except jobs.JobAborted:
        raise
    except Exception as e:   # noqa: BLE001 - no printer: the breaks of the file stay
        log.warning("could not set the page breaks %s on the Answer sheet: %s", breaks, e)


def _file_draw(t, q: int, pre: Prepared, book, parked=None, left: list | None = None) -> dict:
    """The answer's sheets built as a file (Answer look and charts, Data, Lines), moved in after the question's
    place and finished there (tables, formulas, chart labels, calculated); then the read-back, status, Key point
    heights, Sources, spec, sheet order and layout check - all of it with the user's Excel quiet, and
    shown twice on the way when the workbook is on screen: the formatted Answer once it is in, its numbers and
    charts once they are calculated; the Answer stays active. book (movein.Book): the workbook's sheets as read once
    for the job, kept up to date by every step. parked (a follow-up, movein.Parked): the new sheets go where the
    parked old ones are, and the old ones are dropped after the read-back, before Sources and the spec (a drop that
    fails leaves both as they were). left: temporary workbooks that would not close (movein.close_left). The caller
    owns rollback and save; the temporary file is deleted whatever happens."""
    from .xl import movein, sheets
    fetched = datetime.now().astimezone()
    user = datasheet.windows_user()
    stamp = datasheet.fetched_text(fetched, user)
    path = _temp_path()
    try:
        with stage("build"):
            base = movein.base(t)               # the destination's Normal style: the moved sheets measure as drawn
            header = header_rows(q, pre.spec, pre, fetched, user)
            plan = _answer_plan(q, pre, fetched)
            lines = [_lines_sheet(q, ls, stamp) for ls in pre.frame.lines.values()]
            data = xlsx.DataSheet(f"Q{q} Data", header, data_grid(pre.frame))
            try:
                built = xlsx.build(path, f"Q{q} Answer", plan, data, lines, base=base)
            except xlsx.TooManyRows as e:
                raise GuardError(str(e), "narrow the query or the lines cell (a shorter period, a lower account or "
                                 "entity) so it fits on one sheet") from None
        answer = built.sheets[0]
        with movein.quiet(t.app, t.wb) as screen:
            with stage("move in"):
                after, first = _place(book, q) if parked is None else _parked_place(book, parked)
                movein.move_in(t, built, after, first, book=book, left=left)
                if parked is not None:
                    _take_places(t, parked, built.sheets, book)
                if screen.active:               # the active workbook (on screen or not) ends on its new Answer
                    screen.land = answer
                if screen.shown:                # the formatted, still empty answer on screen
                    movein.show(t.app, t.wb, answer)
            with stage("finish"):
                movein.finish(t, built)
                movein.calculate(t, built.sheets)
                if screen.shown:                # its numbers and charts fill in
                    movein.show(t.app, t.wb, answer)
            with stage("read back"):
                ws_answer, ws_data = _answer_sheets(t, q)
                block = read_block(ws_answer, plan)
                mismatches = compare(plan.addr, pre.result.cells, block, plan.fractions)
                write_status(ws_answer, plan, mismatches)
                _fit_text_rows(ws_answer, plan, block)
            if parked is not None:
                movein.drop(t, parked, book, alerts_off=True)
            with stage("sources"):
                made = len(t.sheets_added)
                logged = _write_sources(t, q, pre, ws_data, header, stamp)
                book.sheets += [movein.Sheet(n, None, -1, False) for n in t.sheets_added[made:]]   # a new log, last
            if _save_spec(t, q, pre.raw):
                book.sheets.append(movein.Sheet(sheets.META_SHEET, None, 0, False))
            with stage("arrange"):
                arranged = _arrange(t, answer, book, answer if screen.active else None)
                found = _check(plan, ws_answer, ws_data, not (logged and any(x in logged for x in SOURCES_PROBLEMS)),
                               header)
    finally:
        _drop_file(path)
    return reply(pre, q, str(t.path), list(built.sheets), mismatches,
                 built.chart_lint + [w for w in (logged, arranged) if w], layout=found, charts=built.charts)


def _parked_place(book, parked) -> tuple[str | None, bool]:
    """(the sheet the new answer goes after, first) so it takes the old one's place: the sheet before the first
    parked one (first: the parked sheets come first, the new ones then go before them)."""
    names = book.names()
    parked_names = {n.casefold() for n in parked.sheets.values()}
    i = next(i for i, n in enumerate(names) if n.casefold() in parked_names)
    return (names[i - 1], False) if i else (None, True)


def wanted_order(names: list[str], new: list[str], parked: dict[str, str]) -> list[str]:
    """The workbook's order once the parked sheets are gone, as it was before the follow-up: names (the order after
    the copy) with each new sheet in its parked namesake's place (parked: old name -> parked name; a namesake the
    new answer has no sheet for just goes), and a new sheet without one (an added lines block) right after the new
    sheet before it in new (built order). The first new sheet, the Answer, always has a namesake."""
    back = {p: n for n, p in parked.items() if n in new}
    order = [back.get(n, n) for n in names if n not in new and (n in back or n not in parked.values())]
    for i, n in enumerate(new):
        if n not in order:
            order.insert(order.index(new[i - 1]) + 1 if i else 0, n)
    return order


def _take_places(t, parked, sheets: list[str], book) -> None:
    """The new sheets moved to their old namesakes' places, so the workbook's order is the one before the follow-up
    (one copy lands them side by side; a question's sheets may not be). Only new sheets move, the fewest
    (sheetorder.plan_moves over the order without the parked sheets: a move before or after a sheet that stays is
    the same with them there): none when the old ones stood side by side. Positional Move, one call each; the Book
    (which knows where the copy landed) gives the order and is told the new one."""
    from .xl import sheetorder
    from .xl.jobs import com
    names = book.names()
    staying = [n for n in names if n not in parked.sheets.values()]
    wanted = wanted_order(names, sheets, parked.sheets)
    moves = sheetorder.plan_moves(staying, wanted, set(staying) - set(sheets))
    for n, where, anchor in moves:
        if where == "before":
            com(lambda n=n, a=anchor: t.wb.Worksheets(n).Move(t.wb.Worksheets(a)))
        else:
            com(lambda n=n, a=anchor: t.wb.Worksheets(n).Move(None, t.wb.Worksheets(a)))
    book.order(sheetorder.simulate(names, moves))


def _dropped(t, parked) -> bool:
    """True when none of the parked sheets is in the workbook any more (the drop's Delete went through); False also
    when the workbook cannot be read (then the restore is tried, and says what it could not put back)."""
    from .xl.jobs import com
    try:
        have = com(lambda: {str(s.Name).casefold() for s in t.wb.Worksheets}, limit=30, check=False)
    except Exception:   # noqa: BLE001
        return False
    return not any(n.casefold() in have for n in parked.sheets.values())


KEPT = "; the new answer is in place (the old one is gone)"


def _kept(t, e: BaseException, save: bool = True) -> None:
    """A follow-up that failed once its old sheets were dropped: the new answer is in place and the error says so;
    our own file (never the user's) is saved first, as on success - not when it was the save that failed (save
    False). The error keeps its class."""
    from .xl.jobs import com
    note = KEPT
    if save and _is_ours(t):
        try:
            com(lambda: t.wb.Save(), limit=30, check=False)
            note += " and saved"
        except Exception as e2:   # noqa: BLE001 - the draw's error stays the one raised
            log.warning("could not save %s after a failed follow-up: %s", t.wb.Name, e2)
    if e.args and isinstance(e.args[0], str):
        e.args = (e.args[0] + note, *e.args[1:])
    else:
        e.args = (*e.args, note.lstrip("; "))


def _question_parts(book, q: int) -> tuple[list[str], list[str]]:
    """(sheets, tables) a follow-up replaces, from the Book: 'Q<q> Answer' and 'Q<q> Data' (always named, so park
    refuses a missing one by its name), every 'Q<q> Lines b<i>' present, and the answer's tables on those sheets
    (Data_Q<q>, Lines_Q<q>_<alias>: sheets.answer_sheets / answer_tables)."""
    from .xl import sheets
    names = sheets.answer_sheets(book.names(), q)
    want = [f"Q{q} Answer", f"Q{q} Data"] + [n for n in names if n not in (f"Q{q} Answer", f"Q{q} Data")]
    return want, sheets.answer_tables(q, [n for n, _, s in book.tables() if s in names])


def excel_job(h, pre: Prepared, target: str = "session") -> dict:
    """A new answer (_file_draw) as the next Q<n> of the target workbook. The workbook's sheets are read once. On
    any failure the sheets it brought in are removed (they are in t.sheets_added) and the error is raised. A
    temporary workbook that would not close is tried once more at the end (a warning says so)."""
    from .xl import movein
    with stage("open target"):
        t = _target(h, target, pre.spec.question[:60])
    left: list = []
    with autosave_paused(t) as notes:
        try:
            with stage("draw"):
                book = _book(t)
                out = _file_draw(t, _next_q_in(book), pre, book, left=left)
            with stage("save"):
                _save(t)
        except Exception:
            _rollback(t)
            raise
        finally:
            notes += movein.close_left(left)
    remember(pre, out)
    return _with_notes(out, notes)


def redraw_job(h, pre: Prepared, q: int, target: str = "session") -> dict:
    """A follow-up or refresh: the sheets of Q<q> (Answer, Data, every Lines b<i>) and their tables are parked under
    'old ...' names, the new answer is drawn by _file_draw where they were, and the old ones are dropped; the Sources
    log gets its new rows and _svmeta's row for Q<q> is rewritten. A drawn range an earlier version stored with the
    spec is not used: the sheets are replaced whole. The workbook's sheets are read once. A missing Answer or Data
    sheet is an InputError naming it, before anything is renamed. On a failure the new sheets are removed first
    (rollback), then the old ones get their names back (restore; never a rollback after it, as the old names are then
    the ones in t.sheets_added); RestoreError only when the restore failed. A failure once the old sheets are dropped
    keeps the new answer (a rollback then would leave neither): our own file is saved and the error, of its own
    class, says the new answer is in place (the Sources rows and the spec may be missing then); so does a failing
    save of our file, which is not tried again. A temporary workbook that would not close is tried once more at the
    end (a warning says so)."""
    from .xl import movein
    with stage("open target"):
        t = _target(h, target, pre.spec.question[:60])
    left: list = []
    with autosave_paused(t) as notes:
        with stage("draw"):
            book = _book(t)
            parked = movein.park(t, *_question_parts(book, q), book=book)   # all or nothing: a failed park renamed nothing
            try:
                out = _file_draw(t, q, pre, book, parked, left)
            except Exception as e:
                if _dropped(t, parked):
                    _kept(t, e)
                    raise
                _rollback(t)
                try:
                    movein.restore(t, parked)
                except Exception as e2:   # noqa: BLE001 - the draw's error stays the cause
                    raise RestoreError(f"{e}; the restore was incomplete - {e2}") from e
                raise
            finally:
                notes += movein.close_left(left)
        with stage("save"):
            try:
                _save(t)
            except Exception as e:
                _kept(t, e, save=False)
                raise
        out["followup"] = True
    remember(pre, out)
    return _with_notes(out, notes)


def _stock(pre: Prepared) -> set[str]:
    """Stock metrics (FTE): shown with one decimal (tall1), never summed over months."""
    return {m for m, agg in pre.aggregation.items() if agg == "last"}
