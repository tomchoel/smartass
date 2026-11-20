# engine/smartass/pipeline.py
"""fin_analyze in three steps. prepare(): validate, guard, fetch every grid request in parallel (HTTP, no Excel),
compute the model. excel_job(): runs in the user's Excel (v1 host job) - target, Q<n>, the answer's sheets
built as an .xlsx without Excel and brought in by four batches of Excel operations (xl/inproc: run inside Excel by
its add-in, or over COM): look, move in (the file's sheets inserted with one call), finish (tables, formulas,
calculation, the read-back), close out (status, heights, labels, Sources, spec saved in _svmeta, sheet order); then
a fifth saves our own file. redraw_job() replaces a question's sheets the same way. reply(): the result dict the
agent quotes from (key_values only - never numbers it computed itself)."""
from __future__ import annotations

import contextlib
import logging
import re
import time
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


def _save_spec(t, q: int, raw: dict, b, rows) -> bool:
    """The spec recorded into the close-out batch b, in _svmeta's row for Q<q> (rows: sheets.SpecRows from the look;
    the hidden sheet made first when there is none, as save_meta makes it): True when it is made. No drawn range is
    stored: a follow-up replaces the question's sheets whole."""
    from .xl import sheets
    return sheets.save_spec(t, q, {"v2": raw}, ids={}, create=True, b=b, rows=rows)


def _arrange(t, sheet: str, book, b, wbn, active: bool, sources_ours: bool) -> dict | str:
    """Sheet order like v1 tidy, recorded into the close-out batch b: our own files fully, the user's workbook only
    our sheets among themselves (sheetorder.record_arrange); the order and tabs from the job's Book, no read. active:
    the workbook is the active one (then the Answer is the sheet that is active, and is landed on after a move);
    sources_ours: its 'Sources' sheet is our log. Cosmetic: the dict record_arrange returns (its moves' operations
    may fail), or a warning text when it could not be recorded."""
    from .xl import sheetorder
    try:
        mode = "full" if _is_ours(t) else "active"
        rows = [(s.name, s.visible == -1, s.tab) for s in book.sheets]
        out = sheetorder.record_arrange(b, wbn, rows, sheet if active else None, mode, sheet if active else None,
                                        sources_ours)
        book.order(out["order"])
    except Exception as e:   # noqa: BLE001
        return f"could not arrange the sheets ({e}); the answer is built"
    return out


def _is_ours(t) -> bool:
    """The engine's own analysis file (created by us, or sitting in the analyser folder - a follow-up targets the
    stored path), never the user's workbook."""
    from .core import serverkit as kit
    return bool(getattr(t, "created", False)) or kit._is_our_analysis_file(str(t.path))


AUTOSAVE_NOTE = ("AutoSave was switched off in {} so the answer is not saved for you; save yourself or switch it back "
                 "on.")


def _save(job) -> None:
    """The save batch (not quiet: Excel saves with the user's settings): the print setup's name of a new Sources log
    (once its setup went through), then our own analysis file saved - never the user's workbook - and AutoSave back
    on in it when the look switched it off (a log line when Excel refuses). The user's workbook keeps AutoSave off,
    with a note (AutoSave would otherwise save our edits into their file; v1 does the same). AutoSave stays off while
    the job writes: left on, it wrote our file between the batches (live, 20.11.2026: a write 0.8 s before the save
    batch). No batch at all when there is nothing to do (the user's workbook, no new log): a finished answer is
    never put at risk by a batch that does nothing. A failing save raises with AutoSave still off: the caller's
    clean-up turns it on after its rollback, as before."""
    from .xl import styling
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    ours = _is_ours(job.t)
    if ours or job.setups:
        b = Batch(quiet=False)
        wbn = b.workbook(job.name)
        for sheet, key in job.setups:
            with b.attempt():
                styling.setup_name(wbn.Worksheets(sheet), key)
        on_at = None
        if ours:
            wbn.Save()
            if job.paused:
                on_at = b.position
                with b.attempt():
                    wbn.AutoSaveOn = True
        res = run(job.t.app, b)
        if on_at is not None and any(i >= on_at for i, _, _ in res.skipped):
            log.warning("could not switch AutoSave back on in %s: %s", job.name,
                        next(m for i, _, m in res.skipped if i >= on_at))
    if job.paused:
        job.paused = False
        if not ours:
            job.notes.append(AUTOSAVE_NOTE.format(job.name))


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


def _record_status(b, wbn, q: int, pre: Prepared, header, key: str = "status"):
    """The calculated OK/DIFF the Sources log and the layout check need, read in the finish batch b (operations that
    may fail: an unreadable cell is no status): each query's Data header status cell and each drill's tie-out on its
    lines sheet. The reader gives ({query id: (cell, value)}, {alias: value})."""
    with b.attempt():
        if header.status:
            data = wbn.Worksheets(f"Q{q} Data")
            for qid, cell in header.status.items():
                b.ret(data.Range(cell).Value2, f"{key}.query {qid}")
        for a in pre.frame.lines:
            b.ret(wbn.Worksheets(f"Q{q} Lines {a}").Range(datasheet.LINES_STATUS).Value2, f"{key}.lines {a}")

    def status(res) -> tuple[dict, dict]:
        v = res.values
        return ({qid: (cell, v.get(f"{key}.query {qid}")) for qid, cell in header.status.items()},
                {a: v.get(f"{key}.lines {a}") for a in pre.frame.lines})
    return status


def _write_sources(t, q: int, pre: Prepared, b, wbn, log, header, fetched: str, status: dict, ties: dict,
                   look, book):
    """The pull logged on the Sources sheet (xl/sources.py), recorded into the close-out batch b: each query's OK/DIFF
    from the calculated Data header and each drill's tie-out from its lines sheet (status, ties: _record_status), and
    the header's Sources cells linked to the rows (every draw: the header rewrite cleared them). A new log's sheet
    joins the Book (last). Provenance, not the answer: its operations may fail. Returns (sources.Written, notes), or
    a warning text when it could not be recorded (log None: the finish could not read the log)."""
    from .xl import movein, sources
    if log is None:
        return "could not log the pull on the Sources sheet (the log could not be read); the answer is built"
    try:
        notes: list[str] = []
        made = len(t.sheets_added)
        written = sources.record_write(b, wbn, t, q, pre, log, header, fetched,
                                       {qid: v if v in ("OK", "DIFF") else "" for qid, (_, v) in status.items()},
                                       {a: v if isinstance(v, str) else "" for a, v in ties.items()},
                                       look.separators, look.screen, notes)
        book.sheets += [movein.Sheet(n, -1, False) for n in t.sheets_added[made:]]      # a new log, last
    except Exception as e:   # noqa: BLE001
        return f"could not log the pull on the Sources sheet ({e}); the answer is built"
    return written, notes


def _sources_warning(pre: Prepared, recorded, res) -> str | None:
    """What the close-out did with the Sources log as a warning text (an operation Excel refused, a query or drill
    without its row, a note such as a v1 log left as it is); None = every pull is logged."""
    if isinstance(recorded, str):
        return recorded
    written, notes = recorded
    refused = written.refused(res.skipped)
    if refused is not None:
        return f"could not log the pull on the Sources sheet ({refused}); the answer is built"
    missing = [x for x in [*(qq.id for qq in pre.spec.queries), *pre.frame.lines] if x not in written.rows]
    if missing:
        notes.append(f"the Sources sheet has no row for {', '.join(missing)}; the answer is built")
    return "; ".join(notes) or None


SOURCES_PROBLEMS = ("could not log", "has no row")      # _write_sources warnings that mean rows are missing


def _check(plan, block, charts, status, sources_ok: bool) -> list[str]:
    """The fast layout check (render/check.py) over what the finish batch read: findings for the reply, never
    raised."""
    from .render.check import check
    return check(plan, block, charts, status, sources_ok)


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


def _next_q_in(book) -> int:
    from .xl import sheets
    return sheets.next_question_in(book.names())


def _place(book, q: int) -> tuple[str | None, bool]:
    """(the sheet the new answer's sheets go after, first): where sheets.add_sheet would put them (None: first)."""
    from .xl import sheets
    after = sheets.place_after_in(book.names(), q)
    return after, after is None


def _lines_sheet(q: int, ls, fetched: str) -> xlsx.LinesSheet:
    """One line set as the built file's 'Q<q> Lines <alias>' sheet: the title, source, cube value and F1:K3 block
    the COM writer gave it before 0.3.0."""
    from .xl.drill_xl import table_name
    table = table_name(q, ls.alias)
    cell = " / ".join(ls.cell.values()) if ls.cell else ls.alias
    return xlsx.LinesSheet(f"Q{q} Lines {ls.alias}", table, f"Ledger lines behind {cell}", LINES_SOURCE,
                           list(ls.columns), ls.rows, ls.cube_value, _lines_grid(table, ls.cell, fetched))


def _temp_path():
    """A new name under config.HOME / tmp (never OneDrive): a file of its own for every answer (move_in inserts its
    sheets by its path)."""
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


@dataclass
class Breaks:
    """The page breaks _fit_text_rows recorded: rows - the new ones, keep - the file's (set again when Excel refuses
    one of the new), span - the range of their operations in the batch."""
    rows: list[int]
    keep: list[int]
    span: range


def _fit_text_rows(ws, plan, block: tuple, b) -> Breaks | None:
    """The Key point rows to the height of their calculated text (look.text_heights over the read-back block; one
    operation per run of rows with the same new height), then the page breaks again when those heights move them,
    recorded into b (ws: the Answer's node there). The breaks may fail (a printer that refuses them): the caller sets
    the file's again (_refused_breaks); a break never fails an answer."""
    from .render import look
    lk = look.answer_look(plan)
    heights = look.text_heights(plan, lk, block)
    for h, a, c in _height_runs(heights):
        ws.Range(f"{a}:{c}").RowHeight = h
    breaks = look.page_breaks(plan, lk, heights)
    if breaks == lk.breaks:
        return None
    with b.attempt():
        start = b.position
        ws.ResetAllPageBreaks()
        for r in breaks:
            ws.HPageBreaks.Add(ws.Range(f"A{r}"))
    return Breaks(list(breaks), list(lk.breaks), range(start, b.position))


def _refused_breaks(t, sheet: str, breaks: Breaks | None, res) -> None:
    """When Excel refused a page break of the close-out, the file's breaks are set again in a batch of their own
    (a refusal of those too leaves the breaks Excel took), logged."""
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    from .xl.jobs import book_name
    why = next((m for i, _, m in res.skipped if breaks is not None and i in breaks.span), None)
    if why is None:
        return
    log.warning("could not set the page breaks %s on the Answer sheet: %s", breaks.rows, why)
    b = Batch(calculation=False)
    ws = b.workbook(book_name(t)).Worksheets(sheet)
    with b.attempt():
        ws.ResetAllPageBreaks()
        for r in breaks.keep:
            ws.HPageBreaks.Add(ws.Range(f"A{r}"))
    try:
        run(t.app, b)
    except Exception as e:   # noqa: BLE001 - the breaks of the file stay
        log.warning("could not set the page breaks of the file again: %s", e)


@dataclass
class Look:
    """What the look batch read (_look): the Book, the Excel around (movein.Screen), the destination's base, the
    defined names a table could want, the kind of a 'Sources' sheet (sources_unread: there is one and its kind could
    not be read - the log is then not written, with a warning, rather than taken for ours), where the specs are
    (sheets.SpecRows), and the separators of the user's number formats (Application.International 3 and 4)."""
    book: object
    screen: object
    base: object
    names: list
    kind: str | None
    spec_rows: object
    separators: tuple
    sources_unread: bool = False
    autosave: bool = False                       # AutoSave was on, and the look switched it off


def _look(t, name: str) -> Look:
    """The look batch: every read the job needs before it builds the file (not quiet: it only reads)."""
    from .xl import movein, sheets, sources
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    b = Batch(quiet=False)
    wbn = b.workbook(name)
    book, screen = movein.record_book(b, wbn), movein.record_screen(b, wbn)
    base, names = movein.record_base(b, wbn), movein.record_names(b, wbn)
    kind, spec_rows = sources.record_kind(b, wbn), sheets.record_spec_rows(b, wbn)
    b.ret(b.app.International(3), "decimal separator")
    b.ret(b.app.International(4), "thousands separator")
    with b.attempt():                                   # last: a look that fails before leaves AutoSave alone
        b.ret(wbn.AutoSaveOn, "autosave")               # a local file or an Excel without AutoSave: nothing
        off_at = b.position
        wbn.AutoSaveOn = False                          # off while the job writes (see _save)
    res = run(t.app, b)
    v = res.values
    got = book(res)
    has_sources = "Sources" in got.names()
    found = kind(res) if has_sources else None
    return Look(got, screen(res), base(res), names(res), found, spec_rows(res),
                (v["decimal separator"], v["thousands separator"]), has_sources and found is None,
                bool(v.get("autosave")) and not any(i == off_at for i, _, _ in res.skipped))


@dataclass
class _Job:
    """One file job's state, for its batches and its clean-up: t, q (None: a new answer, numbered by the look),
    name - the workbook's name; look; parked (a follow-up: the plan of its old sheets' names); paused - the look
    switched AutoSave off; moved - the move-in batch ran, in part or whole (the user's sheet is landed on again after a
    failure); inserted / dropped - whether the new sheets came in and the old ones went (None: unknown, the workbook
    says); setups - (sheet, key) of a print setup whose name the save batch adds; notes - for the reply, or the
    error's text when the job fails (final review M2)."""
    t: object
    q: int | None = None
    name: str = ""
    look: Look | None = None
    parked: object = None
    paused: bool = False
    moved: bool = False
    inserted: bool | None = False
    dropped: bool | None = False
    built: object = None
    setups: list = field(default_factory=list)
    notes: list = field(default_factory=list)

    @staticmethod
    def after(e: BaseException, at: int) -> bool | None:
        """Whether a batch's failure e came at or after its operation at; None when that cannot be told (the add-in
        gave no index, or the job was abandoned part-way: the workbook says). Any other failure (a busy Excel)
        refused the batch before it began."""
        from .xl import jobs
        from .xl.inproc.runner import BatchError
        if isinstance(e, BatchError):
            return None if e.index is None else e.index >= at
        return None if isinstance(e, jobs.JobAborted) else False

    def not_inserted(self) -> None:
        """Before a rollback: the new sheets' names leave t.sheets_added when the insert never ran (those names may
        still be the old sheets' own when a park did not finish); unknown: told by the parked names in the
        workbook (the insert comes after the whole park)."""
        if self.built is None or self.inserted is True:
            return
        if self.inserted is None and self.parked is not None:
            have = _sheet_names(self.t)
            if have is not None and all(n.casefold() in have for n in self.parked.sheets.values()):
                return
        elif self.inserted is None:
            return                                      # a new answer: those names were not there before
        for n in self.built.sheets:
            if n in self.t.sheets_added:
                self.t.sheets_added.remove(n)

    def settle(self) -> None:
        """After a failure (and its rollback / restore): the sheet the job found active is active again once the
        move-in ran, and AutoSave back on in our own file (a note for the user's workbook)."""
        from .xl import movein
        from .xl.inproc.batch import Batch
        from .xl.inproc.runner import run
        if self.moved and self.look is not None:
            movein.land(self.t, self.look.screen)
        if not self.paused:
            return
        self.paused = False
        if not _is_ours(self.t):
            self.notes.append(AUTOSAVE_NOTE.format(self.name))
            return
        b = Batch(quiet=False)
        b.workbook(self.name).AutoSaveOn = True
        try:
            run(self.t.app, b, limit=30, check=False)
        except Exception as e:   # noqa: BLE001 - the job's error is the one raised
            log.warning("could not switch AutoSave back on in %s: %s", self.name, e)

    def fail(self, e: BaseException) -> BaseException:
        """The error a failed job raises: a busy Excel inside a batch as the HostError jobs.com gives, and the
        job's notes (AutoSave left off in the user's workbook) at the end of its text."""
        from .xl import jobs
        from .xl.host import HostError
        from .xl.inproc.runner import BatchError
        if isinstance(e, BatchError) and jobs.BUSY in e.message:
            e = HostError(jobs.BUSY)
        if self.notes:
            note = "; " + "; ".join(self.notes)
            e.args = (e.args[0] + note, *e.args[1:]) if e.args and isinstance(e.args[0], str) else (*e.args, note)
        return e


def _sheet_names(t) -> set[str] | None:
    """The workbook's sheet names (casefold), read in a batch of their own (also for an abandoned job); None when
    they cannot be read."""
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    from .xl.jobs import book_name
    try:
        b = Batch(quiet=False)
        b.ret(b.verb("each", b.workbook(book_name(t)).Worksheets, ["Name"], None), "names")
        return {str(n).casefold() for (n,) in run(t.app, b, limit=30, check=False).values["names"]}
    except Exception:   # noqa: BLE001
        return None


def _note_marks(res, stages: tuple = ("",), split: str | None = None, overhead: float = 0.0) -> None:
    """Where a batch's time inside Excel went, noted under the current stage (core/timing.note): each mark is recorded
    right after its step, so a step's time runs from the mark before it - '<stage>excel.<step>'; 'excel.quiet' is the
    settings turned off (up to the first mark), 'excel.end' the time after the last mark (the settings put back, the
    result written). With two stages and split (a mark), the batch's time is booked as two stages: the first up to
    that mark plus overhead (the run's time outside Excel), the second the rest, each with its steps."""
    from .core import timing
    prev, stage, upto = 0.0, stages[0], 0.0
    for name, ms in res.marks.items():
        timing.note(f"{stage}excel.{name}", max(ms - prev, 0.0) / 1000)
        prev = ms
        if name == split:
            upto, stage = ms, stages[1]
    timing.note(f"{stage}excel.end", max(res.ms - prev, 0.0) / 1000)
    if split is not None:
        timing.note(stages[0].rstrip("."), upto / 1000 + overhead)
        timing.note(stages[1].rstrip("."), max(res.ms - upto, 0.0) / 1000)


def _paint(look) -> None:
    """One pause (movein.PAINT_PAUSE, SV_PAINT_PAUSE; 0: none) for Excel to paint the new answer - with its numbers and
    charts - before the close-out batch, when the workbook is the active one of a visible Excel."""
    from .core.timing import stage
    from .xl import movein
    if look.screen.shown and movein.PAINT_PAUSE > 0:
        with stage("paint"):
            time.sleep(movein.PAINT_PAUSE)


def _bring_in(job, built, plan, pre: Prepared, header, book) -> tuple:
    """The move-in and the finish in ONE batch (no Python decision lies between them, and a boundary there painted a
    half-drawn window in the owner's Excel): (a follow-up: the old sheets parked), the file's sheets inserted, moved
    to their old places (a follow-up), then the tables, the formula blocks, the labels' position and colour, the
    calculation, the labels' lifts, the reads the close-out needs (the read-back block, the status cells, the chart
    count, the Sources log), the Answer active (in the active workbook); quiet with Calculation manual (it writes the
    formulas). Its time is booked as the stages 'move in' and 'finish' (split at the last move-in mark), then one
    paint. Returns (block, (status, ties), chart count, the log or None). A failure: inserted is decided by the
    insert's index, as a separate move-in did."""
    from .xl import movein, sources
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    t, q, look, parked = job.t, job.q, job.look, job.parked
    b = Batch()
    b.mark("quiet")
    wbn = b.workbook(job.name)
    if parked is not None:
        movein.record_park(b, wbn, parked, book)
        b.mark("park")
    after, first = _place(book, q) if parked is None else _parked_place(book, parked)
    at = movein.move_in(t, built, after, first, book=book, names=look.names, b=b)
    b.mark("insert")
    job.built = built
    split = "insert"
    if parked is not None:
        _take_places(parked, built.sheets, book, wbn)
        b.mark("moves")
        split = "moves"
    labels = movein.finish(t, built, b)
    b.mark("tables, formulas, labels, calculate")
    answer = wbn.Worksheets(built.sheets[0])
    block = read_block(answer, plan, b)
    status = _record_status(b, wbn, q, pre, header)
    with b.attempt():
        b.ret(answer.ChartObjects().Count, "charts")
    log_ = None
    if not look.sources_unread:
        log_ = sources.log_of(look.kind, book.names(), f"Q{q} Data")
        with b.attempt():                               # the log is provenance: unreadable, it is a warning
            read_log = sources.record_read(b, wbn, log_)
    if look.screen.active:                              # the active workbook (on screen or not) shows its new Answer
        with b.attempt():
            answer.Activate()
    b.mark("reads")
    job.moved = True
    t0 = time.perf_counter()
    try:
        res = run(t.app, b, limit=movein.INSERT_WAIT_S)
    except BaseException as e:
        job.inserted = job.after(e, at)
        raise
    job.inserted = True
    _note_marks(res, ("move in.", "finish."), split, time.perf_counter() - t0 - res.ms / 1000)
    movein.label_skips(res, labels)
    _paint(look)
    got_log = None
    if log_ is not None:
        try:
            got_log = read_log(res)
        except Exception:   # noqa: BLE001
            got_log = None
    return block(res), status(res), res.values.get("charts"), got_log


def _close_out(job, built, plan, pre: Prepared, header, stamp: str, read: tuple, mismatches: list[str],
               book) -> tuple:
    """The close-out batch: status, the Key point heights and page breaks, a follow-up's old
    sheets dropped, the Sources rows, the spec, the sheet order, the Answer active; then what it could not do as
    warnings (labels and breaks logged) and the layout check over the finish's reads. Returns (the Sources warning,
    the order warning, the layout findings)."""
    from .xl import movein, sheets, sources
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    t, q, look = job.t, job.q, job.look
    block, (status, ties), charts, log_ = read
    answer = built.sheets[0]
    b = Batch(calculation=False)
    b.mark("quiet")
    wbn = b.workbook(job.name)
    ws = wbn.Worksheets(answer)
    write_status(ws, plan, mismatches)
    breaks = _fit_text_rows(ws, plan, block, b)
    b.mark("status, heights")
    drop_at = movein.drop(t, job.parked, book, b) if job.parked is not None else None
    if drop_at is not None:
        b.mark("drop")
    recorded = _write_sources(t, q, pre, b, wbn, log_, header, stamp, status, ties, look, book)
    b.mark("sources")
    if _save_spec(t, q, pre.raw, b, look.spec_rows):
        book.sheets.append(movein.Sheet(sheets.META_SHEET, 0, False))
    b.mark("spec")
    ours = log_.name == "Sources" if log_ is not None else look.kind in ("v1", "v2")
    arranged = _arrange(t, answer, book, b, wbn, look.screen.active, ours)
    if look.screen.active:                              # the active workbook ends on its new Answer
        with b.attempt():
            ws.Activate()
    b.mark("order")
    try:
        res = run(t.app, b)
    except BaseException as e:
        job.dropped = False if drop_at is None else None if getattr(e, "index", None) == drop_at else job.after(
            e, drop_at)
        raise
    job.dropped = drop_at is not None
    _note_marks(res, ("",))
    _refused_breaks(t, answer, breaks, res)
    logged = _sources_warning(pre, recorded, res)
    if not isinstance(recorded, str):
        sources.unfreeze(t, recorded[0].sheet, recorded[0], res, look.screen, answer)
    if not isinstance(recorded, str) and recorded[0].setup is not None:
        key, span = recorded[0].setup
        if not any(i in span for i, _, _ in res.skipped):
            job.setups.append((recorded[0].sheet, key))     # the save batch names it (styling.setup_name)
    if isinstance(arranged, dict):
        refused = next((m for i, _, m in res.skipped if i in arranged["moves"]), None)
        arranged = None if refused is None else f"could not arrange the sheets ({refused}); the answer is built"
    found = _check(plan, block, charts, status, not (logged and any(x in logged for x in SOURCES_PROBLEMS)))
    return logged, arranged, found


def _file_draw(job, pre: Prepared) -> dict:
    """The answer's sheets built as a file (Answer look and charts, Data, Lines) and brought into the workbook in three
    batches (xl/inproc): look (every read the build needs: the sheets, the base, the screen, the kind of a Sources
    sheet, the stored specs' rows, the defined names; not quiet); bring in (a follow-up: the old sheets parked; the
    insert, the moves, tables, formula blocks, labels, calculation and the reads the close-out needs; quiet with
    Calculation manual), then one pause for Excel to paint the new answer when the workbook is on screen; close out
    (status, Key point heights and page breaks, a follow-up's old sheets dropped, Sources, spec, sheet order, the
    Answer active; quiet, Calculation left alone). Each quiet batch puts the user's settings back itself. job: _Job
    (q None: a new answer, the next Q<n>; else a follow-up of Q<q>, whose sheets are parked, replaced and dropped
    after the read-back, before Sources and the spec). The caller owns rollback, restore and save; the temporary
    file is deleted whatever happens."""
    from .xl import jobs, movein
    t = job.t
    job.name = jobs.book_name(t)
    fetched = datetime.now().astimezone()
    user = datasheet.windows_user()
    stamp = datasheet.fetched_text(fetched, user)
    path = _temp_path()
    try:
        with stage("look"):
            look = job.look = _look(t, job.name)
            job.paused = look.autosave
        book = look.book
        if job.q is None:
            job.q = _next_q_in(book)
        else:                                           # a missing Answer or Data sheet: an InputError, here
            job.parked = movein.parking(book, *_question_parts(book, job.q))
        q = job.q
        with stage("build"):
            header = header_rows(q, pre.spec, pre, fetched, user)
            plan = _answer_plan(q, pre, fetched)
            lines = [_lines_sheet(q, ls, stamp) for ls in pre.frame.lines.values()]
            data = xlsx.DataSheet(f"Q{q} Data", header, data_grid(pre.frame))
            try:
                built = xlsx.build(path, f"Q{q} Answer", plan, data, lines, base=look.base)
            except xlsx.TooManyRows as e:
                raise GuardError(str(e), "narrow the query or the lines cell (a shorter period, a lower account or "
                                 "entity) so it fits on one sheet") from None
        read = _bring_in(job, built, plan, pre, header, book)     # booked as 'move in' and 'finish' (+ 'paint')
        mismatches = compare(plan.addr, pre.result.cells, read[0], plan.fractions)
        with stage("close out"):
            logged, arranged, found = _close_out(job, built, plan, pre, header, stamp, read, mismatches, book)
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
    the insert) with each new sheet in its parked namesake's place (parked: old name -> parked name; a namesake the
    new answer has no sheet for just goes), and a new sheet without one (an added lines block) right after the new
    sheet before it in new (built order). The first new sheet, the Answer, always has a namesake."""
    back = {p: n for n, p in parked.items() if n in new}
    order = [back.get(n, n) for n in names if n not in new and (n in back or n not in parked.values())]
    for i, n in enumerate(new):
        if n not in order:
            order.insert(order.index(new[i - 1]) + 1 if i else 0, n)
    return order


def _take_places(parked, sheets: list[str], book, wbn) -> None:
    """The new sheets moved to their old namesakes' places, so the workbook's order is the one before the follow-up
    (one insert lands them side by side; a question's sheets may not be), recorded on the workbook's node wbn of the
    move-in batch. Only new sheets move, the fewest (sheetorder.plan_moves over the order without the parked sheets:
    a move before or after a sheet that stays is the same with them there): none when the old ones stood side by
    side. Positional Move; the Book (which knows where the insert landed) gives the order and is told the new one."""
    from .xl import sheetorder
    from .xl.inproc.batch import MISSING
    names = book.names()
    staying = [n for n in names if n not in parked.sheets.values()]
    wanted = wanted_order(names, sheets, parked.sheets)
    moves = sheetorder.plan_moves(staying, wanted, set(staying) - set(sheets))
    for n, where, anchor in moves:
        if where == "before":
            wbn.Worksheets(n).Move(wbn.Worksheets(anchor))
        else:
            wbn.Worksheets(n).Move(MISSING, wbn.Worksheets(anchor))
    book.order(sheetorder.simulate(names, moves))


def _dropped(t, parked) -> bool:
    """True when none of the parked sheets is in the workbook any more (the drop's Delete went through); False also
    when the workbook cannot be read (then the restore is tried, and says what it could not put back)."""
    have = _sheet_names(t)
    return have is not None and not any(n.casefold() in have for n in parked.sheets.values())


KEPT = "; the new answer is in place (the old one is gone)"


def _kept(t, e: BaseException, save: bool = True) -> None:
    """A follow-up that failed once its old sheets were dropped: the new answer is in place and the error says so;
    our own file (never the user's) is saved first, as on success - not when it was the save that failed (save
    False). The error keeps its class."""
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    from .xl.jobs import book_name
    note = KEPT
    if save and _is_ours(t):
        try:
            b = Batch(quiet=False)
            b.workbook(book_name(t)).Save()
            run(t.app, b, limit=30, check=False)
            note += " and saved"
        except Exception as e2:   # noqa: BLE001 - the draw's error stays the one raised
            log.warning("could not save %s after a failed follow-up: %s", book_name(t), e2)
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
    return want, sheets.answer_tables(q, [n for n, s in book.tables() if s in names])


def excel_job(h, pre: Prepared, target: str = "session") -> dict:
    """A new answer (_file_draw) as the next Q<n> of the target workbook. On any failure the sheets it brought in
    are removed (they are in t.sheets_added), the sheet the job found active is active again, AutoSave back on in
    our own file, and the error is raised with the job's notes in its text."""
    with stage("open target"):
        t = _target(h, target, pre.spec.question[:60])
    job = _Job(t)
    try:
        with stage("draw"):
            out = _file_draw(job, pre)
        with stage("save"):
            _save(job)
    except Exception as e:
        job.not_inserted()
        _rollback(t)
        job.settle()
        _raise(job, e)
    remember(pre, out)
    return _with_notes(out, job.notes)


def _raise(job, e: BaseException):
    """Raise the failed job's error (_Job.fail): e itself with its text completed, or the HostError of a busy Excel
    from e."""
    err = job.fail(e)
    if err is e:
        raise e
    raise err from e


def redraw_job(h, pre: Prepared, q: int, target: str = "session") -> dict:
    """A follow-up or refresh: the sheets of Q<q> (Answer, Data, every Lines b<i>) and their tables are parked under
    'old ...' names, the new answer is drawn by _file_draw where they were, and the old ones are dropped; the Sources
    log gets its new rows and _svmeta's row for Q<q> is rewritten. A drawn range an earlier version stored with the
    spec is not used: the sheets are replaced whole. A missing Answer or Data sheet is an InputError naming it,
    before anything is renamed. On a failure the new sheets are removed first (rollback; not when the insert never
    ran: their names were still the old sheets'), then the old ones get their names back (restore; never a rollback
    after it, as the old names are then the ones in t.sheets_added); RestoreError only when the restore failed. A
    failure once the old sheets are dropped keeps the new answer (a rollback then would leave neither): our own file
    is saved and the error, of its own class, says the new answer is in place (the Sources rows and the spec may be
    missing then); so does a failing save of our file, which is not tried again. After any failure the sheet the job
    found active is active again and AutoSave back on in our own file; the job's notes end the error's text."""
    from .xl import movein
    with stage("open target"):
        t = _target(h, target, pre.spec.question[:60])
    job = _Job(t, q)
    with stage("draw"):
        try:
            out = _file_draw(job, pre)
        except Exception as e:
            if job.parked is not None and (job.dropped is True or (job.dropped is None and _dropped(t, job.parked))):
                _kept(t, e)
                job.settle()
                _raise(job, e)
            job.not_inserted()
            _rollback(t)
            if job.moved:                               # the park may have run: the old names go back
                try:
                    movein.restore(t, job.parked)
                except Exception as e2:   # noqa: BLE001 - the draw's error stays the cause
                    job.settle()
                    raise job.fail(RestoreError(f"{e}; the restore was incomplete - {e2}")) from e
            job.settle()
            _raise(job, e)
    with stage("save"):
        try:
            _save(job)
        except Exception as e:
            _kept(t, e, save=False)
            job.settle()
            _raise(job, e)
    out["followup"] = True
    remember(pre, out)
    return _with_notes(out, job.notes)


def _stock(pre: Prepared) -> set[str]:
    """Stock metrics (FTE): shown with one decimal (tall1), never summed over months."""
    return {m for m, agg in pre.aggregation.items() if agg == "last"}
