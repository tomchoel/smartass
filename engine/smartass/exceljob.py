# engine/smartass/exceljob.py
"""The Excel job of fin_analyze (pipeline.py prepares the answer and builds the reply). excel_job(): runs in the
user's Excel (v1 host job) - target, Q<n>, the answer's sheets built as an .xlsx without Excel and brought in by four
batches of Excel operations (xl/inproc: run inside Excel by its add-in, or over COM): look, move in (the file's sheets
inserted with one call), finish (tables, formulas, calculation, the read-back), close out (status, heights, labels,
Sources, spec saved in _svmeta, sheet order); then a fifth saves our own file. redraw_job() replaces a question's
sheets the same way. The first answer of a new workbook builds the whole workbook as its file instead (its Sources
log furnished, the sensitivity label set) and Excel only opens it: no insert, no Sources set-up in Excel. Moved out of
pipeline.py as it was; its log lines keep the logger smartview.pipeline."""
from __future__ import annotations

import dataclasses
import logging
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from functools import partial

from .core import timing
from .core.timing import Timer, stage
from .model.spec import LinesBlock
from .model.validate import GuardError
from .pipeline import Prepared, _stock, reply
from .render import datasheet, xlsx
from .render.datasheet import header_rows
from .render.excel import answer_footer, data_grid, read_block, write_status
from .render.layout import answer_subtitle, constant_columns, layout
from .render.keypoints import free_points, key_points
from .render.readback import compare

log = logging.getLogger("smartview.pipeline")


class RestoreError(RuntimeError):
    """A follow-up failed AND its restore was incomplete: the workbook may be damaged, nothing more should be drawn."""


def _target(h, target: str, title: str, build_new: bool = False):
    """The job's workbook (xl/target.resolve_target). h: the job's runner, which every job seam takes; the resolve
    needs none of it."""
    from .xl import target as xltarget
    return xltarget.resolve_target(target, allow_archive=False, title=title, build_new=build_new)


def _new_book() -> tuple | None:
    """(the base, the theme) a new workbook's file is built in when the first answer's file IS the new workbook:
    what a workbook added from the TN template has (target.new_base, styling.template_theme). None - then a new
    workbook is made from the template as before - without our label id and site (the file could not be saved), without a
    stored base or the template's theme, or with a standard column XlsxWriter has no digit width for (the widths would
    not measure as the template's)."""
    from .core import config
    from .xl import styling
    from .xl import target as xltarget
    if not (config.LABEL_ID and config.LABEL_SITE):
        return None
    base, theme = xltarget.new_base(), styling.template_theme()
    if base is None or theme is None or base.col_px not in xlsx.STANDARD_COLUMNS:
        return None
    return base, theme


def _save_spec(t, q: int, raw: dict, b, rows) -> bool:
    """The spec recorded into the close-out batch b, in _svmeta's row for Q<q> (rows: sheets.SpecRows from the look;
    the hidden sheet made first when there is none): True when it is made. No drawn range is
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


AUTOSAVE_LEFT = ("AutoSave could not be switched back on in {} ({}): the next answer in this Excel switches it on, "
                 "or switch it on yourself")
AUTOSAVE_YOURSELF = "AutoSave could not be switched back on in {} ({}): switch it on yourself"   # nothing recorded it
AUTOSAVE_AFTER = "AutoSave could not be switched back on in {} after the save after the last answer ({})"
ON_WITH_THIS = "; this answer switches it on"
ON_NOW = "; it is switched on now"
NOT_ON_NOW = "; it could not be switched on now either ({}): switch it on yourself"
NOT_KNOWN_ON = "; which file it was is not known: switch it on yourself"
ANOTHER_LEFT = "; the workbook open by that name now is another file, so it was left alone"
AUTOSAVE_NOTE = ("AutoSave was switched off in {} so the answer is not saved for you; save yourself or switch it back "
                 "on.")
SAVE_FAILED = "the save of {} after the last answer failed ({})"
SAVE_CLOSED = ("{} was closed before smartass saved it after the last answer (Excel asks whether to save it when it "
               "closes)")
SAVED_NOW = "; it is saved now"
NOT_SAVED_NOW = "; it could not be saved now either ({}): save it yourself"
NOT_OPEN_NOW = "; it is not open any more"
NO_PATH = "; which file it was is not known, so it was not saved: save it yourself"
SAVE_UNKNOWN = ("smartass could not tell how its save after the last answer went ({}): this answer is saved in the "
                "job")
ANOTHER_NOW = "; the workbook open by that name now is another file, so it was not saved"
OTHER_SERVER = " (asked by another smartass server in this Excel)"     # an outcome another build's add-in told


def _save(job) -> None:
    """The save step (not quiet: Excel saves with the user's settings). Our own analysis file - never the user's
    workbook - is saved by Excel itself right after the reply when the add-in is in that Excel (xl/inproc/later.py:
    one Run asks for it; the save, then AutoSave back on when the look switched it off, run once Excel has painted the
    answer); otherwise (the COM executor, SV_SAVE=sync, a workbook this request created - target.Target.new: its
    first answer is on disk before the session points at it, so a crash right after the reply leaves no hollow file
    (final review I1) -, a request the add-in did not take, or a save after the last answer that failed -
    _earlier_save) the save batch saves it here and switches AutoSave back on (a note when Excel refuses). The save batch also names the print setup of a new Sources log (once its setup went through),
    before the save is asked for. Another workbook of ours whose save after the last answer failed is saved after
    this one's (_resave; a note says how that went). The user's workbook keeps AutoSave off, with a note (AutoSave
    would otherwise save our edits into their file; v1 does the same). AutoSave stays off while the job writes: left
    on, it wrote our file during the job's batches over COM (live, 20.11.2026, scripts/check_followups.py --settle 10
    with SV_INPROC=0: a write 1.75 s before the save batch; in-process the file was written only by the save). No
    batch at all when there is nothing to do (the user's workbook, no new log, no earlier save to repeat): a finished
    answer is never put at risk by a batch that does nothing. A failing save raises with AutoSave still off: the
    caller's clean-up turns it on after its rollback, as before (and _Job.fail puts a failed earlier save back for
    the next job)."""
    from .xl.inproc import later
    ours = _is_ours(job.t)
    again = job.resave
    same = again is not None and _same_file(again[2].get("full_name"), job)
    later_ok = (ours and again is None and not job.sync_save and not getattr(job.t, "new", False)
                and later.available(job.t.app))
    if job.setups or (ours and not later_ok):
        _save_batch(job, save=ours and not later_ok)
    if later_ok and later.queue(job.t.app, job.name, job.paused) is None:
        _save_batch(job, save=True)
    if again is not None:
        if same and ours:
            job.notes[again[1]] += SAVED_NOW              # this job's own save was that workbook's
        elif not same:
            _resave(job, *again)
        job.resave = None
    if job.paused:
        job.paused = False
        if not ours:
            job.notes.append(AUTOSAVE_NOTE.format(job.name))


def _save_batch(job, save: bool) -> None:
    """The save batch of _save: the print setups' names (once), then our own file saved and AutoSave back on in it
    (save)."""
    from .xl import styling
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    b = Batch(quiet=False)
    wbn = b.workbook(job.name)
    for sheet, key in job.setups:
        with b.attempt():
            styling.setup_name(wbn.Worksheets(sheet), key)
    job.setups = []
    on_at = None
    if save:
        wbn.Save()
        if job.paused:
            on_at = b.position
            with b.attempt():
                wbn.AutoSaveOn = True
    res = run(job.t.app, b)
    if on_at is not None and any(i >= on_at for i, _, _ in res.skipped):
        job.autosave_left(next(m for i, _, m in res.skipped if i >= on_at))


def _same_file(full: str | None, job) -> bool:
    """The failed save's file (its FullName as the request found it) is this job's workbook (the look's FullName)."""
    mine = job.look.full_name if job.look is not None else ""
    return bool(full) and bool(mine) and str(full).casefold() == mine.casefold()


def _resave(job, name: str, note: int, out: dict) -> None:
    """Another workbook of ours whose save after the last answer failed (out: its outcome), saved now - only when the
    workbook open by that name is still that file (its FullName as the request found it): one batch reads it, one
    saves it; an outcome without that FullName saves nothing (the user's own workbook can have our file's name).
    job.notes[note] says how it went. Never raises (this job's answer is in place and saved)."""
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    if not out.get("full_name"):
        job.notes[note] += NO_PATH
        return
    try:
        b = Batch(quiet=False)
        with b.attempt():
            b.ret(b.workbook(name).FullName, "full")
        now = run(job.t.app, b).values.get("full")
        full = out["full_name"]
        if now is None:
            job.notes[note] += NOT_OPEN_NOW
            return
        if str(now).casefold() != full.casefold():
            job.notes[note] += ANOTHER_NOW
            return
        b = Batch(quiet=False)
        b.workbook(name).Save()
        run(job.t.app, b)
        job.notes[note] += SAVED_NOW
    except Exception as e:   # noqa: BLE001 - said in the note
        job.notes[note] += NOT_SAVED_NOW.format(getattr(e, "message", None) or e)


def _earlier_save(job) -> None:
    """The save the job before left to Excel in this Excel (xl/inproc/later.py), read once right after the look (whose
    Run ran it if it was still waiting): its seconds go to this job's timings as 'previous save' (outside its stages);
    one that failed is a note of the reply, and this job saves in the foreground (_save) - that workbook too when it
    is another one (_resave); if this job fails before its save, the outcome is put back for the next job
    (_Job.fail). One whose workbook was closed first is a note only (Excel asked the user whether to save it)."""
    from .core import timing
    from .xl.inproc import later
    out = later.take(job.t.app)
    if out is None:
        return
    timing.aside("previous save", float(out.get("seconds") or 0.0))
    if out.get("unknown"):
        job.notes.append(SAVE_UNKNOWN.format(out.get("error")) + (OTHER_SERVER if out.get("other_server") else ""))
        job.sync_save = True                            # which file it was is not known: this one is saved in the job
        return
    book = out["workbook"] + (OTHER_SERVER if out.get("other_server") else "")
    if out.get("autosave_error"):
        _autosave_again(job, out, book)
    if out["ok"]:
        return
    if out.get("closed"):
        job.notes.append(SAVE_CLOSED.format(book))
        return
    job.notes.append(SAVE_FAILED.format(book, out.get("error") or "no reason given"))
    job.resave = (out["workbook"], len(job.notes) - 1, out)


def _autosave_again(job, out: dict, book: str) -> None:
    """The save after the last answer could not switch AutoSave back on in its workbook (the outcome's autosave_error;
    final review M6): a note, and AutoSave switched on - by this job's own save when it is this job's workbook (as if
    its look had switched it off: also recorded for the next job, and its clean-up switches it on), else now, only in
    the workbook open by that name when it is still that file (its FullName as the request found it). Never raises."""
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    log.warning("could not switch AutoSave back on in %s: %s", book, out["autosave_error"])
    job.notes.append(AUTOSAVE_AFTER.format(book, out["autosave_error"]))
    note = len(job.notes) - 1
    full = out.get("full_name")
    if not full:
        job.notes[note] += NOT_KNOWN_ON
        return
    if _same_file(full, job):
        job.paused = True
        job.notes[note] += ON_WITH_THIS
        return
    try:
        b = Batch(quiet=False)
        with b.attempt():
            b.ret(b.workbook(out["workbook"]).FullName, "full")
        now = run(job.t.app, b).values.get("full")
        if now is None:
            job.notes[note] += NOT_OPEN_NOW
            return
        if str(now).casefold() != full.casefold():
            job.notes[note] += ANOTHER_LEFT
            return
        b = Batch(quiet=False)
        b.workbook(out["workbook"]).AutoSaveOn = True
        run(job.t.app, b)
        job.notes[note] += ON_NOW
    except Exception as e:   # noqa: BLE001 - said in the note
        job.notes[note] += NOT_ON_NOW.format(getattr(e, "message", None) or e)


def _settings_left(job) -> None:
    """Before the look reads the settings: what a job in this Excel that did not finish left changed is put back
    (xl/jobstate.put_back), said in the notes."""
    from .xl import jobstate
    job.excel = jobstate.key_of(job.t)
    note = jobstate.put_back(job.t.app, job.excel)
    if note:
        job.notes.append(note)


def _record_settings(job) -> None:
    """After the look, before the first quiet batch: the settings it found and what the job sets, with our own file
    whose AutoSave the look switched off (xl/jobstate.begin; the job's end forgets them)."""
    from .xl import jobstate
    ours = (job.name, job.look.full_name) if job.paused and job.look.full_name and _is_ours(job.t) else None
    job.state = jobstate.begin(job.excel, job.look.settings, ours)


def _rollback(t) -> None:
    from .xl import sheets
    sheets.rollback(t)


def _forget(job) -> None:
    """A new answer that failed once its close-out ran, in whole or in part: its spec's row of _svmeta and its rows of
    the Sources log (when they were there before the job: one it made went with its sheets) cleared, so no question is
    left that has no sheets (final review M5: a default follow-up picked it, and the next answer's Supersedes linked
    to it). One batch of operations that may fail, also for an abandoned job; logged, never raised (the job's error
    is the one raised). Known limits keep a line the answer may have added."""
    from .xl import sheets, sources
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    if job.spec_row is None and job.logged is None:
        return
    b = Batch(calculation=False)
    with b.attempt():
        wbn = b.workbook(job.name)
        if job.spec_row is not None:
            wbn.Worksheets(sheets.META_SHEET).Range(f"{job.spec_row}:{job.spec_row}").ClearContents()
        if job.logged is not None:
            sheet, first, last = job.logged
            last_col = sources._col(len(sources.LOG_COLUMNS))
            wbn.Worksheets(sheet).Range(f"A{first}:{last_col}{last}").ClearContents()
    try:
        res = run(job.t.app, b, limit=30, check=False)
    except Exception as e:   # noqa: BLE001 - the job's error stays the one raised
        log.warning("could not remove the spec and Sources rows of the failed Q%s: %s", job.q, e)
        return
    for _, _, why in res.skipped:
        log.warning("could not remove the spec and Sources rows of the failed Q%s: %s", job.q, why)


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


def _style_findings(plan, block) -> list[str]:
    """The report style's text findings over the read-back block (render/look.style_findings: a text cut off or
    that does not fit, apply_report's words) for the reply's warnings, first, as before 0.3.0 (final review M1);
    never raised."""
    from .render import look
    try:
        return look.style_findings(plan, block)
    except Exception as e:   # noqa: BLE001 - a finding is never worth a failed answer
        log.warning("the style findings could not be made: %s", e)
        return []


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
    kind of a 'Sources' sheet (sources_unread: there is one and its kind could not be read - the log is then not
    written, with a warning, rather than taken for ours), where the specs are (sheets.SpecRows), and the separators
    of the user's number formats (Application.International 3 and 4)."""
    book: object
    screen: object
    base: object
    kind: str | None
    spec_rows: object
    separators: tuple
    sources_unread: bool = False
    autosave: bool = False                       # AutoSave was on, and the look switched it off
    full_name: str = ""                          # the workbook's FullName (which file a failed earlier save was)
    settings: dict = field(default_factory=dict)     # the quiet settings as found (xl/jobstate.record_found)


def _look(job) -> Look:
    """The look batch: every read the job needs before it builds the file (not quiet: it only reads). Whether it
    switched AutoSave off goes to the job (job.paused) before any of its results is read, so a reader that fails
    leaves the clean-up to switch it back on. A run whose outcome is unknown (the add-in may have run the whole batch,
    the AutoSave-off included) counts as paused in our own file, so the clean-up switches it on there; in the user's
    workbook it is left as the look may have left it, without a note (whether it went off cannot be told)."""
    from .xl import jobstate, movein, sheets, sources
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import BatchError, run
    t = job.t
    b = Batch(quiet=False)
    wbn = b.workbook(job.name)
    book, screen = movein.record_book(b, wbn), movein.record_screen(b, wbn)
    found = jobstate.record_found(b)
    base = movein.record_base(b, wbn)
    kind, spec_rows = sources.record_kind(b, wbn), sheets.record_spec_rows(b, wbn)
    b.ret(b.app.International(3), "decimal separator")
    b.ret(b.app.International(4), "thousands separator")
    with b.attempt():
        b.ret(wbn.FullName, "full name")
    with b.attempt():                                   # last: a look that fails before leaves AutoSave alone
        b.ret(wbn.AutoSaveOn, "autosave")               # a local file or an Excel without AutoSave: nothing
        off_at = b.position
        wbn.AutoSaveOn = False                          # off while the job writes (see _save)
    try:
        res = run(t.app, b)
    except BatchError as e:
        if (e.index is None or e.index >= off_at) and _is_ours(t):
            job.paused = True                           # may be off now: the clean-up switches it on (our file)
        raise
    v = res.values
    job.paused = bool(v.get("autosave")) and not any(i == off_at for i, _, _ in res.skipped)
    got = book(res)
    has_sources = "sources" in {n.casefold() for n in got.names()}     # as Excel: names without case
    kind_found = kind(res) if has_sources else None
    return Look(got, screen(res), base(res), kind_found, spec_rows(res),
                (v["decimal separator"], v["thousands separator"]), has_sources and kind_found is None, job.paused,
                str(v.get("full name") or ""), found(res))


@dataclass
class _Job:
    """One file job's state, for its batches and its clean-up: t, q (None: a new answer, numbered by the look),
    name - the workbook's name; look; parked (a follow-up: the plan of its old sheets' names); paused - the look
    switched AutoSave off; moved - the move-in batch ran, in part or whole (the user's sheet is landed on again after a
    failure); inserted / dropped - whether the new sheets came in and the old ones went (None: unknown, the workbook
    says); mentions / repointed - a follow-up's formulas elsewhere that point at its sheets, and whether the
    close-out re-pointed them (None: unknown; a failure before the drop points them back); setups - (sheet, key) of
    a print setup whose name the save batch adds; notes - for the reply, or the error's text when the job fails
    (final review M2); pointing - what the re-pointing could not do, for the reply's warnings."""
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
    new_book: tuple | None = None               # (base, theme) a new workbook the job builds is built in (_new_book)
    mentions: object = None                     # a follow-up's formulas elsewhere to re-point (movein.Mentions)
    repointed: bool | None = False              # the close-out re-pointed them (None: unknown - pointed back)
    setups: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    resave: tuple | None = None                 # (workbook, its note's index, outcome): a save after the reply failed
    sync_save: bool = False                     # how the save after the last reply went is not known: save in the job
    pointing: list = field(default_factory=list)    # what the re-pointing could not do, for the reply's warnings
    excel: tuple | None = None                  # (pid, start time) of its Excel (xl/jobstate.key_of)
    state: tuple | None = None                  # the settings it recorded before its first quiet batch (jobstate)
    keep_state: bool = False                    # AutoSave could not be switched back on: the next job does it
    spec_stored: bool | None = False            # the close-out stored the spec (None: unknown - the workbook says)
    spec_row: int | None = None                 # the row of _svmeta the close-out writes the spec in (_forget)
    logged: tuple | None = None                 # (sheet, first row, last row) of the Sources rows it writes (_forget)

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
            self.autosave_left(getattr(e, "message", None) or str(e))

    def autosave_left(self, why: str) -> None:
        """AutoSave could not be switched back on in our own file: said in the notes, and the recorded settings kept
        for the next job in this Excel, which switches it on (xl/jobstate) - the note promises that only when the
        state recorded this workbook (final review M7: not after a look that failed before the state was written, or
        whose read of the full name failed)."""
        from .xl import jobstate
        log.warning("could not switch AutoSave back on in %s: %s", self.name, why)
        full = self.look.full_name if self.look is not None else ""
        said = AUTOSAVE_LEFT if jobstate.records_autosave(self.state, full) else AUTOSAVE_YOURSELF
        self.notes.append(said.format(self.name, why))
        self.keep_state = True

    def end(self) -> None:
        """The job is over (its clean-up done): the settings it recorded are forgotten (xl/jobstate.end) - kept when
        AutoSave could not be switched back on (autosave_left)."""
        from .xl import jobstate
        jobstate.end(self.state, keep=self.keep_state)
        self.state = None

    def fail(self, e: BaseException) -> BaseException:
        """The error a failed job raises: a busy Excel inside a batch as the HostError jobs.com gives, and the
        job's notes (AutoSave left off in the user's workbook, a save after the last answer that failed) at the end of
        its text. A failed save after the last answer this job took but did not repeat is put back for the next job
        (xl/inproc/later.py give_back) - not after a RestoreError (the workbook may be damaged: nothing saves it).
        The job is over then (end)."""
        from .xl import jobs
        from .core.errors import HostError
        from .xl.inproc import later
        from .xl.inproc.runner import BatchError
        self.end()
        if self.resave is not None and not isinstance(e, RestoreError):
            later.give_back(self.t.app, self.resave[2])
            self.resave = None
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
        b.ret(b.verb("each", b.workbook(book_name(t)).Worksheets, ["Name"]), "names")
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
    count, the Sources log, a follow-up's formulas elsewhere that point at the parked sheets: movein.record_mentions),
    the Answer active (in the active workbook); quiet with Calculation manual (it writes the formulas). Its time is
    booked as the stages 'move in' and 'finish' (split at the last move-in mark), then one paint. Returns (block,
    (status, ties), chart count, the log or None). A failure: inserted is decided by the insert's index, as a
    separate move-in did. A new workbook the job built (t.made) has its sheets already: no name check, no park and no
    insert - the batch is the finish."""
    from .xl import movein, sources
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    t, q, look, parked = job.t, job.q, job.look, job.parked
    new = t.made is not None                            # the built file is the workbook: its sheets are there
    b = Batch()
    b.mark("quiet")
    wbn = b.workbook(job.name)
    free_at = at = None
    if new:
        split = "quiet"                                 # nothing to move in: 'move in' is the settings turned off
    else:
        free_at = movein.record_free(b, wbn, [spec.name for spec in built.tables])     # first: nothing renamed yet
        if parked is not None:
            movein.record_park(b, wbn, parked, book)
            b.mark("park")
        after, first = _place(book, q) if parked is None else _parked_place(book, parked)
        at = movein.move_in(t, built, after, first, book=book, b=b, free_at=free_at)
        b.mark("insert")
        split = "insert"
    job.built = built
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
    mentions = movein.record_mentions(b, wbn, parked, built) if parked is not None else None
    if look.screen.active:                              # the active workbook (on screen or not) shows its new Answer
        with b.attempt():
            answer.Activate()
    b.mark("reads")
    job.moved = True
    t0 = time.perf_counter()
    try:
        res = run(t.app, b, limit=movein.INSERT_WAIT_S)
    except BaseException as e:
        taken = movein.name_taken(e, free_at, job.name)
        if taken is not None:                           # refused before its park and insert: nothing ran
            job.inserted, job.moved = False, False
            try:
                raise taken from None
            finally:
                taken = None            # held by this frame, which its traceback holds: a cycle with the job's objects
        job.inserted = True if new else job.after(e, at)
        raise
    job.inserted = True
    if mentions is not None:
        job.mentions = mentions(res)
        job.pointing += _mentions_notes(q, job.mentions)
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
    """The close-out batch: status, the Key point heights and page breaks, a follow-up's formulas elsewhere re-pointed
    at the new sheets (each a compare-and-set that may fail, with Calculation manual) right before its old sheets
    are dropped, the Sources rows, the
    spec, the sheet order, the Answer active; then what it could not do as warnings (labels and breaks logged) and
    the layout check over the finish's reads. Returns (the Sources warning,
    the order warning, the layout findings)."""
    from .xl import movein, sheets, sources
    from .xl.inproc.batch import Batch
    from .xl.inproc.runner import run
    t, q, look = job.t, job.q, job.look
    block, (status, ties), charts, log_ = read
    answer = built.sheets[0]
    rewrites = job.mentions.plan if job.parked is not None and job.mentions is not None else []
    b = Batch(calculation=bool(rewrites))               # with re-pointing writes: manual (each would recalculate)
    b.mark("quiet")
    wbn = b.workbook(job.name)
    ws = wbn.Worksheets(answer)
    write_status(ws, plan, mismatches)
    breaks = _fit_text_rows(ws, plan, block, b)
    b.mark("status, heights")
    unread = job.mentions.unread if job.parked is not None and job.mentions is not None else []
    counted = bool(unread) and look.settings.get("Calculation") == AUTOMATIC    # the user's manual: not calculated
    if counted:                                         # before anything of the follow-up is written
        _record_refs(b, wbn, unread, "refs before")
    writes = movein.record_repoint(b, wbn, rewrites) if rewrites else None
    point_at = writes.start if writes is not None else None
    if writes is not None:
        b.mark("re-point")
    drop_at = movein.drop(t, job.parked, book, b) if job.parked is not None else None
    if drop_at is not None:
        b.mark("drop")
        if counted:
            _record_refs(b, wbn, unread, "refs after", calculate=bool(rewrites))
    recorded = _write_sources(t, q, pre, b, wbn, log_, header, stamp, status, ties, look, book)
    if not isinstance(recorded, str) and log_.exists and recorded[0].rows:     # a new log goes with t.sheets_added
        job.logged = (recorded[0].sheet, min(recorded[0].rows.values()), max(recorded[0].rows.values()))
    b.mark("sources")
    if _save_spec(t, q, pre.raw, b, look.spec_rows):
        book.sheets.append(movein.Sheet(sheets.META_SHEET, 0, False))
        t.sheets_added.append(sheets.META_SHEET)        # made by this job: it goes with the job's sheets
    else:
        job.spec_row = look.spec_rows.row(f"SPEC Q{q}")[0]
    spec_end = b.position
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
        job.repointed = False if point_at is None else job.after(e, point_at)
        job.spec_stored = job.after(e, spec_end)        # its operations are no attempts: past them, it is stored
        raise
    job.dropped = drop_at is not None
    job.repointed = point_at is not None
    job.spec_stored = True
    if writes is not None:
        refused, changed = writes.outcome(res.values, res.skipped)
        job.pointing += _refused_notes(q, refused, "point", "at the new") + _changed_notes(q, changed)
    job.pointing += _unread_notes(q, unread, res.values if counted else None)
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
    sheet, the stored specs' rows; not quiet); bring in (the answer's table names looked up among the defined names
    first - a new answer's number comes from the look - then a follow-up's old sheets parked; the insert, the moves,
    tables, formula blocks, labels, calculation and the reads the close-out needs; quiet with Calculation manual),
    then one pause for Excel to paint the new answer when the workbook is on screen; close out (status, Key point
    heights and page breaks, a follow-up's formulas elsewhere re-pointed at the new sheets and its old sheets
    dropped, Sources, spec, sheet order, the Answer active; quiet, Calculation left alone). Each quiet batch puts the
    user's settings back itself. job: _Job (q None: a new answer, the next Q<n>; else a follow-up of Q<q>, whose
    sheets are parked, replaced and dropped after the read-back, before Sources and the spec). The caller owns
    rollback, restore and save; the temporary file is deleted whatever happens. A new workbook the job builds
    (job.t.made, the first answer: Q1) is built first - at its reserved name, as the workbook itself (xlsx.NewBook) in
    the base _new_book gave and the date format of the chosen Excel - and opened (stage 'open'); then the same look,
    a bring-in that is the finish only, the close-out."""
    from .xl import jobs, movein
    from .xl import target as xltarget
    t = job.t
    job.name = jobs.book_name(t)
    fetched = datetime.now().astimezone()
    user = datasheet.windows_user()
    stamp = datasheet.fetched_text(fetched, user)
    new = t.made is not None
    path = None if new else _temp_path()
    try:
        if new:                                         # the first answer of a new workbook: its file IS the workbook
            job.q = 1
            base, theme = job.new_book
            built, plan, header = _build(1, pre, t.made, fetched, user, stamp, _dated(t, base),
                                         xlsx.NewBook(xltarget.label_properties(), theme))
            with stage("open"):
                xltarget.open_new(t)
        with stage("look"):
            _settings_left(job)
            look = job.look = _look(job)
            _earlier_save(job)
        _record_settings(job)
        book = look.book
        if t.from_template:                             # what a new workbook has: the next one is built in it
            xltarget.remember_new_base(look.base)
        if new:
            _check_new_base(job.new_book[0], look.base)
        if not new:
            if job.q is None:
                job.q = _next_q_in(book)
            else:                                       # a missing Answer or Data sheet: an InputError, here
                job.parked = movein.parking(book, *_question_parts(book, job.q))
            built, plan, header = _build(job.q, pre, path, fetched, user, stamp, look.base)
        q = job.q
        read = _bring_in(job, built, plan, pre, header, book)     # booked as 'move in' and 'finish' (+ 'paint')
        mismatches = compare(plan.addr, pre.result.cells, read[0], plan.fractions)
        with stage("close out"):
            logged, arranged, found = _close_out(job, built, plan, pre, header, stamp, read, mismatches, book)
    finally:
        if path is not None:
            _drop_file(path)
    return reply(pre, q, str(t.path), list(built.sheets), mismatches,
                 _style_findings(plan, read[0]) + built.chart_lint + [w for w in (logged, arranged) if w]
                 + job.pointing, layout=found, charts=built.charts)


def _build(q: int, pre: Prepared, path, fetched: datetime, user: str, stamp: str, base, book=None) -> tuple:
    """The answer's sheets built as a file at path in base (stage 'build'), a workbook of its own with book
    (xlsx.NewBook): (BuiltFile, the Answer's plan, the Data header). More rows than a sheet has is a GuardError with
    its fix."""
    with stage("build"):
        header = header_rows(q, pre.spec, pre, fetched, user)
        plan = _answer_plan(q, pre, fetched)
        lines = [_lines_sheet(q, ls, stamp) for ls in pre.frame.lines.values()]
        data = xlsx.DataSheet(f"Q{q} Data", header, data_grid(pre.frame))
        try:
            built = xlsx.build(path, f"Q{q} Answer", plan, data, lines, base=base, book=book)
        except xlsx.TooManyRows as e:
            raise GuardError(str(e), "narrow the query or the lines cell (a shorter period, a lower account or "
                             "entity) so it fits on one sheet") from None
    return built, plan, header


def _check_new_base(built, found) -> None:
    """The base Excel gives a workbook built in the stored base (the look's), against that base: on a mismatch (a
    default font or template the stored base no longer describes) a warning, and the stored base dropped - the next
    new workbook is made from the template once and stores a fresh one. The answer is kept either way."""
    from .xl import target as xltarget
    key = lambda b: (b.font_name, b.font_size, b.row_px, b.col_px)     # noqa: E731 - the date format is Excel's own
    if key(found) != key(built):
        log.warning("a new workbook built in the stored base %s opened with %s: the stored base is dropped (the next "
                    "new workbook is made from the template and stores a fresh one)", key(built), key(found))
        xltarget.forget_new_base()


def _dated(t, base):
    """base with the date format of t's Excel (Application.International, one read): how the Lines sheets' dates are
    stored there (movein.date_format), as the look reads it for an existing workbook."""
    from .xl import jobs, movein
    intl = jobs.com(lambda: t.app.International)
    return dataclasses.replace(base, date_format=movein.date_format(*(intl[i - 1] for i in movein.DATE_SETTINGS)))


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
SPEC_NOT_STORED = ("; its spec could not be stored, so a next follow-up of Q{q} would patch the spec before it: "
                   "repeat this follow-up to store it")
SPEC_UNKNOWN = ("; whether its spec was stored is not known (if not, a next follow-up of Q{q} would patch the spec "
                "before it): repeat this follow-up to be sure")


def _kept(t, e: BaseException, save: bool = True, job=None) -> None:
    """A follow-up that failed once its old sheets were dropped: the new answer is in place and the error says so;
    our own file (never the user's) is saved first, as on success - not when it was the save that failed (save
    False). With job (a failure in the close-out): a spec the close-out did not store, or may not have stored, is
    said too (final review M4: the next follow-up would patch the spec before it). The error keeps its class."""
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
    if job is not None and job.spec_stored is not True:
        note += (SPEC_NOT_STORED if job.spec_stored is False else SPEC_UNKNOWN).format(q=job.q)
    if e.args and isinstance(e.args[0], str):
        e.args = (e.args[0] + note, *e.args[1:])
    else:
        e.args = (*e.args, note.lstrip("; "))


def _mentions_notes(q: int, found) -> list[str]:
    """What the look for the formulas pointing at Q<q> could not do as a whole, as a warning: it does not stop the
    follow-up (an answer never fails because another sheet could not be scanned), and says the formulas were not
    checked. A sheet it could not read is said after the drop, when it could matter (_unread_notes)."""
    if found.failed is not None:
        return [f"formulas elsewhere in the workbook that point at the old Q{q} were not checked ({found.failed}): its "
                "old sheets were dropped anyway, so any such formula may show #REF! now"]
    return []


UNREAD = "could not read the formulas on sheet '{sheet}' ({why}): if one there pointed at Q{q}, it shows #REF! now"
UNREAD_REF = ("could not read the formulas on sheet '{sheet}' ({why}), and {n} cell(s) there show #REF! since the "
              "old Q{q} was dropped ({cells}): a formula there pointed at it")
AUTOMATIC = -4105                       # xlCalculationAutomatic
REF_ERROR = -2146826265                 # #REF! as Value2 gives it in both executors: told by its code, not its text


def _record_refs(b, wbn, unread: list, key: str, calculate: bool = False) -> None:
    """Where the cells that show #REF! are on each sheet the look could not read (a protected sheet's values can be
    read when its formulas cannot): the verb errors, bounded as the scan is, recorded in b - after the sheet is
    calculated (calculate: the batch runs with Calculation manual); each sheet operations that may fail."""
    for sheet, _ in unread:
        with b.attempt():
            ws = wbn.Worksheets(sheet)
            if calculate:
                ws.Calculate()
            b.ret(b.verb("errors", ws, REF_ERROR), f"{key}.{sheet}")


def _unread_notes(q: int, unread: list, values: dict | None) -> list[str]:
    """A sheet whose formulas the look could not read, as a warning only when it could matter: a cell there shows
    #REF! after the drop that did not before - compared by where, so a #REF! the follow-up cleared cannot hide a new
    one (phase 3b review: a protected sheet with a hidden formula warned on every follow-up); when that cannot be
    told (values None, or a read that failed) as before."""
    out = []
    for sheet, why in unread:
        before = (values or {}).get(f"refs before.{sheet}")
        after = (values or {}).get(f"refs after.{sheet}")
        if isinstance(before, tuple) and isinstance(after, tuple):
            seen = set(before)                          # once per sheet (final review M3), not per cell
            new = [a for a in after if a not in seen]
            if new:
                more = f" and {len(new) - POINTED_MAX} more" if len(new) > POINTED_MAX else ""
                out.append(UNREAD_REF.format(sheet=sheet, why=why, n=len(new), q=q,
                                             cells=", ".join(new[:POINTED_MAX]) + more))
            continue
        out.append(UNREAD.format(sheet=sheet, why=why, q=q))
    return out


POINTED_MAX = 10                        # formulas named in one warning


def _refused_notes(q: int, failed: list, verb: str, which: str) -> list[str]:
    """The formulas Excel refused to point at the new (or back at the old) Q<q>, as one warning."""
    if not failed:
        return []
    where = [r.where_text for r, _ in failed]
    more = f" and {len(where) - POINTED_MAX} more" if len(where) > POINTED_MAX else ""
    return [f"could not {verb} {len(where)} formula(s) {which} Q{q} ({failed[0][1]}): "
            f"{', '.join(where[:POINTED_MAX])}{more} show #REF!"]


def _changed_notes(q: int, changed: list) -> list[str]:
    """The formulas pointing at Q<q> that were not what the bring-in read any more when the close-out came to write
    them (the user edited them, or inserted a row, a chart or a name in between): left alone, as one warning."""
    if not changed:
        return []
    where = [r.where_text for r in changed]
    more = f" and {len(where) - POINTED_MAX} more" if len(where) > POINTED_MAX else ""
    return [f"{len(where)} formula(s) pointing at Q{q} changed while the answer was drawn: not re-pointed "
            f"({', '.join(where[:POINTED_MAX])}{more}); one that still points at the old Q{q} shows #REF!"]


def _point_back(job) -> None:
    """A follow-up that failed before its old sheets were dropped, once the close-out may have re-pointed the formulas
    elsewhere: they get their old texts back (movein.unpoint) while the parked sheets are still there; what Excel
    refused is a note of the error."""
    from .xl import movein
    if job.repointed is False or job.mentions is None or not job.mentions.plan:
        return
    job.notes += _refused_notes(job.q, movein.unpoint(job.t, job.mentions.plan), "point", "back at the old")


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
    are removed (they are in t.sheets_added; a hidden _svmeta it made too), then its spec row and Sources rows
    (_forget), the sheet the job found active is active again, AutoSave back on in our own file, and the error is
    raised with the job's notes in its text. The first answer of a new workbook (a
    target that resolves to one in a running Excel, when _new_book knows its base and theme) is built as the workbook
    itself and only opened in Excel (target.Target.made); on any failure that workbook is closed without saving and
    the file the job reserved deleted, and the error says so. A workbook the job created is saved inside the job
    (_save), and becomes the session's only once that save went through."""
    from .xl import target as xltarget
    book = _new_book()
    with stage("open target"):
        t = _target(h, target, pre.spec.question[:60], build_new=book is not None)
    job = _Job(t, new_book=book)
    try:
        with stage("draw"):
            out = _file_draw(job, pre)
        with stage("save"):
            _save(job)
    except Exception as e:
        if t.made is not None:
            job.notes.append(xltarget.discard_new(t))
        else:
            try:
                job.not_inserted()
                _rollback(t)
                _forget(job)
            finally:
                job.settle()
        _raise(job, e)
    job.end()
    if t.made is not None:
        from .core import session
        session.set_workbook(t.path)
    remember(pre, out)
    return _with_notes(out, job.notes)


def _raise(job, e: BaseException):
    """Raise the failed job's error (_Job.fail): e itself with its text completed, or the HostError of a busy Excel
    from e."""
    err = job.fail(e)
    try:
        if err is e:
            raise e
        raise err from e
    finally:
        e = err = None                  # the error's traceback holds this frame: no cycle with the job's objects


def redraw_job(h, pre: Prepared, q: int, target: str = "session") -> dict:
    """A follow-up or refresh: the sheets of Q<q> (Answer, Data, every Lines b<i>) and their tables are parked under
    'old ...' names, the new answer is drawn by _file_draw where they were, and the old ones are dropped; the Sources
    log gets its new rows and _svmeta's row for Q<q> is rewritten. A formula elsewhere in the workbook that points at
    them - a cell, a legacy array, a chart series, a defined name, also on another answer's sheet - is re-pointed at
    the new sheets right before the drop (movein.record_mentions / record_repoint); what Excel refused is a warning.
    A drawn range an earlier version stored with the spec is not used: the sheets are replaced whole. A missing
    Answer or Data sheet is an InputError naming it, before anything is renamed. On a failure the re-pointed formulas
    are pointed back first (when the close-out may have re-pointed them), then the new sheets are removed (rollback;
    not when the insert never ran: their names were still the old sheets'), then the old ones get their names back
    (restore; never a rollback after it, as the old names are then the ones in t.sheets_added); RestoreError only
    when the restore failed. A failure once the old sheets are dropped keeps the new answer (a rollback then would
    leave neither): our own file is saved and the error, of its own class, says the new answer is in place (the
    Sources rows may be missing then; a spec the close-out did not store, or may not have, is said too: _kept); so
    does a failing save of our file, which is not tried again.
    After any failure the sheet the job found active is active again and AutoSave back on in our own file; the job's
    notes end the error's text."""
    from .xl import movein
    with stage("open target"):
        t = _target(h, target, pre.spec.question[:60])
    job = _Job(t, q)
    with stage("draw"):
        try:
            out = _file_draw(job, pre)
        except Exception as e:
            failure = e
            try:
                if job.parked is not None and (job.dropped is True
                                               or (job.dropped is None and _dropped(t, job.parked))):
                    _kept(t, e, job=job)
                else:
                    _point_back(job)                    # before the rollback: the new sheets take them along
                    job.not_inserted()
                    _rollback(t)
                    if job.moved:                       # the park may have run: the old names go back
                        try:
                            movein.restore(t, job.parked)
                        except Exception as e2:   # noqa: BLE001 - the draw's error stays the cause
                            failure = RestoreError(f"{e}; the restore was incomplete - {e2}")
            finally:
                job.settle()
            try:
                if failure is e:
                    _raise(job, e)
                raise job.fail(failure) from e
            finally:
                failure = None          # held by this frame, which its traceback holds: a cycle with the job's objects
    with stage("save"):
        try:
            _save(job)
        except Exception as e:
            try:
                _kept(t, e, save=False)
            finally:
                job.settle()
            _raise(job, e)
    job.end()
    out["followup"] = True
    remember(pre, out)
    return _with_notes(out, job.notes)
