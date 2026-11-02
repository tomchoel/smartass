"""Analyses in the user's workbook: sv_analyse, sv_ws_pull, sv_ws_answer, sv_ws_build, sv_concept. Registered on the shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

from typing import Literal

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError
from pydantic import BaseModel, Field, field_validator

from . import qlog
from .grid import Axis, to_markdown
from .host import HostError, InputError, rewrap
from .server import WRITE, AxesIn, AxisIn, Connection, axes_in, _call_user, _doc, _err, _is_our_analysis_file, tool

# ---------------------------------------------------------------- analysis in the user's workbook
from . import answer as answer_mod            # noqa: E402
from . import ops, templates, userexcel   # noqa: E402
from .timing import Timer, stage                    # noqa: E402

from .core.reply import _sessions, fin  # noqa: E402,F401
from .core.reply import answer_markdown as _answer_markdown  # noqa: E402,F401
from .core.reply import finish as _finish  # noqa: E402
from .core.reply import load_session as _load_session  # noqa: E402
from .core.reply import result as _result  # noqa: E402
from .core.reply import session as _session  # noqa: E402
from .core.reply import tidy as _tidy  # noqa: E402

Target = Literal["active", "new"] | str


@tool(annotations=WRITE)
async def sv_analyse(question: str, template: Literal["cost_base", "monthly_trend", "actual_vs_plan"],
                     params: dict, target: str = "session", allow_archive: bool = False, q: int | None = None,
                     ctx: Context | None = None) -> dict:
    """Answer a financial question as an analysis IN THE USER'S EXCEL: live Smart View data sheets ('Q<n> FinMax' ...),
    a formatted answer sheet 'Q<n> Svar' (TN report style, live formulas, key points, checks, variance drivers, chart)
    and a 'Kilder' log. target: 'session' (default: this session's own analysis workbook in
    the analyser folder (config.ANALYSER_DIR, set in config.toml)\\<project>, created on its first question and
    reused/reopened for every later one),
    'active' (the workbook the user is working in - only when they ask for it), 'new' (a new file that becomes the
    session's workbook), or a path. Archive files are refused unless allow_archive=true. q: extend an existing
    question number (follow-up), else the next free number. Templates and params:
      cost_base:      parent ('un.SY.2TEC'), accounts (['ac.C50S','ac.C60','ac.C70','ac.STA20Y']), period, top (25)
      monthly_trend:  entity, accounts, through ('Sep 26')
      actual_vs_plan: entity (FinPlan alias, e.g. '2TEC Technology'), month ('Sep'), year ('FY26'), basis, forecast,
                      finmax_entity (FinMax entity for the EBITDA reconciliation; default: mapped via the atlas,
                      'BANORWAY' -> 'un.SY'; no match -> no reconciliation, with a note)
    Write the chat reply from the returned key_points/table/checks."""
    fn = templates.TEMPLATES[template]

    def job(h):
        tm = Timer.start()
        with stage("plan (template)"):
            plan = fn(h, **params)
        with stage("check members (atlas)"):
            templates.check_pulls(plan.pulls, confirm=ops.live_confirm(h))
        with stage("open target"):
            t = userexcel.resolve_target(h, target, allow_archive, title=plan.title)
            qn = q or userexcel.next_question(t)
        pulls_by_id, sess = {}, _session(t.path, qn)
        try:
            for ps in plan.pulls:
                with stage("pull"):
                    p = userexcel.pull(h, t, qn, ps.connection, ps.pov, ps.rows, ps.cols,
                                       title=f"Q{qn} {ps.connection}" + (f" {ps.title_suffix}" if ps.title_suffix else ""))
                pulls_by_id[ps.id] = p
                sess[p.sheet] = p
            spec = plan.answer(pulls_by_id)
            with stage("answer"):
                res = answer_mod.build(t, qn, spec, sess)
        except Exception as e:
            removed = userexcel.rollback(t)
            for name in removed:
                sess.pop(name, None)
            raise rewrap(e, f"{e} (the analysis was not completed; removed the partial sheets: "
                         f"{', '.join(removed) or 'none'})") from None
        notes = "; ".join(spec.get("notes", [])[:1])
        with stage("log Kilder"):
            for p in pulls_by_id.values():
                try:
                    userexcel.log_source(t, qn, question, p, notes)
                except Exception as e:   # noqa: BLE001 - the source log must never undo a finished analysis
                    res["warnings"].append(f"could not log {p.sheet} in 'Kilder': {e}")
        warns = [w for p in pulls_by_id.values() for w in p.parsed.pov_problems] + \
                [f"unknown members in {p.sheet}: {', '.join(p.parsed.unknown_members)}"
                 for p in pulls_by_id.values() if p.parsed.unknown_members]
        t.created = t.created or _is_our_analysis_file(t.path)     # an analysis file of ours: save it too
        with stage("save + activate"):
            _finish(t, qn, spec, res, saved_new=t.created)
        return {**_result(t, qn, spec, res, list(pulls_by_id.values()), warns), "timings": tm.summary()}
    try:
        out = await _call_user(job, what=f"sv_analyse {template}")
    except (HostError, ToolError) as e:
        qlog.write({"tool": "sv_analyse", "template": template, "params": params, "error": str(e)})
        raise _err(e)
    qlog.write({"tool": "sv_analyse", "template": template, "params": params, "question": question,
                "workbook": out["workbook"], "q": out["question"], "checks": out["checks"]})
    return out


@tool(annotations=WRITE)
async def sv_ws_pull(question: str, connection: Connection, rows: AxesIn, columns: AxesIn,
                     pov: dict[str, str] | None = None, target: str = "session", allow_archive: bool = False,
                     q: int | None = None, sheet_title: str | None = None, ctx: Context | None = None) -> dict:
    """Free-form analysis, step 1 (use sv_analyse when a template fits): put a LIVE Smart View grid (financial NOK
    format) into the user's workbook as 'Q<n> <cube>' and log it in 'Kilder'.
    For several pulls of ONE question: call again with target=<returned 'workbook'> and q=<returned 'q'> (otherwise
    target='new' would create a new file and q would advance). Then call sv_ws_answer(q, spec, target=<workbook>).
    Returns workbook, q, sheet, a data preview and the exact row_keys/col_keys to use in the answer spec."""
    rows, columns = axes_in(rows), axes_in(columns)
    rows_ax = [Axis(a.dimension, a.members) for a in rows]
    cols_ax = [Axis(a.dimension, a.members) for a in columns]

    def job(h):
        tm = Timer.start()
        with stage("open target"):
            t = userexcel.resolve_target(h, target, allow_archive, title=question[:50])
            qn = q or userexcel.next_question(t)
        try:
            with stage("pull"):
                p = userexcel.pull(h, t, qn, connection, pov or {}, rows_ax, cols_ax,
                                   title=sheet_title or f"Q{qn} {connection}")
        except Exception as e:
            removed = userexcel.rollback(t)
            raise rewrap(e, f"{e} (removed the partial sheet: {', '.join(removed) or 'none'})") from None
        _session(t.path, qn)[p.sheet] = p
        with stage("log Kilder"):
            try:
                userexcel.log_source(t, qn, question, p)
            except Exception:   # noqa: BLE001 - the source log must never undo a finished pull
                pass
        _tidy(t, p.sheet)
        t.status(None)
        if t.created:
            with stage("save"):
                userexcel.com(t.wb.Save)
        return {"timings": tm.summary(), "workbook": t.path, "q": qn, "sheet": p.sheet,
                "table": to_markdown(p.parsed.columns, p.parsed.rows, p.parsed.row_labels, limit=40),
                "row_keys": [list(k) for k in p.layout.row_tuples][:200],
                "col_keys": [list(k) for k in p.layout.col_tuples],
                "warnings": p.parsed.pov_problems + ([f"unknown members: {', '.join(p.parsed.unknown_members)}"]
                                                     if p.parsed.unknown_members else []),
                "retrieve_seconds": round(p.seconds, 2)}
    try:
        return await _call_user(job, what="sv_ws_pull")
    except (HostError, ToolError) as e:
        raise _err(e)


SPEC_DOC = """Spec format (all numbers become live formulas):
      {title, subtitle ('NOK · FinMax · Actual · ...'), key_points: [your interpretation],
       auto_points: true (default: factual points drafted from the numbers come first) | false,
       columns: [{header, kind: heltall|tall1|prosent1}],
       sections: [{title, lines: [{key, label, kind: line|sum|grand, num?: tall1,
                   cells: [{pull: '<pull id or sheet>', row: [...], col: [...]} | {var: [i, j]} | {varpct: [i, j]} |
                           {expr: 'rev+cogs'} | {sum: [keys] | 'm*'} | {avg: [keys] | 'f*'} | {ref: {key, col}} | {rowsum: [cols]} |
                           {value} | null]}],
                   repeat: {over: ['Jan', ...] | 'months:Jan..Dec' | [{m: 'Jan', ...}, ...],
                            line: {key: 'm{i}', label: '{v}', cells: [{pull: 'a', row: [...], col: ['{v} 26']}, ...]}}}],
       checks: [{label, expr, columns: [..]} | {label, static: 'OK'|'DIFF', detail}], notes: [..], source,
       chart: {type: column|bar|line|waterfall, lines: [keys] | 'm*', column} | {type, row: key, columns: [..]} |
              {type, lines: [keys] | 'm*', columns: [i, j]} (one series per column, e.g. this year vs last year)}
    In a repeat, a pull cell whose member is not in the pull (e.g. 'Oct 26' before it closes) is left empty, and so
    are variances that depend on it. Column indexes in var/varpct/ref/rowsum are 0-based ANSWER columns.
    Formatting: full NOK #,##0, negatives in parentheses; FinMax signs kept; variance = current − comparison."""
SPEC_DOC += """
    Drill lines (after sv_drill): {drill: 'd', where: {COLUMN: 'exact value', ...}, agg: sum|count|debit|credit}
    sums the question's Linjer table (SUMIFS; an exact-match FILTER where Excel would coerce the key). where {} =
    all lines. Reconcile against the drill's live grid: {pull: '<cube_pull sheet>', row: [...], col: [...]}.
    Mixed signs (profile.signs_mixed): no share-of-total column; show debit and credit instead."""


def _answer_job(t, q: int, raw_spec: dict, ids: dict, sess: dict, tm, extra_warnings: list | None = None) -> dict:
    from . import specs
    from .tools_drill import _drill_rows, _has_drill_cells
    with stage("prepare spec"):
        spec = specs.prepare(raw_spec, ids, sess,
                             drills=_drill_rows(t.path, q) if _has_drill_cells(raw_spec) else None)
    # lines Smart View rejected (rc=1000) are missing from the answer: always say so, in the result and on the sheet
    rejected = list(dict.fromkeys(w for p in sess.values() for w in p.parsed.pov_problems
                                  if w.startswith("Smart View rejected")))
    extra_warnings = list(dict.fromkeys((extra_warnings or []) + rejected))
    if rejected:
        spec = {**spec, "notes": list(spec.get("notes") or []) + [
            "Utelatt fordi Smart View avviste medlemmet: " + w.split(": ", 1)[-1].split(". Pick")[0]
            for w in rejected]}
    with stage("answer"):
        res = answer_mod.build(t, q, spec, sess)
    try:
        userexcel.save_spec(t, q, raw_spec, ids)
    except Exception:   # noqa: BLE001 - only needed for patch follow-ups
        pass
    t.created = t.created or _is_our_analysis_file(t.path)     # made by an earlier pull of ours: ours to save
    with stage("save + activate"):
        _finish(t, q, spec, res, saved_new=t.created)
    return {**_result(t, q, spec, res, list(sess.values()), extra_warnings), "timings": tm.summary()}


@tool(annotations=WRITE)
@_doc(SPEC_DOC)
async def sv_ws_answer(q: int, spec: dict | None = None, spec_patch: dict | None = None, target: str = "session",
                       allow_archive: bool = False, ctx: Context | None = None) -> dict:
    """Free-form analysis, step 2 (sv_ws_build does pulls + answer in ONE call): build or rebuild 'Q<n> Svar' from a
    spec whose cells reference the pulls of question q (same target as the pulls, ideally their 'workbook').
    Follow-up: pass spec_patch instead of spec - it is deep-merged into the question's last spec (dicts merge,
    lists/values replace, null deletes), e.g. {"notes": [...]} or {"key_points": [...]}."""
    from . import specs

    def job(h):
        tm = Timer.start()
        with stage("open target"):
            t = userexcel.resolve_target(h, target, allow_archive)
        sess = _load_session(t, q)
        if not sess:
            raise InputError(f"No pulls for Q{q} in '{t.name}'. Run sv_ws_pull first with the same target.")
        stored = userexcel.load_spec(t, q)
        if stored and stored[0].get("version") == 1:
            raise InputError(f"Q{q} was answered by sv_answer (answer spec v1); sv_ws_answer cannot patch it. Call "
                             f"sv_answer again with the changed spec (follow-ups on v1 answers come with the new "
                             f"pipeline).")
        ids = stored[1] if stored else {}
        if spec is None:
            if not stored:
                raise InputError(f"No earlier answer spec for Q{q}; pass spec.")
            raw = specs.merge(stored[0], spec_patch or {})
        else:
            raw = specs.merge(spec, spec_patch or {})
        return _answer_job(t, q, raw, ids, sess, tm)
    try:
        return await _call_user(job, what="sv_ws_answer")
    except (HostError, ToolError) as e:
        raise _err(e)


def _build_job(h, tm, question: str, pulls: list, spec: dict, target: str, allow_archive: bool, q: int | None,
               replace: bool = False, before_answer=None) -> dict:
    """All pulls as live Smart View sheets, then the answer (shared by sv_ws_build, sv_concept and fin_pnl).
    replace=True (a whole question re-run, e.g. sv_concept with q): the question's old data sheets are parked first
    and deleted only after the new build succeeded (restored if it fails).
    before_answer(spec, {pull id: Pull}) -> spec runs after the pulls and before the answer sheet (fin_pnl adds its
    warehouse = source sheet checks there)."""
    with stage("open target"):
        t = userexcel.resolve_target(h, target, allow_archive, title=question[:50])
        qn = q or userexcel.next_question(t)
    parked = {}
    if replace and q:
        with stage("park old data sheets"):
            parked = userexcel.park_question_data(t, qn)
        _sessions.get(t.path.lower(), {}).pop(qn, None)
    sess, ids, warns = _session(t.path, qn), {}, []
    try:
        for i, pi in enumerate(pulls):
            default = f"Q{qn} {pi.connection}" + (f" {pi.id}" if i else "")
            with stage("pull"):
                p = userexcel.pull(h, t, qn, pi.connection, pi.pov or {},
                                   [Axis(a.dimension, a.members) for a in pi.rows],
                                   [Axis(a.dimension, a.members) for a in pi.columns], title=pi.title or default)
            sess[p.sheet] = p
            ids[pi.id] = p.sheet
            warns += p.parsed.pov_problems + ([f"unknown members in {p.sheet}: {', '.join(p.parsed.unknown_members)}"]
                                              if p.parsed.unknown_members else [])
        if before_answer is not None:
            spec = before_answer(spec, {pi.id: sess[ids[pi.id]] for pi in pulls})
            if not isinstance(spec, dict):
                raise InputError("before_answer must return the spec")
        out = _answer_job(t, qn, spec, ids, sess, tm, warns)
    except Exception as e:
        removed = userexcel.rollback(t)
        for name in removed:
            sess.pop(name, None)
        if parked:
            userexcel.restore_parked(t, parked)
            _sessions.get(t.path.lower(), {}).pop(qn, None)      # reload the restored sheets from the workbook
        raise rewrap(e, f"{e} (the analysis was not completed; removed the partial sheets: "
                     f"{', '.join(removed) or 'none'}" + ("; the previous version was kept" if parked else "")
                     + ")") from None
    if parked:
        with stage("drop old data sheets"):
            userexcel.drop_parked(t, parked)
    with stage("log Kilder"):
        for pi in pulls:
            try:
                userexcel.log_source(t, qn, question, sess[ids[pi.id]])
            except Exception as e:   # noqa: BLE001 - the source log must never undo a finished analysis
                out["warnings"].append(f"could not log {ids[pi.id]} in 'Kilder': {e}")
    _tidy(t, out["answer_sheet"])                     # again: the Kilder log may have added a sheet
    if t.created:
        userexcel.com(t.wb.Save)
    out["timings"] = tm.summary()
    return out


class PullIn(BaseModel):
    id: str = Field(description="Short id used in the spec's pull cells, e.g. 'm' (monthly) or 'ytd'")
    connection: Connection
    rows: list[AxisIn]
    columns: list[AxisIn]
    pov: dict[str, str] | None = None
    title: str | None = Field(default=None, description="Sheet name; default 'Q<n> <cube>' (+ ' <id>' if needed)")

    @field_validator("rows", "columns", mode="before")
    @classmethod
    def _dict_axes(cls, v):
        return [a.model_dump() for a in axes_in(v)] if isinstance(v, dict) else v


@tool(annotations=WRITE)
@_doc(SPEC_DOC)
async def sv_ws_build(question: str, pulls: list[PullIn], spec: dict, target: str = "session",
                      allow_archive: bool = False, q: int | None = None, ctx: Context | None = None) -> dict:
    """ONE call for a free-form analysis: every pull as a live Smart View sheet, then 'Q<n> Svar' from the spec, in
    the session's workbook (target 'session' default: one file per session, Q1, Q2 ... | 'active' = the user's open
    workbook, only when asked | 'new' | path; archive refused unless allow_archive). Pull cells in the
    spec use the pull ids. Follow-ups: sv_ws_answer(q, spec_patch=..., target=<workbook>); new data for the same
    question: sv_ws_build(..., q=<q>, target=<workbook>). Returns key points, table, checks, sheets and timings."""

    def job(h):
        return _build_job(h, Timer.start(), question, pulls, spec, target, allow_archive, q)
    try:
        out = await _call_user(job, what="sv_ws_build")
    except (HostError, ToolError) as e:
        qlog.write({"tool": "sv_ws_build", "question": question, "error": str(e)})
        raise _err(e)
    qlog.write({"tool": "sv_ws_build", "question": question, "workbook": out["workbook"], "q": out["question"],
                "checks": out["checks"], "timings": out["timings"]})
    return out


@tool(annotations=WRITE)
async def sv_concept(concept: str, period: str | None = None, compare: str | None = None, entity: str | None = None,
                     target: str = "session", allow_archive: bool = False, q: int | None = None,
                     ctx: Context | None = None) -> dict:
    """ONE call for a business-concept analysis (e.g. 'splitt' = Splitt profitability): builds every block of the
    concept from knowledge\\concepts.json - booked lines, estimated lines with their keys, attributed lines and what is
    not in FinMax - as live Smart View sheets + 'Q<n> Svar' in the session's workbook, YTD `period` (default: last
    closed month) vs `compare` (default: same month last year). Use sv_scope first to find the concept."""
    from . import concepts as cmod

    def job(h):
        tm = Timer.start()
        c = cmod.get(concept)
        with stage("plan (concept)"):
            p = period or templates.closed_month(h, "FinMax")
            mo, yy = templates.split_period(p)
            ly = compare or f"{templates.MONTHS[mo - 1]} {yy - 1:02d}"
            pulls, spec = cmod.plan(c, p, ly, entity)
        pull_objs = [PullIn.model_validate({k: v for k, v in pl.items() if v is not None}) for pl in pulls]
        return _build_job(h, tm, f"{c.get('title', concept)} {p}", pull_objs, spec, target, allow_archive, q,
                          replace=True)
    try:
        out = await _call_user(job, what=f"sv_concept {concept}")
    except (HostError, ToolError, ValueError) as e:
        qlog.write({"tool": "sv_concept", "concept": concept, "error": str(e)})
        raise _err(e)
    qlog.write({"tool": "sv_concept", "concept": concept, "workbook": out["workbook"], "q": out["question"],
                "checks": out["checks"], "timings": out["timings"]})
    return out
