"""Analyses in the user's workbook: sv_analyse, sv_ws_pull, sv_ws_answer, sv_ws_build, sv_concept. Registered on the shared `mcp` object by importing this module (see server.py)."""
from __future__ import annotations

from typing import Literal

from mcp.server.mcpserver.exceptions import ToolError
from pydantic import BaseModel, Field

from . import qlog
from .grid import Axis, to_markdown
from .host import HostError
from .server import WRITE, AxisIn, Connection, _call_user, _doc, _err, _is_our_analysis_file, log, tool

# ---------------------------------------------------------------- analysis in the user's workbook
from . import answer as answer_mod            # noqa: E402
from . import templates, userexcel   # noqa: E402
from .timing import Timer, stage                    # noqa: E402

Target = Literal["active", "new"] | str
_sessions: dict[str, dict[int, dict[str, "userexcel.Pull"]]] = {}   # workbook path -> q -> sheet -> Pull


def _session(path: str, q: int) -> dict:
    return _sessions.setdefault(path.lower(), {}).setdefault(q, {})


def fin(v, kind: str = "heltall") -> str:
    """Financial text for chat, matching the sheet: 1 000 000 / (1 000 000) / – ; % with sign; FTE 1 decimal."""
    if v is None or v == "":
        return ""
    if not isinstance(v, (int, float)):
        return str(v)
    if kind.startswith("prosent"):
        return "–" if abs(v) < 5e-5 else f"{v * 100:+.1f} %".replace(".", ",")
    if abs(v) < 0.5 and kind != "tall1":
        return "–"
    s = f"{abs(v):,.1f}" if kind == "tall1" else f"{abs(v):,.0f}"
    s = s.replace(",", " ").replace(".", ",")
    return f"({s})" if v < 0 else s


def _answer_markdown(spec: dict, res: dict) -> str:
    cols = spec.get("columns", [])
    head = ["Line"] + [c.get("header", "") for c in cols]
    out = ["| " + " | ".join(h.replace("|", "/") for h in head) + " |", "|---|" + "---:|" * len(cols)]
    n = 0
    for sec in spec.get("sections", []):
        out.append(f"| **{sec.get('title', '')}** |" + " |" * len(cols))
        for ln in sec.get("lines", []):
            k = ln.get("key")
            if not k or k not in res["values"]:
                continue
            vals = res["values"][k]
            cells = []
            for j, v in enumerate(vals):
                kind = ln.get("num") or cols[j].get("kind") or "heltall"
                if j < len(ln.get("cells", [])) and isinstance(ln["cells"][j], dict) and "varpct" in ln["cells"][j]:
                    kind = "prosent1"
                cells.append(fin(v, kind))
            label = ln.get("label", k)
            if ln.get("kind") in ("sum", "grand"):
                label = f"**{label}**"
            out.append("| " + label + " | " + " | ".join(cells) + " |")
            n += 1
            if n >= 45:
                out.append("| … see the answer sheet for the rest |" + " |" * len(cols))
                return "\n".join(out)
    return "\n".join(out)


def _result(t, q: int, spec: dict, res: dict, pulls: list, warnings: list) -> dict:
    try:
        if userexcel.com(lambda: t.wb.ReadOnly, limit=10):
            warnings = [f"'{t.name}' is open READ-ONLY (someone or something else holds the file) - the analysis "
                        "is built but cannot be saved under this name."] + warnings
    except Exception:   # noqa: BLE001 - informational only
        pass
    return {
        "workbook": t.path, "question": f"Q{q}", "answer_sheet": res["sheet"],
        "data_sheets": [p.sheet for p in pulls] + ["Kilder"],
        "key_points": spec.get("key_points", []),
        "table": _answer_markdown(spec, res),
        "checks": res["checks"],
        "warnings": warnings + res["warnings"] + [f"formula error on line '{k}'" for k in res["formula_errors"]],
        "style_lint": res["lint"],
        # tn-formelstil: 'block' findings must be fixed (spec_patch) before the answer is presented as done
        "formula_style": res.get("formula_style"),
        "note": "Your workbook was NOT saved - review and save it yourself." if not t.created
        else f"New analysis workbook saved: {t.path}",
    }


def _tidy(t, sheet: str | None = None) -> None:
    """Fixed, neat sheet order (sheetorder: Svar first, other sheets, data per question, Kilder, hidden) and land
    on `sheet`. Our own analysis files get the full order; in a user's own workbook only our sheets are reordered
    among themselves. Cosmetic: never fails the tool."""
    from . import sheetorder
    try:
        with stage("arrange sheets"):
            mode = "full" if (t.created or _is_our_analysis_file(t.path)) else "active"
            sheetorder.arrange_sheets(t.wb, mode, activate=sheet)
    except Exception as e:   # noqa: BLE001
        log.warning("arrange sheets failed in %s: %s", t.name, e)
        if sheet:
            try:
                userexcel.com(t.wb.Worksheets(sheet).Activate)
            except Exception:   # noqa: BLE001
                pass


def _finish(t, q: int, spec: dict, res: dict, saved_new: bool):
    _tidy(t, res["sheet"])
    t.status(None)
    if saved_new:
        userexcel.com(t.wb.Save)


@tool(annotations=WRITE)
async def sv_analyse(question: str, template: Literal["division_pnl", "cost_base", "monthly_trend", "actual_vs_plan"],
                     params: dict, target: str = "session", allow_archive: bool = False, q: int | None = None) -> dict:
    """Answer a financial question as an analysis IN THE USER'S EXCEL: live Smart View data sheets ('Q<n> FinMax' ...),
    a formatted answer sheet 'Q<n> Svar' (TN report style, live formulas, key points, checks, variance drivers, chart)
    and a 'Kilder' log. target: 'session' (default: this session's own analysis workbook in
    the analyser folder (config.ANALYSER_DIR, set in config.toml)\\<project>, created on its first question and
    reused/reopened for every later one),
    'active' (the workbook the user is working in - only when they ask for it), 'new' (a new file that becomes the
    session's workbook), or a path. Archive files are refused unless allow_archive=true. q: extend an existing
    question number (follow-up), else the next free number. Templates and params:
      division_pnl:   entity (e.g. 'un.SY.2TEC'), period ('Sep 26'; default last closed), basis 'ytd'|'month',
                      compare 'ly'|None, drivers bool
      cost_base:      parent ('un.SY.2TEC'), accounts (['ac.C50S','ac.C60','ac.C70','ac.STA20Y']), period, top (25)
      monthly_trend:  entity, accounts, through ('Sep 26')
      actual_vs_plan: entity (FinPlan alias, e.g. '2TEC Technology'), month ('Sep'), year ('FY26'), basis, forecast
    Write the chat reply from the returned key_points/table/checks."""
    fn = templates.TEMPLATES[template]

    def job(h):
        tm = Timer.start()
        with stage("plan (template)"):
            plan = fn(h, **params)
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
            raise HostError(f"{e} (the analysis was not completed; removed the partial sheets: "
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
async def sv_ws_pull(question: str, connection: Connection, rows: list[AxisIn], columns: list[AxisIn],
                     pov: dict[str, str] | None = None, target: str = "session", allow_archive: bool = False,
                     q: int | None = None, sheet_title: str | None = None) -> dict:
    """Free-form analysis, step 1 (use sv_analyse when a template fits): put a LIVE Smart View grid (financial NOK
    format) into the user's workbook as 'Q<n> <cube>' and log it in 'Kilder'.
    For several pulls of ONE question: call again with target=<returned 'workbook'> and q=<returned 'q'> (otherwise
    target='new' would create a new file and q would advance). Then call sv_ws_answer(q, spec, target=<workbook>).
    Returns workbook, q, sheet, a data preview and the exact row_keys/col_keys to use in the answer spec."""
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
            raise HostError(f"{e} (removed the partial sheet: {', '.join(removed) or 'none'})") from None
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
                           {expr: 'rev+cogs'} | {sum: [keys] | 'm*'} | {ref: {key, col}} | {rowsum: [cols]} |
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


def _load_session(t, q: int) -> dict:
    sess = _sessions.get(t.path.lower(), {}).get(q)
    if not sess:                                       # e.g. after a server restart: rebuild from the workbook
        with stage("reload pulls from workbook"):
            sess = userexcel.load_pulls(t, q)
        if sess:
            _sessions.setdefault(t.path.lower(), {})[q] = sess
    return sess or {}


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
                       allow_archive: bool = False) -> dict:
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
            raise HostError(f"No pulls for Q{q} in '{t.name}'. Run sv_ws_pull first with the same target.")
        stored = userexcel.load_spec(t, q)
        ids = stored[1] if stored else {}
        if spec is None:
            if not stored:
                raise HostError(f"No earlier answer spec for Q{q}; pass spec.")
            raw = specs.merge(stored[0], spec_patch or {})
        else:
            raw = specs.merge(spec, spec_patch or {})
        return _answer_job(t, q, raw, ids, sess, tm)
    try:
        return await _call_user(job, what="sv_ws_answer")
    except (HostError, ToolError) as e:
        raise _err(e)


def _build_job(h, tm, question: str, pulls: list, spec: dict, target: str, allow_archive: bool, q: int | None,
               replace: bool = False) -> dict:
    """All pulls as live Smart View sheets, then the answer (shared by sv_ws_build and sv_concept).
    replace=True (a whole question re-run, e.g. sv_concept with q): the question's old data sheets are parked first
    and deleted only after the new build succeeded (restored if it fails)."""
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
        out = _answer_job(t, qn, spec, ids, sess, tm, warns)
    except Exception as e:
        removed = userexcel.rollback(t)
        for name in removed:
            sess.pop(name, None)
        if parked:
            userexcel.restore_parked(t, parked)
            _sessions.get(t.path.lower(), {}).pop(qn, None)      # reload the restored sheets from the workbook
        raise HostError(f"{e} (the analysis was not completed; removed the partial sheets: "
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


@tool(annotations=WRITE)
@_doc(SPEC_DOC)
async def sv_ws_build(question: str, pulls: list[PullIn], spec: dict, target: str = "session",
                      allow_archive: bool = False, q: int | None = None) -> dict:
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
                     target: str = "session", allow_archive: bool = False, q: int | None = None) -> dict:
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
