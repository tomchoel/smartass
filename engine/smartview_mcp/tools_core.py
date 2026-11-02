"""sv_answer: one question through the new pipeline (design 2026-11-02 §3) - the walking skeleton beside the old
tools, which stay until the benchmarks reach parity."""
from __future__ import annotations

from typing import Literal

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError

from . import config, templates, tools_ws, userexcel
from .core import pipeline, prefs, runlog, scope
from .core import period as periods
from .host import HostError, InputError
from .server import LOCAL, READ, WRITE, _call, _call_user, _err, _is_our_analysis_file, tool
from .timing import Timer, stage


async def _closed_month() -> tuple[str, str | None]:
    """(last closed month 'Sep 26', warning or None). HTTP first (cached); if EPM cannot be reached, the month before
    today, with a warning, so the gate still answers (it pulls nothing)."""
    try:
        return await _call(lambda h: templates.closed_month(h), timeout=30, what="closed month"), None
    except Exception as e:   # noqa: BLE001 - the gate must answer without EPM
        import datetime as dt
        d = dt.date.today().replace(day=1) - dt.timedelta(days=1)
        return periods.label(d.year, d.month), f"closed month not read from EPM ({str(e)[:120]}); assumed last month"


@tool(annotations=READ)
async def sv_gate(question: str) -> dict:
    """NEW PIPELINE step 1 - call FIRST for every finance question, before any pull (procedure:
    knowledge/advisor/gate.md). Runs the fixed scope checklist (period, scenario, entity, measure, currency/sign,
    granularity) against the catalog, the capability map and the user's learned defaults, and suggests one outcome:
    - clear: proceed; state every checklist item with source convention/learned (and each warning) as an assumption.
    - ambiguous: pull nothing; ask the returned questions (max 3) in ONE message, in English; then call sv_gate again
      with the answers folded into the question and offer sv_remember to save an answer as the default.
    - infeasible: say why (reason) in one or two sentences and offer the proxy; build it labelled as a proxy.
    Copy the result into the answer spec's gate (outcome, checklist, questions/reason/proxy) and assumptions.
    Returns {suggested_outcome, checklist{item: {value, source, gap, ask}}, questions, reason, proxy, warnings,
    entities, measures, concepts, closed_month, next}."""
    if not question or not question.strip():
        raise _err(InputError("question is the user's question, as asked"))
    closed, warn = await _closed_month()
    g = scope.gate(question.strip(), closed)
    if warn:
        g["warnings"].append(warn)
    return g


@tool(annotations=LOCAL)
async def sv_remember(item: Literal["measure", "period", "scenario", "entity"], phrase: str, value: str) -> dict:
    """Save the user's answer to a clarifying question as their default, so sv_gate states it next time instead of
    asking. ONLY after the user agreed ("Save 'costs = opex' as your default?"). item + phrase = the gap and the words
    in the question (e.g. measure 'costs', period 'this year'); value = a measure key (sv_gate measures), ytd|fy, a
    catalog scenario key (actual, budget, forecast, ...) or a catalog entity key. Replaces an earlier default for the
    same phrase. Stored in the user's readable preferences.toml (edit or delete freely)."""
    try:
        entry = prefs.remember(item, phrase, value)
    except InputError as e:
        raise _err(e) from None
    return {"saved": entry, "file": str(config.PREFERENCES)}


@tool(annotations=WRITE)
async def sv_answer(spec: dict, target: str = "session", allow_archive: bool = False,
                    ctx: Context | None = None) -> dict:
    """NEW PIPELINE (walking skeleton): answer ONE question from an answer spec v1 (core/contract.py):
    {version: 1, question, gate: {outcome: clear|ambiguous|infeasible, checklist: {period, scenario, entity, measure,
    currency_sign, granularity}, questions (ambiguous, 1-3), reason + proxy (infeasible)}, assumptions: [...],
    pulls: [{id, connection, rows, columns, pov, title?}], layout: {title, columns, sections, checks, ...} (the
    sv_ws_build spec grammar; pull cells use the pull ids)}.
    ambiguous -> returns the questions, pulls nothing. Otherwise: the answer sheet 'Q<n> Answer' appears first with
    the question, assumptions and a status line; each pull is fetched over HTTP and written as a source sheet in one
    block (Smart View ad hoc grid if HTTP fails); then the answer is rendered (live formulas) and checked.
    Returns the answer table, checks, review findings (Critical/Warning/Info), routes and timings."""
    try:
        parsed = pipeline.parse_spec(spec)
    except InputError as e:
        raise _err(e) from None
    early = pipeline.gate_reply(parsed)
    if early is not None:
        runlog.append(runlog.entry(parsed, None, [], None, {}))
        return early

    def job(h):
        tm = Timer.start()
        try:
            r = pipeline.run(h, spec, target, allow_archive)
        except Exception as e:
            runlog.append(runlog.entry(parsed, None, [], None, tm.summary(), error=str(e)[:500]))
            raise
        r.t.created = r.t.created or _is_our_analysis_file(r.t.path)
        with stage("log sources"):
            for s in r.sources:
                try:
                    userexcel.log_source(r.t, r.q, parsed.question, s.pull, notes=s.note or s.route)
                except Exception as e:   # noqa: BLE001 - the source log must never undo a finished answer
                    r.res["warnings"].append(f"could not log {s.pull.sheet} in 'Kilder': {e}")
        e = runlog.entry(parsed, r.q, r.sources, r.review, tm.summary(), workbook=r.t.path)
        try:
            runlog.write_sheet(r.t, e)
        except Exception as ex:   # noqa: BLE001
            r.res["warnings"].append(f"run log sheet not written: {ex}")
        runlog.append(e)
        with stage("save + activate"):
            tools_ws._finish(r.t, r.q, r.prepared, r.res, saved_new=r.t.created)
        out = tools_ws._result(r.t, r.q, r.prepared, r.res, [s.pull for s in r.sources], [])
        return {**out, "gate": parsed.gate.outcome, "routes": {s.pull.sheet: s.route for s in r.sources},
                "review": r.review, "timings": tm.summary()}
    try:
        return await _call_user(job, what="sv_answer")
    except (HostError, ToolError) as e:
        raise _err(e)
