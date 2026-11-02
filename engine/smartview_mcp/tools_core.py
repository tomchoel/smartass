"""sv_answer: one question through the new pipeline (design 2026-11-02 §3) - the walking skeleton beside the old
tools, which stay until the benchmarks reach parity."""
from __future__ import annotations

import contextlib
import datetime as dt
from typing import Literal

from mcp.server.mcpserver import Context
from mcp.server.mcpserver.exceptions import ToolError
from pydantic import ValidationError

from . import catalog, config, templates, tools_ws, userexcel
from .core import contract, pipeline, prefs, recipes, runlog, scope
from .core import followup as followup_mod
from .core import period as periods
from .host import HostError, InputError
from .server import LOCAL, READ, WRITE, _call_user, _err, _is_our_analysis_file, tool
from .timing import Timer, stage


class _NoExcel:
    """A host stand-in for templates.closed_month: the gate never starts Excel, so the COM fallback fails fast."""
    def on_host(self, fn):
        raise HostError("Smart View HTTP route unavailable (the gate does not start Excel)")


async def _closed_month() -> tuple[str, str | None]:
    """(last closed month 'Sep 26', warning or None): the cached or HTTP answer (~0.2 s), never Excel. If EPM cannot
    be reached, the month before today, with a warning, so the gate still answers (it pulls nothing)."""
    import asyncio
    try:
        return await asyncio.wait_for(asyncio.to_thread(templates.closed_month, _NoExcel()), 30), None
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


def _followup_spec(t, patch: dict, q: int | None) -> tuple[int, dict, dict]:
    """(q, stored answer, patched spec) for a follow-up on answer q (default: the current answer)."""
    q = q or followup_mod.current_question(t)
    if q is None:
        raise InputError("There is no answer to follow up in this workbook: ask it as a new question with spec=...")
    old = userexcel.load_answer(t, q)
    if old is None or old["spec"].get("version") != 1:
        raise InputError(f"Q{q} Answer was not drawn by sv_answer, so it has no answer spec to patch: ask it as a "
                         "new question with spec=..., or use sv_ws_answer for answers drawn by sv_ws_build.")
    try:
        raw = contract.patched(old["spec"], patch)
    except ValueError as e:          # pydantic ValidationError is a ValueError
        raise InputError(f"The follow-up does not give a valid answer spec: {e}") from None
    return q, old, raw


def _check_recipe_spec(spec: dict) -> None:
    """The early check of a recipe spec (no pulls/layout yet): the gate and the recipe field."""
    try:
        contract.Gate.model_validate(spec.get("gate") or {})
        contract.Recipe.model_validate(spec["recipe"])
    except ValidationError as e:
        raise InputError("This is not a valid answer spec v1: " + "; ".join(
            f"{'.'.join(map(str, x['loc'])) or 'spec'}: {x['msg']}" for x in e.errors()[:8])) from None


def _pnl_facts(h, area: str, period: str | None, compare: tuple[str, ...]):
    """(res, mode, fills, notes, entity) for recipe pnl: the warehouse facts, exactly as fin_pnl gets them."""
    from . import atlas, tools_fin   # tools_fin is a tool module too: import it when the recipe runs
    cat = catalog.load()
    with contextlib.closing(atlas.connect()) as adb:
        entity = tools_fin.resolve_scope(cat, area, adb)
    res, mode, fills, notes = tools_fin.prepare(h, cat, entity, period, compare)
    return res, mode, fills, notes, entity


def _expand_recipe(h, spec: dict) -> tuple[dict, dict]:
    """(the full spec, extra result keys) for a recipe spec; runs inside the Excel job (closed month, warehouse)."""
    ctx = recipes.RecipeContext(closed=templates.closed_month(h), cat=catalog.load(), label=templates.member_label,
                                pnl_facts=lambda area, period, compare: _pnl_facts(h, area, period, compare),
                                now=dt.datetime.now().strftime("%Y-%m-%d %H:%M"))
    try:
        return recipes.expand(spec, ctx)
    except InputError:
        raise
    except ValueError as e:          # pydantic ValidationError is a ValueError
        raise InputError(f"The recipe does not give a valid answer spec: {e}") from None


@tool(annotations=WRITE)
async def sv_answer(spec: dict | None = None, followup: dict | None = None, q: int | None = None,
                    target: str = "session", allow_archive: bool = False, ctx: Context | None = None) -> dict:
    """NEW PIPELINE: answer ONE question from an answer spec v1 (core/contract.py) - or change the current answer
    with a follow-up. Give exactly one of spec and followup.
    spec (a new question -> a new 'Q<n> Answer'): {version: 1, question, gate: {outcome: clear|ambiguous|infeasible,
    checklist: {period, scenario, entity, measure, currency_sign, granularity}, questions (ambiguous, 1-3), reason +
    proxy (infeasible)}, assumptions: [...], pulls: [{id, connection, rows, columns, pov, title?}], layout: {title,
    columns, sections, checks, ...} (the sv_ws_build spec grammar; pull cells use the pull ids)}.
    ambiguous -> returns the questions, pulls nothing. Otherwise: the answer sheet 'Q<n> Answer' appears first with
    the question, assumptions and a status line; each pull is fetched over HTTP and written as a source sheet in one
    block (Smart View ad hoc grid if HTTP fails); then the answer is rendered (live formulas) and checked.
    A recipe drafts pulls + layout server-side: spec = {version: 1, question, gate, assumptions?, recipe: {name,
    params, overrides?}} without pulls/layout. Recipes: pnl {area, period?, compare?}; actual_vs_plan {entity
    (FinPlan), month?, year?, basis?, forecast?, finmax_entity?}; monthly_trend {entity?, accounts?, through?};
    cost_base {parent?, accounts?, period?, top?}; concept {concept, period?, compare?, entity?}. The period
    defaults to the closed month (returned as period). overrides = a follow-up-shaped patch on the draft; recipe
    params cannot be changed by a follow-up (ask a new question).
    followup: a patch on the current answer (the active Q<n> Answer, else the latest; or answer q). Use it unless
    the user asks a new question. Dicts merge, null deletes; pulls merge by id, layout.sections by title, lines by
    key, checks by label ({..., "drop": true} removes; new ones are appended); {"append": [...]} adds to a list
    (e.g. notes); any other list replaces. Unchanged pulls are not fetched again; the answer is redrawn in its own
    block, so the user's cells and charts outside it survive.
    Returns the answer table, checks, review findings (Critical/Warning/Info: knowledge/advisor/findings.md),
    material lines that need a driver in commentary.drivers (knowledge/advisor/variance-commentary.md), routes,
    timings and an outline (pull ids, columns, section titles, line keys, check labels) to address follow-ups.
    Call sv_gate first for a new question (knowledge/advisor/gate.md)."""
    if (spec is None) == (followup is None):
        raise _err(InputError("Give exactly one of spec or followup: spec = new question, followup = change the "
                              "current answer."))
    parsed = None
    recipe = spec is not None and bool(spec.get("recipe")) and not spec.get("layout")
    if spec is not None:
        try:
            if recipe and (spec.get("gate") or {}).get("outcome") != "ambiguous":
                _check_recipe_spec(spec)            # pulls + layout come from the recipe, inside the Excel job
            else:
                parsed = pipeline.parse_spec(spec)
        except InputError as e:
            raise _err(e) from None
        early = pipeline.gate_reply(parsed) if parsed is not None else None
        if early is not None:
            runlog.append(runlog.entry(parsed, None, [], None, {}))
            return early

    def job(h):
        tm = Timer.start()
        p = parsed
        extra = {}
        try:
            if followup is None and recipe:
                raw, extra = _expand_recipe(h, spec)
                p = pipeline.parse_spec(raw)
                r = pipeline.run(h, raw, target, allow_archive)
            elif followup is None:
                r = pipeline.run(h, spec, target, allow_archive)
            else:
                with stage("open target"):
                    t = userexcel.resolve_target(h, target, allow_archive)
                fq, old, raw = _followup_spec(t, followup, q)
                p = pipeline.parse_spec(raw)
                early = pipeline.gate_reply(p)
                if early is not None:
                    return early
                r = pipeline.run(h, raw, target, allow_archive, q=fq, old=old, t=t)
        except Exception as e:
            if p is not None:
                runlog.append(runlog.entry(p, None, [], None, tm.summary(), error=str(e)[:500]))
            raise
        r.t.created = r.t.created or _is_our_analysis_file(r.t.path)
        with stage("log sources"):
            for s in r.sources:
                if s.route == "reused":
                    continue
                try:
                    userexcel.log_source(r.t, r.q, p.question, s.pull, notes=s.note or s.route)
                except Exception as e:   # noqa: BLE001 - the source log must never undo a finished answer
                    r.res["warnings"].append(f"could not log {s.pull.sheet} in 'Kilder': {e}")
        e = runlog.entry(p, r.q, r.sources, r.review, tm.summary(), workbook=r.t.path)
        try:
            runlog.write_sheet(r.t, e)
        except Exception as ex:   # noqa: BLE001
            r.res["warnings"].append(f"run log sheet not written: {ex}")
        runlog.append(e)
        with stage("save + activate"):
            tools_ws._finish(r.t, r.q, r.prepared, r.res, saved_new=r.t.created)
        out = tools_ws._result(r.t, r.q, r.prepared, r.res, [s.pull for s in r.sources], [])
        out = {**out, "gate": p.gate.outcome, "routes": {s.pull.sheet: s.route for s in r.sources},
               "review": r.review, "material": r.material or [], "outline": contract.outline(r.raw or {}),
               "timings": tm.summary()}
        if followup is not None:
            out.update(followup_of=r.q, reused=r.reused)
        if extra:
            warn = list(dict.fromkeys(list(extra.get("warnings") or []) + list(out.get("warnings") or [])))
            out = {**out, **{k: v for k, v in extra.items() if k != "warnings"}, "warnings": warn}
        return out
    try:
        return await _call_user(job, what="sv_answer")
    except (HostError, ToolError, InputError) as e:
        raise _err(e)
