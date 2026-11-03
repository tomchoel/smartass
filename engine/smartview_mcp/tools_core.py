"""The pipeline tools: sv_gate (scope), sv_remember (learned defaults) and sv_answer (one question, a follow-up or a
recipe; design 2026-11-02 §3). gate() and answer() are also the HTTP API's POST /gate and /answer."""
from __future__ import annotations

import contextlib
import datetime as dt
import json
import os
import time
from typing import Literal

from mcp.server.mcpserver import Context
from pydantic import ValidationError

from . import catalog, config, templates
from .excel import sheets, target as xltarget
from .core import contract, drillstep, drilltime, pipeline, prefs, progress, recipes, reply, runlog, scope
from .core import followup as followup_mod
from .core import period as periods
from .host import HostError, InputError
from .server import LOCAL, READ, WRITE, _err, _is_our_analysis_file, _with_progress, host, tool
from .timing import COM_PREFIX, Timer, stage


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
        d = dt.date.today().replace(day=1) - dt.timedelta(days=1)
        return periods.label(d.year, d.month), f"closed month not read from EPM ({str(e)[:120]}); assumed last month"


@tool(annotations=READ)
async def sv_gate(question: str, answers: dict[str, str] | None = None) -> dict:
    """Optional: sv_answer(question=...) runs the gate itself; call sv_gate alone only to inspect the scope.
    The gate (knowledge/advisor/gate.md) runs the fixed scope checklist (period, scenario, entity, measure, currency/sign,
    granularity) against the catalog, the capability map and the user's learned defaults, and suggests one outcome:
    - clear: proceed; state every checklist item with source convention/learned (and each warning) as an assumption,
      then call sv_answer with the named recipe (+ params) now; never ask whether to build a sheet.
    - ambiguous: pull nothing; ask the returned questions (max 3) in ONE message, in the user's language; then call
      sv_gate again with the same question and answers={item: answer} (e.g. {"measure": "opex"}) and offer
      sv_remember to save an answer as the default.
    - infeasible: say why (reason) in one or two sentences and offer the proxy; build it labelled as a proxy.
    Copy the result into the answer spec's gate (outcome, checklist, questions/reason/proxy) and assumptions.
    Returns {suggested_outcome, checklist{item: {value, source, gap, ask}}, questions, reason, proxy, warnings,
    entities, measures, concepts, closed_month, recipe: {name, params} | None (None: a free-form spec), next}."""
    try:
        return await gate(question, answers)
    except InputError as e:
        raise _err(e) from None


async def gate(question: str, answers: dict | None = None) -> dict:
    """sv_gate without the MCP wrapper (also POST /gate); raises InputError."""
    if not question or not question.strip():
        raise InputError("question is the user's question, as asked")
    if answers is not None and not isinstance(answers, dict):
        raise InputError("answers is {checklist item: the user's answer}, e.g. {\"measure\": \"opex\"}")
    closed, warn = await _closed_month()
    g = scope.gate(question.strip(), closed, answers=answers)
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
    old = sheets.load_answer(t, q)
    if old is None or old["spec"].get("version") != 1:
        raise InputError(f"Q{q} Answer was not drawn by sv_answer, so it has no answer spec to patch: ask it as a "
                         "new question with spec=...")
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
    """(res, mode, fills, notes, entity) for recipe pnl: the facts read live by pnl_facts.prepare."""
    from . import atlas, pnl_facts
    cat = catalog.load()
    with contextlib.closing(atlas.connect()) as adb:
        entity = pnl_facts.resolve_scope(cat, area, adb)
    res, mode, fills, notes = pnl_facts.prepare(h, cat, entity, period, compare)
    return res, mode, fills, notes, entity


def _recipe_ctx(h, lines: dict | None = None) -> recipes.RecipeContext:
    return recipes.RecipeContext(closed=templates.closed_month(h), cat=catalog.load(), label=templates.member_label,
                                 pnl_facts=lambda area, period, compare: _pnl_facts(h, area, period, compare),
                                 now=dt.datetime.now().strftime("%Y-%m-%d %H:%M"), lines=lines)


def _expand_recipe(h, spec: dict, lines: dict | None = None) -> tuple[dict, dict]:
    """(the full spec, extra result keys) for a recipe spec; runs inside the Excel job (closed month, live reads).
    lines: the drilled rows by alias (lines_ranking)."""
    ctx = _recipe_ctx(h, lines)
    try:
        return recipes.expand(spec, ctx)
    except InputError:
        raise
    except ValueError as e:          # pydantic ValidationError is a ValueError
        raise InputError(f"The recipe does not give a valid answer spec: {e}") from None


def _closed_cached() -> str:
    """The closed month for answer()'s re-gate: templates.closed_month's disk cache, else the month before today
    (never EPM, never Excel: the re-gate must be instant)."""
    try:
        e = json.loads((config.SVKIT / "closed_month.json").read_text(encoding="utf-8")).get("FinMax")
        month = e.get("month") if isinstance(e, dict) else e
        if isinstance(month, str) and month:
            return month
    except (OSError, ValueError, AttributeError):
        pass
    d = dt.date.today().replace(day=1) - dt.timedelta(days=1)
    return periods.label(d.year, d.month)


def _precheck(spec: dict) -> dict | None:
    """Before any Excel job (D5, D4, D7): the gate re-run on the question (still ambiguous -> its questions, nothing
    pulled), departures from the standard analysis without a deviation, and sheet text that is not English."""
    question = str(spec.get("question") or "").strip()
    if not question:
        raise InputError("spec.question is the user's question, as asked")
    answers = (spec.get("gate") or {}).get("answers") or None
    g = scope.gate(question, _closed_cached(), answers=answers)
    if g["suggested_outcome"] == "ambiguous":
        return {"gate": "ambiguous", "questions": g["questions"],
                "checklist": {k: v["value"] or v["gap"] for k, v in g["checklist"].items()},
                "note": "Nothing was pulled: the gate is still ambiguous for this question. Ask the user these "
                        "questions in their language, in ONE message; then call sv_gate again with the same question "
                        "and answers={item: answer}, and pass the same answers as spec.gate.answers."}
    named, rec = (g.get("recipe") or {}).get("name"), spec.get("recipe") or {}
    dep = []
    if named and not rec:
        dep.append(f"a free-form spec where sv_gate named recipe {named}")
    elif named and rec.get("name") != named:
        dep.append(f"recipe {rec.get('name')} instead of {named}")
    if rec:
        dep += recipes.departures(rec.get("name"), rec.get("params") or {})
    dep += recipes.layout_departures(spec.get("layout"))
    if dep and not spec.get("deviation"):
        raise InputError("This departs from the standard analysis: " + "; ".join(dep) + ". Use the recipe and "
                         "defaults sv_gate named, or pass deviation='<the user's explicit request>' (it is written "
                         "on the sheet).")
    bad = [(path, s) for path, s in contract.sheet_texts(spec) if contract.non_english(s)]
    if bad:
        raise InputError("Sheet text that is not English:\n" + "\n".join(f"  {p}: {s}" for p, s in bad[:8]) +
                         "\nWrite sheet text in English (the chat stays in the user's language); names and member "
                         "labels may stay as they are.")
    return None


@tool(annotations=WRITE)
async def sv_answer(spec: dict | None = None, followup: dict | None = None, q: int | None = None,
                    target: str = "session", allow_archive: bool = False, question: str | None = None,
                    answers: dict[str, str] | None = None, detail: str = "lean",
                    ctx: Context | None = None) -> dict:
    """Answer ONE finance question. New question: sv_answer(question=...) - it runs the gate itself; ambiguous ->
    ask its questions in ONE message, then sv_answer(question=..., answers={item: answer}). Follow-up:
    sv_answer(followup=...). A hand-written spec (spec=...) still works. Give exactly one of question, spec and
    followup.
    question: the user's question as asked; the gate (knowledge/advisor/gate.md) runs first. ambiguous / infeasible
    -> returns the gate (questions / reason + proxy), pulls nothing. clear with a recipe -> answered at once with
    that recipe and the gate's assumptions. clear without a recipe -> asks for a spec (pulls + layout).
    answers = {checklist item: the user's answer} to the questions of an earlier ambiguous reply (same question).
    spec = the answer spec v1 (core/contract.py), below.
    spec (a new question -> a new 'Q<n> Answer'): {version: 1, question, gate: {outcome: clear|ambiguous|infeasible,
    checklist: {period, scenario, entity, measure, currency_sign, granularity}, questions (ambiguous, 1-3), reason +
    proxy (infeasible)} (or sv_gate's output as is: it is normalised, and its recipe used when spec has none),
    assumptions: [...], pulls: [{id, connection, rows, columns, pov, title?}], layout: {title,
    columns, sections, checks, ...} (the answer spec grammar; pull cells use the pull ids)}.
    ambiguous -> returns the questions, pulls nothing. Otherwise: the answer sheet 'Q<n> Answer' appears first with
    the question, assumptions and a status line; each pull is fetched over HTTP and written as a source sheet in one
    block (Smart View ad hoc grid if HTTP fails); then the answer is rendered (live formulas) and checked.
    A recipe drafts pulls + layout server-side: spec = {version: 1, question, gate, assumptions?, recipe: {name,
    params, overrides?}} without pulls/layout. Recipes: pnl {area, period?, compare?}; actual_vs_plan {entity
    (FinPlan), month?, year?, basis?, forecast?, finmax_entity?}; monthly_trend {entity?, accounts?, through?,
    compare? (default with last year)}; cost_base {parent?, accounts?, period?, top? (10 + Other)}; concept {concept,
    period?, compare?, entity?}; lines_ranking {accounts, entity?, period?, by, projects?, top?, compare?,
    group_separate?} (drills the ledger lines itself). Use the recipe + params sv_gate named; a departure needs
    spec.deviation = the user's explicit request (written on the sheet). Sheet text must be English. The period
    defaults to the closed month (returned as period). overrides = a follow-up-shaped patch on the draft; recipe
    params cannot be changed by a follow-up (ask a new question).
    followup: a patch on the current answer (the active Q<n> Answer, else the latest; or answer q). Use it unless
    the user asks a new question. Dicts merge, null deletes; pulls merge by id, layout.sections by title, lines by
    key, checks by label ({..., "drop": true} removes; new ones are appended); {"append": [...]} adds to a list
    (e.g. notes); any other list replaces. Unchanged pulls are not fetched again; the answer is redrawn in its own
    block, so the user's cells and charts outside it survive.
    Returns the answer table, checks, review findings (Critical/Warning/Info: knowledge/advisor/findings.md),
    material lines that need a driver in commentary.drivers (knowledge/advisor/variance-commentary.md), routes,
    timings and an outline (pull ids, columns, section titles, line keys, check labels) to address follow-ups;
    key_values (line -> column -> value as the sheet calculated them: quote these), followups (at most 3 next
    questions of the recipe) and sheets (every sheet written).
    detail="full" returns the whole table and every recipe extra (default "lean": rounded key_values, summarised
    checks, big extras trimmed).
    Use spec=... when the gate named no recipe (or the user asked for a departure: spec.deviation)."""
    try:
        return await answer(spec, followup, q, target, allow_archive, question, answers, detail)
    except (HostError, InputError) as e:
        raise _err(e) from None


def _assumptions(g: dict) -> list[str]:
    """The gate's convention / learned checklist items as 'Item: value - source.', then its warnings."""
    out = []
    for key, item in (g.get("checklist") or {}).items():
        if isinstance(item, dict) and item.get("source") in ("convention", "learned"):
            label = key.replace("_", " ").capitalize()
            out.append(f"{label}: {item.get('value')} - {item['source']}.")
    return out + [str(w) for w in g.get("warnings") or []]


async def answer(spec: dict | None = None, followup: dict | None = None, q: int | None = None,
                 target: str = "session", allow_archive: bool = False, question: str | None = None,
                 answers: dict | None = None, detail: str = "lean") -> dict:
    """sv_answer without the MCP wrapper (also POST /answer); raises InputError / HostError."""
    if sum(x is not None for x in (question, spec, followup)) != 1:
        raise InputError("Give exactly one of question (a new question: the gate runs first), spec (a hand-written "
                         "answer spec) or followup (change the current answer).")
    if question is not None:
        g = await gate(question, answers)
        if g["suggested_outcome"] != "clear":
            return g
        if not g.get("recipe"):
            return {**g, "next": "Clear, but no recipe fits: send sv_answer(spec=...) with pulls and layout."}
        if answers:
            g = {**g, "answers": {k: str(v) for k, v in answers.items()}}
        spec = {"version": 1, "question": question, "gate": g, "assumptions": _assumptions(g),
                "recipe": g["recipe"]}
        return await _answer_spec(spec, None, None, target, allow_archive, detail)
    return await _answer_spec(spec, followup, q, target, allow_archive, detail)


async def _answer_spec(spec: dict | None, followup: dict | None, q: int | None, target: str, allow_archive: bool,
                       detail: str = "lean") -> dict:
    """Answer a spec (new question) or apply a follow-up."""
    parsed = None
    if spec is not None and isinstance(spec.get("gate"), dict):
        raw_gate = spec["gate"]
        spec = {**spec, "gate": contract.normalize_gate(raw_gate)}
        named = raw_gate.get("recipe") if "suggested_outcome" in raw_gate else None
        if named and not spec.get("recipe") and not spec.get("layout") and not spec.get("pulls"):
            spec["recipe"] = named                  # the gate's own recipe when the caller passed the gate as is
    recipe = spec is not None and bool(spec.get("recipe")) and not spec.get("layout")
    if spec is not None:
        if recipe and (spec.get("gate") or {}).get("outcome") != "ambiguous":
            _check_recipe_spec(spec)            # pulls + layout come from the recipe, inside the Excel job
        else:
            parsed = pipeline.parse_spec(spec)
        early = pipeline.gate_reply(parsed) if parsed is not None else None
        if early is not None:
            runlog.append(runlog.entry(parsed, None, [], None, {}))
            return early
        early = _precheck(spec)
        if early is not None:
            return early

    def job(h, pre: dict | None = None):
        tm = Timer.start()
        p = parsed
        extra, drawn = {}, []
        try:
            if followup is None and recipe:
                t, fq = _reopen(h, pre, allow_archive) if pre else \
                    _frame(h, spec, target, allow_archive, "drafting the analysis ...")
                try:
                    lines, drawn = _drilled(t, fq, spec, pre) if pre else (None, [])
                    raw, extra = _expand_recipe(h, spec, lines)
                    p = pipeline.parse_spec(raw)
                except Exception:
                    sheets.rollback(t)
                    raise
                extra["warnings"] = (pre or {}).get("warnings", []) + list(extra.get("warnings") or [])
                r = pipeline.run(h, raw, target, allow_archive, q=fq, t=t, frame_done=True, drills=lines,
                                 value_findings=extra.pop("findings", None))
            elif followup is None:
                r = pipeline.run(h, spec, target, allow_archive)
            else:
                with stage("open target"):
                    t = xltarget.resolve_target(h, target, allow_archive)
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
                    sheets.log_source(r.t, r.q, p.question, s.pull, notes=s.note or s.route)
                except Exception as e:   # noqa: BLE001 - the source log must never undo a finished answer
                    r.res["warnings"].append(f"could not log {s.pull.sheet} in 'Sources': {e}")
        e = runlog.entry(p, r.q, r.sources, r.review, tm.summary(), workbook=r.t.path)
        try:
            runlog.write_sheet(r.t, e)
        except Exception as ex:   # noqa: BLE001
            r.res["warnings"].append(f"run log sheet not written: {ex}")
        runlog.append(e)
        with stage("save + activate"):
            reply.finish(r.t, r.q, r.prepared, r.res, saved_new=r.t.created)
        out = reply.result(r.t, r.q, r.prepared, r.res, [s.pull for s in r.sources], [])
        out = {**out, "gate": p.gate.outcome, "routes": {s.pull.sheet: s.route for s in r.sources},
               "review": r.review, "material": r.material or [], "outline": contract.outline(r.raw or {}),
               "timings": tm.summary()}
        if followup is not None:
            out.update(followup_of=r.q, reused=r.reused)
        if drawn:                                   # the Lines sheets after the answer's own sheets
            n = 1 + len(r.res.get("extra_sheets") or [])
            out["sheets"] = out["sheets"][:n] + drawn + out["sheets"][n:]
            out["data_sheets"] = drawn + out["data_sheets"]
        if pre:
            out["timings"] = {**out["timings"], "first frame": pre["frame_s"], "drills": pre["drill_s"],
                              **pre.get("stages", {})}
        fill = extra.pop("fill", None) if extra else None
        if extra:
            warn = list(dict.fromkeys(list(extra.get("warnings") or []) + list(out.get("warnings") or [])))
            out = {**out, **{k: v for k, v in extra.items() if k != "warnings"}, "warnings": warn}
        rec = (r.raw or {}).get("recipe") or {}
        params = rec.get("params") or {}
        out["followups"] = recipes.followups(rec.get("name"), params, out.get("period") or params.get("period"), fill)
        return reply.lean(out, detail)
    if followup is None and recipe and recipes.drilling(spec["recipe"].get("name")):
        return await _answer_with_drills(spec, target, allow_archive, job)
    return await _call_user(job, what="sv_answer")


async def _answer_with_drills(spec: dict, target: str, allow_archive: bool, job) -> dict:
    """A recipe that drills (lines_ranking, D21): job A draws the first frame, the drills run on the event loop
    (no Excel held), job B writes the Lines sheets and renders the answer - a cube-only answer if d1 failed."""
    t0, rec = time.monotonic(), spec["recipe"]

    def frame_job(h):
        # the frame first: the closed month (EPM, or an Excel fallback) and the plan are not needed to draw it
        ts, tm = {"wait for Excel": time.monotonic() - t0}, Timer.start()     # job A's stages in the log + timings
        accounts = ", ".join(str(a) for a in (rec.get("params") or {}).get("accounts") or []) or "the accounts"
        t, q = _frame(h, spec, target, allow_archive, f"Drilling ledger lines for {accounts} (about 60 s) ...", ts)
        frame_s = time.monotonic() - t0
        ctx = _recipe_ctx(h)
        plan = recipes.drills(rec.get("name"), rec.get("params") or {}, ctx)
        ts["closed month + plan"] = time.monotonic() - t0 - frame_s
        return {"target": t.path if os.path.isabs(str(t.path)) else "active", "created": bool(t.created),
                "added": list(t.sheets_added), "q": q, "plan": plan, "closed": ctx.closed, "frame_s": round(frame_s, 1),
                "stages": {f"first frame.{k}": round(v, 1) for k, v in {**tm.summary(), **ts}.items()
                           if k != "_total" and not k.startswith(COM_PREFIX)}}
    pre = await _call_user(frame_job, what="sv_answer")
    t1, _ = time.monotonic(), drilltime.take()     # drop what earlier drills (sv_drill) left
    pre["results"], warnings, pre["failed"] = await drillstep.drill_all(pre.pop("plan"))
    pre["drill_s"] = round(time.monotonic() - t1, 1)
    pre["stages"].update({f"drills.{k}": v for k, v in drilltime.take().items()})
    if pre["failed"]:
        warnings.append(f"Drill failed: {pre['failed']}")
    pre["warnings"] = warnings
    return await _call_user(lambda h: job(h, pre), what="sv_answer")


def _frame(h, spec: dict, target: str, allow_archive: bool, status: str, ts: dict | None = None):
    """(t, q): open the target and draw the first frame of 'Q<q> Answer' before the recipe drafts or drills (ts gets
    'open workbook' and 'draw frame', seconds)."""
    t0 = time.monotonic()
    with stage("open target"):
        t = xltarget.resolve_target(h, target, allow_archive, title=str(spec.get("question") or "")[:50])
        q = sheets.next_question(t)
    t1 = time.monotonic()
    with stage("answer skeleton"):
        progress.show(t, q, progress.frame(spec), status)
    if ts is not None:
        ts.update({"open workbook": t1 - t0, "draw frame": time.monotonic() - t1})
    return t, q


def _reopen(h, pre: dict, allow_archive: bool):
    """Job B: the workbook job A framed (COM objects do not cross jobs), with job A's new sheets for a rollback."""
    with stage("open target"):
        t = xltarget.resolve_target(h, pre["target"], allow_archive)
    t.created = t.created or pre["created"]
    t.sheets_added[:0] = [s for s in pre["added"] if s not in t.sheets_added]
    return t, pre["q"]


def _drilled(t, q: int, spec: dict, pre: dict) -> tuple[dict | None, list[str]]:
    """({alias: rows} or None if d1 failed, the Lines sheets written): each drill as 'Q<q> Lines <alias>', with
    the recipe's derived columns."""
    if pre["failed"]:
        progress.set_status(t, q, f"Drill failed: {pre['failed']}; cube totals only")
        return None, []
    rec, lines, names = spec["recipe"], {}, []
    for alias, res in pre["results"].items():
        res.columns = recipes.mark(rec.get("name"), rec.get("params") or {}, res.rows, res.columns)
        with stage(f"write lines {alias}"):
            info = drillstep.write(t, q, alias, res, title=f"Ledger lines {alias}: {spec['question']}"[:120],
                                   question=spec["question"])
        lines[alias] = res.rows
        names += [info["sheet"]] if info else []
    progress.set_status(t, q, "drafting the analysis ...")
    return lines, names


async def _call_user(fn, what: str):
    """server._call_user without the MCP error mapping: HTTP needs InputError/HostError as they are."""
    return await host().call(_with_progress(fn), what=what, user_excel=True)
