"""One question through the new core (design §3): scope -> pull -> answer -> check. The scope gate's verdict
arrives in the answer spec (the agent applies the checklist); ambiguous questions stop in tools_core before any
Excel work. A follow-up (phase 5) patches the stored spec and redraws the same answer in its block."""
from __future__ import annotations

import copy
from dataclasses import dataclass, field

from pydantic import ValidationError

from .. import answer as answer_mod
from .. import catalog, drill_xl, fin_spec, specs
from ..excel import jobs, sheets, target as xltarget
from ..host import InputError, rewrap
from ..timing import stage
from . import check, commentary, contract, findings, followup, progress, recipes, source
from . import owned as owned_mod


@dataclass
class Run:
    t: object
    q: int
    spec: contract.AnswerSpec
    prepared: dict
    res: dict
    sources: list
    review: dict
    material: list = None
    raw: dict = None                                   # the stored spec (what a follow-up patches)
    owned: dict | None = None                          # the tool-owned block of the answer sheet
    reused: list[str] = field(default_factory=list)    # follow-up: pull ids whose sheets were kept


def parse_spec(raw: dict) -> contract.AnswerSpec:
    try:
        return contract.validate(raw)
    except ValidationError as e:
        raise InputError("This is not a valid answer spec v1: " + "; ".join(
            f"{'.'.join(map(str, x['loc'])) or 'spec'}: {x['msg']}" for x in e.errors()[:8])) from None


def gate_reply(spec: contract.AnswerSpec) -> dict | None:
    if spec.gate.outcome != "ambiguous":
        return None
    return {"gate": "ambiguous", "questions": spec.gate.questions, "checklist": spec.gate.checklist,
            "note": "Nothing was pulled. Ask the user these questions in ONE message, then call sv_answer again "
                    "with the answers applied (gate 'clear', assumptions stated)."}


def answer_layout(raw_layout: dict, spec: contract.AnswerSpec) -> dict:
    """The layout answer.build renders, with the question, assumptions and (for a proxy) what is answered instead
    leading the notes. The raw layout is not changed: it is what save_spec stores."""
    lay = copy.deepcopy(raw_layout)
    lead = [f"Question (as asked): {spec.question}"] + [f"Assumption: {a}" for a in spec.assumptions]
    if spec.deviation:
        lead.append(f"Deviation from the standard analysis (user's request): {spec.deviation}")
    if spec.gate.outcome == "infeasible":
        lead.append(f"Not answerable as asked ({spec.gate.reason}); answered instead: {spec.gate.proxy}")
    lay["notes"] = lead + list(lay.get("notes") or [])
    drivers = commentary.sheet_lines([d.model_dump() for d in spec.commentary.drivers])
    if drivers:
        lay["key_points"] = list(lay.get("key_points") or []) + drivers
    return lay


def drill_aliases(spec) -> set[str]:
    """The drill aliases of every drill cell in the spec's layout and extras (lines and repeat templates)."""
    lay = spec.layout.model_dump() if getattr(spec, "layout", None) is not None else {}
    lays = [lay] + [x.layout.model_dump() for x in getattr(spec, "extras", None) or []]
    out = set()
    for sec in (s for lay in lays for s in lay.get("sections") or []):
        lines = list(sec.get("lines") or []) + ([sec["repeat"].get("line") or {}] if sec.get("repeat") else [])
        out |= {str(c["drill"]) for ln in lines for c in ln.get("cells") or [] if isinstance(c, dict) and "drill" in c}
    return out


def question_for(t, spec) -> int:
    """A new spec with drill cells answers the pending drill (Lines sheets without an Answer); otherwise the next q."""
    if drill_aliases(spec):
        pending = sheets.pending_drill_question(t)
        if pending is not None:
            return pending
    return sheets.next_question(t)


def check_drill_tables(t, q: int, spec) -> None:
    """Refuse drill cells whose Lines table is missing before anything is written: Excel rejects the formula
    (0x800A03EC). Tables that cannot be read are not checked."""
    aliases = drill_aliases(spec)
    if not aliases:
        return
    try:
        have = drill_xl.lines_tables(t.wb)
    except InputError:
        raise
    except Exception:   # noqa: BLE001 - cannot check: the formulas are written as before
        return
    known = {h.lower() for h in have}
    missing = [n for n in (drill_xl.table_name(q, a) for a in sorted(aliases)) if n.lower() not in known]
    if missing:
        raise InputError(f"The drill table(s) {', '.join(missing)} do not exist in this workbook (Lines tables here: "
                         f"{', '.join(have) or 'none'}). Drill first with sv_drill(q={q}, id='<alias>'), or use the "
                         "recipe the gate named (lines_ranking drills itself).")


def run(h, raw: dict, target: str = "session", allow_archive: bool = False, *,
        q: int | None = None, old: dict | None = None, t=None, frame_done: bool = False,
        drills: dict[str, list[dict]] | None = None, value_findings: list[dict] | None = None) -> Run:
    """A new question (old None) -> a new Q<n> Answer; a follow-up (q + old = sheets.load_answer, raw already
    patched) -> the same Q<q> Answer redrawn in its block, unchanged pulls not fetched again.
    frame_done: the caller already drew the first frame of Q<q> Answer (a recipe that drafts or drills first).
    drills: alias -> drilled rows (drafted key points); value_findings: the recipe's findings on calculated values."""
    spec = parse_spec(raw)
    if spec.gate.outcome == "ambiguous":
        raise InputError("An ambiguous question pulls nothing: ask its questions first.")
    pf = source.Prefetch(spec.pulls) if old is None else None   # grids arrive while Excel opens the workbook
    with stage("open target"):
        t = t or xltarget.resolve_target(h, target, allow_archive, title=spec.question[:50])
        if q is None:
            q = question_for(t, spec)
    check_drill_tables(t, q, spec)
    if old is None:
        return _new(h, t, q, raw, spec, frame_done, drills, value_findings, pf=pf)
    return _follow(h, t, q, raw, spec, old)


def _new(h, t, q: int, raw: dict, spec: contract.AnswerSpec, frame_done: bool = False,
         drills: dict | None = None, value_findings: list[dict] | None = None, pf: source.Prefetch | None = None) -> Run:
    sources: list[source.Source] = []
    try:
        if not frame_done:
            with stage("answer skeleton"):
                progress.show(t, q, spec, f"fetching {len(spec.pulls)} source grid(s) ...")
        for i, pin in enumerate(spec.pulls):
            progress.set_status(t, q, f"fetching {pin.connection} '{pin.id}' ({i + 1}/{len(spec.pulls)}) ...")
            sources.append(source.pull(h, t, q, i, pin, pf=pf))
        progress.set_status(t, q, "building the answer ...")
        ids = {pin.id: s.pull.sheet for pin, s in zip(spec.pulls, sources)}
        sess = {s.pull.sheet: s.pull for s in sources}
        res, prepared = _draw(t, q, raw, spec, ids, sess, owned=None, drills=drills)
    except Exception as e:
        if pf is not None:
            pf.cancel()
        removed = sheets.rollback(t)
        raise rewrap(e, f"{e} (the answer was not completed; removed the partial sheets: "
                     f"{', '.join(removed) or 'none'})") from None
    return _finish(t, q, raw, spec, ids, prepared, res, sources, value_findings=value_findings)


def _exists(t, name: str) -> bool:
    return sheets._sheet_exists(t, name)


def _delete(t, names: list[str]) -> None:
    if names:
        sheets.drop_parked(t, {n: n for n in names})


def _follow(h, t, q: int, raw: dict, spec: contract.AnswerSpec, old: dict) -> Run:
    """Design §3.3 follow-up: reuse the sheets of unchanged pulls, park + refetch changed ones, redraw Q<q> Answer
    over its own block (a hidden backup restores it if the redraw fails), drop the sheets of removed pulls.
    The status goes to Excel's status bar only: the skeleton would clear the user's sheet."""
    plan = followup.plan_sources(old["spec"], old["ids"], raw, exists=lambda s: _exists(t, s))
    stored = sheets.load_pulls(t, q) if plan.reuse else {}
    for pid, sheet in list(plan.reuse.items()):
        if stored.get(sheet) is None:                       # no meta row (sheet copied by hand): fetch it again
            del plan.reuse[pid]
            plan.replace[pid] = sheet
    raw = {**raw, "layout": followup.drop_stale_expect(raw["layout"], list(plan.replace))}
    name = f"Q{q} {progress.ANSWER}"
    parked, bname, sources, ids, sess = {}, None, [], {}, {}
    try:
        if plan.replace:
            parked = sheets.park_sheets(t, list(plan.replace.values()))
        fresh = [p for p in spec.pulls if p.id not in plan.reuse]
        pf = source.Prefetch(spec.pulls, only={i for i, p in enumerate(spec.pulls) if p.id not in plan.reuse})
        for i, pin in enumerate(spec.pulls):
            if pin.id in plan.reuse:
                sheet = plan.reuse[pin.id]
                src = source.Source(stored[sheet], "reused", [])
            else:
                t.status(f"fetching {pin.connection} '{pin.id}' ({fresh.index(pin) + 1}/{len(fresh)}) ...")
                src = source.pull(h, t, q, i, pin, pf=pf)
                sheet = src.pull.sheet
            sources.append(src)
            ids[pin.id], sess[sheet] = sheet, src.pull
        t.status("redrawing the answer ...")
        if _exists(t, name):
            bname = owned_mod.backup(t, name)
        res, prepared = _draw(t, q, raw, spec, ids, sess, owned=old.get("owned"))
    except Exception as e:
        sheets.rollback(t)                               # removes freshly added source sheets
        if parked:
            sheets.restore_parked(t, parked)
        if bname:
            owned_mod.restore(t, name, bname)
        raise rewrap(e, f"{e} (the follow-up was not applied; {name} and its sources are unchanged)") from None
    _delete(t, [n for n in [bname, *plan.drop] if n])
    if parked:
        sheets.drop_parked(t, parked)
    return _finish(t, q, raw, spec, ids, prepared, res, sources, reused=list(plan.reuse))


def expect_checks(layout: dict, by_id: dict) -> dict:
    """layout.expect (recipe pnl) -> one static '<source> = source sheet' check per pulled source; expect removed."""
    exp = layout.get("expect")
    if not exp:
        return layout
    out = {k: v for k, v in layout.items() if k != "expect"}
    extra = [fin_spec.reconcile_pull(by_id[pid], [tuple(c) for c in cells], source=exp["source"])
             for pid, cells in (exp.get("cells") or {}).items() if pid in by_id]
    return {**out, "checks": list(out.get("checks") or []) + extra}


def _draw(t, q: int, raw: dict, spec: contract.AnswerSpec, ids: dict, sess: dict, owned: dict | None,
          drills: dict | None = None):
    by_id = {pid: sess[sheet] for pid, sheet in ids.items() if sheet in sess}
    layout = {k: v for k, v in expect_checks(raw["layout"], by_id).items() if k != "expect"}
    with stage("prepare spec"):
        prepared = specs.prepare(answer_layout(layout, spec), ids, sess, drills=drills)
    with stage("answer"):
        res = answer_mod.build(t, q, prepared, sess, suffix=progress.ANSWER, owned=owned, reserve=findings.RESERVE)
    _extras(t, q, raw, ids, sess, res, drills)
    return res, prepared


def _extras(t, q: int, raw: dict, ids: dict, sess: dict, res: dict, drills: dict | None = None) -> None:
    """Each extra as its own sheet 'Q<q> <topic>' (after the answer), linked both ways; its checks, formula errors
    and warnings join the answer's (checks as '<topic>: <label>') so the review covers them."""
    topics = []
    for x in raw.get("extras") or []:
        topic = x["topic"]
        with stage(f"extra '{topic}'"):
            prep = specs.prepare(x["layout"], ids, sess, drills=drills)
            rx = answer_mod.build(t, q, prep, sess, suffix=topic)
        res.setdefault("checks", {}).update({f"{topic}: {k}": v for k, v in (rx.get("checks") or {}).items()})
        res.setdefault("formula_errors", []).extend(f"{topic}: {k}" for k in rx.get("formula_errors") or [])
        res.setdefault("warnings", []).extend(rx.get("warnings") or [])
        res.setdefault("extras", {})[topic] = {"sheet": rx["sheet"], "values": rx.get("values") or {}}
        topics.append(topic)
    if topics:
        res["extra_sheets"] = [res["extras"][x]["sheet"] for x in topics]
        progress.link(t, q, topics)


def _finish(t, q: int, raw: dict, spec: contract.AnswerSpec, ids: dict, prepared: dict, res: dict, sources: list,
            reused: list[str] | None = None, value_findings: list[dict] | None = None) -> Run:
    flagged = commentary.material(prepared, res.get("values") or {}, catalog.load())
    review = check.review(res, sources, extra=commentary.findings(
        flagged, [d.model_dump() for d in spec.commentary.drivers])
        + recipes.value_findings(value_findings, res.get("values") or {}),
        warn_checks={c.label for c in spec.layout.checks if c.severity == "Warning"})
    ext = res.get("extent")
    if ext:
        try:
            with stage("findings table"):
                ws = jobs.com(lambda: t.wb.Worksheets(res["sheet"]))
                ext = {**ext, "rows": findings.render(ws, ext["rows"] + 2, review["findings"])}
        except Exception as e:   # noqa: BLE001 - the table never undoes a finished answer
            res.setdefault("warnings", []).append(f"findings table not written on the sheet: {e}")
    try:
        sheets.save_spec(t, q, raw, ids, owned=ext)
    except Exception:   # noqa: BLE001 - only needed for follow-ups
        pass
    return Run(t, q, spec, prepared, res, sources, review, flagged, raw=raw, owned=ext, reused=reused or [])
