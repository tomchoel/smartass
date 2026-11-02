"""The advisor gate (design §3.1): every question passes the same fixed checklist before anything is pulled.

    gate(question, closed, answers=None) -> {suggested_outcome: clear | ambiguous | infeasible, checklist, questions,
                               reason, proxy, warnings, entities, measures, concepts, recipe, next}

recipe = the standard recipe (+ params) for a clear question (first match in _recipe), or None = a free-form spec.
answers = the user's answers by checklist item; each resolves its item with source "user".

Each checklist item is resolved from the question, a learned default (core/prefs), the catalog, or a stated convention,
or left as a gap. Ask policy 'balanced': only gaps that change the numbers are asked (measure, 'this year'); the rest
are conventions stated on the answer sheet. Infeasible = the cubes cannot answer it as asked; the proxy is offered.
The agent may overrule the suggestion, but must say why (the run log records both)."""
from __future__ import annotations

import re
from typing import Callable

from .. import catalog, concepts
from ..host import InputError
from . import period as periods
from . import prefs, semantic
from .contract import CHECKLIST

Learned = Callable[[str, str], "str | None"]

CONCEPT_MIN_SCORE = 2          # a concept counts as the measure only when at least two question words match it
_MEMBER_RX = re.compile(r"(?<![\w.])(un\.SY(?:\.[\w-]+)*)", re.I)
_SCENARIO_RX = {
    "budget": re.compile(r"\b(budget|budgets|budsjett|budsjettet|target|plan|mål)\b", re.I),
    "forecast": re.compile(r"\b(forecast|forecasts|prognose|prognosen|fcst|outlook|estimate)\b", re.I),
    "actual": re.compile(r"\b(actual|actuals|faktisk|regnskap|realised|realized|outcome)\b", re.I),
    "last_year": re.compile(r"\b(last year|prior year|previous year|i fjor|fjoråret|ly|yoy|year over year)\b", re.I),
}
_ASSUME = ("Put every checklist item whose source is convention or learned on the answer sheet as a stated "
           "assumption (and each warning). ")
NEXT = {
    "clear": _ASSUME + "Call sv_answer with spec {version: 1, question, gate (copy), recipe: {name: '%s', params: "
                       "the recipe params above}} now; it builds the sheets. Never ask whether to build a sheet. "
                       "Quote key_values from its reply.",
    "free": _ASSUME + "No standard recipe fits: write a free-form spec and call sv_answer now; it builds the sheets. "
                      "Never ask whether to build a sheet. Quote key_values from its reply.",
    "ambiguous": "Do not pull yet. Ask these questions in the user's language, in one message; then call sv_gate again "
                 "with the same question and answers={item: answer}; offer sv_remember for an answer they want as "
                 "the default next time.",
    "infeasible": "Do not pull what was asked. Tell the user the reason in one or two sentences, then offer the proxy "
                  "(build it when they agree, labelled as a proxy on the answer sheet).",
}


def _item(value, source, gap=None, ask=False) -> dict:
    return {"value": value, "source": source, "gap": gap, "ask": ask}


def _entity(cat, question: str, warnings: list[str]) -> tuple[dict, list[dict]]:
    literal = _MEMBER_RX.search(question)
    if literal:
        member = literal.group(1)
        known = next((e for e in cat.entities if e.finmax and e.finmax.lower() == member.lower()), None)
        if known:
            found = (known,)
        else:
            warnings.append(f"{member} has no FinPlan member: no plan/budget comparison is possible, last year only")
            ent = {"key": member, "label": member, "finmax": member, "finplan": None}
            return _item(member, "question"), [ent]
    else:
        found = catalog.resolve_entities(cat, question)
    if not found:
        company = next(e for e in cat.entities if e.key == "company")
        found, source = (company,), "convention"
    else:
        source = "question"
    for e in found:
        if e.abolished:
            succ = ", ".join(x.label for x in cat.entities if x.key in e.successors)
            warnings.append(f"{e.label} was abolished {e.abolished}; its business is now in {succ}")
    ents = [{"key": e.key, "label": e.label, "finmax": e.finmax, "finplan": e.finplan} for e in found]
    label = " + ".join(f"{e.label} ({e.finmax})" for e in found)
    return _item(label, source), ents


def _measure(question: str, learned: Learned) -> tuple[dict, list[semantic.Measure], list[str]]:
    found = semantic.measures(question)
    vague = [m for m in found if m.vague]
    if vague:
        for m in vague:
            key = learned("measure", m.matched.lower())
            if key and semantic.measure(key):
                chosen = semantic.measure(key)
                return _item(semantic.describe(chosen), "learned"), [chosen], []
        ask = next((m for m in vague if m.key != "numbers"), vague[0])
        q = f"Which measure do you mean by '{ask.matched}': " + "; ".join(ask.options) + "?"
        return _item(None, "question", gap=q, ask=True), [], []
    if found:
        return _item("; ".join(semantic.describe(m) for m in found), "question"), found, []
    hits = [c for s, c in concepts.scored(question) if s >= CONCEPT_MIN_SCORE][:3]
    if hits:
        return _item(hits[0].get("title") or hits[0]["id"], "catalog"), [], [c["id"] for c in hits]
    q = _NO_MEASURE
    return _item(None, "question", gap=q, ask=True), [], []


_NO_MEASURE = "Which measure: revenue, opex, EBITDA, or the P&L to EBITDA?"


def _scenario(cat, question: str, p: periods.Period) -> tuple[dict, list[str]]:
    pos = {k: m.start() for k, rx in _SCENARIO_RX.items() if (m := rx.search(question))}
    named = sorted(pos, key=pos.get)                          # in the question's order: 'forecast vs budget'
    years = {y for y, _ in periods._months(question, 0) if y}
    if len(years) > 1 and "last_year" not in named:          # 'Sep 26 vs Sep 25' compares with last year
        named.append("last_year")
    labels = {"last_year": "Actual last year (comparison)"}
    if not [k for k in named if k != "last_year"]:
        base = "forecast" if p.kind == "fy" else "actual"
        named = [base, *named]
        source = "convention"
    else:
        source = "question"
    text = " vs ".join(labels.get(k) or cat.scenarios[k].label for k in named)
    return _item(text, source), named


def gate(question: str, closed: str, learned: Learned | None = None, answers: dict | None = None) -> dict:
    """answers: the user's answers to the gate's questions, by checklist item ({"measure": "opex"}); each one
    resolves that item (source "user") without rewording the question."""
    learned = learned or prefs.lookup
    answers = {k: str(v).strip() for k, v in (answers or {}).items() if str(v or "").strip()}
    unknown = set(answers) - set(CHECKLIST)
    if unknown:
        raise InputError(f"unknown answers {sorted(unknown)}; answer by checklist item: {', '.join(CHECKLIST)}")
    cat = catalog.load()
    warnings: list[str] = []
    reasons: list[str] = []
    proxies: list[str] = []
    checklist: dict[str, dict] = {}

    # period ------------------------------------------------------------------
    p = periods.parse(answers.get("period") or question, closed)
    if "period" in answers:
        checklist["period"] = _item(p.label, "user")
    elif p.kind == "this_year":
        val = learned("period", "this year")
        if val in ("ytd", "fy"):
            p = periods.parse(question.replace("this year", "YTD" if val == "ytd" else "full year"), closed)
            checklist["period"] = _item(p.label, "learned")
        else:
            checklist["period"] = _item(None, "question", gap=p.gap, ask=True)
    else:
        checklist["period"] = _item(p.label, "convention" if p.convention else "question")
    cy = periods.parse_month(closed)[0]
    if "period" not in answers and p.kind == "year" and p.end and p.end[0] == cy \
            and not _SCENARIO_RX["forecast"].search(question) \
            and not _SCENARIO_RX["budget"].search(question) and len(set(periods._BARE_YEAR_RX.findall(question))) == 1:
        # a bare current year ('i 2026') with actuals = year to date through the closed month, by convention
        p = periods.Period("ytd", f"YTD {closed}", convention=True, end=periods.parse_month(closed))
        checklist["period"] = _item(f"YTD {closed} ({cy} to date)", "convention")

    # scenario ----------------------------------------------------------------
    checklist["scenario"], scen = _scenario(cat, answers.get("scenario") or question, p)
    if p.open and "actual" in scen:
        reasons.append(f"actuals for {p.label} do not exist yet (last closed month {closed})")
        proxies.append(f"the latest forecast for {p.label}")

    # entity ------------------------------------------------------------------
    checklist["entity"], ents = _entity(cat, answers.get("entity") or question, warnings)

    # measure -----------------------------------------------------------------
    if "measure" in answers:
        checklist["measure"], found, concept_ids = _answered_measure(answers["measure"])
    else:
        checklist["measure"], found, concept_ids = _measure(question, learned)
    if not concept_ids:
        concept_ids = [c["id"] for s, c in concepts.scored(question) if s >= CONCEPT_MIN_SCORE][:3]
    if checklist["measure"]["gap"] == _NO_MEASURE and "actual" in scen and {"budget", "forecast"} & set(scen) \
            and any(e.get("finplan") for e in ents):
        # nothing named, compared with plan: actual_vs_plan shows the P&L to EBITDA, so asking changes nothing
        checklist["measure"] = _item("the P&L to EBITDA (actual vs plan shows every line)", "convention")

    # currency / sign ---------------------------------------------------------
    code, ok = semantic.currency(answers.get("currency_sign") or question)
    if ok:
        checklist["currency_sign"] = _item("NOK; FinMax signs (revenue positive, costs negative)", "convention")
    else:
        checklist["currency_sign"] = _item(code, "question", gap=f"{code} is not reported in the cubes")
        reasons.append(f"the cubes hold no {code} values (FinMax reports NOK and local currency only)")
        proxies.append(semantic.currency_proxy())

    # granularity -------------------------------------------------------------
    gs = semantic.grains(answers.get("granularity") or question)
    for g in gs:
        for m in found:
            ok, proxy = semantic.feasible(m.key, g)
            if not ok:
                reasons.append(f"{m.matched or m.key} is not available {g.matched}")
                proxies.append(proxy)
        if g.note:
            warnings.append(g.note)
    checklist["granularity"] = _item(", ".join(g.matched for g in gs) if gs else "total (no breakdown)",
                                     "question" if gs else "convention")
    for m in found:
        if m.note and m.note not in (checklist["measure"]["value"] or ""):
            warnings.append(m.note)

    # outcome -----------------------------------------------------------------
    for k, said in answers.items():
        checklist[k] = _item(checklist[k]["value"] or said, "user")
    questions = [checklist[k]["gap"] for k in CHECKLIST if checklist[k]["ask"]][:3]
    outcome = "infeasible" if reasons else "ambiguous" if questions else "clear"
    recipe = _recipe(question, found, gs, p, scen, ents, concept_ids) if outcome == "clear" else None
    return {
        "question": question,
        "suggested_outcome": outcome,
        "checklist": checklist,
        "questions": questions if outcome == "ambiguous" else [],
        "reason": "; ".join(reasons) or None,
        "proxy": "; ".join(dict.fromkeys(proxies)) or None,
        "warnings": list(dict.fromkeys(warnings)),
        "entities": ents,
        "measures": [{"key": m.key, "concept": m.concept, "kpi": m.kpi, "account": m.account,
                      "accounts": list(m.accounts), "definition": semantic.describe(m)} for m in found],
        "concepts": concept_ids,
        "closed_month": closed,
        "recipe": recipe,
        "next": NEXT["clear"] % recipe["name"] if recipe else NEXT["free" if outcome == "clear" else outcome],
    }


def _answered_measure(said: str) -> tuple[dict, list[semantic.Measure], list[str]]:
    """The user's answer to 'which measure': a measure key ('opex') or words naming one ('operating costs')."""
    m = semantic.measure(said) or next((x for x in semantic.measures(said) if not x.vague), None)
    if m:
        return _item(semantic.describe(m), "user"), [m], []
    hits = [c["id"] for s, c in concepts.scored(said) if s >= CONCEPT_MIN_SCORE][:3]
    return _item(said, "user"), [], hits


_TREND_RX = re.compile(r"\b(trend|trends|utvikling|month by month|måned for måned)\b", re.I)


def _recipe(question, found, gs, p, scen, ents, concept_ids) -> dict | None:
    """The standard recipe (+ params) for a clear question: the first matching rule (design D2), else None."""
    grain = {g.key for g in gs}
    ent = ents[0] if ents else {}
    member = ent.get("finmax") or "un.SY"
    month = periods.label(*p.end) if p.end else None
    m = found[0] if found else None
    accounts = list(m.accounts) if m and m.accounts else []
    if not accounts and m and m.concept:
        line = catalog.load().line_by_concept("FinMax", m.concept)
        accounts = [line.account] if line else []

    def drop_none(**kw):
        return {k: v for k, v in kw.items() if v not in (None, [])}

    if grain & {"supplier", "project"} and m and m.accounts:
        return {"name": "lines_ranking", "params": {"accounts": list(m.accounts), "entity": member, "period": p.label,
                                                    "by": "supplier", "projects": "project" in grain}}
    if "month" in grain or _TREND_RX.search(question):
        return {"name": "monthly_trend", "params": drop_none(entity=member, accounts=accounts, through=month)}
    if "cost_centre" in grain:
        return {"name": "cost_base", "params": drop_none(parent=member, accounts=accounts, period=month)}
    if {"budget", "forecast"} & set(scen) and ent.get("finplan"):
        return {"name": "actual_vs_plan", "params": drop_none(
            entity=ent["finplan"], month=periods._ABBR[p.end[1] - 1] if p.end else None,
            year=f"FY{p.end[0] % 100:02d}" if p.end else None)}
    if m and (m.concept or m.kpi):
        return {"name": "pnl", "params": drop_none(area=member, period=month)}
    if concept_ids:
        return {"name": "concept", "params": drop_none(concept=concept_ids[0], period=month, entity=member)}
    if accounts:
        return {"name": "cost_base", "params": drop_none(parent=member, accounts=accounts, period=month, top=10)}
    return None


