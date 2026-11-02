"""One question through the new core (design §3): scope -> pull -> answer -> check. The scope gate's verdict
arrives in the answer spec (the agent applies the checklist); ambiguous questions stop in tools_core before any
Excel work. Walking skeleton: new questions only (follow-ups are phase 5)."""
from __future__ import annotations

import copy
from dataclasses import dataclass

from pydantic import ValidationError

from .. import answer as answer_mod
from .. import catalog, specs, userexcel
from ..host import InputError, rewrap
from ..timing import stage
from . import check, commentary, contract, progress, source


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
    lead = [f"Question: {spec.question}"] + [f"Assumption: {a}" for a in spec.assumptions]
    if spec.gate.outcome == "infeasible":
        lead.append(f"Not answerable as asked ({spec.gate.reason}); answered instead: {spec.gate.proxy}")
    lay["notes"] = lead + list(lay.get("notes") or [])
    drivers = commentary.sheet_lines([d.model_dump() for d in spec.commentary.drivers])
    if drivers:
        lay["key_points"] = list(lay.get("key_points") or []) + drivers
    return lay


def run(h, raw: dict, target: str = "session", allow_archive: bool = False) -> Run:
    spec = parse_spec(raw)
    if spec.gate.outcome == "ambiguous":
        raise InputError("An ambiguous question pulls nothing: ask its questions first.")
    with stage("open target"):
        t = userexcel.resolve_target(h, target, allow_archive, title=spec.question[:50])
        q = userexcel.next_question(t)
    sources: list[source.Source] = []
    try:
        with stage("answer skeleton"):
            progress.show(t, q, spec, f"fetching {len(spec.pulls)} source grid(s) ...")
        for i, pin in enumerate(spec.pulls):
            progress.set_status(t, q, f"fetching {pin.connection} '{pin.id}' ({i + 1}/{len(spec.pulls)}) ...")
            sources.append(source.pull(h, t, q, i, pin))
        progress.set_status(t, q, "building the answer ...")
        ids = {pin.id: s.pull.sheet for pin, s in zip(spec.pulls, sources)}
        sess = {s.pull.sheet: s.pull for s in sources}
        with stage("prepare spec"):
            prepared = specs.prepare(answer_layout(raw["layout"], spec), ids, sess)
        with stage("answer"):
            res = answer_mod.build(t, q, prepared, sess, suffix=progress.ANSWER)
    except Exception as e:
        removed = userexcel.rollback(t)
        raise rewrap(e, f"{e} (the answer was not completed; removed the partial sheets: "
                     f"{', '.join(removed) or 'none'})") from None
    flagged = commentary.material(prepared, res.get("values") or {}, catalog.load())
    review = check.review(res, sources, extra=commentary.findings(
        flagged, [d.model_dump() for d in spec.commentary.drivers]))
    try:
        userexcel.save_spec(t, q, raw, ids)
    except Exception:   # noqa: BLE001 - only needed for follow-ups
        pass
    return Run(t, q, spec, prepared, res, sources, review, flagged)
