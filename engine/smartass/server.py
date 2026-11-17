# engine/smartass/server.py
"""smartass: the MCP server (stdio; started by run_server.py through child_server.py). Tools: fin_analyze (new
question or follow-up patch), fin_lookup, fin_status, fin_refresh, fin_drill, and xl_snapshot / xl_revert / xl_tidy
for edits to open workbooks made with Excel MCP. Error text (guard), Excel jobs (_call_user, progress notes) and the
lifespan (warm-up, closing the hidden Excel) come from core.serverkit."""
from __future__ import annotations

import asyncio
import contextlib
import datetime as dt

from mcp.server.mcpserver import Context, MCPServer

from . import __version__
from .knowledge import atlas
from .core import serverkit as kit
from .xl.host import InputError

from .data.adhoc import adhoc_fetch
from .data.provider import ComboProvider
from .followup import find_question, merge_patch, v2_questions
from .model.chain import restore_chain
from .model.lines import LineSet
from .model.spec import AnalysisSpec, LinesBlock
from .model.validate import GuardError, guard_chart, guard_key_points
from .pipeline import RestoreError, excel_job, read_stored, redraw_job, run_prepare

TOOL_NAMES = ("fin_analyze", "fin_lookup", "fin_status", "fin_refresh", "fin_drill", "fin_critic", "xl_snapshot",
              "xl_revert", "xl_tidy")

INSTRUCTIONS = """smartass (read-only towards Oracle EPM: FinMax/FinLite actuals, FinPlan plan/forecast).
1) fin_lookup first when a word is not an exact member: metrics, entities (business areas, FinPlan division totals),
KPIs, vague words (-> the options to ask), grains ('per customer' -> what is answerable), rules (kind 'rule').
Every hit carries caveats; 'via' = inherited from that member. Pass them on.
2) New question: fin_analyze(spec) with an AnalysisSpec v2 - plan (depth, purpose, hypotheses), queries (cube,
metrics, entity, periods e.g. 'closed-ytd', compare ly | budget | forecast; budget/forecast only in FinPlan),
blocks (table, kpi_tiles, trend, matrix, bridge, ranking, drivers, commentary). An EBITDA bridge uses every P&L
step: Revenue, COGS, Opex, Other items (FinMax) - otherwise 'Other / unexplained' is large.
3) Follow-up: fin_analyze(patch=..., q=n) - a JSON merge patch on the stored spec (lists like blocks are replaced
whole); default q = the active Q<n> Answer, else the latest. 4) Monthly packs: fin_refresh() re-runs every v2
answer in the workbook to the latest closed month (relative periods roll forward), no AI needed.
5) Quote numbers only from key_values; explain material lines from drivers evidence or label them a hypothesis.
Every answer has at least one chart block (kind by message); the sheet shows Key points first, the plan is not
drawn. Fix each reply.layout finding by patch.
6) Editing a workbook the user has open with Excel MCP: xl_snapshot first, then edit, xl_revert to undo; finish
with xl_tidy (zero BLOCK findings). 7) After the server reloads, Claude Code may keep an old tool schema: if a
parameter seems missing, ask the user to run /mcp and reconnect smartass. Repeated errors: fin_status(restart=True).
Ledger lines: a lines block (cell with All Accounts, All Entities, All Time Periods; Periodic; FinLite) or
fin_drill(cell) ranks suppliers, projects or accounts and ties them to the cube.
Critic: after a standard/deep answer run fin_critic(q) (brief + rubric), store the six
criterion scores with fin_critic(q, scores, notes) (the server computes the score), patch and re-score while it
returns passed false (at most 2 rounds); the method is in MCP prompts (plan,
critic, monthly-pack, variance-commentary, gl-extract, prodreg-input) and past answers in fin_lookup(kind='answer').
Sheet text is English; answer the user in their language. The user's own workbooks are never saved."""

mcp = MCPServer("smartass", instructions=INSTRUCTIONS, lifespan=kit.lifespan)


def _register(annotations):
    def deco(fn):
        mcp.tool(annotations=annotations)(kit.guard(fn))
        return fn
    return deco


def _resolver(forecast_year: int | None = None):
    """A KnowledgeResolver (knowledge.db, built when missing, + the atlas); the caller closes it."""
    from .knowledge.resolver import open_resolver
    year = forecast_year if forecast_year is not None else dt.date.today().year % 100
    return open_resolver(year)


ANSWER_KEY_VALUES = 5


def _past(row: dict) -> dict:
    kv = row.get("key_values") or {}
    return {"workbook": row["workbook"], "q": row["q"], "question": row["question"], "when": row["at"],
            "score": row.get("score"), "notes": row.get("notes") or "",
            "key_values": dict(list(kv.items())[:ANSWER_KEY_VALUES])}


def lookup(query: str, kind: str = "any", cube: str | None = None, limit: int = 15) -> dict:
    """{metrics: [id, name, aliases, description, differs, members, caveats ...], members: [... with caveats],
    concepts: [id, title, kind, blocks, notes ...], entities, kpis, vocab, rules};
    kind 'answer': {answers: past answers from answers.db (workbook, q, question, when, score, top key values)}."""
    if kind == "answer":
        from . import answers
        return {"answers": [_past(r) for r in answers.search(query, limit)]}
    from .knowledge.resolver import freshness_notes
    with contextlib.closing(_resolver()) as r:
        out = r.search(query, kind, cube, limit)
        warn = freshness_notes(r.db)
    return {**out, "warnings": warn} if warn else out


@_register(kit.LOCAL)
def fin_lookup(query: str, kind: str = "any", cube: str | None = None) -> dict:
    """Look up before you build. kind 'metric' (name or alias, e.g. 'driftskostnader': members per cube, how the
    cubes differ), 'member' (cube members, each with its meaning note and a behaviour
    card: time balance, account type, children ignored/subtracted in the consolidation), 'entity' (a business area by name or alias:
    FinMax/FinPlan members, FinPlan division total, successors of abolished areas), 'kpi' (ratios: numerator,
    denominator, basis), 'vocab' (vague words -> the options to offer; 'per customer' etc. -> what is answerable),
    'concept' (product concepts such as 'splitt'), 'rule' (the always-true cube rules; query '' lists them),
    'answer' (past answers). 'any' = metric + member + concept + entity + kpi + vocab. Every hit carries its caveats
    (reorgs, abolished units, account traps; 'via' names the member a caveat was inherited from): pass them on.
    cube: FinMax | FinLite | FinPlan to narrow. Local, instant."""
    return lookup(query, kind, cube)


def _provider():
    from .epm import svhttp

    from .data.provider import SvHttpProvider
    return SvHttpProvider(svhttp.shared())


def _clean(fn):
    """Refusals raised inside a host job are input errors, not 'Unexpected ... run fin_status(restart=True)'."""
    def job(h):
        try:
            return fn(h)
        except GuardError as e:
            raise InputError(f"{e.reason} - {e.fix}") from None
    return job


async def _user(fn, what: str):
    return await kit._call_user(_clean(fn), what=what)


async def _drill(cell: dict):
    """The FinLite ledger lines behind one cube cell (v1 drill: Smart View link, hidden browser, reconciliation)."""
    from .epm import drillstep
    return await drillstep.drill(cell, connection="FinLite")


async def _line_sets(spec: AnalysisSpec) -> tuple[dict, list[str]]:
    """Drill every lines block: {'b<i>': LineSet} and the drills' warnings. Project grouping marks each line's
    booked project (INITIATIVE) first. A drill with 0 lines still gives a LineSet (prepare refuses it, with the reason)."""
    from .model.lines import add_initiative
    blocks = [(i, b) for i, b in enumerate(spec.blocks) if isinstance(b, LinesBlock)]
    results = await asyncio.gather(*[_drill(b.cell) for _, b in blocks])      # independent cells: drilled concurrently
    out: dict[str, LineSet] = {}
    notes: list[str] = []
    for (i, b), res in zip(blocks, results):
        rows, cols = res.rows, list(res.columns)
        if b.group_by == "project":
            cols = add_initiative(rows, cols)
        out[f"b{i}"] = LineSet(f"b{i}", rows, cols, res.cell_value, reconciliation=res.reconciliation,
                                warnings=list(res.warnings), cell=b.cell)
        notes += res.warnings
    return out, notes


def _has_lines(raw: dict) -> bool:
    return any(isinstance(b, dict) and b.get("type") == "lines" for b in raw.get("blocks") or [])


async def _prepare(raw: dict, provider, spec: AnalysisSpec | None = None, target: str = "session",
                   complete: bool = False):
    """(Prepared, drill warnings). A stored or patched raw spec with lines blocks is re-drilled, so lines are fresh.
    Grids the HTTP route could not fetch are pulled as ad hoc grids in the user's Excel (one job), then the model is
    prepared again with those results fixed; this fallback runs at most once. complete (new questions and follow-up
    patches, not fin_refresh): the P&L chain is completed in the spec (pipeline.prepare), and the completed spec is
    what gets stored."""
    notes: list[str] = []
    kw = {"complete": True} if complete else {}
    lines: dict = {}
    if spec is not None or _has_lines(raw):
        lines, notes = await _line_sets(spec or AnalysisSpec.model_validate(raw))
    pre = await asyncio.to_thread(run_prepare, raw, provider, lines=lines, **kw)
    failed = getattr(pre, "failed", None)
    if failed:
        fixed = await _user(lambda h: adhoc_fetch(h, failed, target), "ad hoc fallback")
        pre = await asyncio.to_thread(run_prepare, raw, ComboProvider(provider, {**pre.fetched, **fixed}),
                                      lines=lines, **kw)
    return pre, notes


TEXT_FIX_PREFIX = "Text broken in EPM"


def _broken_texts(warning: str) -> list[str]:
    """'Text broken in EPM (letters lost at load): a; b - add corrections to ...' -> ['a', 'b']."""
    body = warning.split("): ", 1)[-1].split(" - add corrections", 1)[0]
    return [t.strip() for t in body.split(";") if t.strip()]


def _with_notes(out: dict, notes: list[str]) -> dict:
    """Drill notes join the warnings; the broken-text ones become reply['text_fixes'] (candidates for text_fixes.json)."""
    fixes = [t for n in notes if n.startswith(TEXT_FIX_PREFIX) for t in _broken_texts(n)]
    notes = [n for n in notes if not n.startswith(TEXT_FIX_PREFIX)]
    if notes:
        out["warnings"] = list(out.get("warnings", [])) + notes
    if fixes:
        out["text_fixes"] = list(dict.fromkeys(out.get("text_fixes", []) + fixes))
    return out


def _default_target(target: str) -> str:
    """Follow-ups, fin_refresh and fin_critic: target 'session' (the default) means the workbook of the latest
    stored answer (answers.db), so a follow-up lands where the user's last answer is (an open 'Book1', the active
    workbook) and not in an older session file. Explicit targets win; no stored answer (or the latest one in the
    session workbook itself) keeps 'session'."""
    if target != "session":
        return target
    from . import answers
    from .core import session
    try:
        row = answers.latest()
    except Exception:   # noqa: BLE001 - answer memory unreadable: the session workbook, as before
        return target
    wb = (row or {}).get("workbook")
    if not wb:
        return target
    own = session.workbook()
    if own and answers._key(own) == answers._key(wb):
        return target                                # the session file stays 'ours' (saved after each answer)
    return wb


async def _new_question(spec: AnalysisSpec, target: str) -> dict:
    guard_chart(spec)                  # new questions only: every other path gets an implicit chart
    guard_key_points(spec)             # new questions only: a stored spec keeps its first 5 (a warning)
    pre, notes = await _prepare(spec.model_dump(), _provider(), spec, target, complete=True)
    return _with_notes(await _user(lambda h: excel_job(h, pre, target), "fin_analyze"), notes)


@_register(kit.WRITE)
async def fin_analyze(spec: AnalysisSpec | None = None, patch: dict | None = None, q: int | None = None,
                      target: str = "session", ctx: Context | None = None) -> dict:
    """Run one analysis. New question: spec = AnalysisSpec v2. Follow-up: patch = JSON merge patch on the stored
    spec of Q<q> (default: the active Q<n> Answer, else the latest); the same sheets are rebuilt in place.
    Fetches over Smart View HTTP, computes, and builds Q<n> Answer (live formulas on table Data_Q<n>) + Q<n> Data in
    the session workbook (target 'session', 'new' or 'active', an absolute path or the name of an open workbook).
    A follow-up's target default: the workbook of your latest answer. Returns key_values (quote only these),
    material lines, findings, checks, warnings, layout (check findings, each with its fix: patch them) and charts
    (the chart titles drawn)."""
    if (spec is None) == (patch is None):
        raise GuardError("give spec (a new question) or patch (a follow-up), not both",
                         "pass spec=... for a new question, or patch=... (and q) to change an existing answer")
    if spec is not None:
        return await _new_question(spec, target)
    provider = _provider()
    target = _default_target(target)
    stored, active, path = await _user(lambda h: read_stored(h, target), "fin_analyze follow-up")
    q2, raw, owned = find_question(stored, q, active)
    pre, notes = await _prepare(restore_chain(merge_patch(raw, patch)), provider, target=path, complete=True)
    return _with_notes(await _user(lambda h: redraw_job(h, pre, q2, owned, path), "fin_analyze follow-up"), notes)


@_register(kit.WRITE)
async def fin_drill(cell: dict, group_by: str = "supplier", top: int = 10, target: str = "session",
                    ctx: Context | None = None) -> dict:
    """Ledger lines behind one cube cell (FinLite, Periodic): ranked by supplier, project or account, with Other,
    lines without the group, the lines total, the cube value and the gap. cell: {"All Accounts": "ac.C60",
    "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"} (a parent period such as 'Yr 26' works).
    Same sheets and checks as a lines block in fin_analyze."""
    spec = AnalysisSpec.model_validate({
        "version": 2, "question": f"Ledger lines behind {' / '.join(cell.values())}", "plan": {"depth": "quick"},
        "queries": [], "blocks": [{"type": "lines", "cell": cell, "group_by": group_by, "top": top}]})
    return await _new_question(spec, target)


@_register(kit.WRITE)
async def fin_refresh(target: str = "session", ctx: Context | None = None) -> dict:
    """Re-run every v2 answer in the workbook with its stored spec (no patch): relative periods such as
    'closed-ytd' roll forward to the latest closed month. A failing question is reported and does not stop the rest.
    target default: the workbook of your latest answer."""
    provider = _provider()
    target = _default_target(target)
    stored, _, path = await _user(lambda h: read_stored(h, target), "fin_refresh")
    done, workbook, broken = [], None, None
    for q in v2_questions(stored):
        if broken is not None:                  # a failed restore may have left the workbook damaged: draw nothing more
            done.append({"q": q, "ok": False, "skipped": True,
                         "error": f"skipped: the restore of Q{broken} failed - check the workbook first"})
            continue
        try:
            pre, notes = await _prepare(stored[q]["spec"]["v2"], provider, target=path)
            out = _with_notes(await _user(lambda h, q=q, pre=pre: redraw_job(h, pre, q, stored[q].get("owned"), path),
                                          "fin_refresh"), notes)
            workbook = out.get("workbook", workbook)
            done.append({"q": q, "ok": True, "closed": out.get("closed"),
                         **({"critic_brief": out["critic_brief"]} if "critic_brief" in out else {})})
        except RestoreError as e:
            broken = q
            done.append({"q": q, "ok": False, "error": str(e)})
        except Exception as e:   # noqa: BLE001 - one question never stops the pack
            done.append({"q": q, "ok": False, "error": str(e)})
    return {"refreshed": done, "workbook": workbook}


def _rubric() -> str:
    from .core import config
    return (config.KNOWLEDGE_DIR / "agent" / "critic.md").read_text(encoding="utf-8")


async def _answer_for(q: int | None, target: str = "session") -> tuple[str, int]:
    """(workbook, q) of the answer the critic reviews: in the target workbook (default: the workbook of the latest
    answer), found as a follow-up finds it (the explicit q, else the active Q<n> Answer, else the latest v2 answer).
    Without a session workbook, or when the defaulted workbook is not open any more: the latest stored run (in that
    workbook when known). Answers are keyed by the path, or by the name of an unsaved book, as answers.store keys."""
    from . import answers
    explicit = target != "session"
    target = _default_target(target)
    try:
        stored, active, path = await _user(lambda h: read_stored(h, target), "fin_critic")
    except (GuardError, InputError) as e:
        if isinstance(e, InputError) and (explicit or target == "session"):
            raise
        row = answers.latest(q, None if target == "session" else target)
        if row is None:
            raise GuardError("no stored answer to review" + (f" for Q{q}" if q is not None else ""),
                             "ask a question with fin_analyze(spec) first") from None
        return row["workbook"], row["q"]
    q2, _, _ = find_question(stored, q, active)
    return path, q2


def _critic_score(scores: list[int]) -> tuple[int, float]:
    """(score, mean) of the six rubric scores: the mean rounded half up, so 4.5 is 5 and 3.5 is 4."""
    if len(scores) != 6:
        raise GuardError(f"{len(scores)} scores given, the rubric has 6 criteria",
                         "pass one score per criterion, in the rubric's order 1-6")
    if any(not 1 <= s <= 5 for s in scores):
        raise GuardError(f"scores {scores} are outside 1-5", "score each rubric criterion 1-5")
    mean = sum(scores) / 6
    return int(mean + 0.5), round(mean, 2)


@_register(kit.LOCAL)
async def fin_critic(q: int | None = None, scores: list[int] | None = None, notes: str = "",
                     target: str = "session", ctx: Context | None = None) -> dict:
    """The critic pass for a standard or deep answer or a pack question. target: the workbook of the answer
    (default: the workbook of your latest answer; 'active', an absolute path or an open workbook's name). Without
    scores: {brief, rubric} of the stored answer Q<q> (default: the active Q<n> Answer, else the latest) - score it
    on the rubric's 6 criteria. With scores (six integers 1-5, criteria 1-6 in order) and notes (the weakest
    criterion and why): the server computes the score (the mean rounded half up), stores it with the answer
    (answer memory, fin_lookup kind 'answer') and returns it with passed (score >= 4). passed false: fix by
    fin_analyze(patch=...) and score again (at most 2 rounds). Never compute the score yourself."""
    from . import answers
    rated = _critic_score(list(scores)) if scores is not None else None
    workbook, q2 = await _answer_for(q, target)
    if rated is not None:
        score, mean = rated
        stored_notes = (f"Scores {' '.join(f'{i}={s}' for i, s in enumerate(scores, 1))} (mean {mean:.2f})."
                        + (f" {notes.strip()}" if notes.strip() else ""))
        if not answers.set_score(workbook, q2, float(score), stored_notes):
            raise GuardError(f"Q{q2} has no stored answer to score", "re-run it (fin_analyze patch or fin_refresh) "
                             "so it is stored, then score it")
        return {"stored": True, "q": q2, "score": score, "passed": score >= 4, "mean": mean, "scores": list(scores)}
    row = answers.get(workbook, q2)
    if row is None:
        raise GuardError(f"Q{q2} has no stored answer to review", "re-run it (fin_analyze patch or fin_refresh) "
                         "so it is stored with its brief")
    return {"q": q2, "workbook": workbook, "question": row["question"], "brief": row.get("critic_brief"),
            "score": row.get("score"), "rubric": _rubric(),
            **({} if row.get("critic_brief") else {"note": "no brief: the answer is depth 'quick' or older than "
                                                   "the critic; patch plan.depth to 'standard' to get one"})}


@_register(kit.ADMIN)
async def fin_status(restart: bool = False) -> dict:
    """Health: version, atlas age, knowledge build (metrics, facts, drafts awaiting approval), where answer files go,
    and the doctor checks (config, Excel, Smart View, VBA access, EPM, drill browser, atlas, knowledge, output
    folder, HTTP API, update; only problems listed).
    restart=True after repeated errors or to load new code: the relay starts a fresh server process (its hidden
    Excel closes); run without the relay, only this server's own Excel is closed (it restarts on the next pull).
    A stale atlas: run `smartass atlas build` in a terminal."""
    import os

    from .core import config
    out: dict = {"version": __version__}
    if restart:
        if os.environ.get("SVKIT_PROXY"):
            out["restarted"] = "relay"             # run_server.py replaced the child before forwarding this call
        else:
            old, kit._host = kit._host, None
            if old is not None:
                await asyncio.to_thread(old.shutdown)
            out["restarted"] = "excel"
    with contextlib.closing(atlas.connect()) as db:
        st = atlas.status(db)
    from .knowledge.build import summary
    from .knowledge.refresh import due
    kn = summary(config.KNOWLEDGE_DB)
    refresh = due(config.HOME, config.KNOWLEDGE_DB)
    from .ops import doctor
    checks = await asyncio.to_thread(doctor.run)         # network checks (EPM 5 s, HTTP 2 s) off the event loop
    return {**out, "behind_relay": bool(os.environ.get("SVKIT_PROXY")), "analyser": str(config.ANALYSER_DIR),
            "atlas": {"oldest_age_days": st["oldest_age_days"], "stale": st["stale"],
                      "dimensions_built": len(st["built"])},
            "metrics": kn.get("metrics", 0),
            "knowledge_layer": {k: kn.get(k) for k in ("rules", "entities", "kpis", "vocab", "notes",
                                                       "atlas_built_at")},
            "knowledge": {"built_at": kn.get("built_at"), "metrics": kn.get("metrics", 0),
                          "facts": kn.get("facts", 0), "drafts": kn.get("drafts", 0),
                          "refresh_month": refresh["refresh_month"], "refresh_due": refresh["due"],
                          **({"note": f"knowledge refresh due (closed {refresh['closed_seen']}, last refresh "
                                      f"{refresh['refresh_month'] or 'never'}): run `smartass knowledge refresh` "
                                      "in a terminal, then `smartass knowledge approve`"} if refresh["due"] else {})},
            "doctor": {"overall": doctor.overall(checks),
                       "problems": [c.__dict__ for c in checks if c.status != "OK"]}}


@_register(kit.WRITE)
async def xl_snapshot(workbook: str | None = None, sheet: str | None = None, range: str | None = None,
                      allow_archive: bool = False) -> dict:
    """Save the current state of a range (default: the sheet's used range) of an open workbook, changing nothing.
    Call before any Excel MCP edit to a workbook the user has open; xl_revert(change_id) undoes the edit.
    allow_archive=true only when the user explicitly asked to edit a file in the Produktregnskapet archive."""
    from .workbench import snapshot
    return await _user(lambda h: snapshot(workbook, sheet, range, allow_archive=allow_archive), "xl_snapshot")


@_register(kit.WRITE)
async def xl_revert(change_id: str | None = None, workbook: str | None = None, force: bool = False) -> dict:
    """Put a range back as xl_snapshot (or a v1 xl_* edit) recorded it. Default: the latest change. Not saved."""
    from .workbench import revert
    return await _user(lambda h: revert(change_id, workbook, force), "xl_revert")


@_register(kit.WRITE)
async def xl_tidy(workbook: str | None = None, sheet: str | None = None, rows: dict | None = None,
                  last_col: str | None = None, allow_archive: bool = False) -> dict:
    """TN report style (when rows + last_col are given) and the tn-formelstil formula check; zero BLOCK findings
    before work is done. rows: {"1": "title", "4": "header", "5-12": "line", "13": "sum", ...}; last_col e.g. "F".
    Kinds: title, subtitle, section, text, header, line, sum, grand, note, blank (subtotal, total, heading accepted).
    allow_archive=true only when the user explicitly asked to edit a file in the Produktregnskapet archive."""
    from .workbench import tidy
    return await _user(lambda h: tidy(workbook, sheet, rows, last_col, allow_archive=allow_archive), "xl_tidy")


# ---------------------------------------------------------------- playbooks: MCP prompts and resources (sub-spec 2 §6)
PLAYBOOKS = {"monthly-pack": "Monthly pack: build once with relative periods, refresh, critic, commentary patch.",
             "variance-commentary": "Variance commentary: drivers from data, caveats, FinMax vs FinPlan.",
             "gl-extract": "GL extract: ledger lines per account and cost centre, reconciled exactly to the cube.",
             "prodreg-input": "Produktregnskapet input: drill per division node, sum to YTD, support workbook rule."}
GUIDES = {"plan": ("agent/plan.md", "How to write the plan of an AnalysisSpec v2: purpose, hypotheses, blocks."),
          "critic": ("agent/critic.md", "The critic rubric: score a standard or deep answer 1-5 on 6 criteria."),
          **{n: (f"playbooks/{n}.md", d) for n, d in PLAYBOOKS.items()}}
FINANCE_RULES = "## 6. Finance rules"


def _knowledge_text(rel: str) -> str:
    from .core import config
    return (config.KNOWLEDGE_DIR / rel).read_text(encoding="utf-8")


def _guide_prompt(name: str, rel: str, desc: str):
    def render() -> str:
        return (f"smartass {name} ({desc.split(':')[0].lower()}): apply this with the fin_* tools; quote numbers "
                f"only from key_values.\n\n" + _knowledge_text(rel))
    render.__name__ = name.replace("-", "_")
    mcp.prompt(name=name, description=desc)(render)


def _playbook_resource(name: str, desc: str):
    def read() -> str:
        return _knowledge_text(f"playbooks/{name}.md")
    read.__name__ = f"playbook_{name.replace('-', '_')}"
    mcp.resource(f"smartass://playbooks/{name}", name=name, description=desc, mime_type="text/markdown")(read)


for _n, (_rel, _d) in GUIDES.items():
    _guide_prompt(_n, _rel, _d)
for _n, _d in PLAYBOOKS.items():
    _playbook_resource(_n, _d)


SKILL_REFS = ("spec", "charts", "concepts", "cubes", "drill", "operations")


def _skill_dirs() -> list:
    """Where agents keep their copy of the smartass skill (Claude Code, Copilot/agents). Some clients (Copilot) read a
    skill's relative link 'reference/spec.md' as an MCP resource of this server, by its file:// URI."""
    from pathlib import Path
    return [Path.home() / d / "skills" / "smartass" for d in (".agents", ".claude", ".copilot")]


def _reference_resource(name: str) -> None:
    def read() -> str:
        from .core import config
        return (config.REPO_DIR / "skills" / "smartass" / "reference" / f"{name}.md").read_text(encoding="utf-8")
    read.__name__ = f"reference_{name}"
    desc = f"smartass skill reference: {name}.md"
    mcp.resource(f"smartass://reference/{name}.md", name=f"reference/{name}.md", description=desc,
                 mime_type="text/markdown")(read)
    for d in _skill_dirs():
        mcp.resource((d / "reference" / f"{name}.md").as_uri(), name=f"reference/{name}.md ({d.parent.parent.name})",
                     description=desc, mime_type="text/markdown")(read)


for _n in SKILL_REFS:
    _reference_resource(_n)


@mcp.resource("smartass://conventions", name="conventions", mime_type="text/markdown",
              description="Finance conventions: cube choice, NOK format, signs, variance, P&L chain, members.")
def conventions() -> str:
    """The skill's finance rules section (one source: skills/smartass/SKILL.md)."""
    from .core import config
    skill = (config.REPO_DIR / "skills" / "smartass" / "SKILL.md").read_text(encoding="utf-8")
    body = skill.split(FINANCE_RULES, 1)[1].split("\n## ", 1)[0]
    return "# smartass finance conventions\n" + body


def main():
    mcp.run()
