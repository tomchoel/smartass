"""The monthly knowledge refresh (sub-spec 2 §8) and the approval of drafts.

`smartass knowledge refresh`: rebuild the atlas (live EPM, optional), diff it against the snapshot of the last refresh,
draft facts for the moved or removed members that a metric or fact references, rebuild knowledge.db, run the offline
test set, write the impact report and save the new snapshot. Nothing is approved here: drafts reach answers only as
the "unapproved knowledge change" caveat until `smartass knowledge approve` flips them.

Home folder layout: knowledge/atlas_snapshot.json (member -> [parent, label] per cube and dimension),
knowledge/reports/<yyyy-mm>.md, knowledge/state.json (the latest closed month an answer saw). The refresh month is
in knowledge.db meta (`refresh_month`, kept by every rebuild)."""
from __future__ import annotations

import contextlib
import datetime as dt
import json
import re
import sqlite3
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from ..model.periods import Month, label
from .build import SEARCH_DIMENSIONS, BuildReport, _atlas, build, summary
from .semantic import Fact, Metric, load_all, load_semantic, to_toml

SNAPSHOT = "atlas_snapshot.json"
STATE = "state.json"


def _ym(m: Month) -> str:
    return f"{m[0]:04d}-{m[1]:02d}"


# ---- snapshot and diff

def snapshot(db: sqlite3.Connection) -> dict:
    """{cube: {dimension: {member: [parent, label]}}} of the account and entity trees (what metrics and facts name).
    A shared member keeps its shallowest parent."""
    marks = ",".join("?" * len(SEARCH_DIMENSIONS))
    out: dict = {}
    for cube, dim, name, lab, parent in db.execute(
            f"SELECT connection, dimension, name, label, parent FROM member WHERE dimension IN ({marks}) "
            "ORDER BY connection, dimension, depth DESC, ord DESC", SEARCH_DIMENSIONS):
        out.setdefault(cube, {}).setdefault(dim, {})[name] = [parent, lab]        # the shallowest row is written last
    return out


@dataclass
class AtlasDiff:
    added: list[dict] = field(default_factory=list)
    removed: list[dict] = field(default_factory=list)
    moved: list[dict] = field(default_factory=list)          # parent changed: from / to = parents
    relabelled: list[dict] = field(default_factory=list)     # from / to = labels

    def counts(self) -> dict[str, int]:
        return {k: len(getattr(self, k)) for k in ("added", "removed", "moved", "relabelled")}


def diff(old: dict, new: dict) -> AtlasDiff:
    d = AtlasDiff()
    for cube in sorted(set(old) | set(new)):
        for dim in sorted(set(old.get(cube, {})) | set(new.get(cube, {}))):
            a, b = old.get(cube, {}).get(dim, {}), new.get(cube, {}).get(dim, {})
            where = {"cube": cube, "dimension": dim}
            for m in sorted(set(b) - set(a)):
                d.added.append({**where, "member": m, "parent": b[m][0]})
            for m in sorted(set(a) - set(b)):
                d.removed.append({**where, "member": m, "parent": a[m][0]})
            for m in sorted(set(a) & set(b)):
                if a[m][0] != b[m][0]:
                    d.moved.append({**where, "member": m, "from": a[m][0], "to": b[m][0]})
                if a[m][1] != b[m][1]:
                    d.relabelled.append({**where, "member": m, "from": a[m][1], "to": b[m][1]})
    return d


# ---- drafts

def _references(metrics: list[Metric], facts: list[Fact]) -> tuple[dict, dict]:
    """(metric refs, fact refs): (cube, member casefolded) -> ids that name the member."""
    mref: dict[tuple[str, str], list[str]] = {}
    fref: dict[tuple[str, str], list[str]] = {}
    for m in metrics:
        for cube, ms in m.members().items():
            for x in ms:
                mref.setdefault((cube, x.casefold()), []).append(m.id)
    for f in facts:
        for cube, ms in f.scope_members.items():
            for x in ms:
                fref.setdefault((cube, x.casefold()), []).append(f.id)
    return mref, fref


def _fact_id(month: Month, member: str) -> str:
    return f"refresh_{month[0]:04d}_{month[1]:02d}_" + re.sub(r"[^a-z0-9]+", "_", member.casefold()).strip("_")


def draft_facts(d: AtlasDiff, metrics: list[Metric], facts: list[Fact], month: Month) -> tuple[list[Fact], list[str]]:
    """A draft fact per moved or removed member that a metric or fact references (an added child of a metric's
    member needs nothing: metrics use parents). A removed metric member also gets a report note: the build fails
    until the metric is fixed. Returns (drafts, notes)."""
    mref, fref = _references(metrics, facts)
    drafts: dict[str, Fact] = {}
    notes: list[str] = []
    source = f"smartass knowledge refresh {_ym(month)} (atlas diff)"
    for kind, items in (("moved", d.moved), ("removed", d.removed)):
        for x in items:
            key = (x["cube"], x["member"].casefold())
            users_m, users_f = mref.get(key, []), fref.get(key, [])
            if not users_m and not users_f:
                continue
            if kind == "moved":
                text = f"{x['member']} moved from {x['from']} to {x['to']} in {label(month)}."
                severity = "info"
            else:
                text = f"{x['member']} was removed from {x['cube']} / {x['dimension']} in {label(month)}."
                severity = "warn"
            if users_m:
                text += f" Used by metric(s) {', '.join(sorted(set(users_m)))}."
            fid = _fact_id(month, x["member"])
            drafts.setdefault(fid, Fact(id=fid, text=text, scope_members={x["cube"]: [x["member"]]},
                                        scope_metrics=sorted(set(users_m)), valid_from=_ym(month),
                                        severity=severity, source=source, status="draft"))
            if kind == "removed" and users_m:
                notes.append(f"metric {', '.join(sorted(set(users_m)))}: {x['cube']} member {x['member']} was "
                             "removed from the atlas - the build fails until the metric file is fixed")
    return list(drafts.values()), notes


def write_drafts(semantic_dir: Path, drafts: list[Fact]) -> list[Fact]:
    """Draft TOML files into semantic/facts; an existing file (an earlier refresh, or already approved) is kept."""
    out = []
    folder = Path(semantic_dir) / "facts"
    folder.mkdir(parents=True, exist_ok=True)
    for f in drafts:
        p = folder / f"{f.id}.toml"
        if p.exists():
            continue
        p.write_text(to_toml(f), encoding="utf-8")
        out.append(f)
    return out


# ---- state: closed month seen, refresh due

def _state(home: Path) -> dict:
    p = Path(home) / "knowledge" / STATE
    with contextlib.suppress(OSError, ValueError):
        return json.loads(p.read_text(encoding="utf-8"))
    return {}


def _save_state(home: Path, **kw) -> None:
    p = Path(home) / "knowledge" / STATE
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps({**_state(home), **kw}, indent=1), encoding="utf-8")


def note_closed(home: Path, month: Month) -> None:
    """Remember the latest closed month an answer saw (never moves back); written only when it changes."""
    ym = _ym(month)
    if (_state(home).get("closed_seen") or "") < ym:
        _save_state(home, closed_seen=ym)


def due(home: Path, knowledge_db: Path) -> dict:
    """{refresh_month, closed_seen, due}: due when a closed month newer than the last refresh has been seen."""
    rm = summary(knowledge_db).get("refresh_month") or None
    seen = _state(home).get("closed_seen") or None
    return {"refresh_month": rm, "closed_seen": seen, "due": bool(seen and (rm is None or seen > rm))}


# ---- refresh

@dataclass
class RefreshResult:
    month: Month
    report: Path
    diff: AtlasDiff
    drafts: list[Fact]
    notes: list[str]
    build: BuildReport
    testset: list[dict]
    atlas: list[dict] | None = None
    first: bool = False

    @property
    def ok(self) -> bool:
        return self.build.ok and all(r.get("ok") is not False for r in self.testset)


def _affected(d: AtlasDiff, metrics: list[Metric]) -> list[str]:
    mref, _ = _references(metrics, [])
    hits = set()
    for x in [*d.moved, *d.removed, *d.relabelled]:
        hits.update(mref.get((x["cube"], x["member"].casefold()), []))
    return sorted(hits)


def _report(r: RefreshResult, metrics: list[Metric]) -> str:
    ok = sum(1 for t in r.testset if t.get("ok") is True)
    lines = [f"# Knowledge refresh {_ym(r.month)}", "",
             f"Run {dt.datetime.now():%Y-%m-%d %H:%M}. Atlas: "
             + ("not rebuilt (--no-atlas)" if r.atlas is None else
                f"rebuilt, {sum(1 for a in r.atlas if not a.get('error'))} dimension(s), "
                f"{sum(1 for a in r.atlas if a.get('error'))} failed"), ""]
    lines += ["## Atlas diff", ""]
    if r.first:
        lines += ["No snapshot from an earlier refresh: this first refresh saves the baseline.", ""]
    else:
        lines += [f"- {k}: {v}" for k, v in r.diff.counts().items()] + [""]
        for kind in ("moved", "removed", "relabelled"):
            for x in getattr(r.diff, kind)[:50]:
                extra = f" ({x['from']} -> {x['to']})" if "from" in x else ""
                lines.append(f"- {kind}: {x['cube']} / {x['dimension']} {x['member']}{extra}")
        lines.append("")
    aff = _affected(r.diff, metrics)
    lines += ["## Metrics affected", "", (", ".join(aff) if aff else "none"), ""]
    lines += ["## Drafts written (status draft; approve with `smartass knowledge approve <id|all>`)", ""]
    lines += [f"- {f.id}: {f.text}" for f in r.drafts] or ["none"]
    lines += [""] + [f"- NOTE {n}" for n in r.notes] + ([""] if r.notes else [])
    lines += ["## Build", "", "```", r.build.to_text(), "```", ""]
    lines += ["## Test set (offline)", "", f"{ok}/{len(r.testset)} ok"]
    lines += [f"- {t['id']}: {'ok' if t.get('ok') else 'FAIL'} {'; '.join(t.get('diffs') or [])}".rstrip()
              for t in r.testset if t.get("ok") is not True]
    return "\n".join(lines) + "\n"


def refresh(semantic_dir: Path, atlas_db: Path, knowledge_db: Path, home: Path, month: Month,
            rebuild_atlas: Callable[[], list[dict]] | None = None,
            run_testset: Callable[[], list[dict]] | None = None) -> RefreshResult:
    """One monthly refresh. rebuild_atlas: None skips it (--no-atlas); the CLI passes the `smartass atlas build`
    function. run_testset: the offline test set (runner.run_offline over every case)."""
    atlas_done = rebuild_atlas() if rebuild_atlas is not None else None
    folder = Path(home) / "knowledge"
    snap_path = folder / SNAPSHOT
    with contextlib.closing(_atlas(atlas_db)) as db:
        new = snapshot(db)
    old = None
    with contextlib.suppress(OSError, ValueError):
        old = json.loads(snap_path.read_text(encoding="utf-8"))
    metrics, facts = load_semantic(semantic_dir)
    d = diff(old, new) if old is not None else AtlasDiff()
    drafts, notes = draft_facts(d, metrics, facts, month)
    written = write_drafts(semantic_dir, drafts)
    report = build(semantic_dir, atlas_db, knowledge_db, extra_meta={"refresh_month": _ym(month)})
    tests = run_testset() if run_testset is not None else []
    res = RefreshResult(month, folder / "reports" / f"{_ym(month)}.md", d, written, notes, report, tests,
                        atlas_done, first=old is None)
    res.report.parent.mkdir(parents=True, exist_ok=True)
    res.report.write_text(_report(res, metrics), encoding="utf-8")
    snap_path.write_text(json.dumps(new, ensure_ascii=False), encoding="utf-8")
    _save_state(home, last_report=str(res.report), last_refresh=_ym(month))
    return res


# ---- approve and status

def _items(semantic_dir: Path) -> dict:
    """{id: (item, folder)} of every kind with an id (ids are unique across kinds)."""
    k = load_all(semantic_dir)
    kinds = {"metrics": k.metrics, "facts": k.facts, "concepts": k.concepts, "rules": k.rules,
             "entities": k.entities, "kpis": k.kpis, "vocab": k.vocab}
    return {x.id: (x, folder) for folder, xs in kinds.items() for x in xs}


def approve(semantic_dir: Path, atlas_db: Path, knowledge_db: Path, ids: list[str], by: str,
            on: str | None = None) -> tuple[list[str], BuildReport]:
    """Set status approved and approved = {by, on} on the named drafts ('all' = every draft), then rebuild.
    An unknown id raises KeyError before anything is written."""
    on = on or dt.date.today().isoformat()
    items = _items(semantic_dir)
    if "all" in ids:
        chosen = [i for i, (x, _) in items.items() if x.status == "draft"]
    else:
        unknown = [i for i in ids if i not in items]
        if unknown:
            raise KeyError(f"unknown id(s): {', '.join(unknown)}")
        chosen = [i for i in ids if items[i][0].status == "draft"]
    for i in chosen:
        x, kind = items[i]
        y = x.model_copy(update={"status": "approved", "approved": {"by": by, "on": on}})
        (Path(semantic_dir) / kind / f"{i}.toml").write_text(to_toml(y), encoding="utf-8")
    return sorted(chosen), build(semantic_dir, atlas_db, knowledge_db)


def status(semantic_dir: Path, knowledge_db: Path, home: Path) -> dict:
    st = _state(home)
    return {"knowledge_db": str(knowledge_db), "meta": summary(knowledge_db),
            "drafts": sorted(i for i, (x, _kind) in _items(semantic_dir).items() if x.status == "draft"),
            "last_report": st.get("last_report"), **due(home, knowledge_db)}
