"""Record, approve and run the test set. record needs EPM (live); approve and run are offline. A case's grids and
lines are stored in <HOME>/<id>/ (grids.json, lines.json); candidate.json holds the closed months and the key values
of the last recording, golden.json the approved copy. Relative periods make a golden belong to the closed month it
was recorded in, so a run replays that same recording."""
from __future__ import annotations

import contextlib
import datetime as dt
import json
import shutil
from pathlib import Path

from ..core import config as v1config

from ..data.recorder import RecordingProvider, ReplayProvider, load_lines, save_lines
from ..pipeline import prepare

HOME = Path(v1config.SVKIT) / "v2" / "testset"
CASES = Path(__file__).parent / "cases"


def load_cases() -> list[dict]:
    return [json.loads(p.read_text(encoding="utf-8")) for p in sorted(CASES.glob("*.json"))]


def _resolver_for(closed: dict):
    """An AtlasResolver on the local atlas (as run_prepare builds it); its connection is closed by _close."""
    from ..knowledge import atlas

    from ..knowledge.atlas_resolver import AtlasResolver
    from ..knowledge.interim import catalog_metrics
    db = atlas.connect()
    r = AtlasResolver(db, catalog_metrics(), forecast_year=closed["FinPlan"][0] % 100)
    r._db = db
    return r


def _close(resolver) -> None:
    db = getattr(resolver, "_db", None)
    if db is not None:
        with contextlib.suppress(Exception):
            db.close()


def _prepare(spec: dict, provider, closed: dict, lines: dict):
    r = _resolver_for(closed)
    try:
        return prepare(spec, r, provider, closed=closed, lines=lines)
    finally:
        _close(r)


def record(case: dict, provider, lines: dict) -> dict:
    """Run the case on the provider (live), store grids + lines + the candidate key values; returns the candidate."""
    rec = RecordingProvider(provider)
    closed = rec.closed_months()
    pre = _prepare(case["spec"], rec, closed, lines)
    d = HOME / case["id"]
    rec.save(d / "grids.json")
    save_lines(d / "lines.json", lines)
    cand = {"closed": {c: list(m) for c, m in closed.items()}, "key_values": pre.result.key_values,
            "recorded_at": dt.datetime.now().isoformat(timespec="seconds")}
    (d / "candidate.json").write_text(json.dumps(cand, ensure_ascii=False, indent=1), encoding="utf-8")
    return cand


def approve(case_id: str) -> Path:
    d = HOME / case_id
    if not (d / "candidate.json").exists():
        raise FileNotFoundError(f"{case_id}: no candidate - record it first")
    shutil.copyfile(d / "candidate.json", d / "golden.json")
    return d / "golden.json"


def _tol(a: float, tol_abs: float, tol_frac: float) -> float:
    return tol_abs if abs(a) >= 1 else tol_frac


def compare(golden: dict, new: dict, tol_abs: float = 1.0, tol_frac: float = 1e-6) -> list[str]:
    diffs: list[str] = []
    for line, heads in golden.items():
        if line not in new:
            diffs.append(f"{line}: missing")
            continue
        for h, a in heads.items():
            b = new[line].get(h)
            if a is None or b is None:
                if a is not b:
                    diffs.append(f"{line} / {h}: golden {a} vs now {b}")
            elif abs(a - b) > _tol(a, tol_abs, tol_frac):
                diffs.append(f"{line} / {h}: golden {a:,.6g} vs now {b:,.6g}")
    return diffs


def run_offline(cases: list[dict]) -> list[dict]:
    out = []
    for c in cases:
        d = HOME / c["id"]
        if not (d / "golden.json").exists():
            out.append({"id": c["id"], "ok": None, "diffs": ["no golden"]})
            continue
        golden = json.loads((d / "golden.json").read_text(encoding="utf-8"))
        replay = ReplayProvider(d / "grids.json")
        try:
            pre = _prepare(c["spec"], replay, replay.closed_months(), load_lines(d / "lines.json"))
        except Exception as e:   # noqa: BLE001 - a broken case is a failed case, not a crashed run
            out.append({"id": c["id"], "ok": False, "diffs": [f"{type(e).__name__}: {e}"]})
            continue
        diffs = compare(golden["key_values"], pre.result.key_values)
        out.append({"id": c["id"], "ok": not diffs, "diffs": diffs})
    return out
