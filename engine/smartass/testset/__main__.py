"""uv run python -m smartass.testset record [ids] | candidate --offline [ids] | approve ids | run
record is live (EPM over the hidden browser, never Excel); candidate --offline (a candidate from the replay, with its
diff to the golden), approve and run are offline. run exits 1 on any failure."""
from __future__ import annotations

import asyncio
import os
import sys

# Before core.config is imported (runner imports it): own browser profile slots and no warm browser, so a test-set
# run never collides with the running servers' slots (an explicit setting wins).
os.environ.setdefault("SV_DRILL_PROFILE_SLOTS", "6")
os.environ.setdefault("SV_DRILL_KEEP_WARM", "0")

from ..ops.cli import utf8_stdio  # noqa: E402
from . import runner  # noqa: E402


def _pick(ids: list[str]) -> list[dict]:
    cases = runner.load_cases()
    if not ids:
        return cases
    known = {c["id"] for c in cases}
    if set(ids) - known:
        raise SystemExit(f"unknown case(s): {', '.join(sorted(set(ids) - known))}; known: {', '.join(sorted(known))}")
    return [c for c in cases if c["id"] in ids]


def _record(ids: list[str]) -> int:
    from .. import server
    from ..model.spec import AnalysisSpec
    provider = server._provider()
    bad = 0
    for c in _pick(ids):
        try:
            spec = AnalysisSpec.model_validate(c["spec"])
            lines = asyncio.run(server._line_sets(spec))[0]
            cand = runner.record(c, provider, lines)
            print(f"{c['id']}: recorded, {len(cand['key_values'])} key lines (closed {cand['closed']['FinMax']})")
            for k, v in list(cand["key_values"].items())[:12]:
                print(f"    {k}: {v}")
        except Exception as e:   # noqa: BLE001 - report each refusal, go on with the next case
            bad += 1
            print(f"{c['id']}: FAILED - {type(e).__name__}: {e}")
    return 1 if bad else 0


def main(argv: list[str]) -> int:
    utf8_stdio()
    cmd, ids = (argv[0] if argv else ""), argv[1:]
    if cmd == "record":
        return _record(ids)
    if cmd == "candidate":
        if "--offline" not in ids:
            raise SystemExit("candidate needs --offline (a replay of the recording); a live candidate is `record`")
        for c in _pick([i for i in ids if i != "--offline"]):
            r = runner.candidate_offline(c)
            n = len(r["diffs"])
            print(f"{r['id']:<6}{f'{n} diff(s)' if n else 'same as golden'}")
            for d in r["diffs"]:
                print(f"      {d}")
        return 0
    if cmd == "approve":
        if not ids:
            raise SystemExit("approve needs case ids, e.g. approve b01 b02")
        for i in ids:
            print(f"{i}: approved -> {runner.approve(i)}")
        return 0
    if cmd == "run":
        report = runner.run_offline(_pick(ids))
        for r in report:
            mark = {True: "ok", False: "FAIL", None: "no golden"}[r["ok"]]
            print(f"{r['id']:<6}{mark}")
            for d in r["diffs"] if r["ok"] is False else []:
                print(f"      {d}")
        return 1 if any(r["ok"] is False for r in report) else 0
    raise SystemExit(__doc__)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
