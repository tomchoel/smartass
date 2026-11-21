"""Draw every test-set case (offline replay, no EPM) into one new, unsaved workbook in a hidden Excel and render each
sheet to PNG, for a visual audit (memory: review-every-answer-visually). Never touches the user's Excel or workbooks;
nothing is saved (Excel saves silently fail without a sensitivity label, and an audit needs no file).

Usage (from engine/):  uv run --with pymupdf python -B ../scripts/audit_testset.py <png dir> [case ids]   < /dev/null
Prints per case the reply's warnings and layout findings; then Read the PNGs (audit_<sheet>_<page>.png).
"""
import json
import os
import sys

import pythoncom
import win32com.client as w
import win32process

sys.path.insert(0, os.path.dirname(__file__))
from render_answers import export  # noqa: E402

from smartass import exceljob, pipeline  # noqa: E402
from smartass.testset import runner  # noqa: E402
from smartass.xl.target import Target  # noqa: E402


def _synthetic():
    """(title, Prepared) for the block types the test set has no case for: matrix, a table by entities, a ranking
    by products. Figures from tests/v2 fixtures (invented)."""
    tests = os.path.join(os.path.dirname(__file__), "..", "engine", "tests", "v2")
    sys.path[:0] = [os.path.abspath(tests), os.path.abspath(os.path.join(tests, ".."))]
    from test_v2_pipeline import FakeProvider
    from test_v2_products import pfacts, presolver, pspec
    from v2util import facts, resolver, spec_dict
    out = []
    for title, blocks in (("Matrix of Opex", [{"type": "matrix", "query": "fm", "metric": "Opex"}]),
                          ("Table by entities", [{"type": "table", "query": "fm", "rows": "entities",
                                                  "metric": "Opex"}])):
        d = spec_dict(blocks=blocks, question=title)
        if title.startswith("Matrix"):         # a long message title (live 04.11: it ran off the page)
            d["plan"]["title"] = ("Explain the COGS (+23m) and Opex (+28m) underspend vs Target for 2WHS Wholesale "
                                  "YTD Oct 26: which accounts, which cost centres and which suppliers drive it")
        out.append((title, pipeline.prepare(d, resolver(), FakeProvider(facts()))))
    d = pspec()
    d["question"] = "Ranking by products"
    out.append(("Ranking by products", pipeline.prepare(d, presolver(), FakeProvider(pfacts()))))
    return out


def main() -> None:
    out_dir = os.path.abspath(sys.argv[1])
    ids = set(sys.argv[2:])          # --charts: also b01 with the chart kinds an agent picks
    os.makedirs(out_dir, exist_ok=True)
    pythoncom.CoInitialize()
    app = w.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Add()
        pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        t = Target(app=app, wb=wb, pid=pid, hwnd=app.Hwnd, created=False)
        exceljob._target = lambda h, target, title, build_new=False: t            # the audit book, never a session/user workbook
        exceljob._save = lambda t: None                          # never saved
        exceljob.remember = lambda pre, out: None                # no answer memory from an audit
        cases = runner.load_cases()
        if "--charts" in ids:              # the chart kinds an agent picks (the test set only has implicit charts)
            ids.discard("--charts")
            b01 = next(c for c in cases if c["id"] == "b01")
            extra = json.loads(json.dumps(b01))
            extra["id"], extra["rec"] = "b01+charts", "b01"          # replays the b01 recording
            extra["spec"]["question"] = "Chart kinds on the Wholesale P&L"
            extra["spec"]["blocks"] += [
                {"type": "chart", "kind": "column", "title": "Opex stands out", "source": {"block": 1},
                 "highlight": ["Opex"]},
                {"type": "chart", "kind": "variance", "title": "Variance by line", "source": {"block": 1}},
                {"type": "chart", "kind": "combo", "title": "Actual and variance %", "source": {"block": 1}},
                {"type": "chart", "kind": "stacked", "title": "Parts AC and PY", "source": {"block": 1},
                 "series": ["value", "ly"]},
                {"type": "chart", "kind": "stacked100", "title": "Mix AC and PY", "source": {"block": 1},
                 "series": ["value", "ly"]}]
            cases = [extra] if not ids else cases + [extra]
            ids.add("b01+charts")
        if "--synthetic" in ids:           # block types the test set lacks, on the unit-test fixtures (no EPM)
            ids.discard("--synthetic")
            for title, pre in _synthetic():
                try:
                    res = exceljob.excel_job(None, pre, "session")
                    print(f"synthetic Q{res['q']}: {title}")
                    for k in ("warnings", "layout", "findings", "mismatches"):
                        for x in res.get(k) or []:
                            print(f"    {k}: {x}")
                except Exception as e:   # noqa: BLE001
                    print(f"synthetic {title}: FAILED {type(e).__name__}: {e}")
            if not ids:
                cases = []
        for case in cases:
            if ids and case["id"] not in ids:
                continue
            try:
                _, pre = runner._replay({**case, "id": case.get("rec", case["id"])})
                res = exceljob.excel_job(None, pre, "session")
                print(f"{case['id']} Q{res['q']}: {case['spec']['question']}")
                print('    keys:', sorted(res))
                for k, v in res.items():
                    if k in ("warnings", "layout", "lint", "findings", "mismatches") and v:
                        for x in v:
                            print(f"    {k}: {x}")
            except Exception as e:   # noqa: BLE001 - report and go on
                print(f"{case['id']}: FAILED {type(e).__name__}: {e}")
        export(wb, out_dir, "audit")
        wb.Close(False)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
