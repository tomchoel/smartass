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

from smartass import pipeline  # noqa: E402
from smartass.testset import runner  # noqa: E402
from smartass.xl.target import Target  # noqa: E402


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
        pipeline._target = lambda h, target, title: t            # the audit book, never a session/user workbook
        pipeline._save = lambda t: None                          # never saved
        pipeline.remember = lambda pre, out: None                # no answer memory from an audit
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
        for case in cases:
            if ids and case["id"] not in ids:
                continue
            try:
                _, pre = runner._replay({**case, "id": case.get("rec", case["id"])})
                res = pipeline.excel_job(None, pre, "session")
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
