"""Run log per answer (design recommendation 6): one JSON line per run in config.RUN_LOG and one row in the
workbook's hidden '_runlog' sheet - question, gate, assumptions, pulls (cube, route, seconds), timings, result."""
from __future__ import annotations

import datetime as dt
import json

from smartass.core import config
from smartass.xl import jobs, sheets

SHEET = "_runlog"
HEADER = ("When", "Question no", "Question", "Gate", "Pulls", "Result", "Details")


def entry(spec, q: int | None, sources: list, review: dict | None, timings: dict, error: str | None = None,
          workbook: str | None = None) -> dict:
    return {"ts": dt.datetime.now().isoformat(timespec="seconds"), "workbook": workbook,
            "q": f"Q{q}" if q else None, "question": spec.question, "gate": spec.gate.outcome,
            "assumptions": list(spec.assumptions), "deviation": spec.deviation,
            "pulls": [{"id": pin.id, "cube": pin.connection, "sheet": s.pull.sheet, "route": s.route,
                       "seconds": round(s.pull.seconds, 2)} for pin, s in zip(spec.pulls, sources)],
            "ok": None if review is None else review["ok"],
            "findings": [] if review is None else [f"{x['severity']}: {x['text']}" for x in review["findings"]],
            "timings": timings, "error": error}


def sheet_row(e: dict) -> tuple:
    result = "error" if e["error"] else {True: "OK", False: "check findings"}.get(e["ok"], e["gate"])
    pulls = "; ".join(f"{p['id']}={p['cube']} ({p['route']}, {p['seconds']} s)" for p in e["pulls"])
    return (e["ts"], e["q"] or "", e["question"], e["gate"], pulls, result,
            json.dumps(e, ensure_ascii=False, default=str)[:30000])


def append(e: dict) -> None:
    try:
        config.RUN_LOG.parent.mkdir(parents=True, exist_ok=True)
        with open(config.RUN_LOG, "a", encoding="utf-8") as f:
            f.write(json.dumps(e, ensure_ascii=False, default=str) + "\n")
    except OSError:
        pass                      # the log must never break an answer (OneDrive lock etc.)


def write_sheet(t, e: dict) -> None:
    com = jobs.com
    if not sheets._sheet_exists(t, SHEET):
        ws = com(lambda: t.wb.Worksheets.Add(None, t.wb.Worksheets(t.wb.Worksheets.Count)))
        com(lambda: setattr(ws, "Name", SHEET))
        com(lambda: setattr(ws, "Visible", 0))                 # xlSheetHidden
        com(lambda: setattr(ws.Range("A1:G1"), "Value", (HEADER,)))
    ws = com(lambda: t.wb.Worksheets(SHEET))
    r = com(lambda: ws.Cells(ws.Rows.Count, 1).End(-4162).Row) + 1
    com(lambda: setattr(ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)), "Value",
                        (tuple("'" + str(x) for x in sheet_row(e)),)))
