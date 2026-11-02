"""The answer sheet appears first (design §3.3): title, question, status line and assumptions, before any pull.
answer.build later clears it and renders the full answer in one block."""
from __future__ import annotations

from .. import answer as answer_mod
from ..excel.jobs import com
from .contract import AnswerSpec

ANSWER = "Answer"            # sheet 'Q<n> Answer'
STATUS_ROW = 3


def skeleton(spec: AnswerSpec, status: str) -> list[str]:
    rows = [spec.layout.title if spec.layout else spec.question, f"Question (as asked): {spec.question}",
            f"Status: {status}"]
    if spec.assumptions:
        rows += ["", "Assumptions"] + [f"•  {a}" for a in spec.assumptions]
    if spec.deviation:
        rows += ["", f"Deviation from the standard analysis (user's request): {spec.deviation}"]
    if spec.gate.outcome == "infeasible":
        rows += ["", f"Not answerable as asked: {spec.gate.reason}", f"Answered instead: {spec.gate.proxy}"]
    return rows


def show(t, q: int, spec: AnswerSpec, status: str) -> str:
    ws = answer_mod.answer_sheet(t, q, suffix=ANSWER)
    rows = skeleton(spec, status)
    rng = ws.Range(ws.Cells(1, 2), ws.Cells(len(rows), 2))
    com(lambda: setattr(rng, "Value", tuple((("'" + r) if r else "",) for r in rows)))
    com(lambda: setattr(ws.Cells(1, 2).Font, "Bold", True))
    com(lambda: setattr(ws.Cells(1, 2).Font, "Size", 14))
    com(lambda: setattr(ws.Cells(STATUS_ROW, 2).Font, "Italic", True))
    com(t.wb.Activate)
    com(ws.Activate)
    t.status(status)
    return ws.Name


def set_status(t, q: int, text: str) -> None:
    ws = com(lambda: t.wb.Worksheets(f"Q{q} {ANSWER}"))
    com(lambda: setattr(ws.Cells(STATUS_ROW, 2), "Value", "'Status: " + text))
    t.status(text)
