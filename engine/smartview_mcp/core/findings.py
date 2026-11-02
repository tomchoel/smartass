"""The review on the sheet (design §3.4): the Critical/Warning/Info findings of check.review as a small block below
the answer's notes, inside the tool-owned range. Written after the build because it needs the calculated checks."""
from __future__ import annotations

from ..userexcel import com as _com

LIMIT = 15
RESERVE = LIMIT + 3          # header + rows + tail + a blank row: room made before the answer is written
_STYLE = {"section": "TN Seksjon", "critical": "TN Note", "note": "TN Note"}
RED = 0x0000C0               # BGR for Excel: dark red


def rows(fs: list[dict], limit: int = LIMIT) -> list[tuple[str, str]]:
    out = [("Findings", "section")]
    if not fs:
        return out + [("No findings: all checks OK.", "note")]
    for f in fs[:limit]:
        out.append((f"{f['severity']} — {f['text']}", "critical" if f["severity"] == "Critical" else "note"))
    if len(fs) > limit:
        out.append((f"… and {len(fs) - limit} more (see the tool result)", "note"))
    return out


def render(ws, row: int, fs: list[dict], col: int = 2, com=_com) -> int:
    """Write the block from `row` in column B as one 2-D write, then a style per row; returns its last row."""
    rs = rows(fs)
    last = row + len(rs) - 1
    com(lambda: setattr(ws.Range(ws.Cells(row, col), ws.Cells(last, col)), "Value", tuple((t,) for t, _ in rs)))
    for i, (_, kind) in enumerate(rs):
        c = ws.Cells(row + i, col)
        com(lambda c=c, k=kind: setattr(c, "Style", _STYLE[k]))
        if kind == "critical":
            com(lambda c=c: setattr(c.Font, "Color", RED))
    return last
