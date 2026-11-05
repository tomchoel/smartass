# engine/smartass/render/datasheet.py
"""The provenance header above table Data_Q<n> (sub-spec 5 §2.1): one label/value block per query, so the sheet
says what was fetched, from where, when and by whom, and checks itself against EPM (Rows, Control total).

Layout, from row 1 (label in A, value in B):
  1  Title                (the title text itself sits in A1)
  2  Question
  per query, in spec order:
     Query · POV · Periods · Route · Fetched · Rows · Control total (B static, C live SUMIFS, D OK/DIFF) · Sources
  then Tolerance (NOK) (one, used by every status) and Refresh.
One blank row, then the table header at row len(header) + 2. The formulas carry fixed addresses: the header always
starts in row 1. The Sources value is left empty here; plan 5 Task 11 links it to the Sources sheet."""
from __future__ import annotations

import getpass
import os
from datetime import datetime

from ..model.periods import closed_for, label, resolve

REFRESH_TEXT = "Refresh: fin_refresh, or ask a follow-up. Do not edit this table by hand."
TOLERANCE = 1.0                    # NOK: EPM's static total against the table's live sum
ROUTES = {"http": "HTTP", "adhoc": "Smart View ad hoc", "replay": "Replay"}


class Header(list):
    """[(label, value)] in row order (row = index + 1). value is text, a number, a formula or a tuple of them
    (one per column from B). sources / status: query id -> the 'Sources' value cell and the control-total status
    cell."""

    def __init__(self, rows=(), sources=None, status=None):
        super().__init__(rows)
        self.sources: dict[str, str] = dict(sources or {})
        self.status: dict[str, str] = dict(status or {})

    @property
    def table_row(self) -> int:
        return len(self) + 2


def windows_user() -> str:
    try:
        return os.getlogin()
    except OSError:
        try:
            return getpass.getuser()
        except Exception:   # noqa: BLE001 - no user to name
            return ""


def fetched_text(at: datetime, user: str) -> str:
    """'2026-11-05 09:30 +0100 · tom': static (%z, not the long Windows zone name)."""
    text = f"{at:%Y-%m-%d %H:%M}" + (f" {at:%z}" if at.utcoffset() is not None else "")
    return f"{text} · {user}" if user else text


def _span(expr: str, closed) -> str:
    ms = resolve(expr, closed)
    text = label(ms[0]) if len(ms) == 1 else f"{label(ms[0])} – {label(ms[-1])}"
    return f"{text} ({expr})"


def _query_rows(q: int, query, pre, fetched: str, first: int, tol: str):
    """(rows, sources cell, status cell) of one query's block starting at sheet row `first`."""
    t = f"Data_Q{q}"
    reqs = [r for r in pre.requests if r.query == query.id]
    scen = list(dict.fromkeys(f"{r.scenario}/{r.version}" if r.version else r.scenario for r in reqs))
    balances = list(dict.fromkeys(r.balance for r in reqs))
    name = pre.result.entity_labels.get(query.id, query.entity)
    pov = " · ".join(x for x in (name if name == query.entity else f"{name} ({query.entity})", "NOK",
                                 ", ".join(balances)) if x)
    closed = closed_for(pre.closed, query.cube)
    periods = _span(query.periods, closed)
    if query.plan_periods:
        periods += f" · plan {_span(query.plan_periods, closed)}"
    got = [g for g in pre.fetched.values() if g.request.query == query.id]
    route = " + ".join(dict.fromkeys(ROUTES.get(g.route, g.route) for g in sorted(got, key=lambda g: g.route != "http")))
    failed = [r for r in pre.failed if r.query == query.id]
    if failed:
        route = " · ".join(x for x in (route, f"HTTP failed for {len(failed)} grid(s)") if x)
    static = pre.frame.total({"Query": (query.id,)}) or 0.0
    ct = first + 6
    rows = [("Query", f"{query.id}: {query.cube}" + (f" · {', '.join(scen)}" if scen else "")),
            ("POV", pov),
            ("Periods", periods),
            ("Route", route),
            ("Fetched", fetched),
            ("Rows", f'=COUNTIFS({t}[Query],"{query.id}")'),
            ("Control total", (static, f'=SUMIFS({t}[Value],{t}[Query],"{query.id}")',
                               f'=IF(ABS(B{ct}-C{ct})<={tol},"OK","DIFF")')),
            ("Sources", "")]
    return rows, f"B{first + 7}", f"D{ct}"


def header_rows(q: int, spec, pre, fetched_at: datetime, user: str) -> Header:
    """The Data_Q<q> header rows (see the module doc). pre: the Prepared run (requests, routes, frame, closed)."""
    tol_row = 3 + 8 * len(spec.queries)
    tol = f"$B${tol_row}"
    fetched = fetched_text(fetched_at, user)
    head = Header([("Title", f"Q{q} Data — {spec.question}"), ("Question", f"Q{q}: {spec.question}")])
    for query in spec.queries:
        rows, src, status = _query_rows(q, query, pre, fetched, len(head) + 1, tol)
        head += rows
        head.sources[query.id], head.status[query.id] = src, status
    head += [("Tolerance (NOK)", TOLERANCE), ("Refresh", REFRESH_TEXT)]
    assert len(head) == tol_row + 1, "the Tolerance row moved: fix tol_row"
    return head


# ------------------------------------------------------------------ the lines sheets (sub-spec 5 §2.3)
LINES_HEADER = "F1:K3"     # rows 1-3 only: the lines table header is row 4 (drill_xl.write_lines); A1/A2, C2/D2 stay


def lines_header(table: str, cell: dict, fetched: str) -> list[list]:
    """The provenance block of a 'Q<n> Lines <alias>' sheet, as the 3 x 6 grid for F1:K3 (label, value pairs):
      F1 Drill cell  · G1 the drilled cell's dimensions
      F2 Fetched     · G2 static time and user
      F3 Lines total · G3 =SUM(<table>[AMOUNT]) · H3 Tie-out · I3 OK/DIFF against the cube cell · J3/K3 tolerance.
    The amount column is the one LinePick sums (model.lines.AMOUNT); a missing cube value says so, never DIFF."""
    from ..model.lines import AMOUNT, CUBE_CELL
    dims = " · ".join(f"{k}: {v}" for k, v in (cell or {}).items())
    tie = (f'=IF(NOT(ISNUMBER({CUBE_CELL})),"no cube value",'
           f'IF(ABS($G$3-{CUBE_CELL})<=$K$3,"OK","DIFF"))')
    return [["Drill cell", dims, "", "", "", ""],
            ["Fetched", fetched, "", "", "", ""],
            ["Lines total", f"=SUM({table}[{AMOUNT}])", "Tie-out", tie, "Tolerance (NOK)", TOLERANCE]]
