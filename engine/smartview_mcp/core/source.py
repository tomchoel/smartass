"""pull, form C (design §3.2): fetch a grid over Smart View's HTTP provider and write it as a source sheet in ONE
block, with a provenance panel and the pull spec stored in _svmeta; a Smart View ad hoc grid in Excel (form A) when
the HTTP route fails. The grid starts at A1 in Smart View's own shape (POV row, column headers, row labels), so the
label-keyed lookups of answer.py work unchanged - which is why provenance is a panel to the RIGHT of the grid: a
header above it would move every label, and row 1 is the POV row."""
from __future__ import annotations

import datetime as dt
import math
from dataclasses import dataclass

from .. import config, svhttp, userexcel
from ..grid import MISSING, Axis, Layout, parse
from ..timing import stage
from .contract import Pull as PullSpec

GREY = 0x808080
ROUTES = {"http": "HTTP value grid; refresh = ask smartass to re-fetch",
          "adhoc": "Smart View ad hoc grid; refresh in Smart View"}


@dataclass
class Source:
    pull: userexcel.Pull
    route: str                          # http | adhoc
    mismatches: list[str]               # sheet cells that differ from the fetch (must be empty)
    note: str | None = None             # why the fallback was used


def cells(values: list[list], data_from: tuple[int, int] | None = None) -> tuple[tuple, ...]:
    """The fetched grid as Excel values: text stays text ('Jan 26' would become a date), blanks stay empty. In the
    data block (rows/columns from `data_from`, 0-based) a missing cell is 0, as Smart View's #NumericZero writes it
    on ad hoc sheets: a '#Missing' text would make every var/sum formula on that line #VALUE!."""
    r0, c0 = data_from or (len(values), 0)

    def one(v, r, c):
        if r >= r0 and c >= c0 and v in MISSING:
            return 0
        if v is None or v == "":
            return ""
        return "'" + v if isinstance(v, str) else v
    w = max(len(r) for r in values)
    return tuple(tuple(one(row[c], r, c) if c < len(row) else "" for c in range(w)) for r, row in enumerate(values))


def provenance(connection: str, lay: Layout, route: str, fetched: dt.datetime) -> list[tuple[str, str]]:
    pov = " · ".join(m for d, m in lay.pov if m != d) or "top members"
    return [("Source", f"Oracle EPM Cloud · {connection}"), ("POV", pov),
            ("Fetched", fetched.strftime("%Y-%m-%d %H:%M")), ("Route", ROUTES[route])]


def mismatches(values: list[list], read, limit: int = 5) -> list[str]:
    """Every fetched number must be on the sheet unchanged (design §3.4: sheet value = fetched value)."""
    out = []
    for r, row in enumerate(values):
        for c, v in enumerate(row):
            if isinstance(v, bool) or not isinstance(v, (int, float)):
                continue
            got = read[r][c] if r < len(read) and c < len(read[r]) else None
            if isinstance(got, bool) or not isinstance(got, (int, float)) or abs(got - v) > 1e-6 * max(1.0, abs(v)):
                out.append(f"row {r + 1} column {c + 1}: fetched {v!r}, sheet shows {got!r}")
                if len(out) == limit:
                    return out
    return out


def sheet_title(q: int, i: int, pin: PullSpec) -> str:
    return pin.title or f"Q{q} {pin.connection}" + (f" {pin.id}" if i else "")


def write(t, q: int, title: str, pin: PullSpec, f: svhttp.Fetched) -> Source:
    from .. import styling
    from ..ops import _used_values
    com, lay = userexcel.com, f.layout
    name = userexcel.unique_sheet_name(t, title)
    ws = userexcel.add_sheet(t, name, after_q=q, color=0xA6A6A6)
    grid = cells(f.values, data_from=(1 + lay.nc, lay.nr))
    h, w = len(grid), len(grid[0])
    parsed = parse(lay, f.values, False, False)
    p = userexcel._make_pull(name, pin.connection, lay, parsed, f.seconds)
    now = dt.datetime.now()
    with userexcel.fast_excel(t.app):
        com(lambda: setattr(ws.Range(ws.Cells(1, 1), ws.Cells(h, w)), "Value", grid))
        styling.apply_data_sheet(ws, header_rows=1 + lay.nc, label_cols=lay.nr,
                                 fte_rows=userexcel.fte_rows(lay, parsed))
        panel = provenance(pin.connection, lay, "http", now)
        prng = ws.Range(ws.Cells(2, w + 2), ws.Cells(1 + len(panel), w + 3))
        com(lambda: setattr(prng, "Value", tuple((k, "'" + v) for k, v in panel)))
        com(lambda: setattr(prng.Font, "Color", GREY))
        com(lambda: setattr(prng.Font, "Italic", True))
        userexcel.save_meta(t, p, extra={"route": "http", "fetched": now.isoformat(timespec="seconds"),
                                         "pull": pin.model_dump()})
    return Source(p, "http", mismatches(f.values, com(lambda: _used_values(ws))))


def pull(h, t, q: int, i: int, pin: PullSpec) -> Source:
    """Pull i of question q: HTTP first, the ad hoc grid if the HTTP route fails or the grid is too big for one
    request. A wrong query (InputError) is never retried through Excel."""
    title = sheet_title(q, i, pin)
    rows = [Axis(a.dimension, a.members) for a in pin.rows]
    cols = [Axis(a.dimension, a.members) for a in pin.columns]
    n = math.prod(len(a.members) for a in rows + cols)
    if n / 200 > config.SVHTTP_TIMEOUT_S:
        note = f"{n} cells is too big for one HTTP request; fetched through Smart View in Excel."
    else:
        try:
            with stage("fetch (HTTP)"):
                f = svhttp.shared().fetch(pin.connection, dict(pin.pov), rows, cols)
        except svhttp.SvHttpError as e:
            note = f"HTTP route failed ({e}); fetched through Smart View in Excel instead."
        else:
            with stage("write source sheet"):
                return write(t, q, title, pin, f)
    with stage("pull (Smart View ad hoc)"):
        p = userexcel.pull(h, t, q, pin.connection, dict(pin.pov), rows, cols, title=title)
    return Source(p, "adhoc", [], note)
