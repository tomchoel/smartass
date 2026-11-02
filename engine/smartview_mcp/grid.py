"""Pure grid logic: build a free-form ad hoc layout and parse what Smart View returns.

Layout (proven on EPM Cloud Planning, 2026-10-13), for nr row dims and nc column dims:
  row 1              : POV members, starting in column nr+1
  rows 2 .. 1+nc     : column headers (cross product of the column member lists), from column nr+1
  rows 2+nc ..       : row headers (cross product of the row member lists) in columns 1..nr
Smart View keeps the order, so results are mapped by position. Suppression is done here, not in
Smart View, because server-side suppression would shift positions.
"""
from __future__ import annotations

import itertools
import re
from dataclasses import dataclass, field

MISSING = {"#Missing", "#NumericZero", "#Missing ", None, ""}
NOACCESS = {"#NoAccess", "#No Access"}
INVALID = {"#Invalid"}

# "3WHS_W Wholesale (un.SY)" -> code 3WHS_W, prefix un.SY  =>  member name "un.SY.3WHS_W"
_LABEL = re.compile(r"^\s*(?P<code>\S+)\s+(?P<desc>.*?)\s*\((?P<prefix>[A-Za-z][\w.]*)\)\s*$")


def name_from_label(label: str) -> str | None:
    """Derive the member name from Smart View's 'code alias (prefix)' display, if it follows that pattern."""
    if not isinstance(label, str):
        return None
    m = _LABEL.match(label)
    if not m:
        return None
    return f"{m['prefix']}.{m['code']}"


def indent_of(label) -> int:
    return len(label) - len(label.lstrip(" ")) if isinstance(label, str) else 0


@dataclass
class Axis:
    dimension: str
    members: list[str]


@dataclass
class Layout:
    pov: list[tuple[str, str]]                  # (dimension, member)
    rows: list[Axis]
    cols: list[Axis]
    row_tuples: list[tuple[str, ...]] = field(default_factory=list)
    col_tuples: list[tuple[str, ...]] = field(default_factory=list)

    def __post_init__(self):
        self.row_tuples = list(itertools.product(*[a.members for a in self.rows]))
        self.col_tuples = list(itertools.product(*[a.members for a in self.cols]))

    @property
    def nr(self) -> int:
        return len(self.rows)

    @property
    def nc(self) -> int:
        return len(self.cols)

    @property
    def n_cells(self) -> int:
        return len(self.row_tuples) * len(self.col_tuples)

    def writes(self) -> list[tuple[int, int, str]]:
        """(row, col, text) cell writes, 1-based."""
        w = [(1, self.nr + 1 + j, m) for j, (_, m) in enumerate(self.pov)]
        for k, combo in enumerate(self.col_tuples):
            for d, m in enumerate(combo):
                w.append((2 + d, self.nr + 1 + k, m))
        for k, combo in enumerate(self.row_tuples):
            for d, m in enumerate(combo):
                w.append((2 + self.nc + k, 1 + d, m))
        return w

    def matrix(self) -> list[list]:
        """The whole layout as a dense 2D list (for one bulk Range.Value write)."""
        height = 1 + self.nc + len(self.row_tuples)
        width = self.nr + max(len(self.pov), len(self.col_tuples))
        m = [[None] * width for _ in range(height)]
        for r, c, v in self.writes():
            m[r - 1][c - 1] = v
        return m


def validate_axes(dimensions: list[str], pov: dict[str, str], rows: list[Axis], cols: list[Axis]) -> list[str]:
    """Return problems (empty = fine). `dimensions` are the cube's dimension names (top members)."""
    problems = []
    used = [a.dimension for a in rows + cols] + list(pov)
    known = set(dimensions)
    for d in used:
        if d not in known:
            problems.append(f"unknown dimension '{d}' (cube has: {', '.join(dimensions)})")
    dupes = {d for d in used if used.count(d) > 1}
    for d in sorted(dupes):
        problems.append(f"dimension '{d}' is used more than once (pov/rows/columns)")
    for a in rows + cols:
        if not a.members:
            problems.append(f"no members given for dimension '{a.dimension}'")
    if not rows or not cols:
        problems.append("need at least one row dimension and one column dimension")
    return problems


def full_pov(dimensions: list[str], pov: dict[str, str], rows: list[Axis], cols: list[Axis]) -> list[tuple[str, str]]:
    """POV for every dimension not on an axis: user's choice, else the dimension's top member."""
    on_axis = {a.dimension for a in rows + cols}
    return [(d, pov.get(d, d)) for d in dimensions if d not in on_axis]


@dataclass
class Parsed:
    columns: list[str]
    rows: list[list]
    row_labels: list[list[str]]
    col_labels: list[list[str]]
    unknown_members: list[str]
    pov_problems: list[str]
    no_access_cells: int
    total_rows: int
    suppressed_rows: list[str] = field(default_factory=list)   # known members whose row was all #Missing


def _clean(v):
    if v in MISSING:
        return None
    if isinstance(v, str) and v.strip() in MISSING:
        return None
    return v


def parse(layout: Layout, values: tuple | list, suppress_missing: bool, suppress_zeros: bool) -> Parsed:
    """Map Smart View's returned UsedRange.Value2 (starting at A1) back onto the requested layout."""
    grid = [list(r) for r in values] if isinstance(values, (tuple, list)) else [[values]]
    nr, nc = layout.nr, layout.nc
    width = max(len(r) for r in grid) if grid else 0

    def cell(r, c):
        return grid[r][c] if r < len(grid) and c < len(grid[r]) else None

    # POV: Smart View appends a dimension's top member at the end of row 1 when a typed POV member was
    # not recognised (it then treats the typed text as a comment).
    pov_problems = []
    returned_pov = [cell(0, nr + j) for j in range(width - nr)]
    returned_pov = [v for v in returned_pov if v not in (None, "")]
    if len(returned_pov) > len(layout.pov):
        extra = returned_pov[len(layout.pov):]
        pov_problems.append("Smart View did not recognise a POV member and appended: " + ", ".join(map(str, extra)))

    col_labels = [[cell(1 + d, nr + k) for d in range(nc)] for k in range(len(layout.col_tuples))]
    out_rows, row_labels, unknown, suppressed = [], [], set(), []
    no_access = 0
    last = [None] * nr
    first_data_row = 1 + nc
    for k, combo in enumerate(layout.row_tuples):
        r = first_data_row + k
        labels = []
        for d in range(nr):
            v = cell(r, d)
            if v not in (None, ""):
                last[d] = v
            labels.append(last[d] if v in (None, "") else v)
        data_raw = [cell(r, nr + j) for j in range(len(layout.col_tuples))]
        if all(v is None for v in data_raw):
            # a recognised member always yields a value or a #Missing/#NoAccess label; all-None = comment row
            for d in range(nr):
                raw = cell(r, d)
                if isinstance(raw, str) and raw.strip() == combo[d]:
                    unknown.add(f"{layout.rows[d].dimension}: {combo[d]}")
        data = []
        for v in data_raw:
            if isinstance(v, str) and v.strip() in NOACCESS:
                no_access += 1
                data.append("#NoAccess")
            elif isinstance(v, str) and v.strip() in INVALID:
                data.append("#Invalid")
            else:
                data.append(_clean(v))
        nums = [v for v in data if isinstance(v, (int, float))]
        if suppress_missing and all(v is None for v in data):
            if not all(v is None for v in data_raw):    # a real member with no data (not an unknown name)
                suppressed.append(" / ".join(combo))
            continue
        if suppress_zeros and nums and all(v == 0 for v in nums) and len(nums) == len([v for v in data if v is not None]):
            continue
        out_rows.append(list(combo) + data)
        row_labels.append([str(x).strip() if x is not None else None for x in labels])
    for j, combo in enumerate(layout.col_tuples):
        for d in range(nc):
            raw = cell(1 + d, nr + j)
            if isinstance(raw, str) and raw.strip() == combo[d]:
                # unchanged text may simply be a member whose alias equals its name; only flag if whole column empty
                col_vals = [cell(first_data_row + k, nr + j) for k in range(len(layout.row_tuples))]
                if all(v is None for v in col_vals):
                    unknown.add(f"{layout.cols[d].dimension}: {combo[d]}")
    columns = [a.dimension for a in layout.rows] + [" / ".join(c) for c in layout.col_tuples]
    return Parsed(columns, out_rows, row_labels, col_labels, sorted(unknown), pov_problems, no_access,
                  len(layout.row_tuples), suppressed)


def verify_labels(lay: Layout, parsed) -> list[str]:
    """Each returned label must be the requested member (where Smart View's 'code alias (prefix)' display lets us
    derive the name). A mismatch means rows/columns moved - never build formulas on that. Used by both routes
    (COM and svhttp): `parsed` is parse(lay, values, False, False)."""
    bad = []
    for kind, tuples, labels in (("row", lay.row_tuples, parsed.row_labels),
                                 ("column", lay.col_tuples, parsed.col_labels)):
        for i, combo in enumerate(tuples):
            labs = labels[i] if i < len(labels) else ()
            for d, want in enumerate(combo):
                got = labs[d] if d < len(labs) else None
                name = name_from_label(got) if got else None
                if name and name.lower() != want.lower() and str(got).strip().lower() != want.lower():
                    bad.append(f"{kind} {i + 1}: asked {want}, grid shows {got}")
    return bad


