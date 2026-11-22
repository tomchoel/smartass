"""A neat, fixed sheet order for analysis workbooks (applied at the end of every tool that adds or rebuilds sheets).

Full mode (our own analysis workbook):
  1. 'Q<n> Answer' by question number (numeric: Q10 after Q9)
  2. other visible sheets (the user's own, e.g. 'Intervjuguide'), in their current relative order
  3. data sheets per question: FinMax, FinLite, FinPlan pulls (plain name first, then suffixes), then 'Q<n> Lines*'
  4. 'Sources' (and 'Sources SV')
  5. hidden sheets ('_svmeta', parked) - never moved, they end up last
Active mode (the user's own workbook): only our sheets (Answer, data, Sources) are reordered among the slots they
already occupy; every other sheet stays exactly where it is.

Moves are planned so that as few sheets as possible move (a weighted longest increasing subsequence keeps the rest);
hidden sheets and, in active mode, the user's sheets are pinned and never moved. Moving sheets does not break formulas.
"""
from __future__ import annotations

import logging
import re

from .inproc.batch import MISSING

log = logging.getLogger("smartview.sheetorder")

_OURS = re.compile(r"^Q(\d+) (Answer|Data|FinMax|FinLite|FinPlan|Lines)(?: (.+))?$")
_RANK = {"Data": -1, "FinMax": 0, "FinLite": 1, "FinPlan": 2, "Lines": 3}   # v2 Q<n> Data before the pulls
SOURCES = ("Sources", "Sources SV")
ANSWER_TAB = 0x64381F          # TN navy (BGR of #1F3864), as answer.answer_sheet
DATA_TAB = 0xA6A6A6          # grey, as add_sheet(color=...)


def _suffix_key(sfx: str | None) -> tuple[int, int]:
    if not sfx:
        return (0, 0)                     # 'Q1 FinLite' before 'Q1 FinLite 2' / 'Q1 FinLite w'
    if sfx.isdigit():
        return (1, int(sfx))
    return (2, 0)                         # text suffix (pull id / drill alias): current order


def classify(name: str) -> tuple[str, tuple]:
    """('answer' | 'data', sort key) for our sheets, ('other', ()) otherwise."""
    m = _OURS.match(name)
    if not m:
        return "other", ()
    q, kind, sfx = int(m.group(1)), m.group(2), m.group(3)
    if kind == "Answer":
        return "answer", (q, *_suffix_key(sfx))
    return "data", (q, _RANK[kind], *_suffix_key(sfx))


def _is_visible(name: str, visible) -> bool:
    if isinstance(visible, dict):
        return bool(visible.get(name, True))
    return name in visible


def default_sources(names: list[str], mode: str) -> set[str]:
    if mode == "full":
        return set(SOURCES)
    return {"Sources SV"} if "Sources SV" in names else {"Sources"}   # a user's own 'Sources' pushed ours to 'Sources SV'


def sheet_order(names: list[str], visible, mode: str = "full", sources: set[str] | None = None) -> list[str]:
    """Target order of the sheets `names` (current order). visible: {name: bool} or a set of visible names.
    mode 'full' (our analysis workbook) or 'active' (the user's own workbook: only our sheets move, within the
    slots they already occupy)."""
    sources = default_sources(names, mode) if sources is None else set(sources)
    answers, data, srcs, other, hidden = [], [], [], [], []
    for i, n in enumerate(names):
        if not _is_visible(n, visible):
            hidden.append(n)
            continue
        kind, key = classify(n)
        if kind == "answer":
            answers.append((key, i, n))
        elif kind == "data":
            data.append((key, i, n))
        elif n in sources:
            srcs.append(((SOURCES.index(n) if n in SOURCES else 9,), i, n))
        else:
            other.append(n)
    ours = [n for _, _, n in sorted(answers)] + [n for _, _, n in sorted(data)] + [n for _, _, n in sorted(srcs)]
    if mode == "full":
        return [n for _, _, n in sorted(answers)] + other + [n for _, _, n in sorted(data)] + \
            [n for _, _, n in sorted(srcs)] + hidden
    if mode != "active":
        raise ValueError(f"mode must be 'full' or 'active', got {mode!r}")
    mine = set(ours)
    it = iter(ours)
    return [next(it) if n in mine else n for n in names]          # our sheets refill our own slots, in order


def pinned_sheets(names: list[str], visible, mode: str = "full", sources: set[str] | None = None) -> set[str]:
    """Sheets that must never move: hidden ones; in active mode also every sheet that is not ours."""
    sources = default_sources(names, mode) if sources is None else set(sources)
    pins = {n for n in names if not _is_visible(n, visible)}
    if mode == "active":
        pins |= {n for n in names if classify(n)[0] == "other" and n not in sources}
    return pins


def plan_moves(current: list[str], target: list[str], pinned: set[str] = frozenset()) -> list[tuple[str, str, str]]:
    """Fewest moves from `current` to `target`: [(sheet, 'before' | 'after', anchor sheet)], applied in order.
    Keeps a heaviest increasing subsequence in place (pinned sheets weigh more than all others together)."""
    if sorted(current) != sorted(target):
        raise ValueError("current and target must hold the same sheets")
    pos = {n: i for i, n in enumerate(current)}
    seq = [pos[n] for n in target]
    big = len(target) + 1
    w = [big if n in pinned else 1 for n in target]
    best, prev = [0] * len(seq), [-1] * len(seq)
    for i in range(len(seq)):
        best[i] = w[i]
        for j in range(i):
            if seq[j] < seq[i] and best[j] + w[i] > best[i]:
                best[i], prev[i] = best[j] + w[i], j
    keep, i = set(), (max(range(len(seq)), key=best.__getitem__) if seq else -1)
    while i >= 0:
        keep.add(target[i])
        i = prev[i]
    moves, order = [], list(current)
    for i in range(len(target) - 1, -1, -1):          # back to front: each sheet goes right before its successor
        n = target[i]
        if n in keep:
            continue
        mv = (n, "before", target[i + 1]) if i + 1 < len(target) else (n, "after", order[-1])
        moves.append(mv)
        order = simulate(order, [mv])
    if order != list(target):                          # cannot happen; never leave a half-planned order
        raise RuntimeError(f"sheet order plan failed: {order} != {target}")
    return moves


def simulate(order: list[str], moves) -> list[str]:
    order = list(order)
    for n, where, anchor in moves:
        if n == anchor:
            continue
        order.remove(n)
        k = order.index(anchor)
        order.insert(k if where == "before" else k + 1, n)
    return order


# ------------------------------------------------------------------ in Excel (a batch)
def record_arrange(b, wbn, rows: list[tuple], active: str | None, mode: str = "full", activate: str | None = None,
                   sources_ours: bool | None = None) -> dict:
    """Record into b what puts the workbook's sheets in the fixed order (see module doc): positional Move (keyword
    Before=/After= are ignored through late-bound COM), the Answer/data tab colours, and activate `activate` (else
    the sheet that was active). rows: [(name, visible, tab colour)] in order and active: the active sheet (None when
    the workbook is not the active one), as the caller knows them (the file job's Book: nothing is read);
    sources_ours: in active mode, a 'Sources' sheet (without 'Sources SV') is smartass's log (A1 says so), else the
    user's and not moved. Every operation may fail (cosmetic). Returns {'moved': [...], 'order': [...], 'moves':
    range of the moves' operations} (a skipped one there means the order was not reached)."""
    names = [n for n, _, _ in rows]
    visible = {n: v for n, v, _ in rows}
    tabs = {n: c for n, _, c in rows}
    sources = None
    if mode == "active" and "Sources" in names and "Sources SV" not in names:
        sources = {"Sources"} if sources_ours else set()        # a user's own 'Sources' is not ours to move
    target = sheet_order(names, visible, mode, sources)
    moves = plan_moves(names, target, pinned_sheets(names, visible, mode, sources))
    with b.attempt():
        start = b.position
        for n, where, anchor in moves:
            if where == "before":                   # positional Before
                wbn.Worksheets(n).Move(wbn.Worksheets(anchor))
            else:                                   # positional After
                wbn.Worksheets(n).Move(MISSING, wbn.Worksheets(anchor))
        span = range(start, b.position)
        for n in target:
            kind = classify(n)[0] if visible.get(n) else "other"
            want = ANSWER_TAB if kind == "answer" else DATA_TAB if kind == "data" else None
            if want is not None and tabs.get(n) != want:
                wbn.Worksheets(n).Tab.Color = want
        show = activate or active
        if show and (moves or show != active):      # a Move can change the active sheet: land again after it
            wbn.Worksheets(show).Activate()
    return {"moved": [m[0] for m in moves], "order": target, "moves": span}


def arrange_sheets(wb, mode: str = "full", activate: str | None = None) -> dict:
    """arrange in two batches of its own: one reads the order (names, visibility, tab colours, the active sheet,
    whether a 'Sources' sheet is ours), one moves, colours and activates (record_arrange; quiet). Raises only if Excel
    refuses a move. Returns {'moved': [...], 'order': [...]}."""
    from .inproc.batch import Batch
    from .inproc.runner import run
    from .jobs import com
    from .sources import OURS_A1
    app, name = com(lambda: (wb.Application, str(wb.Name)))
    rb = Batch(quiet=False)
    wbn = rb.workbook(name)
    rb.ret(rb.verb("each", wbn.Worksheets, ["Name", "Visible", "Tab.Color"]), "sheets")
    rb.ret(wbn.FullName, "book")
    with rb.attempt():
        rb.ret(wbn.ActiveSheet.Name, "active")
        rb.ret(rb.app.ActiveWorkbook.FullName, "active book")
        rb.ret(wbn.Worksheets("Sources").Range("A1").Value2, "sources a1")
    v = run(app, rb).values
    rows = [(str(n), vis == -1, tab) for n, vis, tab in v["sheets"]]
    active = v.get("active") if v.get("active book") == v["book"] else None
    b = Batch(calculation=False)
    out = record_arrange(b, b.workbook(name), rows, active, mode, activate, v.get("sources a1") in OURS_A1)
    res = run(app, b)
    refused = [m for i, _, m in res.skipped if i in out["moves"]]
    if refused:
        raise RuntimeError(f"Excel refused a sheet move: {refused[0]}")
    return {"moved": out["moved"], "order": out["order"]}
