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

log = logging.getLogger("smartview.sheetorder")

_OURS = re.compile(r"^Q(\d+) (Answer|FinMax|FinLite|FinPlan|Lines)(?: (.+))?$")
_RANK = {"FinMax": 0, "FinLite": 1, "FinPlan": 2, "Lines": 3}
SOURCES = ("Sources", "Sources SV")
ANSWER_TAB = 0x64381F          # TN navy (BGR of #1F3864), as answer.answer_sheet
DATA_TAB = 0xA6A6A6          # grey, as add_sheet(color=...) / styling.GREY_TAB


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


# ------------------------------------------------------------------ COM
def _cosmetic(read, busy):
    """read(), or None when Excel refuses it (e.g. a chart sheet has no tab colour); a busy Excel is raised so the
    whole read is retried."""
    try:
        return read()
    except Exception as e:   # noqa: BLE001
        if busy(e):
            raise
        return None


def _read_sheets(wb, busy) -> tuple[list[tuple[str, bool, object]], str | None]:
    """[(name, visible, tab colour)] in order + the active sheet's name (None when the workbook is not the active
    one), all in ONE jobs.com call."""
    rows = [(str(s.Name), s.Visible == -1, _cosmetic(lambda s=s: s.Tab.Color, busy)) for s in wb.Worksheets]
    active = _cosmetic(lambda: str(wb.ActiveSheet.Name)
                       if str(wb.Application.ActiveWorkbook.Name) == str(wb.Name) else None, busy)
    return rows, active


def arrange_sheets(wb, mode: str = "full", activate: str | None = None) -> dict:
    """Put the workbook's sheets in the fixed order (see module doc) with positional Move (keyword Before=/After= are
    ignored through late-bound COM), set the Answer/data tab colours, and activate `activate` (else the sheet that was
    active). Returns {'moved': [...], 'order': [...]}; raises only if Excel refuses a move.
    COM calls: one read of the order (names, visibility, tab colours, active sheet), one per sheet that moves or
    whose tab colour is wrong, one Activate unless that sheet is already active - 1 call for an ordered workbook."""
    from .jobs import com, fast_excel, is_busy
    rows, active = com(lambda: _read_sheets(wb, is_busy))
    names = [n for n, _, _ in rows]
    visible = {n: v for n, v, _ in rows}
    tabs = {n: c for n, _, c in rows}
    sources = None
    if mode == "active" and "Sources" in names and "Sources SV" not in names:
        try:                                            # a user's own 'Sources' (not our log) is not ours to move
            a1 = com(lambda: wb.Worksheets("Sources").Range("A1").Value)
        except Exception:   # noqa: BLE001
            a1 = None
        sources = {"Sources"} if a1 == "Fetched" else set()
    target = sheet_order(names, visible, mode, sources)
    moves = plan_moves(names, target, pinned_sheets(names, visible, mode, sources))
    if moves:
        with fast_excel(com(lambda: wb.Application)):
            for n, where, anchor in moves:
                if where == "before":                   # positional Before
                    com(lambda n=n, a=anchor: wb.Worksheets(n).Move(wb.Worksheets(a)))
                else:                                   # positional After
                    com(lambda n=n, a=anchor: wb.Worksheets(n).Move(None, wb.Worksheets(a)))
        got = com(lambda: [str(s.Name) for s in wb.Worksheets])
        if got != target:
            log.warning("sheet order after moves %s != target %s", got, target)
    for n in target:
        kind = classify(n)[0] if visible.get(n) else "other"
        want = ANSWER_TAB if kind == "answer" else DATA_TAB if kind == "data" else None
        if want is None or tabs.get(n) == want:
            continue
        try:
            com(lambda n=n, w=want: setattr(wb.Worksheets(n).Tab, "Color", w))
        except Exception:   # noqa: BLE001 - cosmetic
            pass
    show = activate or active
    if show and (moves or show != active):              # a Move can change the active sheet: land again after it
        try:
            com(lambda: wb.Worksheets(show).Activate())
        except Exception:   # noqa: BLE001 - e.g. the workbook's window is not active: cosmetic
            pass
    return {"moved": [m[0] for m in moves], "order": target}
