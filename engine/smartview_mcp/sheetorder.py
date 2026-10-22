"""A neat, fixed sheet order for analysis workbooks (applied at the end of every tool that adds or rebuilds sheets).

Full mode (our own analysis workbook):
  1. 'Q<n> Svar' by question number (numeric: Q10 after Q9)
  2. other visible sheets (the user's own, e.g. 'Intervjuguide'), in their current relative order
  3. data sheets per question: FinMax, FinLite, FinPlan pulls (plain name first, then suffixes), then 'Q<n> Linjer*'
  4. 'Kilder' (and 'Kilder SV')
  5. hidden sheets ('_svmeta', parked) - never moved, they end up last
Active mode (the user's own workbook): only our sheets (Svar, data, Kilder) are reordered among the slots they
already occupy; every other sheet stays exactly where it is.

Moves are planned so that as few sheets as possible move (a weighted longest increasing subsequence keeps the rest);
hidden sheets and, in active mode, the user's sheets are pinned and never moved. Moving sheets does not break formulas.
"""
from __future__ import annotations

import logging
import re

log = logging.getLogger("smartview.sheetorder")

_OURS = re.compile(r"^Q(\d+) (Svar|FinMax|FinLite|FinPlan|Linjer)(?: (.+))?$")
_RANK = {"FinMax": 0, "FinLite": 1, "FinPlan": 2, "Linjer": 3}
KILDER = ("Kilder", "Kilder SV")
SVAR_TAB = 0x64381F          # TN navy (BGR of #1F3864), as answer.answer_sheet
DATA_TAB = 0xA6A6A6          # grey, as add_sheet(color=...) / styling.GREY_TAB


def _suffix_key(sfx: str | None) -> tuple[int, int]:
    if not sfx:
        return (0, 0)                     # 'Q1 FinLite' before 'Q1 FinLite 2' / 'Q1 FinLite w'
    if sfx.isdigit():
        return (1, int(sfx))
    return (2, 0)                         # text suffix (pull id / drill alias): current order


def classify(name: str) -> tuple[str, tuple]:
    """('svar' | 'data', sort key) for our sheets, ('other', ()) otherwise."""
    m = _OURS.match(name)
    if not m:
        return "other", ()
    q, kind, sfx = int(m.group(1)), m.group(2), m.group(3)
    if kind == "Svar":
        return "svar", (q, *_suffix_key(sfx))
    return "data", (q, _RANK[kind], *_suffix_key(sfx))


def _is_visible(name: str, visible) -> bool:
    if isinstance(visible, dict):
        return bool(visible.get(name, True))
    return name in visible


def default_kilder(names: list[str], mode: str) -> set[str]:
    if mode == "full":
        return set(KILDER)
    return {"Kilder SV"} if "Kilder SV" in names else {"Kilder"}   # a user's own 'Kilder' pushed ours to 'Kilder SV'


def sheet_order(names: list[str], visible, mode: str = "full", kilder: set[str] | None = None) -> list[str]:
    """Target order of the sheets `names` (current order). visible: {name: bool} or a set of visible names.
    mode 'full' (our analysis workbook) or 'active' (the user's own workbook: only our sheets move, within the
    slots they already occupy)."""
    kilder = default_kilder(names, mode) if kilder is None else set(kilder)
    svar, data, kild, other, hidden = [], [], [], [], []
    for i, n in enumerate(names):
        if not _is_visible(n, visible):
            hidden.append(n)
            continue
        kind, key = classify(n)
        if kind == "svar":
            svar.append((key, i, n))
        elif kind == "data":
            data.append((key, i, n))
        elif n in kilder:
            kild.append(((KILDER.index(n) if n in KILDER else 9,), i, n))
        else:
            other.append(n)
    ours = [n for _, _, n in sorted(svar)] + [n for _, _, n in sorted(data)] + [n for _, _, n in sorted(kild)]
    if mode == "full":
        return [n for _, _, n in sorted(svar)] + other + [n for _, _, n in sorted(data)] + \
            [n for _, _, n in sorted(kild)] + hidden
    if mode != "active":
        raise ValueError(f"mode must be 'full' or 'active', got {mode!r}")
    mine = set(ours)
    it = iter(ours)
    return [next(it) if n in mine else n for n in names]          # our sheets refill our own slots, in order


def pinned_sheets(names: list[str], visible, mode: str = "full", kilder: set[str] | None = None) -> set[str]:
    """Sheets that must never move: hidden ones; in active mode also every sheet that is not ours."""
    kilder = default_kilder(names, mode) if kilder is None else set(kilder)
    pins = {n for n in names if not _is_visible(n, visible)}
    if mode == "active":
        pins |= {n for n in names if classify(n)[0] == "other" and n not in kilder}
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
def arrange_sheets(wb, mode: str = "full", activate: str | None = None, colors: bool = True) -> dict:
    """Put the workbook's sheets in the fixed order (see module doc) with positional Move (keyword Before=/After= are
    ignored through late-bound COM), set the Svar/data tab colours, and activate `activate` (else the sheet that was
    active). Returns {'moved': [...], 'order': [...]}; raises only if Excel refuses a move."""
    from .userexcel import com, fast_excel
    sheets = list(com(lambda: wb.Worksheets))
    names = [str(com(lambda s=s: s.Name)) for s in sheets]
    visible = {n: com(lambda s=s: s.Visible) == -1 for n, s in zip(names, sheets)}   # xlSheetVisible
    kilder = None
    if mode == "active" and "Kilder" in names and "Kilder SV" not in names:
        try:                                            # a user's own 'Kilder' (not our log) is not ours to move
            a1 = com(lambda: wb.Worksheets("Kilder").Range("A1").Value)
        except Exception:   # noqa: BLE001
            a1 = None
        kilder = {"Kilder"} if a1 == "Hentet" else set()
    target = sheet_order(names, visible, mode, kilder)
    moves = plan_moves(names, target, pinned_sheets(names, visible, mode, kilder))
    try:
        prev_active = str(com(lambda: wb.ActiveSheet.Name))
    except Exception:   # noqa: BLE001 - e.g. a chart sheet or no window
        prev_active = None
    if moves:
        with fast_excel(com(lambda: wb.Application)):
            for n, where, anchor in moves:
                ws, ref = com(lambda n=n: wb.Worksheets(n)), com(lambda a=anchor: wb.Worksheets(a))
                if where == "before":
                    com(lambda ws=ws, ref=ref: ws.Move(ref))            # positional Before
                else:
                    com(lambda ws=ws, ref=ref: ws.Move(None, ref))      # positional After
        got = [str(com(lambda s=s: s.Name)) for s in com(lambda: wb.Worksheets)]
        if got != target:
            log.warning("sheet order after moves %s != target %s", got, target)
    if colors:
        for n in target:
            kind = classify(n)[0] if visible.get(n) else "other"
            want = SVAR_TAB if kind == "svar" else DATA_TAB if kind == "data" else None
            if want is None:
                continue
            try:
                tab = com(lambda n=n: wb.Worksheets(n).Tab)
                if com(lambda: tab.Color) != want:
                    com(lambda: setattr(tab, "Color", want))
            except Exception:   # noqa: BLE001 - cosmetic
                pass
    show = activate or prev_active
    if show:
        try:
            com(lambda: wb.Worksheets(show).Activate())
        except Exception:   # noqa: BLE001 - e.g. the workbook's window is not active: cosmetic
            pass
    return {"moved": [m[0] for m in moves], "order": target}
