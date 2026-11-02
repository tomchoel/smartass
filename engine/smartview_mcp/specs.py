"""Answer-spec preprocessing (pure Python, no Excel): shorthands that keep specs short, patch-only follow-ups, and
factual key points drafted from the pulled numbers.

Shorthands (all optional; a plain spec passes through unchanged):
  * {"pull": "<id>"}      - a pull id instead of the sheet name.
  * section "repeat"      - one line per value: {"over": [...] | "months:Jan..Dec" | {"top": N, "pull": id,
                            "by": [col members], "among": [row members], "abs": true}, "line": {<line template>}}.
                            Values are strings ("{v}") or dicts ("{name}"); "{i}" is the 0-based index. A "top" over
                            ranks the pulled rows after the pull and yields {"m": member, "label": its label}. A pull
                            cell whose members are not in that pull (e.g. 'Oct 26' before October closes) becomes
                            empty, and var/varpct cells that depend on an empty cell are dropped too.
  * {"sum": "m*"}         - every line key matching the pattern, in spec order ({"avg": "f*"} likewise).
  * spec "auto_points"    - true (default): factual key points drafted from the numbers are put first; the spec's
                            own key_points (interpretation) follow. false: only the spec's own.
"""
from __future__ import annotations

import copy
import fnmatch
import re

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ------------------------------------------------------------------ patches
def merge(base: dict, patch: dict) -> dict:
    """Deep merge for follow-ups: dicts merge key by key, anything else (lists, values) replaces; None deletes."""
    out = copy.deepcopy(base)
    for k, v in (patch or {}).items():
        if v is None:
            out.pop(k, None)
        elif isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = merge(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out


# ------------------------------------------------------------------ expansion
def _over(spec_over, sheets: dict | None = None, pulls: dict | None = None) -> list:
    if isinstance(spec_over, str) and spec_over.startswith("months:"):
        a, _, b = spec_over[7:].partition("..")
        i, j = MONTHS.index(a.strip()[:3].title()), MONTHS.index((b or a).strip()[:3].title())
        return MONTHS[i:j + 1]
    if isinstance(spec_over, dict) and "top" in spec_over:
        return _top(spec_over, sheets or {}, pulls or {})
    return list(spec_over or [])


def _top(o: dict, sheets: dict, pulls: dict) -> list[dict]:
    """{"top": N, "pull": id, "by": [col members], "among": [row members] (default: every row), "abs": true}
    -> [{"m": member, "label": readable label}] for the N largest, ranked on the pulled values."""
    from .templates import _clean                    # lazy: templates imports the atlas
    p = pulls.get(sheets.get(o["pull"], o["pull"]))
    if p is None:
        return []
    among = o.get("among") or [k[0] for k in p.row_labels]
    vals = [(_pull_value(p, [m], o["by"]), m) for m in among]
    key = abs if o.get("abs", True) else (lambda v: v)
    ranked = sorted(((key(v), m) for v, m in vals if isinstance(v, (int, float))), reverse=True)[:int(o["top"])]
    return [{"m": m, "label": _clean((p.row_labels.get((m,)) or (m,))[-1]) or m} for _, m in ranked]


def _fill(obj, v, i):
    if isinstance(obj, str):
        if isinstance(v, dict):
            def rep(m):
                k = m.group(1)
                return "\0" if k in v and v[k] is None else str(v.get(k, i if k == "i" else m.group(0)))
            s = re.sub(r"\{(\w+)\}", rep, obj)
            return None if "\0" in s else s
        return obj.replace("{v}", str(v)).replace("{i}", str(i))
    if isinstance(obj, list):
        return [_fill(x, v, i) for x in obj]
    if isinstance(obj, dict):
        return {k: _fill(x, v, i) for k, x in obj.items()}
    return obj


def expand(spec: dict, sheets: dict[str, str], pulls: dict) -> dict:
    """Resolve pull ids, repeat sections and sum patterns. sheets: pull id -> sheet; pulls: sheet -> Pull."""
    s = copy.deepcopy(spec)
    for sec in s.get("sections", []):
        rep = sec.pop("repeat", None)
        if rep:
            tmpl = rep.get("line", {})
            # repeated lines first, then the section's own lines (totals below the months)
            sec["lines"] = [{**_fill(tmpl, v, i), "_rep": True}
                            for i, v in enumerate(_over(rep.get("over"), sheets, pulls))] + sec.get("lines", [])
    keys = [ln.get("key") for sec in s.get("sections", []) for ln in sec.get("lines", []) if ln.get("key")]

    def fix_cell(c, repeated: bool):
        if not isinstance(c, dict):
            return c
        if "pull" in c:
            c = {**c, "pull": sheets.get(c["pull"], c["pull"])}
            if any(x is None for x in list(c.get("row") or []) + list(c.get("col") or [])):
                return None
            p = pulls.get(c["pull"])
            missing = p is not None and (tuple(c.get("row", [])) not in p.row_labels or
                                         tuple(c.get("col", [])) not in p.col_labels)
            if missing and repeated:                 # e.g. 'Oct 26' in a Jan..Dec repeat: empty, not an error
                return None
        for agg in ("sum", "avg"):
            if isinstance(c.get(agg), str):
                c = {**c, agg: [k for k in keys if fnmatch.fnmatchcase(k, c[agg])]}
        return c

    for sec in s.get("sections", []):
        for ln in sec.get("lines", []):
            repeated = bool(ln.pop("_rep", False))
            cells = [fix_cell(c, repeated) for c in ln.get("cells", [])]
            for j, c in enumerate(cells):            # a variance of an empty cell is empty, not '-x'
                if isinstance(c, dict) and ("var" in c or "varpct" in c):
                    a, b = c.get("var") or c.get("varpct")
                    if a >= len(cells) or b >= len(cells) or cells[a] is None or cells[b] is None:
                        cells[j] = None
            ln["cells"] = cells
    if isinstance(s.get("chart", {}).get("lines"), str):
        s["chart"]["lines"] = [k for k in keys if fnmatch.fnmatchcase(k, s["chart"]["lines"])]
    return s


# ------------------------------------------------------------------ evaluation (for drafted key points)
def _pull_value(p, row, col):
    try:
        j = p.layout.col_tuples.index(tuple(col))
    except ValueError:
        return None
    for r in p.parsed.rows:
        if tuple(r[:p.layout.nr]) == tuple(row):
            v = r[p.layout.nr + j]
            return v if isinstance(v, (int, float)) else 0.0
    return None


def _drill_value(rows: list[dict], where: dict, agg: str):
    """Same rule as drill_xl.drill_formula: match per column like Excel's '=' (text, case-insensitive; '' = blank);
    debit = negatives, credit = positives."""
    if agg not in ("sum", "count", "debit", "credit"):
        raise ValueError(f"unknown drill agg '{agg}'")
    key = lambda x: "" if x is None else str(x).casefold()
    names = {str(c).casefold(): c for c in rows[0]} if rows else {}     # column names: case-insensitive
    want = {names.get(str(k).casefold(), k): key(v) for k, v in where.items()}
    hit = [r for r in rows if all(key(r.get(k)) == v for k, v in want.items())]
    amts = [r["AMOUNT"] for r in hit if isinstance(r.get("AMOUNT"), float)]
    if agg == "count":
        return len(hit)
    if agg == "debit":
        return sum(a for a in amts if a < 0)
    if agg == "credit":
        return sum(a for a in amts if a > 0)
    return sum(amts)


def evaluate(spec: dict, pulls: dict, drills: dict[str, list[dict]] | None = None) -> dict[str, list]:
    """Line values computed in Python with the same rules as the sheet formulas (expanded spec).
    drills: alias -> drill rows (from drill_store) for {"drill": ...} cells."""
    lines = [ln for sec in spec.get("sections", []) for ln in sec.get("lines", []) if ln.get("key")]
    ncol = len(spec.get("columns", []))
    vals: dict[str, list] = {ln["key"]: [None] * ncol for ln in lines}
    for _ in range(6):                               # forward references settle in a few passes
        for ln in lines:
            row = vals[ln["key"]]
            for j in range(ncol):
                c = ln.get("cells", [])[j] if j < len(ln.get("cells", [])) else None
                if not isinstance(c, dict):
                    continue
                try:
                    if "drill" in c:
                        rows = (drills or {}).get(c["drill"])
                        where = c.get("where") or {}
                        have = {str(k).casefold() for k in rows[0]} if rows else set()
                        if rows is None or (rows and any(str(k).casefold() not in have for k in where)):
                            row[j] = None            # like the sheet: a column the lines table lacks stays empty
                        else:
                            row[j] = _drill_value(rows, where, c.get("agg", "sum"))
                    elif "pull" in c:
                        p = pulls.get(c["pull"])
                        row[j] = _pull_value(p, c["row"], c["col"]) if p else None
                    elif "var" in c:
                        a, b = row[c["var"][0]], row[c["var"][1]]
                        row[j] = None if a is None or b is None else a - b
                    elif "varpct" in c:
                        a, b = row[c["varpct"][0]], row[c["varpct"][1]]
                        row[j] = None if a is None or not b else (a - b) / abs(b)
                    elif "sum" in c:
                        xs = [vals[k][j] for k in c["sum"] if k in vals]
                        row[j] = None if any(x is None for x in xs) else sum(xs)
                    elif "avg" in c:
                        xs = [vals[k][j] if k in vals else None for k in c["avg"]]     # an unknown key blanks it
                        row[j] = None if not xs or any(x is None for x in xs) else sum(xs) / len(xs)
                    elif "expr" in c:
                        env = {k: (v[j] if v[j] is not None else float("nan")) for k, v in vals.items()}
                        row[j] = eval(re.sub(r"[^\w+\-*/(). ]", "", c["expr"]), {"__builtins__": {}}, env)
                    elif "ref" in c:
                        row[j] = vals.get(c["ref"]["key"], [None] * ncol)[int(c["ref"].get("col", j))]
                    elif "rowsum" in c:
                        xs = [row[k] for k in c["rowsum"]]
                        row[j] = None if any(x is None for x in xs) else sum(xs)
                    elif "value" in c and isinstance(c["value"], (int, float)):
                        row[j] = c["value"]
                except Exception:   # noqa: BLE001 - a draft value is best effort
                    row[j] = None
    return vals


# ------------------------------------------------------------------ drafted key points
def _fmt(v) -> str:
    s = f"{abs(v):,.0f}".replace(",", " ")
    return f"−{s}" if v < -0.5 else s


def _pct(v) -> str:
    return f"{v * 100:+.1f} %".replace(".", ",").replace("-", "−")


def draft_points(spec: dict, vals: dict) -> list[str]:
    """Factual points from the numbers: each grand total (current vs comparison, variance, %), then the largest
    positive and negative variance among detail lines. Interpretation is left to the spec's own key_points."""
    cols = spec.get("columns", [])
    lines = [ln for sec in spec.get("sections", []) for ln in sec.get("lines", [])
             if ln.get("key") and not ln.get("aux")]          # aux = keys, bases and FinMax basis lines

    def var_idx(ln):
        for j, c in enumerate(ln.get("cells", [])):
            if isinstance(c, dict) and "var" in c:
                return j, c["var"]
        return None
    out = []
    for ln in [x for x in lines if x.get("kind") == "grand"][:2]:
        vi = var_idx(ln)
        v = vals.get(ln["key"], [])
        if not vi:
            continue
        j, (a, b) = vi
        if a < len(v) and b < len(v) and None not in (v[a], v[b], v[j]):
            pct = f", {_pct(v[j] / abs(v[b]))}" if v[b] else ""
            out.append(f"{ln.get('label', ln['key'])}: {_fmt(v[a])} ({cols[a].get('header', '')}) vs {_fmt(v[b])} "
                       f"({cols[b].get('header', '')}), variance {_fmt(v[j])}{pct}.")
    cand = []
    for ln in lines:
        if ln.get("kind") in ("sum", "grand"):
            continue
        vi = var_idx(ln)
        v = vals.get(ln["key"], [])
        if vi and vi[0] < len(v) and v[vi[0]] is not None:
            cand.append((v[vi[0]], ln.get("label", ln["key"])))
    if len(cand) >= 3:
        hi, lo = max(cand), min(cand)
        parts = []
        if hi[0] > 0.5:
            parts.append(f"largest increase {hi[1]} {_fmt(hi[0])}")
        if lo[0] < -0.5:
            parts.append(f"largest decrease {lo[1]} {_fmt(lo[0])}")
        if parts:
            text = "; ".join(parts) + "."
            out.append(text[0].upper() + text[1:])
    return out


def prepare(spec: dict, sheets: dict[str, str], pulls: dict, drills: dict[str, list[dict]] | None = None) -> dict:
    """expand + drafted key points (unless auto_points is false). Returns the spec answer.build consumes.
    drills: alias -> drill rows, so {"drill": ...} cells count in the drafted points."""
    s = expand(spec, sheets, pulls)
    if s.pop("auto_points", True):
        drafted = draft_points(s, evaluate(s, pulls, drills))
        s["key_points"] = drafted + list(s.get("key_points") or [])
    return s
