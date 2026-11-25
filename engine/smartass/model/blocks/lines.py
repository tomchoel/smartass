"""lines: the ledger lines behind one cube cell, grouped by supplier (AP_CUSTOMER_NAME), project (INITIATIVE) or
account (ACCOUNT): top N (biggest absolute amount first; spellings that differ only in case are one group), Other,
lines without a <group> (suppressed when 0 lines; for supplier, top 5 journal/source groups indented underneath),
the lines total, the cube value and the gap (cube - lines), plus Share % (uniform signs) or Debit | Credit (mixed
signs) and Lines count."""
from __future__ import annotations

import re

from ..expr import CubeRef, LinePick, Ref, Sub, Sum, share
from ..lines import AMOUNT, norm
from ..spec import LinesBlock
from .common import BlockOut, Line, ck

COLUMN = {
    "supplier": "AP_CUSTOMER_NAME", "project": "INITIATIVE", "account": "ACCOUNT",
    "cost_centre": "Entity", "entity": "Entity", "product": "Products", "period": "Period",
    "je_source": "JE_SOURCE", "journal": "JOURNAL_NAME", "cp": "CP",
}
NO_GROUP = {
    "supplier": "Lines without a supplier",
    "project": "Lines without a booked project",
    "account": "Lines without an account",
}


def _col_for(tok: str) -> str:
    return COLUMN.get(tok.strip().casefold(), tok)


def _filter_rows(rows: list[dict], where: dict, exclude_reversals: bool) -> list[dict]:
    """Filter raw LineSet rows deterministically in memory when LinesBlock specifies `where` or `exclude_reversals`."""
    from ...epm.drill_lines import FACETS
    from ..periods import parse as parse_period
    out = rows
    if exclude_reversals:
        from ...epm.line_store import _compute_reversal_flags
        flags = _compute_reversal_flags(out)
        out = [r for r, f in zip(out, flags) if not f]
    if not where:
        return out

    def ym(val) -> str | None:
        if not isinstance(val, str) or not val.strip():
            return None
        try:
            y, m = parse_period(val)
            return f"{y:04d}-{m:02d}"
        except ValueError:
            return None

    res = []
    for r in out:
        keep = True
        for k, expected in where.items():
            if k == "period_from":
                cur_ym, lo_ym = ym(r.get("Period")), ym(expected)
                if cur_ym is None or (lo_ym and cur_ym < lo_ym):
                    keep = False
                    break
                continue
            if k == "period_to":
                cur_ym, hi_ym = ym(r.get("Period")), ym(expected)
                if cur_ym is None or (hi_ym and cur_ym > hi_ym):
                    keep = False
                    break
                continue
            col = FACETS.get(k.casefold(), k)
            val = r.get(col)
            if isinstance(expected, (list, tuple)):
                if val not in expected:
                    keep = False
                    break
            elif expected is None or expected == "":
                if val not in (None, ""):
                    keep = False
                    break
            elif str(val or "") != str(expected):
                keep = False
                break
        if keep:
            res.append(r)
    return res
TOP_NO_SUPPLIER_JOURNALS = 5
_REV_PREFIX = re.compile(r"^Reverses\s+", re.IGNORECASE)
_CODE_YYMM = re.compile(r"^([A-Za-zÆØÅæøå]+)2\d{3}$")


def _amt(r: dict) -> float:
    v = r.get(AMOUNT)
    if v is None or v == "" or isinstance(v, bool):
        return 0.0
    return float(v)


def _dominant(stats: dict[str, list[float]]) -> str:
    return max(stats, key=lambda k: (abs(stats[k][0]), stats[k][1], k))


def _record_stat(stats: dict[str, list[float]], key: str, amt: float) -> None:
    if not key:
        return
    entry = stats.setdefault(key, [0.0, 0.0])
    entry[0] += amt
    entry[1] += abs(amt)


def _account_code(acc_stats: dict[str, list[float]], *, multi: bool = False) -> str:
    if not acc_stats:
        return ""
    dom = _dominant(acc_stats)
    if multi and len(acc_stats) > 1:
        return f"{dom} +{len(acc_stats) - 1}"
    return dom


def _journal_group(r: dict, has_src: bool, has_jrn: bool) -> tuple[str, str, str, str | None, tuple[tuple[str, tuple[str, ...]], ...]]:
    src = str(r.get("JE_SOURCE") or "").strip()
    jrn = str(r.get("JOURNAL_NAME") or "").strip()
    if src and src.casefold() != "spreadsheet":
        return src, "JE_SOURCE", src, None, ()
    if jrn:
        s = _REV_PREFIX.sub("", jrn).strip()
        parts = s.split()
        if parts:
            p0 = parts[0]
            m = _CODE_YYMM.match(p0)
            if m:
                label = m.group(1)
            elif p0.isalpha() and len(parts) > 1 and re.fullmatch(r"\d{1,2}", parts[1]):
                label = f"{p0} {parts[1]}"
            else:
                label = p0
        else:
            label = jrn
        return label, "JOURNAL_NAME", jrn, (src if has_src else None), ()
    if src:
        fixed = (("JOURNAL_NAME", ("",)),) if has_jrn else ()
        return src, "JE_SOURCE", src, None, fixed
    if has_jrn:
        fixed = (("JE_SOURCE", ("",)),) if has_src else ()
        return "(no journal/source)", "JOURNAL_NAME", "", None, fixed
    return "(no journal/source)", "JE_SOURCE", "", None, ()


def build_lines(b: LinesBlock, ctx, key: str) -> BlockOut:
    ls = ctx.frame.lines[key]
    filtered = bool(getattr(b, "where", None) or getattr(b, "exclude_reversals", False))
    rows = _filter_rows(ls.rows, getattr(b, "where", None) or {}, bool(getattr(b, "exclude_reversals", False)))
    if filtered:
        ls.rows = rows

    group_list = [b.group_by] if isinstance(b.group_by, str) else list(b.group_by)
    primary = group_list[0]
    col = _col_for(primary)
    extra_group_cols = [_col_for(g) for g in group_list[1:]]
    amts = [_amt(r) for r in rows if r.get(AMOUNT) not in (None, "")]
    signs_mixed = any(x < 0 for x in amts) and any(x > 0 for x in amts)

    groups: dict[str, dict] = {}
    blank_rows: list[dict] = []
    for r in rows:
        name = str(r.get(col) or "").strip()
        if extra_group_cols:
            parts = [name, *(str(r.get(ec) or "").strip() for ec in extra_group_cols)]
            composite = " · ".join(p or "(blank)" for p in parts)
            raw_vals = tuple(str(r.get(c) or "") for c in (col, *extra_group_cols))
        else:
            composite = name
            raw_vals = (name,)
        amt = _amt(r)
        if not name and not extra_group_cols:
            blank_rows.append(r)
            continue
        g = groups.setdefault(norm(composite), {"name": composite, "amt": 0.0, "spellings": set(),
                                                "tuples": set(), "names": {}, "accs": {}, "projs": {}, "descs": {}})
        g["amt"] += amt
        g["spellings"].add(name)
        g["tuples"].add(raw_vals)
        _record_stat(g["names"], str(r.get("ACCOUNT_NAME") or "").strip(), amt)
        _record_stat(g["accs"], str(r.get("ACCOUNT") or "").strip(), amt)
        for dc in ("CC_DESC", "PRODUCT_DESC", "CP"):
            if dc in r and r.get(dc):
                _record_stat(g["descs"], str(r.get(dc)).strip(), amt)
        proj = str(r.get("PROJECT") or "").strip()
        if proj and proj.strip("0"):
            _record_stat(g["projs"], proj, amt)

    ranked = sorted(groups.values(), key=lambda g: -abs(g["amt"]))
    shown, rest = ranked[:b.top], ranked[b.top:]

    if signs_mixed:
        headers = ["Amount", "Debit", "Credit", "Lines"]
        cols = ["amount", "debit", "credit", "lines"]
    else:
        headers = ["Amount", "Share %", "Lines"]
        cols = ["amount", "share", "lines"]
    g_label = ", ".join(group_list) if len(group_list) > 1 else primary
    out = BlockOut(key, "lines", b.title or f"Ledger lines by {g_label}", headers, cols=cols)
    total = Ref(ck(key, "total", 0))
    lines_col = len(cols) - 1

    def row_cells(filter_col: str | None, vals: tuple[str, ...] | None, lkey: str,
                  extra: tuple[tuple[str, tuple[str, ...]], ...] = ()) -> list:
        amt_pick = LinePick(key, filter_col, vals, agg="sum", extra=extra)
        cnt_pick = LinePick(key, filter_col, vals, agg="count", extra=extra)
        if signs_mixed:
            return [amt_pick,
                    LinePick(key, filter_col, vals, agg="debit", extra=extra),
                    LinePick(key, filter_col, vals, agg="credit", extra=extra),
                    cnt_pick]
        return [amt_pick, share(Ref(ck(key, lkey, 0)), total), cnt_pick]

    for i, g in enumerate(shown):
        lkey = f"g{i}"
        spellings = tuple(sorted(g["spellings"]))
        if not extra_group_cols and primary == "account":
            code = _account_code(g["accs"]) or g["name"]
            label = _dominant(g["names"]) if g["names"] else code
        elif not extra_group_cols and primary == "project":
            code = _dominant(g["projs"]) if g["projs"] else ""
            label = g["name"]
        else:
            code = _account_code(g["accs"], multi=True)
            desc = _dominant(g["descs"]) if (not extra_group_cols and g["descs"] and primary not in ("supplier",)) else ""
            label = f"{g['name']} — {desc}" if desc and desc != g["name"] else g["name"]
        if extra_group_cols:
            all_cols = (col, *extra_group_cols)
            first_vals = tuple(sorted({t[0] for t in g["tuples"]}))
            extras = tuple((c, tuple(sorted({t[idx] for t in g["tuples"]}))) for idx, c in enumerate(all_cols[1:], 1))
            out.lines.append(Line(lkey, label, row_cells(all_cols[0], first_vals, lkey, extra=extras), code=code))
        else:
            out.lines.append(Line(lkey, label, row_cells(col, spellings, lkey), code=code))

    has_blank = bool(blank_rows)
    if rest:
        def other_col(j: int):
            named = Sum(tuple(Ref(ck(key, f"g{i}", j)) for i in range(len(shown))))
            tref = Ref(ck(key, "total", j))
            return Sub(Sub(tref, Ref(ck(key, "blank", j))), named) if has_blank else Sub(tref, named)

        if signs_mixed:
            cells = [other_col(0), other_col(1), other_col(2), other_col(3)]
        else:
            cells = [other_col(0), share(Ref(ck(key, "other", 0)), total), other_col(lines_col)]
        out.lines.append(Line("other", "Other", cells))

    if has_blank:
        blank_label = NO_GROUP.get(primary, f"Lines without {primary}")
        out.lines.append(Line("blank", blank_label, row_cells(col, ("",), "blank")))
        if primary == "supplier" and any(str(r.get("JE_SOURCE") or "").strip()
                                         or str(r.get("JOURNAL_NAME") or "").strip() for r in blank_rows):
            has_src = "JE_SOURCE" in ls.columns or any("JE_SOURCE" in r for r in blank_rows)
            has_jrn = "JOURNAL_NAME" in ls.columns or any("JOURNAL_NAME" in r for r in blank_rows)
            jgroups: dict[tuple[str, str, tuple], dict] = {}
            for r in blank_rows:
                jlabel, fcol, raw_val, src_filter, fixed_extra = _journal_group(r, has_src, has_jrn)
                amt = _amt(r)
                jg = jgroups.setdefault((fcol, norm(jlabel), fixed_extra),
                                        {"label": jlabel, "fcol": fcol, "fixed": fixed_extra,
                                         "amt": 0.0, "vals": set(), "srcs": set(), "accs": {}})
                jg["amt"] += amt
                jg["vals"].add(raw_val)
                if src_filter is not None:
                    jg["srcs"].add(src_filter)
                _record_stat(jg["accs"], str(r.get("ACCOUNT") or "").strip(), amt)
            top_j = sorted(jgroups.values(), key=lambda g: -abs(g["amt"]))[:TOP_NO_SUPPLIER_JOURNALS]
            for j, jg in enumerate(top_j):
                ckey = f"blank_{j}"
                extra_parts: list[tuple[str, tuple[str, ...]]] = [(jg["fcol"], tuple(sorted(jg["vals"])))]
                if jg["srcs"]:
                    extra_parts.append(("JE_SOURCE", tuple(sorted(jg["srcs"]))))
                extra_parts.extend(jg["fixed"])
                jcode = _account_code(jg["accs"], multi=True)
                out.lines.append(Line(ckey, jg["label"], row_cells(col, ("",), ckey, extra=tuple(extra_parts)),
                                      level=1, code=jcode))

    if signs_mixed:
        total_cells = [LinePick(key, None, None, agg="sum"),
                       LinePick(key, None, None, agg="debit"),
                       LinePick(key, None, None, agg="credit"),
                       LinePick(key, None, None, agg="count")]
    else:
        total_cells = [LinePick(key, None, None, agg="sum"), None, LinePick(key, None, None, agg="count")]
    pad = [None] * (len(cols) - 1)
    out.lines.append(Line("total", "Filtered lines total" if filtered else "Ledger lines total", total_cells, kind="sum"))
    out.lines.append(Line("cube", "Cube value", [CubeRef(key), *pad]))
    if not filtered:
        out.lines.append(Line("gap", "Gap (cube - lines)", [Sub(Ref(ck(key, "cube", 0)), total), *pad]))
    return out
