"""Pure logic for drill-through lines (no Excel, no browser): the drill page's columns, typed values, repair of
Norwegian text that EPM delivers broken, reconciliation to the cube, and a profile Claude uses to pick a report.

Lines are dicts keyed by the drill page's column LABEL ("AMOUNT", "AP_CUSTOMER_NAME", ...), in page order."""
from __future__ import annotations

import collections
import datetime as dt
import json
import re
from pathlib import Path

AMOUNT = "AMOUNT"
NUM_FIELDS = {"AMOUNTX", "DESC1", "DESC2", "ATTR10", "ATTR30"}      # amount, debit, credit, trans/orig amount
DATE_FIELDS = {"ATTR4", "ATTR17"}                                    # booking date, GL posted date
UTC_FIELDS = {"ATTR17"}          # GL posted date is a UTC timestamp -> Oslo wall clock; booking date is a plain date
KEY_LABELS = ["Period", "GL_POSTED_DATE", "BOOKING_DATE", "GL_VOUCHER_NUMBER", "Entity", "CC_DESC", "ACCOUNT",
              "ACCOUNT_NAME", "AMOUNT", "JE_SOURCE", "JE_CATEGORY_NAME", "AP_CUSTOMER_NAME", "AP_INVOICE_NUMBER",
              "GL_DESCRIPTION_LINE", "Products", "PRODUCT_DESC", "PROJECT", "PROJECT_DESC", "BATCH_NAME",
              "JOURNAL_NAME", "INITIATIVE", "INITIATIVE_DERIVED"]     # the last two: recipe lines_ranking
FACETS = {"supplier": "AP_CUSTOMER_NAME", "account": "ACCOUNT", "cost_centre": "Entity", "month": "Period",
          "je_source": "JE_SOURCE", "je_category": "JE_CATEGORY_NAME", "product": "Products",
          "invoice": "AP_INVOICE_NUMBER"}
FACET_DESC = {"account": "ACCOUNT_NAME", "cost_centre": "CC_DESC", "product": "PRODUCT_DESC",
              "invoice": "AP_CUSTOMER_NAME"}
_MOJIBAKE = re.compile("[\u00c3\u00c2][\u0080-\u00bf]")    # '\u00c3\u00a5' = UTF-8 'å' read as Latin-1
_LOST = "\ufffd"                                            # U+FFFD: a character EPM lost


def columns_from_view(view: dict, name: str = "DrillFilter") -> list[tuple[str, str]]:
    """(label, field) pairs from /aif/ui/model/view?page=DRILL_THROUGH (the page's default view)."""
    items = view.get("items") or []
    v = next((x for x in items if x.get("viewName") == name), items[0] if items else None)
    if v is None:
        raise ValueError("The drill view definition has no views.")
    return [(c["label"], c["value"]) for c in json.loads(v["viewDefinition"])]


def load_text_fixes(*paths: Path) -> dict[str, str]:
    """Merge correction lists in order; a later file wins (shipped list first, then the user's own)."""
    out: dict[str, str] = {}
    for path in paths:
        try:
            out.update(json.loads(Path(path).read_text(encoding="utf-8")).get("fixes", {}))
        except (OSError, ValueError, AttributeError):
            pass
    return out


def repair_text(s: str, fixes: dict[str, str], stats: dict) -> str:
    """Undo double-encoded UTF-8 exactly; replace fragments EPM lost (U+FFFD) from the correction list; report
    whatever is still broken. Good text is returned unchanged."""
    if _MOJIBAKE.search(s):
        try:
            s = s.encode("latin-1").decode("utf-8")
            stats["fixed_mojibake"] += 1
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    if _LOST in s:
        for bad, good in fixes.items():
            s = s.replace(bad, good)
        if _LOST in s:
            stats["unfixed"].add(s)
        else:
            stats["fixed_by_list"] += 1
    return s


def _oslo(d: dt.datetime) -> dt.datetime:
    """UTC -> Europe/Oslo wall clock (naive). Without tzdata (Windows) apply the EU rule: CET, and CEST from the last
    Sunday of March to the last Sunday of October, switching at 01:00 UTC. Never this PC's zone (CI, travel)."""
    try:
        from zoneinfo import ZoneInfo
        return d.replace(tzinfo=dt.timezone.utc).astimezone(ZoneInfo("Europe/Oslo")).replace(tzinfo=None)
    except Exception:   # noqa: BLE001 - ImportError / ZoneInfoNotFoundError when tzdata is missing
        pass

    def last_sunday_0100(month: int) -> dt.datetime:
        last = dt.datetime(d.year, month, 31, 1)
        return last - dt.timedelta(days=(last.weekday() + 1) % 7)
    summer = last_sunday_0100(3) <= d < last_sunday_0100(10)
    return d + dt.timedelta(hours=2 if summer else 1)


def _value(field: str, v):
    if v is None or v == "":
        return None
    if field in NUM_FIELDS:
        try:
            return float(v)
        except (TypeError, ValueError):
            return str(v)
    if field in DATE_FIELDS:
        try:
            d = dt.datetime.fromisoformat(str(v)[:19])
            return _oslo(d) if field in UTC_FIELDS else d
        except ValueError:
            return str(v)
    return v if isinstance(v, str) else str(v)


def normalize(items: list[dict], columns: list[tuple[str, str]], fixes: dict[str, str]) -> tuple[list[dict], dict]:
    """Raw getData items -> rows keyed by label, typed, text repaired. Returns (rows, text_repair_stats)."""
    stats = {"fixed_mojibake": 0, "fixed_by_list": 0, "unfixed": set()}
    rows = []
    for x in items:
        r = {}
        for label, field in columns:
            v = _value(field, x.get(field))
            r[label] = repair_text(v, fixes, stats) if isinstance(v, str) else v
        rows.append(r)
    stats["unfixed"] = sorted(stats["unfixed"])
    return rows, stats


def _amounts(rows: list[dict]) -> list[float]:
    return [r[AMOUNT] for r in rows if isinstance(r.get(AMOUNT), float)]


def reconcile(cube_value, rows: list[dict]) -> dict:
    """Cube cell vs sum of lines. debit = sum of negative amounts, credit = sum of positive (FinMax signs).
    non_numeric = lines whose amount is blank or not a number (left out of the sum)."""
    a = _amounts(rows)
    s = round(sum(a), 2)
    cube = None if cube_value is None else round(float(cube_value), 2)
    gap = None if cube is None else round(cube - s, 2)
    status = "no cube value" if gap is None else ("exact" if abs(gap) < 0.005 else "gap")
    return {"cube_value": cube, "lines_sum": s, "gap": gap, "status": status, "lines": len(rows),
            "debit": round(sum(x for x in a if x < 0), 2), "credit": round(sum(x for x in a if x > 0), 2),
            "non_numeric": len(rows) - len(a)}


def _facet(rows: list[dict], label: str, top: int, desc: str | None = None) -> dict:
    acc: dict = collections.defaultdict(lambda: [0.0, 0, None, 0.0, 0.0])     # amount, lines, desc, debit, credit
    for r in rows:
        e = acc[r.get(label) or ""]
        x = r[AMOUNT] if isinstance(r.get(AMOUNT), float) else 0.0
        e[0] += x
        e[1] += 1
        if desc and not e[2]:
            e[2] = r.get(desc)
        e[3 if x < 0 else 4] += x
    ordered = sorted(acc.items(), key=lambda kv: -abs(kv[1][0]))
    out = []
    for k, (amount, n, d, debit, credit) in ordered[:top]:
        g = {"key": k, "amount": round(amount, 2), "lines": n, "debit": round(debit, 2), "credit": round(credit, 2)}
        if desc:
            g["desc"] = d
        out.append(g)
    return {"groups": len(acc), "top": out}


def profile(rows: list[dict], top: int = 10) -> dict:
    """What Claude needs to choose report sections: top groups per facet by |amount|, lines without supplier and
    the journals behind them (allocations, accruals), whether signs are mixed (then no 'share of total').
    Each group: key, net amount, gross debit (negatives) and credit (positives), lines (+ desc where known)."""
    p = {name: _facet(rows, label, top, FACET_DESC.get(name)) for name, label in FACETS.items()}
    total_amt = sum(_amounts(rows))

    def share(sub: list[dict]) -> dict:              # spec 3: share of the lines and of the (net) amount
        amt = round(sum(_amounts(sub)), 2)
        return {"lines": len(sub), "amount": amt, "share_lines": round(len(sub) / len(rows), 4) if rows else None,
                "share_amount": round(amt / total_amt, 4) if total_amt else None}
    no_sup = [r for r in rows if not r.get("AP_CUSTOMER_NAME")]
    p["no_supplier"] = share(no_sup)
    p["journals_without_supplier"] = _facet(no_sup, "JOURNAL_NAME", top)
    untagged = [r for r in rows if r.get("Products") in (None, "pd.No_Product")]
    p["untagged_product"] = share(untagged)
    a = _amounts(rows)
    p["signs_mixed"] = any(x < 0 for x in a) and any(x > 0 for x in a)
    return p
