"""The lines table of an answer ('Q<n> Lines <alias>', Excel table Lines_Q<n>_<alias>): its name, the look
constants of the built file's Lines sheet (render/xlsx.py), and the formulas answer cells use over it. Formula rules
(tn-formelstil): SUMIFS by default; an exact-match FILTER form where SUMIFS would not match the key literally
('Jul-26' -> a date, digit strings over 15 characters -> a rounded number, '0123' -> 123, * ? ~ -> wildcards, a
leading < > = -> a comparison, '12:30' -> a time, '50%' -> 0.5, 'TRUE' -> a logical, '1,5' -> a number in a Norwegian
locale, over 255 characters -> SUMIFS fails)."""
from __future__ import annotations

import datetime as dt
import re

from ..core.errors import InputError

_DATE_LIKE = re.compile(r"^\s*(\d{1,4}[-./]\d{1,2}([-./]\d{1,4})?|[A-Za-z]{3,9}[-./ ]\d{1,4}"
                        r"|\d{1,2}[-./ ][A-Za-z]{3,9}([-./ ]\d{1,4})?)\s*$")
_PATTERN_CHARS = set("*?~")
_COERCED = re.compile(r"^\s*(\d{1,2}:\d{2}.*|.*%\s*|true|false|-?\d+,\d+)\s*$", re.IGNORECASE)
AGGS = ("sum", "count", "debit", "credit")


def table_name(q: int, alias: str = "d") -> str:
    return f"Lines_Q{q}" if alias == "d" else f"Lines_Q{q}_{alias}"


def _lit(v) -> str:
    return '"' + str(v).replace('"', '""') + '"'


def safe_for_sumifs(v) -> bool:
    """True when a SUMIFS/COUNTIFS criterion "<v>" matches exactly the cells holding the text v."""
    s = str(v)
    if len(s) > 255 or _DATE_LIKE.match(s) or _COERCED.match(s) or _PATTERN_CHARS & set(s) \
            or s[:1] in ("<", ">", "="):
        return False
    return not (s.isdigit() and (len(s) > 15 or (len(s) > 1 and s[0] == "0")))


def drill_formula(table: str, where: dict[str, str | list[str]], agg: str = "sum") -> str:
    """Answer-cell formula over a lines table. where: column label -> exact value ('' = blank), or a list of values
    (any of them: an OR via ISNUMBER(MATCH(...)) in the FILTER form).
    agg: sum | count | debit (sum of negatives) | credit (sum of positives)."""
    if agg not in AGGS:
        raise ValueError(f"unknown drill agg '{agg}' (use one of {', '.join(AGGS)})")
    where = {k: "" if v is None else v for k, v in where.items()}     # None = blank, like specs._drill_value
    empty = [k for k, v in where.items() if isinstance(v, list) and not v]
    if empty:
        raise InputError(f"drill where on {', '.join(empty)}: an empty list of values matches nothing")
    amt = f"{table}[AMOUNT]"
    sign = {"debit": '"<0"', "credit": '">0"'}.get(agg)
    if all(not isinstance(v, list) and safe_for_sumifs(v) for v in where.values()):
        crit = "".join(f",{table}[{k}],{_lit(v)}" for k, v in where.items())
        if agg == "count":
            # ROWS = every line, like COUNTIFS and Python's count (a lines table is never written with 0 rows)
            return f"=COUNTIFS({crit[1:]})" if where else f"=ROWS({amt})"
        if sign:
            return f"=SUMIFS({amt},{amt},{sign}{crit})"
        return f"=SUMIFS({amt}{crit})" if where else f"=SUM({amt})"
    conds = [f"ISNUMBER(MATCH({table}[{k}],{{{','.join(_lit(x) for x in v)}}},0))" if isinstance(v, list)
             else f"{table}[{k}]={_lit(v)}" for k, v in where.items()]
    if sign:
        conds.append(f"{amt}{sign.strip(chr(34))}")
    test = conds[0] if len(conds) == 1 else "*".join(f"({c})" for c in conds)
    if agg == "count":
        return f"=SUM(--({test}))" if len(conds) == 1 else f"=SUM({test})"
    return f"=SUM(FILTER({amt},{test},0))"


# ------------------------------------------------------------------ the lines sheet
AMT_FMT = '#,##0_);(#,##0);"–"_)'
AMOUNT_LABELS = ("AMOUNT", "DEBET_AMOUNT", "CREDIT_AMOUNT", "TRANS_AMOUNT", "Orig_Curr_AMOUNT")
DATE_LABELS = ("GL_POSTED_DATE", "BOOKING_DATE")
TEXT_LABELS = {"ACCOUNT", "AP_INVOICE_NUMBER", "AP_VOUCHER_NUMBER", "GL_VOUCHER_NUMBER", "PROJECT", "LINE",
               "Period", "SYSTEM", "RESERVE", "FUTURE1"}
WIDTHS = {"GL_DESCRIPTION_LINE": 45, "AP_CUSTOMER_NAME": 30, "ACCOUNT_NAME": 34, "BATCH_NAME": 34,
          "JOURNAL_NAME": 34, "PRODUCT_DESC": 30, "CC_DESC": 28, "PROJECT_DESC": 28}
def header_width(name: str) -> int:
    """A column wide enough for its bold header name ('GL_POSTED_DATE' was cut to 'GL_POSTED_DA'): the table's
    filter button takes ~3 units."""
    from .vendor.tn_rapportstil import text_units
    return int(text_units(name, 11, True)) + 4
_EXCEL_EPOCH = dt.datetime(1899, 12, 30)
