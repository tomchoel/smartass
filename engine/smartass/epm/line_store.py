# engine/smartass/epm/line_store.py
"""Session-scoped SQLite store for FinLite drilled ledger lines: auto-caches every `DrillResult` by cell and supports
structured multi-column filtering/aggregation as well as guarded read-only SQLite SELECT queries (`fin_lines`)."""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
import sqlite3
import time
from pathlib import Path
from typing import Any

from ..core import config, session
from ..core.errors import InputError
from ..model.periods import parse as parse_period
from .drill_lines import AMOUNT, FACETS, FACET_DESC

MAX_AGE_DAYS = 7
MAX_QUERY_ROWS = 500
_SAFE_IDENT = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_PII_COLUMNS = {"GL_DESCRIPTION_LINE", "JOURNAL_NAME", "BATCH_NAME"}
_ALLOWED_FUNCTIONS = {
    "ABS", "AVG", "COALESCE", "COUNT", "GROUP_CONCAT", "IFNULL", "IIF", "INSTR", "LENGTH", "LOWER",
    "LTRIM", "MAX", "MIN", "NULLIF", "PRINTF", "REPLACE", "ROUND", "RTRIM", "SUBSTR", "SUBSTRING",
    "SUM", "TOTAL", "TRIM", "TYPEOF", "UPPER",
}


def _cache_dir() -> Path:
    return config.HOME / "line_cache"


def _session_db_path() -> Path:
    raw_id = str(session.info().get("id") or "default")
    safe_id = re.sub(r"[^A-Za-z0-9_.-]+", "_", raw_id)
    d = _cache_dir()
    d.mkdir(parents=True, exist_ok=True)
    _prune_old(d)
    return d / f"lines_{safe_id}.db"


def _prune_old(directory: Path) -> None:
    cutoff = time.time() - MAX_AGE_DAYS * 86400
    for p in directory.glob("lines_*.db"):
        try:
            if p.stat().st_mtime < cutoff:
                p.unlink(missing_ok=True)
        except OSError:
            pass


def _connect(path: Path | None = None) -> sqlite3.Connection:
    db_path = path or _session_db_path()
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.execute(
        """CREATE TABLE IF NOT EXISTS drills (
             drill_id TEXT PRIMARY KEY,
             cell_key TEXT NOT NULL,
             cell_json TEXT NOT NULL,
             fetched_at REAL NOT NULL,
             cube_value REAL,
             lines_sum REAL,
             gap REAL,
             rec_status TEXT,
             row_count INT NOT NULL,
             columns_json TEXT NOT NULL,
             warnings_json TEXT NOT NULL,
             source TEXT NOT NULL,
             table_name TEXT NOT NULL
           )"""
    )
    conn.execute("CREATE INDEX IF NOT EXISTS ix_drills_cell_key ON drills(cell_key)")
    return conn


def cell_key(cell: dict[str, str], connection: str = "FinLite") -> str:
    canon = sorted((str(k), str(v)) for k, v in cell.items())
    return json.dumps({"connection": connection, "cell": canon}, ensure_ascii=False, sort_keys=True)


def make_drill_id(key: str) -> str:
    return "dl_" + hashlib.sha1(key.encode("utf-8")).hexdigest()[:12]


def _period_ym(label: Any) -> str | None:
    if not isinstance(label, str) or not label.strip():
        return None
    try:
        y, m = parse_period(label)
        return f"{y:04d}-{m:02d}"
    except ValueError:
        return None


def _compute_reversal_flags(rows: list[dict]) -> list[int]:
    """Identify deterministic 1:1 exact-opposite pairs sharing (ACCOUNT, Entity, Period, AP_CUSTOMER_NAME)."""
    buckets: dict[tuple, dict[float, list[int]]] = {}
    flags = [0] * len(rows)
    for idx, r in enumerate(rows):
        amt = r.get(AMOUNT)
        if not isinstance(amt, (int, float)) or amt == 0:
            continue
        ctx = (
            str(r.get("ACCOUNT") or ""),
            str(r.get("Entity") or ""),
            str(r.get("Period") or ""),
            str(r.get("AP_CUSTOMER_NAME") or ""),
        )
        rounded = round(float(amt), 2)
        by_amt = buckets.setdefault(ctx, {})
        opp = by_amt.get(-rounded)
        if opp:
            partner = opp.pop(0)
            flags[idx] = 1
            flags[partner] = 1
        else:
            by_amt.setdefault(rounded, []).append(idx)
    return flags


def _sql_value(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, (bool, int, float)):
        return float(v) if isinstance(v, float) else (int(v) if isinstance(v, bool) else v)
    if isinstance(v, (dt.datetime, dt.date)):
        return v.isoformat()
    return str(v)


def _col_type(col: str, rows: list[dict]) -> str:
    for r in rows[:50]:
        v = r.get(col)
        if v is None:
            continue
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return "REAL"
        break
    return "TEXT"


def store_drill(cell: dict[str, str], rows: list[dict], columns: list[str], *,
                cube_value: float | None, reconciliation: dict, warnings: list[str],
                source: str = "", connection: str = "FinLite", db_path: Path | None = None) -> str:
    """Persist a normalized DrillResult into session SQLite and return its `drill_id`."""
    ckey = cell_key(cell, connection=connection)
    did = make_drill_id(ckey)
    table_name = f"lines_{did}"
    raw_cols = list(dict.fromkeys([*columns, *(k for r in rows[:20] for k in r.keys())]))
    # Ensure safe identifier names and avoid collision with synthetic columns
    all_cols = [c for c in raw_cols if _SAFE_IDENT.match(c)]
    extra_cols = []
    if "period_ym" not in all_cols:
        extra_cols.append(("period_ym", "TEXT"))
    if "reversal_pair" not in all_cols:
        extra_cols.append(("reversal_pair", "INTEGER"))

    rev_flags = _compute_reversal_flags(rows)
    col_defs = [f'"{c}" {_col_type(c, rows)}' for c in all_cols] + [f'"{c}" {t}' for c, t in extra_cols]
    col_names = [*all_cols, *(c for c, _ in extra_cols)]

    with _connect(db_path) as conn:
        conn.execute(f'DROP TABLE IF EXISTS "{table_name}"')
        if col_defs:
            conn.execute(f'CREATE TABLE "{table_name}" ({", ".join(col_defs)})')
            placeholders = ", ".join("?" * len(col_names))
            quoted_cols = ", ".join(f'"{c}"' for c in col_names)
            batch = []
            for idx, r in enumerate(rows):
                vals = [_sql_value(r.get(c)) for c in all_cols]
                for ec, _ in extra_cols:
                    vals.append(_period_ym(r.get("Period")) if ec == "period_ym" else rev_flags[idx])
                batch.append(vals)
            if batch:
                conn.executemany(f'INSERT INTO "{table_name}" ({quoted_cols}) VALUES ({placeholders})', batch)
            for idx_col in ("Period", "period_ym", "Entity", "ACCOUNT", "AP_CUSTOMER_NAME", "JE_SOURCE",
                            "PROJECT", "Products"):
                if idx_col in col_names:
                    conn.execute(f'CREATE INDEX IF NOT EXISTS "ix_{did}_{idx_col}" ON "{table_name}"("{idx_col}")')
        else:
            conn.execute(f'CREATE TABLE "{table_name}" ("empty_placeholder" TEXT)')

        conn.execute(
            """INSERT OR REPLACE INTO drills
               (drill_id, cell_key, cell_json, fetched_at, cube_value, lines_sum, gap, rec_status,
                row_count, columns_json, warnings_json, source, table_name)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                did, ckey, json.dumps(cell, ensure_ascii=False), time.time(),
                reconciliation.get("cube_value", cube_value),
                reconciliation.get("lines_sum"),
                reconciliation.get("gap"),
                reconciliation.get("status"),
                len(rows),
                json.dumps(all_cols),
                json.dumps(list(warnings), ensure_ascii=False),
                source,
                table_name,
            ),
        )
    return did


def get_meta(drill_id: str, db_path: Path | None = None) -> dict | None:
    with _connect(db_path) as conn:
        row = conn.execute(
            "SELECT drill_id, cell_json, fetched_at, cube_value, lines_sum, gap, rec_status, row_count, "
            "columns_json, warnings_json, source, table_name FROM drills WHERE drill_id = ?",
            (drill_id,),
        ).fetchone()
    if row is None:
        return None
    return {
        "drill_id": row[0],
        "cell": json.loads(row[1]),
        "fetched_at": row[2],
        "reconciliation": {
            "cube_value": row[3],
            "lines_sum": row[4],
            "gap": row[5],
            "status": row[6],
            "lines": row[7],
        },
        "columns": json.loads(row[8]),
        "warnings": json.loads(row[9]),
        "source": row[10],
        "table_name": row[11],
    }


def find_cached(cell: dict[str, str], connection: str = "FinLite", max_age_s: float = 3600.0,
                db_path: Path | None = None) -> dict | None:
    ckey = cell_key(cell, connection=connection)
    with _connect(db_path) as conn:
        row = conn.execute(
            "SELECT drill_id, fetched_at FROM drills WHERE cell_key = ? ORDER BY fetched_at DESC LIMIT 1",
            (ckey,),
        ).fetchone()
    if not row or (time.time() - row[1]) > max_age_s:
        return None
    return get_meta(row[0], db_path=db_path)


def list_drills(db_path: Path | None = None) -> list[dict]:
    with _connect(db_path) as conn:
        rows = conn.execute(
            "SELECT drill_id, cell_json, fetched_at, row_count, cube_value, rec_status FROM drills "
            "ORDER BY fetched_at DESC LIMIT 50"
        ).fetchall()
    return [
        {
            "drill_id": d,
            "cell": json.loads(c),
            "fetched_at": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ts)),
            "rows": n,
            "cube_value": cv,
            "reconciliation_status": st,
        }
        for d, c, ts, n, cv, st in rows
    ]


def load_rows(drill_id: str, *, where: dict | None = None, exclude_reversals: bool = False,
              db_path: Path | None = None) -> tuple[dict, list[dict]]:
    """Load filtered raw rows for `LinesBlock` integration in Tier 2."""
    meta = get_meta(drill_id, db_path=db_path)
    if meta is None:
        raise InputError(f"unknown drill_id {drill_id!r}; run fin_drill first")
    sql_where, params = _build_where(meta["columns"], where, exclude_reversals)
    table = meta["table_name"]
    cols = meta["columns"]
    if not cols:
        return meta, []
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.execute(
            f'SELECT {", ".join(f"{chr(34)}{c}{chr(34)}" for c in cols)} FROM "{table}" {sql_where}',
            params,
        )
        rows = [dict(r) for r in cur.fetchall()]
    return meta, rows


def _resolve_col(columns: list[str], name: str) -> str:
    """Resolve user-friendly facet aliases (e.g. 'supplier' -> 'AP_CUSTOMER_NAME') or exact/case-insensitive column."""
    if name in columns or name in ("period_ym", "reversal_pair"):
        return name
    low = name.strip().casefold()
    if low in FACETS:
        target = FACETS[low]
        if target in columns:
            return target
    by_low = {c.casefold(): c for c in columns}
    if low in by_low:
        return by_low[low]
    raise InputError(f"unknown ledger column {name!r}; available: {', '.join(columns)}")


def _build_where(columns: list[str], where: dict | None, exclude_reversals: bool) -> tuple[str, list[Any]]:
    clauses: list[str] = []
    params: list[Any] = []
    if exclude_reversals:
        clauses.append('"reversal_pair" = 0')
    for k, v in (where or {}).items():
        if k == "period_from":
            ym = _period_ym(v)
            if not ym:
                raise InputError(f"invalid period_from {v!r} (expected e.g. 'Jan 26')")
            clauses.append('"period_ym" >= ?')
            params.append(ym)
            continue
        if k == "period_to":
            ym = _period_ym(v)
            if not ym:
                raise InputError(f"invalid period_to {v!r} (expected e.g. 'Oct 26')")
            clauses.append('"period_ym" <= ?')
            params.append(ym)
            continue
        if k == "min_abs_amount":
            clauses.append(f'ABS("{AMOUNT}") >= ?')
            params.append(float(v))
            continue
        col = _resolve_col(columns, k)
        if isinstance(v, (list, tuple)):
            vals = list(v)
            if not vals:
                clauses.append("1 = 0")
            else:
                marks = ", ".join("?" * len(vals))
                clauses.append(f'"{col}" IN ({marks})')
                params.extend(vals)
        elif v is None or v == "":
            clauses.append(f'("{col}" IS NULL OR "{col}" = \'\')')
        else:
            clauses.append(f'"{col}" = ?')
            params.append(v)
    return ("WHERE " + " AND ".join(clauses)) if clauses else "", params


def _desc_for(col: str, available: list[str]) -> str | None:
    for alias, c_name in FACETS.items():
        if c_name == col and alias in FACET_DESC:
            d = FACET_DESC[alias]
            if d in available:
                return d
    if col == "PROJECT" and "PROJECT_DESC" in available:
        return "PROJECT_DESC"
    return None


def _run_read_only_sql(conn: sqlite3.Connection, table_name: str, sql: str, limit: int) -> tuple[list[str], list[dict], int]:
    """Execute user SQL strictly against temp view `lines` with SQLite authorizer blocking writes/meta tables."""
    clean = sql.strip().rstrip(";")
    if not clean:
        raise InputError("sql query cannot be empty")
    if ";" in clean:
        raise InputError("only a single SELECT statement is allowed in `sql`")
    low = clean.casefold()
    if not (low.startswith("select") or low.startswith("with")):
        raise InputError("sql must start with SELECT or WITH")

    conn.execute(f'CREATE TEMP VIEW "lines" AS SELECT * FROM "{table_name}"')

    def authorizer(action, arg1, arg2, db_name, trigger_name):
        if action == sqlite3.SQLITE_SELECT:
            return sqlite3.SQLITE_OK
        if action == sqlite3.SQLITE_READ:
            if db_name == "temp" and arg1 == "lines":
                return sqlite3.SQLITE_OK
            if db_name == "main" and arg1 == table_name:
                return sqlite3.SQLITE_OK
            return sqlite3.SQLITE_DENY
        if action == sqlite3.SQLITE_FUNCTION:
            if (arg2 or "").upper() in _ALLOWED_FUNCTIONS:
                return sqlite3.SQLITE_OK
            return sqlite3.SQLITE_DENY
        return sqlite3.SQLITE_DENY

    conn.set_authorizer(authorizer)
    try:
        cur = conn.execute(clean)
        col_names = [d[0] for d in (cur.description or ())]
        fetched = cur.fetchmany(limit + 1)
        truncated = len(fetched) > limit
        rows_out = [dict(zip(col_names, r)) for r in fetched[:limit]]
        return col_names, rows_out, (limit + 1 if truncated else len(rows_out))
    except sqlite3.DatabaseError as e:
        raise InputError(f"invalid or disallowed SQL query: {e}") from None
    finally:
        conn.set_authorizer(None)


def query_lines(drill_id: str, *, sql: str | None = None, where: dict | None = None,
                group_by: str | list[str] | None = None, order_by: str = "-abs_sum",
                limit: int = 50, offset: int = 0, exclude_reversals: bool = False,
                include_descriptions: bool = True, db_path: Path | None = None) -> dict:
    """Query a session-cached drill (`drill_id`) via structured filters/aggregations OR read-only SQLite SELECT."""
    meta = get_meta(drill_id, db_path=db_path)
    if meta is None:
        raise InputError(f"unknown drill_id {drill_id!r}; call fin_drill first to populate the ledger store")

    lim = max(1, min(int(limit), MAX_QUERY_ROWS))
    off = max(0, int(offset))
    table_name = meta["table_name"]
    cols = meta["columns"]
    warnings: list[str] = []

    if sql is not None:
        with _connect(db_path) as conn:
            col_names, out_rows, counted = _run_read_only_sql(conn, table_name, sql, lim)
        truncated = counted > lim
        if any(c in _PII_COLUMNS for c in col_names):
            warnings.append("Journal text columns projected: aggregate personal names or employee IDs before reporting.")
        return {
            "drill_id": drill_id,
            "mode": "sql",
            "cell": meta["cell"],
            "reconciliation": meta["reconciliation"],
            "columns": col_names,
            "rows": out_rows,
            "row_count": len(out_rows),
            "truncated": truncated,
            "warnings": warnings,
        }

    sql_where, params = _build_where(cols, where, exclude_reversals)
    groups = ([group_by] if isinstance(group_by, str) else list(group_by)) if group_by else []
    resolved_groups = [_resolve_col(cols, g) for g in groups]

    with _connect(db_path) as conn:
        tot_row = conn.execute(
            f'SELECT COUNT(*), COALESCE(SUM("{AMOUNT}"), 0.0) FROM "{table_name}" {sql_where}', params
        ).fetchone()
        matched_lines, matched_sum = int(tot_row[0]), round(float(tot_row[1] or 0.0), 2)
        rev_excluded = None
        if exclude_reversals:
            base_where, base_params = _build_where(cols, where, False)
            rw = conn.execute(
                f'SELECT COUNT(*), COALESCE(SUM("{AMOUNT}"), 0.0) FROM "{table_name}" '
                f'{base_where + (" AND " if base_where else "WHERE ")}"reversal_pair" = 1',
                base_params,
            ).fetchone()
            rev_excluded = {"lines": int(rw[0]), "net_amount": round(float(rw[1] or 0.0), 2)}

        if resolved_groups:
            desc_cols = []
            if include_descriptions:
                for gc in resolved_groups:
                    d = _desc_for(gc, cols)
                    if d and d not in resolved_groups and d not in desc_cols:
                        desc_cols.append(d)
            sel_group = [f'COALESCE("{c}", \'\') AS "{c}"' for c in resolved_groups]
            sel_desc = [f'MAX("{d}") AS "{d}"' for d in desc_cols]
            aggs = [
                f'ROUND(COALESCE(SUM("{AMOUNT}"), 0.0), 2) AS "net_amount"',
                f'ROUND(ABS(COALESCE(SUM("{AMOUNT}"), 0.0)), 2) AS "abs_amount"',
                f'ROUND(COALESCE(SUM(CASE WHEN "{AMOUNT}" < 0 THEN "{AMOUNT}" ELSE 0 END), 0.0), 2) AS "debit"',
                f'ROUND(COALESCE(SUM(CASE WHEN "{AMOUNT}" > 0 THEN "{AMOUNT}" ELSE 0 END), 0.0), 2) AS "credit"',
                'COUNT(*) AS "line_count"',
            ]
            order_map = {
                "-abs_sum": '"abs_amount" DESC, "line_count" DESC',
                "abs_sum": '"abs_amount" ASC',
                "-net": '"net_amount" DESC',
                "net": '"net_amount" ASC',
                "-count": '"line_count" DESC',
            }
            order_sql = order_map.get(order_by, '"abs_amount" DESC')
            group_expr = ", ".join(f'COALESCE("{c}", \'\')' for c in resolved_groups)
            query = (
                f'SELECT {", ".join([*sel_group, *sel_desc, *aggs])} FROM "{table_name}" {sql_where} '
                f'GROUP BY {group_expr} ORDER BY {order_sql}'
            )
            cur_all = conn.execute(query + " LIMIT ? OFFSET ?", [*params, lim + 1, off])
            col_names = [d[0] for d in cur_all.description]
            raw_fetched = cur_all.fetchall()
            truncated = len(raw_fetched) > lim
            rows_out = []
            denom = matched_sum if abs(matched_sum) >= 0.01 else None
            for tup in raw_fetched[:lim]:
                rec = dict(zip(col_names, tup))
                rec["share_pct"] = round(rec["net_amount"] / denom, 4) if denom is not None else None
                rows_out.append(rec)
            total_groups = conn.execute(f"SELECT COUNT(*) FROM ({query})", params).fetchone()[0]
            out_cols = [*col_names, "share_pct"]
        else:
            order_sql = f'ABS("{AMOUNT}") DESC' if order_by == "-abs_sum" else f'"{AMOUNT}" DESC'
            sel_cols = [*cols, "reversal_pair"] if "reversal_pair" not in cols else cols
            q_cols = ", ".join(f'"{c}"' for c in sel_cols)
            cur_rows = conn.execute(
                f'SELECT {q_cols} FROM "{table_name}" {sql_where} ORDER BY {order_sql} LIMIT ? OFFSET ?',
                [*params, lim + 1, off],
            )
            col_names = [d[0] for d in cur_rows.description]
            raw_fetched = cur_rows.fetchall()
            truncated = len(raw_fetched) > lim
            rows_out = [dict(zip(col_names, tup)) for tup in raw_fetched[:lim]]
            total_groups = matched_lines
            out_cols = col_names
            if any(c in _PII_COLUMNS for c in out_cols):
                warnings.append("Individual journal lines may contain personal details in description fields: "
                                "aggregate before citing.")

    parent_sum = meta["reconciliation"].get("lines_sum")
    share_of_drill = round(matched_sum / parent_sum, 4) if parent_sum else None
    return {
        "drill_id": drill_id,
        "mode": "grouped" if resolved_groups else "rows",
        "cell": meta["cell"],
        "reconciliation": meta["reconciliation"],
        "filter_summary": {
            "matched_lines": matched_lines,
            "matched_sum": matched_sum,
            "share_of_drill": share_of_drill,
            **({"excluded_reversal_pairs": rev_excluded} if rev_excluded is not None else {}),
        },
        "columns": out_cols,
        "rows": rows_out,
        "total_results": int(total_groups),
        "offset": off,
        "limit": lim,
        "truncated": truncated,
        "warnings": warnings,
    }
