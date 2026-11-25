# engine/smartass/data/slice.py
"""Tier 1 read-only cube exploration: hierarchy browsing (`fin_tree`) and arbitrary multi-dimensional slicing
(`fin_slice`) over the dual EPM engines with closed-month guards, alias resolution, and knowledge caveats."""
from __future__ import annotations

import contextlib
from typing import Any

from ..core.errors import InputError
from ..epm.grid import Axis
from ..knowledge import atlas
from ..model.periods import ABBR, closed_for, label as period_name, parse as parse_period, resolve as resolve_period, shift
from ..model.resolver import UnknownMetric
from ..model.validate import GuardError

CUBES = ("FinMax", "FinLite", "FinPlan")
_FINPLAN_MONTHS = {m: f"FY{y % 100:02d}" for y in range(2020, 2040) for m in ABBR}
_ACCT_DIMS = {"all accounts", "accounts", "account"}
_ENTITY_DIMS = {"all entities", "entities", "entity"}
_PERIOD_DIMS = {"all time periods", "time periods", "period", "periods"}


def _norm_dim(name: str) -> str:
    return name.removeprefix("All ").replace(" ", "").casefold()


def _canon_dim(cube: str, dim: str) -> str:
    """Resolve friendly dimension shorthands ('Accounts', 'Counterparts', 'Entity') per cube."""
    low = dim.strip().casefold()
    if cube == "FinPlan":
        if low in _ACCT_DIMS:
            return "Account"
        if low in _ENTITY_DIMS:
            return "Entity"
        if low in _PERIOD_DIMS:
            return "Period"
    else:
        if low in _ACCT_DIMS:
            return "All Accounts"
        if low in _ENTITY_DIMS:
            return "All Entities"
        if low in _PERIOD_DIMS:
            return "All Time Periods"
    return dim


def _period_dim(cube: str) -> str:
    return "Period" if cube == "FinPlan" else "All Time Periods"


def browse_tree(resolver, *, cube: str | None = None, dimension: str | None = None, parent: str | None = None,
                query: str | None = None, depth: int = 1, leaves_only: bool = False, limit: int = 100) -> dict:
    """Inspect cube dimensions, parent-child subtrees, card attributes, and caveats from atlas.db."""
    if cube is not None and cube not in CUBES:
        raise GuardError(f"unknown cube {cube!r}", f"choose one of {', '.join(CUBES)}")
    db = resolver.db
    if not parent and not query and not dimension:
        return {"dimensions": atlas.dimensions_info(db, cube)}

    if cube is None:
        if parent:
            hits = [c for c in CUBES if atlas.has_member(db, c, _guess_dim(db, c, parent) or "", parent)]
            cube = hits[0] if hits else "FinMax"
        else:
            cube = "FinMax"

    dim = _resolve_dim_name(db, cube, dimension) if dimension else None
    if query and not parent:
        hits = atlas.search(db, query, connection=cube, dimension=dim, limit=limit)
        items = [_enrich_member(resolver, cube, h) for h in hits]
        caveats = resolver.caveats({cube: {h["name"] for h in hits}}) if hits else []
        return {"cube": cube, "dimension": dim, "query": query, "members": items, "caveats": caveats}

    root_in = parent or dim
    canon_root = _resolve_member(resolver, db, cube, dim, root_in) if root_in else root_in
    if canon_root is None:
        raise GuardError(f"{cube}: unknown member or parent {root_in!r}",
                         "check spelling with fin_tree(query=...) or fin_lookup(kind='member')")
    nodes = atlas.tree(db, cube, dimension=dim, parent=canon_root, depth=depth, leaves_only=leaves_only, limit=limit)
    touched = {canon_root, *(n["name"] for n in nodes)}
    caveats = resolver.caveats({cube: touched}) if hasattr(resolver, "caveats") else []
    for n in nodes:
        note = resolver.note(n["name"], cube) if hasattr(resolver, "note") else None
        if note:
            n["note"] = note
    return {"cube": cube, "dimension": dim or (nodes[0]["dimension"] if nodes else None),
            "parent": canon_root, "depth": depth, "count": len(nodes), "members": nodes, "caveats": caveats}


def _guess_dim(db, cube: str, member: str) -> str | None:
    row = db.execute("SELECT dimension FROM member WHERE connection=? AND name=? COLLATE NOCASE LIMIT 1",
                     (cube, member)).fetchone()
    return row[0] if row else None


def _resolve_dim_name(db, cube: str, dim: str) -> str:
    dims = atlas.known_dimensions(db, cube)
    want = _norm_dim(dim)
    for d in dims:
        if _norm_dim(d) == want or d.casefold() == dim.strip().casefold():
            return d
    return _canon_dim(cube, dim)


def _resolve_member(resolver, db, cube: str, dim: str | None, name: str) -> str | None:
    if not dim:
        row = db.execute("SELECT dimension, name FROM member WHERE connection=? AND name=? COLLATE NOCASE LIMIT 1",
                         (cube, name)).fetchone()
        if row:
            return row[1]
        dim = _guess_dim(db, cube, name)
    if dim:
        got = atlas.canonical(db, cube, dim, name)
        if got:
            return got
    if resolver.exists(name, cube):
        return resolver.canonical(name, cube)
    return None


def _enrich_member(resolver, cube: str, h: dict) -> dict:
    out = {
        "cube": h["connection"],
        "dimension": h["dimension"],
        "name": h["name"],
        "label": h["label"],
        "parents": h["parents"],
        "depth": h["depth"],
    }
    if hasattr(resolver, "note"):
        note = resolver.note(h["name"], cube)
        if note:
            out["note"] = note
    if hasattr(resolver, "compact_card"):
        card = resolver.compact_card(h["name"], cube)
        if card:
            out["card"] = card
    return out


def _as_list(val: str | list[str]) -> list[str]:
    return [val] if isinstance(val, str) else list(val)


def _resolve_axis_members(resolver, db, cube: str, dim: str, raw_members: list[str], expand_mode: str | None) -> list[str]:
    low_dim = dim.casefold()
    resolved: list[str] = []
    for item in raw_members:
        if low_dim in _ACCT_DIMS:
            try:
                ms = resolver.members(item, cube)
                resolved.extend(ms)
                continue
            except UnknownMetric:
                pass
        c = atlas.canonical(db, cube, dim, item)
        resolved.append(c or (resolver.canonical(item, cube) if resolver.exists(item, cube) else item))
    if not expand_mode:
        return list(dict.fromkeys(resolved))
    expanded: list[str] = []
    mode = expand_mode.casefold()
    for m in resolved:
        if mode == "children":
            kids = [k["name"] for k in atlas.tree(db, cube, dimension=dim, parent=m, depth=1, limit=500)]
            expanded.extend(kids or [m])
        elif mode == "descendants":
            desc = [k["name"] for k in atlas.descendants(db, cube, m, leaves_only=False, limit=1000)]
            expanded.extend(desc or [m])
        elif mode == "leaves":
            leaves = [k["name"] for k in atlas.descendants(db, cube, m, leaves_only=True, limit=1000)]
            expanded.extend(leaves or [m])
        else:
            raise InputError(f"unknown expand mode {expand_mode!r}: expected 'children', 'descendants', or 'leaves'")
    return list(dict.fromkeys(expanded))


def _resolve_periods(periods: str | list[str] | None, closed_m: tuple[int, int], preliminary: bool) -> tuple[list[str], list[str]]:
    """Resolve periods argument to month labels, enforcing the closed-month guard unless preliminary=True."""
    warnings: list[str] = []
    if periods is None:
        return [period_name(closed_m)], warnings
    exprs = [periods] if isinstance(periods, str) else list(periods)
    months = []
    for e in exprs:
        try:
            months.extend(resolve_period(e, closed_m))
        except ValueError as err:
            raise GuardError(str(err), "use 'closed', 'closed-ytd', 'Sep 26', or 'Jan 26..Sep 26'") from None
    deduped = list(dict.fromkeys(months))
    open_months = [m for m in deduped if m > closed_m]
    if open_months:
        if not preliminary:
            raise GuardError(
                f"requested month(s) {', '.join(period_name(m) for m in open_months)} are after the last closed "
                f"month {period_name(closed_m)}",
                "use 'closed' / 'closed-ytd', or set preliminary=true explicitly if open actuals are intended",
            )
        warnings.append(f"preliminary open month(s) included after {period_name(closed_m)}: "
                        f"{', '.join(period_name(m) for m in open_months)}")
    return [period_name(m) for m in deduped], warnings


def _slice_grids(cube: str, pov: dict[str, str], row_axes: list[Axis], col_axes: list[Axis]):
    """Build concrete grid requests for provider.fetch_slice (splitting FinPlan across years when needed)."""
    if cube != "FinPlan":
        return [(pov, row_axes, col_axes, {c.dimension: c.members for c in col_axes})]

    period_axis = next((a for a in row_axes + col_axes if a.dimension == "Period"), None)
    if period_axis is None or "Years" in pov or any(a.dimension == "Years" for a in row_axes + col_axes):
        return [(pov, row_axes, col_axes, {})]

    by_fy: dict[str, list[tuple[str, str]]] = {}
    for p in period_axis.members:
        y, m = parse_period(p)
        by_fy.setdefault(f"FY{y % 100:02d}", []).append((ABBR[m - 1], p))

    out = []
    for fy, pairs in by_fy.items():
        short_periods = [ab for ab, _ in pairs]
        relabel = {ab: full for ab, full in pairs}
        new_rows = [Axis("Period", short_periods) if a.dimension == "Period" else a for a in row_axes]
        new_cols = [Axis("Period", short_periods) if a.dimension == "Period" else a for a in col_axes]
        out.append(({**pov, "Years": fy}, new_rows, new_cols, relabel))
    return out


def _default_pov(cube: str, resolver, user_pov: dict[str, str] | None) -> dict[str, str]:
    if cube == "FinPlan":
        base = {"Scenario": "Actual", "Version": "Working", "All Balance Types": "Periodic"}
    else:
        base = {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic"}
    out = dict(base)
    for k, v in (user_pov or {}).items():
        if k.casefold() == "scenario" and v.casefold() == "forecast":
            out[k] = resolver.latest_forecast()
        else:
            out[k] = v
    return out


def _base_series_pov_and_periods(cube: str, resolver, compare_name: str, base_periods: list[str],
                                 base_pov: dict[str, str]) -> tuple[dict[str, str], list[str]]:
    name = compare_name.casefold()
    if name == "ly":
        ly_periods = [period_name(shift(parse_period(p), -12)) for p in base_periods]
        return dict(base_pov), ly_periods
    scenarios = resolver.scenarios().get(cube, {})
    if name not in scenarios:
        raise GuardError(f"{cube}: compare {compare_name!r} is not supported",
                         f"choose from {sorted(['ly', *scenarios])}")
    scen_val, ver_val = scenarios[name]
    pov = dict(base_pov)
    if scen_val == "Forecast":
        scen_val = resolver.latest_forecast()
    if cube == "FinPlan":
        pov["Scenario"] = scen_val
        if ver_val:
            pov["Version"] = ver_val
    else:
        pov["All Scenarios"] = scen_val
    return pov, list(base_periods)


def _merge_parsed_rows(grids_with_meta: list[tuple[Any, Any]], row_dims: list[str], relabel_periods: bool) -> tuple[list[str], list[dict]]:
    """Flatten one or more Parsed grids into canonical record dicts keyed by column header."""
    combined: dict[tuple[str, ...], dict[str, Any]] = {}
    col_names: list[str] = []
    nr = len(row_dims)
    for parsed, relabel in grids_with_meta:
        relabel_map = relabel if (relabel_periods and isinstance(relabel, dict)) else {}
        col_heads_raw = parsed.columns[nr:]
        col_heads = [" / ".join(relabel_map.get(part, part) for part in h.split(" / ")) for h in col_heads_raw]
        for ch in col_heads:
            if ch not in col_names:
                col_names.append(ch)
        for raw_row in parsed.rows:
            row_key = tuple(str(x) for x in raw_row[:nr])
            rec = combined.setdefault(row_key, dict(zip(row_dims, row_key)))
            for ch, val in zip(col_heads, raw_row[nr:]):
                rec[ch] = val
    return col_names, list(combined.values())


def run_slice(provider, resolver, closed: dict[str, tuple[int, int]], *, cube: str = "FinMax",
              rows: dict[str, str | list[str]], cols: dict[str, str | list[str]] | None = None,
              pov: dict[str, str] | None = None, periods: str | list[str] | None = None,
              compare: list[str] | None = None, preliminary: bool = False,
              expand: dict[str, str] | None = None, suppress_missing: bool = True, limit: int = 200) -> dict:
    """Execute a read-only multi-dimensional EPM slice, returning exact values, variances (optional), and caveats."""
    if cube not in CUBES:
        raise GuardError(f"unknown cube {cube!r}", f"choose one of {', '.join(CUBES)}")
    if not rows:
        raise GuardError("rows cannot be empty", "pass at least one row dimension mapping, e.g. {'All Entities': [...]}")

    db = resolver.db
    closed_m = closed_for(closed, cube)
    period_dim = _period_dim(cube)

    canon_pov = {_resolve_dim_name(db, cube, k): v for k, v in _default_pov(cube, resolver, pov).items()}
    canon_expand = {_resolve_dim_name(db, cube, k): v for k, v in (expand or {}).items()}
    row_axes: list[Axis] = []
    for k, v in rows.items():
        d = _resolve_dim_name(db, cube, k)
        members = _resolve_axis_members(resolver, db, cube, d, _as_list(v), canon_expand.get(d))
        row_axes.append(Axis(d, members))

    resolved_periods, period_warnings = _resolve_periods(periods, closed_m, preliminary)
    col_axes: list[Axis] = []
    cols_in = dict(cols or {})
    if not cols_in or not any(_resolve_dim_name(db, cube, k) == period_dim for k in cols_in):
        if period_dim not in {a.dimension for a in row_axes} and period_dim not in canon_pov:
            col_axes.append(Axis(period_dim, resolved_periods))
    for k, v in cols_in.items():
        d = _resolve_dim_name(db, cube, k)
        members = resolved_periods if (d == period_dim and periods is not None) else _resolve_axis_members(
            resolver, db, cube, d, _as_list(v), canon_expand.get(d))
        col_axes.append(Axis(d, members))

    if not col_axes:
        raise GuardError("slice has no column axis", "specify `periods` or at least one `cols` dimension")

    row_dims = [a.dimension for a in row_axes]
    main_specs = _slice_grids(cube, canon_pov, row_axes, col_axes)
    main_res = provider.fetch_slice(cube, main_specs, suppress_missing=suppress_missing, label="fin_slice")
    warnings = period_warnings + list(main_res.warnings)
    if main_res.route == "rest":
        warnings.append(f"Fetched over Planning REST fallback ({main_res.fallback}).")

    col_names, records = _merge_parsed_rows(main_res.grids, row_dims, cube == "FinPlan")
    if compare:
        _attach_comparisons(provider, resolver, cube, canon_pov, row_axes, col_axes, resolved_periods,
                            compare, col_names, records, warnings, suppress_missing)

    touched_members = {m for a in row_axes + col_axes for m in a.members} | set(canon_pov.values())
    touched_months = set()
    for p in resolved_periods:
        with contextlib.suppress(ValueError):
            touched_months.add(parse_period(p))
    cavs = []
    if hasattr(resolver, "caveats"):
        cavs = resolver.caveats({cube: touched_members}, months=touched_months or None)
        if hasattr(resolver, "card_caveats"):
            have = {c["id"] for c in cavs}
            cavs += [c for c in resolver.card_caveats({cube: touched_members}) if c["id"] not in have]

    total_rows = len(records)
    lim = max(1, min(int(limit), 1000))
    return {
        "cube": cube,
        "pov": canon_pov,
        "row_dimensions": row_dims,
        "columns": [*row_dims, *col_names],
        "rows": records[:lim],
        "row_count": total_rows,
        "truncated": total_rows > lim,
        "closed": period_name(closed_m),
        "route": main_res.route,
        "warnings": warnings,
        "caveats": cavs,
        "fetch_seconds": round(main_res.seconds, 2),
    }


def _attach_comparisons(provider, resolver, cube: str, pov: dict[str, str], row_axes: list[Axis],
                        col_axes: list[Axis], base_periods: list[str], compare: list[str],
                        col_names: list[str], records: list[dict], warnings: list[str], suppress_missing: bool) -> None:
    row_dims = [a.dimension for a in row_axes]
    by_key = {tuple(str(r[d]) for d in row_dims): r for r in records}
    period_dim = _period_dim(cube)
    for cmp_name in compare:
        cmp_pov, cmp_periods = _base_series_pov_and_periods(cube, resolver, cmp_name, base_periods, pov)
        cmp_cols = [Axis(period_dim, cmp_periods) if a.dimension == period_dim else a for a in col_axes]
        cmp_res = provider.fetch_slice(cube, _slice_grids(cube, cmp_pov, row_axes, cmp_cols),
                                       suppress_missing=suppress_missing, label=f"fin_slice:{cmp_name}")
        warnings.extend(cmp_res.warnings)
        cmp_col_names, cmp_rows = _merge_parsed_rows(cmp_res.grids, row_dims, cube == "FinPlan")
        if cmp_name.casefold() == "ly":
            pair_map = dict(zip(cmp_col_names, col_names))
        else:
            pair_map = {c: c for c in cmp_col_names}
        cmp_by_key = {tuple(str(r[d]) for d in row_dims): r for r in cmp_rows}
        for ac_col in list(col_names):
            cmp_col = next((cc for cc, target in pair_map.items() if target == ac_col), None)
            if cmp_col is None:
                continue
            b_head = f"{ac_col} [{cmp_name}]"
            v_head = f"{ac_col} var vs {cmp_name}"
            p_head = f"{ac_col} var% vs {cmp_name}"
            for h in (b_head, v_head, p_head):
                if h not in col_names:
                    col_names.append(h)
            for rk, rec in by_key.items():
                bval = (cmp_by_key.get(rk) or {}).get(cmp_col)
                aval = rec.get(ac_col)
                rec[b_head] = bval
                if isinstance(aval, (int, float)) and isinstance(bval, (int, float)):
                    diff = round(aval - bval, 4)
                    rec[v_head] = diff
                    rec[p_head] = round(diff / abs(bval), 6) if bval else None
                else:
                    rec[v_head] = None
                    rec[p_head] = None
