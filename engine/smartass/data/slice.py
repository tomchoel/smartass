# engine/smartass/data/slice.py
"""Tier 1 read-only cube exploration: hierarchy browsing (`fin_tree`) and arbitrary multi-dimensional slicing
(`fin_slice`) over the dual EPM engines with closed-month guards, alias resolution, batching, YTD totals,
leaf-to-parent reconciliation, and knowledge caveats."""
from __future__ import annotations

import contextlib
import itertools
from typing import Any

from ..core.errors import InputError
from ..epm.grid import Axis
from ..knowledge import atlas
from ..model.periods import ABBR, closed_for, label as period_name, parse as parse_period, resolve as resolve_period, shift
from ..model.resolver import UnknownMetric
from ..model.validate import GuardError

CUBES = ("FinMax", "FinLite", "FinPlan")
MAX_EXPAND = 5000
MAX_SLICE_LIMIT = 2000
CHUNK_ROW_MEMBERS = 400
_ACCT_DIMS = {"all accounts", "accounts", "account"}
_ENTITY_DIMS = {"all entities", "entities", "entity"}
_PROD_DIMS = {"all products", "products", "product"}
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
    if low in _PROD_DIMS:
        return "All Products"
    return dim


def _period_dim(cube: str) -> str:
    return "Period" if cube == "FinPlan" else "All Time Periods"


def browse_tree(resolver, *, cube: str | None = None, dimension: str | None = None, parent: str | None = None,
                query: str | None = None, depth: int = 1, leaves_only: bool = False, limit: int = 100) -> dict:
    """Inspect cube dimensions, parent-child subtrees, card attributes, truncation metadata, and caveats."""
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
        lim = max(1, min(int(limit), atlas.MAX_TREE_LIMIT))
        hits = atlas.search(db, query, connection=cube, dimension=dim, limit=lim + 1)
        truncated = len(hits) > lim
        hits = hits[:lim]
        items = [_enrich_member(resolver, cube, h) for h in hits]
        caveats = resolver.caveats({cube: {h["name"] for h in hits}}) if hits else []
        return {
            "cube": cube, "dimension": dim, "query": query, "members": items,
            "returned": len(items), "truncated": truncated, "has_more": truncated,
            "effective_limit": lim, "caveats": caveats,
        }

    root_in = parent or dim
    canon_root = _resolve_member(resolver, db, cube, dim, root_in) if root_in else root_in
    if canon_root is None:
        raise GuardError(f"{cube}: unknown member or parent {root_in!r}",
                         "check spelling with fin_tree(query=...) or fin_lookup(kind='member')")
    page = atlas.tree_page(db, cube, dimension=dim, parent=canon_root, depth=depth,
                           leaves_only=leaves_only, limit=limit)
    nodes = page["items"]
    touched = {canon_root, *(n["name"] for n in nodes)}
    caveats = resolver.caveats({cube: touched}) if hasattr(resolver, "caveats") else []
    for n in nodes:
        note = resolver.note(n["name"], cube) if hasattr(resolver, "note") else None
        if note:
            n["note"] = note
    warnings: list[str] = []
    if page["clamped_depth"]:
        warnings.append(f"Requested depth={depth} was clamped to effective_depth={page['effective_depth']}.")
    if page["clamped_limit"]:
        warnings.append(f"Requested limit={limit} was clamped to effective_limit={page['effective_limit']}.")
    if page["truncated"]:
        warnings.append(f"Result truncated at {len(nodes)} members (increase limit or narrow parent).")
    return {
        "cube": cube,
        "dimension": dim or (nodes[0]["dimension"] if nodes else None),
        "parent": canon_root,
        "depth": depth,
        "effective_depth": page["effective_depth"],
        "effective_limit": page["effective_limit"],
        "count": len(nodes),
        "returned": len(nodes),
        "truncated": page["truncated"],
        "has_more": page["has_more"],
        "members": nodes,
        "caveats": caveats,
        **({"warnings": warnings} if warnings else {}),
    }


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
        if dim.casefold() in _PROD_DIMS:
            resolved = atlas.resolve_product(db, cube, name)
            if resolved != name or atlas.canonical(db, cube, dim, resolved):
                return resolved
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


def _resolve_axis_members(resolver, db, cube: str, dim: str, raw_members: list[str],
                          expand_mode: str | None) -> tuple[list[str], list[tuple[str, list[str], str]]]:
    """Resolve members + expand modes cleanly; returns (members, [(parent, children, mode)])."""
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
        if low_dim in _PROD_DIMS:
            item = atlas.resolve_product(db, cube, item)
        c = atlas.canonical(db, cube, dim, item)
        resolved.append(c or (resolver.canonical(item, cube) if resolver.exists(item, cube) else item))
    roots = list(dict.fromkeys(resolved))
    if not expand_mode:
        return roots, []

    mode = expand_mode.casefold()
    if mode not in ("children", "descendants", "leaves"):
        raise InputError(f"unknown expand mode {expand_mode!r}: expected 'children', 'descendants', or 'leaves'")

    expanded: list[str] = []
    expansions: list[tuple[str, list[str], str]] = []
    for m in roots:
        members, total, truncated = atlas.expand_members(db, cube, dim, m, mode=mode, max_members=MAX_EXPAND)
        if truncated:
            raise GuardError(
                f"{cube} {dim}: expanding {m!r} ({mode}) matched more than {MAX_EXPAND} members ({total}+)",
                "pick a narrower parent node so every expanded member can be fetched without dropping rows",
            )
        kids = members or [m]
        expanded.extend(kids)
        if members:
            expansions.append((m, kids, mode))
    return list(dict.fromkeys(expanded)), expansions


def _resolve_periods(periods: str | list[str] | None, closed_m: tuple[int, int],
                     preliminary: bool) -> tuple[list[str], list[str]]:
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


def _chunk_row_axes(row_axes: list[Axis]) -> list[list[Axis]]:
    """Split oversized single-dimension row lists into safe HTTP grid batches without dropping members."""
    longest_idx = max(range(len(row_axes)), key=lambda i: len(row_axes[i].members))
    n = len(row_axes[longest_idx].members)
    if n <= CHUNK_ROW_MEMBERS:
        return [row_axes]
    batches: list[list[Axis]] = []
    big = row_axes[longest_idx]
    for start in range(0, n, CHUNK_ROW_MEMBERS):
        chunk = Axis(big.dimension, big.members[start:start + CHUNK_ROW_MEMBERS])
        batches.append([chunk if i == longest_idx else a for i, a in enumerate(row_axes)])
    return batches


def _slice_grids(cube: str, pov: dict[str, str], row_axes: list[Axis], col_axes: list[Axis]):
    """Build concrete grid requests for provider.fetch_slice (splitting FinPlan years + large row axes)."""
    row_batches = _chunk_row_axes(row_axes)
    if cube != "FinPlan":
        meta = {c.dimension: c.members for c in col_axes}
        return [(pov, rb, col_axes, meta) for rb in row_batches]

    period_axis = next((a for a in row_axes + col_axes if a.dimension == "Period"), None)
    if period_axis is None or "Years" in pov or any(a.dimension == "Years" for a in row_axes + col_axes):
        return [(pov, rb, col_axes, {}) for rb in row_batches]

    by_fy: dict[str, list[tuple[str, str]]] = {}
    for p in period_axis.members:
        y, m = parse_period(p)
        by_fy.setdefault(f"FY{y % 100:02d}", []).append((ABBR[m - 1], p))

    out = []
    for fy, pairs in by_fy.items():
        short_periods = [ab for ab, _ in pairs]
        relabel = {ab: full for ab, full in pairs}
        for rb in row_batches:
            new_rows = [Axis("Period", short_periods) if a.dimension == "Period" else a for a in rb]
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


def _merge_parsed_rows(grids_with_meta: list[tuple[Any, Any]], row_dims: list[str],
                       relabel_periods: bool) -> tuple[list[str], list[dict]]:
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


def _is_level_member(resolver, cube: str, member: str) -> bool:
    if hasattr(resolver, "aggregation") and resolver.aggregation(member) == "last":
        return True
    if hasattr(resolver, "card"):
        c = resolver.card(member, cube)
        if c and c.get("time_balance") in ("balance_last", "average"):
            return True
    db = getattr(resolver, "db", None)
    if db is not None:
        row = db.execute(
            "SELECT 1 FROM card WHERE connection=? AND name=? COLLATE NOCASE "
            "AND time_balance IN ('balance_last', 'average') LIMIT 1",
            (cube, member),
        ).fetchone()
        if row:
            return True
    return False


def _add_ytd_totals(records: list[dict], base_col_names: list[str], all_col_names: list[str],
                    row_dims: list[str], resolver, cube: str, compare: list[str] | None,
                    warnings: list[str]) -> None:
    """Append YTD/Total columns across period columns, never summing non-additive level accounts."""
    period_cols = list(base_col_names)
    if len(period_cols) < 2:
        return
    acct_dims = [d for d in row_dims if d.casefold() in _ACCT_DIMS]

    def row_is_level(rec: dict) -> bool:
        return any(_is_level_member(resolver, cube, str(rec[d])) for d in acct_dims)

    def sum_cols(rec: dict, cols: list[str]) -> float | None:
        nums = [rec.get(c) for c in cols if isinstance(rec.get(c), (int, float))]
        return round(sum(nums), 4) if nums else None

    def end_val(rec: dict, cols: list[str]) -> float | None:
        for c in reversed(cols):
            v = rec.get(c)
            if isinstance(v, (int, float)):
                return round(float(v), 4)
        return None

    has_level = any(row_is_level(r) for r in records)
    all_col_names.append("Total")
    if has_level:
        all_col_names.append("Period End")
        warnings.append("Level/average accounts (e.g. balance_last / FTE) are not summed in 'Total'; see 'Period End'.")
    for r in records:
        if row_is_level(r):
            r["Total"] = None
            r["Period End"] = end_val(r, period_cols)
        else:
            r["Total"] = sum_cols(r, period_cols)
            if has_level:
                r["Period End"] = None

    for cmp_name in (compare or ()):
        cmp_cols = [f"{c} [{cmp_name}]" for c in period_cols]
        btot = f"Total [{cmp_name}]"
        vtot = f"Total var vs {cmp_name}"
        ptot = f"Total var% vs {cmp_name}"
        for h in (btot, vtot, ptot):
            all_col_names.append(h)
        if has_level:
            bend = f"Period End [{cmp_name}]"
            all_col_names.append(bend)
        for r in records:
            is_lvl = row_is_level(r)
            if is_lvl:
                r[btot] = None
                aval = r.get("Period End")
                bval = end_val(r, cmp_cols)
                r[f"Period End [{cmp_name}]"] = bval
            else:
                r[btot] = sum_cols(r, cmp_cols)
                aval = r.get("Total")
                bval = r[btot]
                if has_level:
                    r[f"Period End [{cmp_name}]"] = None
            if isinstance(aval, (int, float)) or isinstance(bval, (int, float)):
                diff = round((aval or 0.0) - (bval or 0.0), 4)
                r[vtot] = diff
                r[ptot] = round(diff / abs(bval), 6) if isinstance(bval, (int, float)) and bval else None
            else:
                r[vtot] = None
                r[ptot] = None


def _reconcile_expansions(provider, resolver, db, cube: str, pov: dict[str, str], row_axes: list[Axis],
                          col_axes: list[Axis], expansions: list[tuple[str, str, list[str], str]],
                          records: list[dict], period_cols: list[str], warnings: list[str]) -> list[dict]:
    """Compare expanded child/leaf sums against parent cube values where additive and single-placement."""
    out: list[dict] = []
    if len(row_axes) != 1:
        for dim, par, kids, mode in expansions:
            out.append({"dimension": dim, "parent": par, "mode": mode, "leaf_count": len(kids),
                        "status": "skipped", "reason": "multi-axis row slice (reconcile on 1 row axis)"})
        return out

    dim, par, kids, mode = expansions[0]
    # Check non-additivity if accounts expanded or account in POV
    acct_pov = pov.get("All Accounts") or pov.get("Account")
    if acct_pov and _is_level_member(resolver, cube, acct_pov):
        return [{"dimension": dim, "parent": par, "mode": mode, "leaf_count": len(kids),
                 "status": "skipped", "reason": f"non-additive account {acct_pov} (level/average)"}]
    # Check if expanded subtree has multiple placements or '-' / '~' operators
    shared = db.execute(
        "SELECT 1 FROM member WHERE connection=? AND dimension=? AND name IN ({}) "
        "GROUP BY name HAVING COUNT(*) > 1 LIMIT 1".format(",".join("?" * len(kids))),
        (cube, dim, *kids),
    ).fetchone()
    if shared:
        return [{"dimension": dim, "parent": par, "mode": mode, "leaf_count": len(kids),
                 "status": "skipped", "reason": f"shared members in {dim} subtree under {par}"}]
    non_plus = db.execute(
        "SELECT 1 FROM card WHERE connection=? AND dimension=? AND parent=? AND operator IN ('-', '~') LIMIT 1",
        (cube, dim, par),
    ).fetchone()
    if non_plus:
        return [{"dimension": dim, "parent": par, "mode": mode, "leaf_count": len(kids),
                 "status": "skipped", "reason": f"non-additive consolidation operators (- or ~) under {par}"}]

    par_axes = [Axis(dim, [par])]
    par_res = provider.fetch_slice(cube, _slice_grids(cube, pov, par_axes, col_axes),
                                   suppress_missing=False, label="fin_slice:reconcile")
    _, par_rows = _merge_parsed_rows(par_res.grids, [dim], cube == "FinPlan")
    par_row = par_rows[0] if par_rows else {}
    for pc in period_cols[:1]:
        pval = par_row.get(pc)
        lsum = sum(r.get(pc) for r in records if isinstance(r.get(pc), (int, float)))
        if isinstance(pval, (int, float)):
            gap = round(pval - lsum, 2)
            status = "exact" if abs(gap) < 1.0 else "gap"
            if status == "gap":
                warnings.append(f"Reconciliation gap for {par} ({pc}): parent {pval:,.2f} vs leaves {lsum:,.2f} (gap {gap:,.2f}).")
            out.append({"dimension": dim, "parent": par, "period": pc, "mode": mode, "leaf_count": len(kids),
                        "parent_value": round(pval, 2), "leaves_sum": round(lsum, 2), "gap": gap, "status": status})
        else:
            out.append({"dimension": dim, "parent": par, "period": pc, "mode": mode, "leaf_count": len(kids),
                        "status": "no parent value"})
    return out


def run_slice(provider, resolver, closed: dict[str, tuple[int, int]], *, cube: str = "FinMax",
              rows: dict[str, str | list[str]], cols: dict[str, str | list[str]] | None = None,
              pov: dict[str, str] | None = None, periods: str | list[str] | None = None,
              compare: list[str] | None = None, preliminary: bool = False,
              expand: dict[str, str] | None = None, suppress_missing: bool = True,
              total: bool = False, limit: int = 200) -> dict:
    """Execute a read-only multi-dimensional EPM slice, returning exact values, comparisons, totals, and caveats."""
    if cube not in CUBES:
        raise GuardError(f"unknown cube {cube!r}", f"choose one of {', '.join(CUBES)}")
    if not rows:
        raise GuardError("rows cannot be empty", "pass at least one row dimension mapping, e.g. {'All Entities': [...]}")

    db = resolver.db
    closed_m = closed_for(closed, cube)
    period_dim = _period_dim(cube)

    canon_pov = {_resolve_dim_name(db, cube, k): v for k, v in _default_pov(cube, resolver, pov).items()}
    for pk, pv in list(canon_pov.items()):
        if pk.casefold() in _PROD_DIMS:
            canon_pov[pk] = atlas.resolve_product(db, cube, pv)

    canon_expand = {_resolve_dim_name(db, cube, k): v for k, v in (expand or {}).items()}
    row_axes: list[Axis] = []
    expansions: list[tuple[str, str, list[str], str]] = []
    for k, v in rows.items():
        d = _resolve_dim_name(db, cube, k)
        members, exp_info = _resolve_axis_members(resolver, db, cube, d, _as_list(v), canon_expand.get(d))
        row_axes.append(Axis(d, members))
        for par, kids, mode in exp_info:
            expansions.append((d, par, kids, mode))

    resolved_periods, period_warnings = _resolve_periods(periods, closed_m, preliminary)
    col_axes: list[Axis] = []
    cols_in = dict(cols or {})
    if not cols_in or not any(_resolve_dim_name(db, cube, k) == period_dim for k in cols_in):
        if period_dim not in {a.dimension for a in row_axes} and period_dim not in canon_pov:
            col_axes.append(Axis(period_dim, resolved_periods))
    for k, v in cols_in.items():
        d = _resolve_dim_name(db, cube, k)
        if d == period_dim and periods is not None:
            members = resolved_periods
        else:
            members, _ = _resolve_axis_members(resolver, db, cube, d, _as_list(v), canon_expand.get(d))
        col_axes.append(Axis(d, members))

    if not col_axes:
        raise GuardError("slice has no column axis", "specify `periods` or at least one `cols` dimension")

    row_dims = [a.dimension for a in row_axes]
    axis_order = {combo: idx for idx, combo in enumerate(itertools.product(*(a.members for a in row_axes)))}

    main_specs = _slice_grids(cube, canon_pov, row_axes, col_axes)
    main_res = provider.fetch_slice(cube, main_specs, suppress_missing=suppress_missing, label="fin_slice")
    warnings = period_warnings + list(main_res.warnings)
    if main_res.route == "rest":
        warnings.append(f"Fetched over Planning REST fallback ({main_res.fallback}).")

    base_col_names, records = _merge_parsed_rows(main_res.grids, row_dims, cube == "FinPlan")
    if not base_col_names and col_axes:
        base_col_names = [" / ".join(c) for c in itertools.product(*(a.members for a in col_axes))]
    all_col_names = list(base_col_names)

    if compare:
        _attach_comparisons(provider, resolver, cube, canon_pov, row_axes, col_axes, resolved_periods,
                            compare, base_col_names, all_col_names, records, axis_order, warnings, suppress_missing)

    records.sort(key=lambda rec: axis_order.get(tuple(str(rec[d]) for d in row_dims), len(axis_order)))

    if total:
        _add_ytd_totals(records, base_col_names, all_col_names, row_dims, resolver, cube, compare, warnings)

    reconciliation = []
    if expansions and suppress_missing and not compare:
        reconciliation = _reconcile_expansions(provider, resolver, db, cube, canon_pov, row_axes, col_axes,
                                               expansions, records, base_col_names, warnings)
    elif expansions and compare:
        # Use non-null current period sum when comparing
        reconciliation = _reconcile_expansions(provider, resolver, db, cube, canon_pov, row_axes, col_axes,
                                               expansions, records, base_col_names, warnings)

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
    lim = max(1, min(int(limit), MAX_SLICE_LIMIT))
    if total_rows > lim:
        warnings.append(f"Returned {lim} of {total_rows} rows (truncated=true; increase limit or narrow filters).")
    return {
        "cube": cube,
        "pov": canon_pov,
        "row_dimensions": row_dims,
        "columns": [*row_dims, *all_col_names],
        "rows": records[:lim],
        "row_count": total_rows,
        "returned": min(lim, total_rows),
        "truncated": total_rows > lim,
        "closed": period_name(closed_m),
        "route": main_res.route,
        "warnings": warnings,
        "caveats": cavs,
        **({"reconciliation": reconciliation} if reconciliation else {}),
        **({"variance_rule": "Missing side treated as 0 when series has data; var_pct is null when base is null or 0."}
           if compare else {}),
        "fetch_seconds": round(main_res.seconds, 2),
    }


def _attach_comparisons(provider, resolver, cube: str, pov: dict[str, str], row_axes: list[Axis],
                        col_axes: list[Axis], base_periods: list[str], compare: list[str],
                        base_col_names: list[str], all_col_names: list[str],
                        records: list[dict], axis_order: dict[tuple[str, ...], int],
                        warnings: list[str], suppress_missing: bool) -> None:
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
            pair_map = dict(zip(cmp_col_names, base_col_names))
        else:
            pair_map = {c: c for c in cmp_col_names}

        cmp_by_key = {tuple(str(r[d]) for d in row_dims): r for r in cmp_rows}
        cmp_empty = all(getattr(parsed, "empty", False) for parsed, _ in cmp_res.grids) and not cmp_rows

        # Union comparison-only keys back into `records` in axis order!
        for rk, cmp_rec in cmp_by_key.items():
            if rk not in by_key:
                new_rec: dict[str, Any] = dict(zip(row_dims, rk))
                for col in all_col_names:
                    new_rec[col] = None
                by_key[rk] = new_rec
                records.append(new_rec)

        for ac_col in list(base_col_names):
            cmp_col = next((cc for cc, target in pair_map.items() if target == ac_col), None)
            b_head = f"{ac_col} [{cmp_name}]"
            v_head = f"{ac_col} var vs {cmp_name}"
            p_head = f"{ac_col} var% vs {cmp_name}"
            for h in (b_head, v_head, p_head):
                if h not in all_col_names:
                    all_col_names.append(h)
            for rk, rec in by_key.items():
                bval = (cmp_by_key.get(rk) or {}).get(cmp_col) if cmp_col is not None else None
                aval = rec.get(ac_col)
                rec[b_head] = bval
                if cmp_empty or isinstance(aval, str) or isinstance(bval, str):
                    rec[v_head] = None
                    rec[p_head] = None
                elif isinstance(aval, (int, float)) or isinstance(bval, (int, float)):
                    diff = round((aval or 0.0) - (bval or 0.0), 4)
                    rec[v_head] = diff
                    rec[p_head] = round(diff / abs(bval), 6) if isinstance(bval, (int, float)) and bval else None
                else:
                    rec[v_head] = None
                    rec[p_head] = None
