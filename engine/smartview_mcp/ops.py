"""Operations executed on the COM thread (each takes the ExcelHost as its first argument)."""
from __future__ import annotations

import datetime as dt
import logging
import os
import re
import shutil
import time
from pathlib import Path

from . import config
from .grid import Axis, Layout, full_pov, indent_of, name_from_label, parse
from .host import ExcelHost, HostError, InputError

log = logging.getLogger("smartview.ops")

_dims_cache: dict[str, list[str]] = {}
_types_cache: dict[str, list[int]] = {}


def _qsheet(conn: str) -> str:
    return f"Q_{conn}"[:31]


def _used_values(ws):
    ur = ws.UsedRange
    last_r = ur.Row + ur.Rows.Count - 1
    last_c = ur.Column + ur.Columns.Count - 1
    v = ws.Range(ws.Cells(1, 1), ws.Cells(last_r, last_c)).Value2
    return v if isinstance(v, tuple) else ((v,),)


def _apply_sheet_options(h: ExcelHost, sheet: str, extra: dict | None = None):
    for k, v in {**config.SHEET_OPTIONS, **(extra or {})}.items():
        h.run("B_SetOption", k, v, sheet)


_aliases: dict[str, dict[str, str]] = {}      # conn -> {raw dimension name or label: grid label}


def _sheet_meta(h: ExcelHost, sheet: str) -> dict:
    r = h._run("B_SheetInfo", sheet)
    return dict(zip(r[1], r[2])) if isinstance(r, tuple) and r[0] == "OK" else {}


def _connect_with_retry(h: ExcelHost, sheet: str, conn: str, attempts: int = 3):
    for i in range(attempts):
        try:
            h.connect_sheet(sheet, conn)
            return
        except HostError as e:
            transient = any(x in str(e) for x in ("752", "Empty response", "Not Connected", "-4", "Cannot connect",
                                                  "did not accept the connection"))
            if i == attempts - 1 or not transient:
                raise
            h.check_signin()                      # a stuck sign-in: no point retrying in this Excel
            time.sleep(4 * (i + 1))


def _grid_labels(vals, types: list[int]) -> list[str] | None:
    """Dimension labels as shown in the default grid (POV row 1, column headers, row headers), in the
    order of `types` (0 row, 1 column, 2 POV). None if the grid is not the standard in-sheet-POV layout."""
    nr, nc = types.count(0), types.count(1)
    grid = [list(r) for r in vals]
    try:
        row_l = [grid[1 + nc][d] for d in range(nr)]
        col_l = [grid[1 + d][nr] for d in range(nc)]
        pov_l = [v for v in grid[0][nr:] if v not in (None, "")]
    except IndexError:
        return None
    if len(pov_l) != types.count(2) or any(not isinstance(x, str) for x in row_l + col_l):
        return None
    it = {0: iter(row_l), 1: iter(col_l), 2: iter(pov_l)}
    return [next(it[t]) for t in types]


_DIMS_FILE = config.SVKIT / "dims_cache.json"
DIMS_MAX_AGE_DAYS = 30


def _save_dims(conn: str) -> None:
    import json
    try:
        data = json.loads(_DIMS_FILE.read_text(encoding="utf-8")) if _DIMS_FILE.exists() else {}
        data[conn] = {"labels": _dims_cache[conn], "types": _types_cache[conn], "aliases": _aliases[conn],
                      "saved": dt.datetime.now().isoformat(timespec="seconds")}
        _DIMS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except (OSError, ValueError, KeyError):
        pass


def dims_known(conn: str) -> bool:
    """Dimension labels/aliases for conn without touching Excel: memory, else the on-disk cache (written after every
    successful default-grid open, refreshed monthly). Lets pulls into the user's Excel skip the hidden server Excel."""
    if conn in _dims_cache and conn in _aliases:
        return True
    import json
    try:
        c = json.loads(_DIMS_FILE.read_text(encoding="utf-8")).get(conn)
        if not c or (dt.datetime.now() - dt.datetime.fromisoformat(c["saved"])).days > DIMS_MAX_AGE_DAYS:
            return False
        _dims_cache[conn], _types_cache[conn], _aliases[conn] = c["labels"], c["types"], c["aliases"]
        return True
    except (OSError, ValueError, KeyError, TypeError):
        return False


def learn_dims(conn: str, raw: list, types: list, labels: list[str]) -> None:
    _dims_cache[conn] = labels
    _types_cache[conn] = types
    amap = {lab: lab for lab in labels}
    amap.update({rw: lab for rw, lab in zip(raw, labels)})
    amap.update({lab.removeprefix("All ").replace(" ", ""): lab for lab in labels})
    _aliases[conn] = amap
    _save_dims(conn)


def ensure_grid_sheet(h: ExcelHost, conn: str, sheet: str | None = None, fresh: bool = False) -> str:
    """A sheet connected to `conn` that holds a retrieved standard ad hoc grid (POV row in the sheet,
    Default alias table) - required before free-form edits."""
    sheet = sheet or _qsheet(conn)
    if not fresh and h.connected.get(sheet) == conn:
        return sheet
    problem = ""
    for attempt in range(3):
        ws = h.sheet(sheet, fresh=True)
        ws.Activate()
        _connect_with_retry(h, sheet, conn)
        if attempt:
            h._run("B_SetAliasTable", sheet, "Default")
        _apply_sheet_options(h, sheet)
        h.run("B_Retrieve", sheet)
        meta = _sheet_meta(h, sheet)
        r = h.run("B_GetDimensions", sheet)
        raw, types = list(r[1]), list(r[2])
        labels = _grid_labels(_used_values(ws), types)
        standard = "Standard" in str(meta.get("Sheet Type", "")) and str(meta.get("Alias Table", "")).lower() != "none"
        if labels and standard:
            learn_dims(conn, raw, types, labels)
            return sheet
        problem = (f"sheet type '{meta.get('Sheet Type')}', alias table '{meta.get('Alias Table')}', "
                   f"grid {'has' if labels else 'lacks'} an in-sheet POV row")
        h._run("B_Disconnect", sheet)
        h.connected.pop(sheet, None)
        time.sleep(2)
    raise HostError(f"Smart View opened {conn} in a non-standard ad hoc mode ({problem}). "
                    "Try sv_restart; if it persists, check the connection in Smart View (alias table Default).")


def canon_dim(conn: str, name: str) -> str:
    """Accept 'All Accounts', 'Accounts', 'GLSources', 'GL Sources' ... and return the grid label."""
    amap = _aliases.get(conn, {})
    if name in amap:
        return amap[name]
    k = name.removeprefix("All ").replace(" ", "")
    return amap.get(k, name)


def dimensions(h: ExcelHost, conn: str) -> dict:
    ensure_grid_sheet(h, conn)
    return {"connection": conn, "dimensions": _dims_cache[conn],
            "default_axes": {d: {0: "row", 1: "column", 2: "pov"}.get(t, t)
                             for d, t in zip(_dims_cache[conn], _types_cache[conn])}}


MAX_BISECT_RETRIEVES = 40


def is_rc1000(msg) -> bool:
    """rc=1000 exactly ('ERR|1000|...' or 'rc=1000') - not 100000 'Unable to connect to the provider'."""
    return re.search(r"(?<!\d)1000(?!\d)", str(msg)) is not None


def is_transient(msg) -> bool:
    """rc=752 'Empty response from provider' or rc=57704 'Unsupported Error' (server says 'H&yperion') on a retrieve:
    seen right after an unfinished sign-in; usually gone on the next try. Not the 57704 of an unsupported member
    query (_explain's 'does not support this member query')."""
    s = str(msg)
    if "does not support this member query" in s:
        return False
    return re.search(r"(?:rc=|ERR\|)(752|57704)(?!\d)", s) is not None or "empty response from provider" in s.lower()


def _rejected_members(h: ExcelHost, conn: str, pov: dict, rows: list[Axis], cols: list[Axis]) -> list[tuple[str, str]]:
    """After an rc=1000 'Unsupported Error': bisect each multi-member axis to find the members Smart View rejects
    (e.g. a name used twice in one dimension, like FinPlan Account 'Other'). Returns [(dimension, member)]."""
    budget = [MAX_BISECT_RETRIEVES]
    axes = rows + cols

    def ok(i: int, members: list[str]) -> bool:
        if budget[0] <= 0:
            raise HostError("Smart View rejected the grid (rc=1000) and the bad member could not be isolated within "
                            f"{MAX_BISECT_RETRIEVES} test retrieves; split the query and retry.")
        budget[0] -= 1
        trial = [Axis(a.dimension, members if j == i else a.members[:1]) for j, a in enumerate(axes)]
        try:
            _query_once(h, conn, pov, trial[:len(rows)], trial[len(rows):], True, False, _retried=True)
            return True
        except HostError as e:
            if is_rc1000(e):
                return False
            raise

    def bad(i: int, members: list[str]) -> list[str]:
        if len(axes[i].members) <= 1 or ok(i, members):
            return []
        if len(members) == 1:
            return members
        mid = len(members) // 2
        return bad(i, members[:mid]) + bad(i, members[mid:])

    out = []
    for i, a in enumerate(axes):
        out += [(a.dimension, m) for m in bad(i, list(a.members))]
    return out


def _explain_reject(conn: str, dim: str, member: str) -> str:
    try:
        from . import atlas
        parents = [r[0] for r in atlas.connect().execute(
            "SELECT DISTINCT parent FROM member WHERE connection=? AND name=?", (conn, member))]
    except Exception:       # noqa: BLE001 - the explanation is best effort
        parents = []
    why = f" (the name occurs under {len(parents)} parents: {', '.join(map(str, parents))})" if len(parents) > 1 else ""
    return f"{dim}: {member}{why}"


def unknown_members(conn: str, pov: dict, rows: list[Axis], cols: list[Axis]) -> list[tuple[str, str]]:
    """(dimension, member) pairs the atlas does not know - one indexed lookup per member. Dimensions whose atlas is
    missing, failed or stale are skipped (no verdict without a trustworthy catalogue)."""
    from . import atlas
    try:
        db = atlas.connect()
    except Exception:       # noqa: BLE001 - no atlas: nothing to check against
        return []
    items = [(canon_dim(conn, d), m) for d, m in pov.items()]
    items += [(canon_dim(conn, a.dimension), m) for a in rows + cols for m in a.members]
    ok: dict[str, bool] = {}
    out = []
    for dim, m in dict.fromkeys(items):
        if dim not in ok:
            ok[dim] = atlas.usable(db, conn, dim)
        if ok[dim] and not atlas.has_member(db, conn, dim, m):
            out.append((dim, m))
    return out


def canonical_pov(conn: str, pov: dict) -> dict:
    """pov with each member in the atlas spelling (Essbase ignores case; the sheets then show the real name)."""
    from . import atlas
    try:
        db = atlas.connect()
    except Exception:       # noqa: BLE001 - no atlas: keep the given spelling
        return dict(pov)
    out = {}
    for d, m in pov.items():
        out[d] = (atlas.canonical(db, conn, canon_dim(conn, d), m) or m) if isinstance(m, str) else m
    return out


def check_members(conn: str, pov: dict, rows: list[Axis], cols: list[Axis], confirm=None) -> None:
    """InputError naming every member the atlas does not know in conn (see unknown_members). The atlas can predate a
    new period or account, so each miss is first confirmed live with confirm(conn, dim, member) -> bool (live_confirm); a
    member is rejected only when the live check says it does not exist (a failed check leaves it unverified)."""
    bad = unknown_members(conn, pov, rows, cols)
    if confirm:
        bad = [(d, m) for d, m in bad if _missing_live(confirm, conn, d, m)]
    if bad:
        raise InputError("Invalid query: unknown member " + "; ".join(f"'{m}' ({conn} / {d})" for d, m in bad)
                         + ". Find the right name with sv_find_member / sv_members.")


def _missing_live(confirm, conn: str, dim: str, member: str) -> bool:
    try:
        return not confirm(conn, dim, member)
    except Exception as e:  # noqa: BLE001 - unverified: let the pull and the error-cell scan decide
        log.warning("live member check of %s / %s failed: %s", conn, member, e)
        return False


def query(h: ExcelHost, conn: str, pov: dict, rows: list[Axis], cols: list[Axis],
          suppress_missing: bool, suppress_zeros: bool):
    """Retrieve a grid. If Smart View rejects it (rc=1000), find the rejected members by bisection, drop them,
    retrieve again and report them in parsed.pov_problems, so one bad member does not sink a large pull.
    A transient rc=752 / rc=57704 is retried once on a fresh grid sheet - after a restart if Smart View's sign-in
    is stuck (h.check_signin raises SignInBlocked; ExcelHost.call restarts Excel and runs the query again)."""
    try:
        return _query_once(h, conn, pov, rows, cols, suppress_missing, suppress_zeros)
    except HostError as e:
        if is_transient(e):
            log.warning("transient Smart View error on %s, retrying once: %s", conn, e)
            h.check_signin()
            time.sleep(config.TRANSIENT_RETRY_WAIT_S)
            ensure_grid_sheet(h, conn, fresh=True)
            return _query_once(h, conn, pov, rows, cols, suppress_missing, suppress_zeros)
        if not is_rc1000(e) or all(len(a.members) <= 1 for a in rows + cols):
            raise
        err = e
    rej = _rejected_members(h, conn, pov, rows, cols)
    if not rej:
        raise err
    drop = set(rej)
    rows2 = [Axis(a.dimension, [m for m in a.members if (a.dimension, m) not in drop]) for a in rows]
    cols2 = [Axis(a.dimension, [m for m in a.members if (a.dimension, m) not in drop]) for a in cols]
    if any(not a.members for a in rows2 + cols2):
        raise InputError("Smart View rejects every member on an axis: "
                         + "; ".join(_explain_reject(conn, d, m) for d, m in rej)
                         + ". Next step: pick valid members with sv_find_member / sv_members.")
    lay, p, secs = _query_once(h, conn, pov, rows2, cols2, suppress_missing, suppress_zeros)
    p.pov_problems.insert(0, "Smart View rejected these members (rc=1000), so they were left out: "
                          + "; ".join(_explain_reject(conn, d, m) for d, m in rej)
                          + ". Pick a more specific member with sv_find_member / sv_members.")
    return lay, p, secs


def _query_once(h: ExcelHost, conn: str, pov: dict, rows: list[Axis], cols: list[Axis],
                suppress_missing: bool, suppress_zeros: bool, _retried: bool = False):
    from .grid import validate_axes
    sheet = ensure_grid_sheet(h, conn)
    dims = _dims_cache[conn]
    pov = {canon_dim(conn, d): m for d, m in pov.items()}
    rows = [Axis(canon_dim(conn, a.dimension), a.members) for a in rows]
    cols = [Axis(canon_dim(conn, a.dimension), a.members) for a in cols]
    problems = validate_axes(dims, pov, rows, cols)
    if problems:
        raise InputError("Invalid query: " + "; ".join(problems))
    lay = Layout(full_pov(dims, pov, rows, cols), rows, cols)
    ws = h.wb.Worksheets(sheet)
    ws.Activate()
    ws.Cells.ClearContents()
    m = lay.matrix()
    height, width = len(m), len(m[0])
    ws.Range(ws.Cells(1, 1), ws.Cells(1 + lay.nc, width)).NumberFormat = "@"
    ws.Range(ws.Cells(2 + lay.nc, 1), ws.Cells(height, lay.nr)).NumberFormat = "@"
    ws.Range(ws.Cells(1, 1), ws.Cells(height, width)).Value = m
    t0 = time.monotonic()
    try:
        h.run("B_Retrieve", sheet)
    except HostError as e:
        if not _retried and (is_rc1000(e) or "values different" in str(e)):
            # Oracle known issue 39163160: refresh after idle can fail once; retry ONCE on a fresh grid sheet
            ensure_grid_sheet(h, conn, fresh=True)
            return _query_once(h, conn, pov, rows, cols, suppress_missing, suppress_zeros, _retried=True)
        raise
    secs = time.monotonic() - t0
    vals = _used_values(ws)
    p = parse(lay, vals, suppress_missing, suppress_zeros)
    from .userexcel import _verify_labels
    bad = _verify_labels(lay, parse(lay, vals, False, False))
    if bad:
        raise HostError("Smart View returned a different layout than requested (rows/columns moved): "
                        + "; ".join(bad[:5]))
    return lay, p, secs


def live_dimension(h: ExcelHost, conn: str, member: str) -> str | None:
    """Dimension of an exact member name, asked live (B_FindMember only); None when EPM does not return it. Smart View
    gives the same kind of error for 'no such member' and a broken session - live_confirm tells them apart."""
    r = h._run("B_FindMember", ensure_grid_sheet(h, conn), member, "Default")
    if isinstance(r, tuple) and r[0] == "OK":
        return str(r[1])
    status = str(r[0] if isinstance(r, tuple) else r)
    if status.startswith("ERR|VBA"):
        raise HostError(f"live member check failed: {status}")
    return None


def live_confirm(h: ExcelHost):
    """confirm(conn, dim, member) for check_members: the member exists live in that dimension. A miss counts only
    when the check itself works right then - a member the atlas has at the top of that dimension is found (asked
    again on every miss: misses are rare); otherwise HostError, so the member stays unverified, not rejected."""
    def confirm(conn: str, dim: str, member: str) -> bool:
        d = live_dimension(h, conn, member)
        if d is not None:
            return canon_dim(conn, d).casefold() == canon_dim(conn, dim).casefold()
        probe = _atlas_top(conn, dim) or dim
        if live_dimension(h, conn, probe) is None:
            raise HostError(f"live member check on {conn} does not work now (even '{probe}' is not found)")
        return False
    return confirm


def _atlas_top(conn: str, dim: str) -> str | None:
    from . import atlas
    try:
        r = atlas.connect().execute("SELECT name FROM member WHERE connection=? AND dimension=? ORDER BY depth, ord "
                                    "LIMIT 1", (conn, dim)).fetchone()
    except Exception:       # noqa: BLE001 - no atlas: probe the dimension name
        return None
    return r[0] if r else None


def find_member(h: ExcelHost, conn: str, member: str) -> dict:
    sheet = ensure_grid_sheet(h, conn)
    r = h._run("B_FindMember", sheet, member, "Default")
    if not (isinstance(r, tuple) and r[0] == "OK"):
        return {"found": False, "member": member, "detail": str(r[0] if isinstance(r, tuple) else r)}
    out = {"found": True, "member": member, "dimension": r[1], "alias": r[2], "generation": r[3], "level": r[4]}
    info = h._run("B_MemberInfoEx", sheet, member)
    if isinstance(info, tuple) and info[0] == "OK":
        props = dict(zip(info[1], info[2]))
        keep = ("Name", "Dimension", "Level", "Generation", "Description", "ParentMbrName", "ChildMbrName",
                "ChildCount", "Consolidation Operator", "Data Storage", "Data Type", "AliasNames", "Udas")
        out["info"] = {k: props.get(k) for k in keep if k in props}
    return out


def dump_dimension(h: ExcelHost, conn: str, dim: str, root: str | None = None) -> list[dict]:
    """Zoom a dimension's tree (all levels) on a grid and return [{label, name, depth, parent}] top-down."""
    sheet = f"A_{conn}"[:31]
    ensure_grid_sheet(h, conn, sheet=sheet, fresh=True)
    dims = _dims_cache[conn]
    other = next(d for d in dims if d != dim)
    lay = Layout(full_pov(dims, {}, [Axis(dim, [root or dim])], [Axis(other, [other])]),
                 [Axis(dim, [root or dim])], [Axis(other, [other])])
    ws = h.wb.Worksheets(sheet)
    ws.Activate()
    ws.Cells.ClearContents()
    m = lay.matrix()
    ws.Range(ws.Cells(1, 1), ws.Cells(len(m), len(m[0]))).NumberFormat = "@"
    ws.Range(ws.Cells(1, 1), ws.Cells(len(m), len(m[0]))).Value = m
    _apply_sheet_options(h, sheet, {5: 1, 33: True})     # indent sub-items; navigate without data
    h.run("B_Retrieve", sheet)
    h.run("B_ZoomIn", sheet, "A3", 1, False)            # all levels
    vals = _used_values(ws)
    labels = [r[0] for r in vals[2:] if r and isinstance(r[0], str) and r[0].strip()]
    _apply_sheet_options(h, sheet, {5: 0, 33: False})
    return tree_from_labels(labels)


def tree_from_labels(labels: list[str]) -> list[dict]:
    """Indentation -> depth/parent. Handles both ancestor-on-top and ancestor-at-bottom ordering."""
    if not labels:
        return []
    ind = [indent_of(x) for x in labels]
    bottom_up = ind[-1] < ind[0]                # top member last = ancestors below their children
    seq = list(zip(labels, ind))
    if bottom_up:
        seq.reverse()
    base = min(i for _, i in seq)
    steps = sorted({i for _, i in seq})
    step = (steps[1] - steps[0]) if len(steps) > 1 else 1
    out, stack = [], []
    for order, (lab, i) in enumerate(seq):
        depth = (i - base) // step if step else 0
        while len(stack) > depth:
            stack.pop()
        parent = stack[-1] if stack else None
        clean = lab.strip()
        name = name_from_label(clean) or clean
        out.append({"label": clean, "name": name, "depth": depth, "parent": parent, "order": order})
        stack.append(name)
    return out


def save_table(h: ExcelHost, path: Path, sheets: list[tuple[str, list[list]]]) -> Path:
    """Write tables into a new workbook, apply the default sensitivity label, SaveAs xlsx (Excel cancels
    unlabelled saves silently, so the label is mandatory)."""
    wb = h.com(h.app.Workbooks.Add)
    try:
        first = True
        for name, rows in sheets:
            ws = wb.Worksheets(1) if first else wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
            first = False
            ws.Name = re.sub(r"[\[\]:*?/\\]", "_", name)[:31]
            if rows:
                width = max(len(r) for r in rows)
                data = [list(r) + [None] * (width - len(r)) for r in rows]
                ws.Range(ws.Cells(1, 1), ws.Cells(len(data), width)).Value = data
        _label(wb)
        from .userexcel import is_archive
        if is_archive(path):
            raise InputError("Refusing to save into the archive.")
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists():
            path.unlink()
        h.com(lambda: wb.SaveAs(str(path), 51))
        if not path.exists():
            raise HostError(f"Excel did not save {path} (sensitivity label rejected?)")
        return path
    finally:
        _close_quietly(h, wb)


def _close_quietly(h: ExcelHost, wb) -> None:
    """Close a workbook of ours without saving. Right after a large SaveAs to OneDrive, Close can fail once with
    0x800A03EC while Excel is still busy with the file - retry, then mark it saved and try again; never let a
    failed close undo a finished save."""
    for attempt in range(4):
        try:
            if attempt >= 2:
                wb.Saved = True
            h.com(lambda: wb.Close(False))
            return
        except Exception:   # noqa: BLE001
            time.sleep(1.5 * (attempt + 1))
    log.warning("could not close a saved workbook; it stays open in the hidden Excel until restart")


def _label(wb):
    try:
        cur = wb.SensitivityLabel.GetLabel()
        if cur.LabelId:
            return
    except Exception:   # noqa: BLE001
        pass
    li = wb.SensitivityLabel.CreateLabelInfo()
    li.LabelId = config.LABEL_ID
    li.SiteId = config.LABEL_SITE
    li.AssignmentMethod = 1
    li.Justification = ""
    wb.SensitivityLabel.SetLabel(li, li)


def refresh_workbook(h: ExcelHost, src: Path, out: Path) -> dict:
    """Copy `src` to the work dir, refresh every Smart View sheet, save the result to `out` (never the source)."""
    config.WORK_DIR.mkdir(parents=True, exist_ok=True)
    work = config.WORK_DIR / f"{dt.datetime.now():%Y%m%d_%H%M%S}{src.suffix}"
    shutil.copyfile(src, work)
    com = h.com
    from .userexcel import is_archive
    if is_archive(out):
        raise InputError("Refusing to save a refreshed copy into the archive.")
    # Macros in the user's workbook must not run in our Excel (an event macro could call a Smart View submit).
    prev_sec = h.app.AutomationSecurity
    h.app.AutomationSecurity = 3                  # msoAutomationSecurityForceDisable, only for this Open
    try:
        wb = com(lambda: h.app.Workbooks.Open(str(work), 0, False), limit_s=300)
    finally:
        h.app.AutomationSecurity = prev_sec
    results, pivots = [], 0
    try:
        # Smart View VBA functions act on the active workbook's sheets by name
        com(wb.Activate)
        for ws in list(wb.Worksheets):
            name = ws.Name
            com(ws.Activate)
            r = h._run("B_SheetInfo", name)
            meta = dict(zip(r[1], r[2])) if isinstance(r, tuple) and r[0] == "OK" else None
            conn = (meta or {}).get("Friendly Name") or (meta or {}).get("Cube")
            if not conn:
                results.append({"sheet": name, "smartview": False})
                continue
            t0 = time.monotonic()
            rr = h._run("B_Retrieve", name)
            if isinstance(rr, str) and rr.startswith("ERR|"):
                h._run("B_Connect", name, conn)
                rr = h._run("B_Retrieve", name)
            results.append({"sheet": name, "smartview": True, "connection": conn, "cube": meta.get("Cube"),
                            "result": rr if rr == "OK" else _explain_status(rr),
                            "seconds": round(time.monotonic() - t0, 1)})
        # Excel pivot tables built on the refreshed grids (no EPM access involved)
        for pc in list(wb.PivotCaches()):
            try:
                com(pc.Refresh)
                pivots += 1
            except Exception:   # noqa: BLE001 - a broken pivot must not block the refresh
                pass
        com(h.app.CalculateFull)
        _label(wb)
        out.parent.mkdir(parents=True, exist_ok=True)
        if out.exists():
            out.unlink()
        com(lambda: wb.SaveAs(str(out), 52 if out.suffix.lower() == ".xlsm" else 51), limit_s=300)
        saved = out.exists()
    finally:
        com(lambda: wb.Close(False))
        com(h.wb.Activate)
        try:
            os.remove(work)
        except OSError:
            pass
    return {"source": str(src), "output": str(out), "saved": saved, "pivot_caches_refreshed": pivots,
            "sheets": results}


def _explain_status(status) -> str:
    from .host import _explain
    return _explain("Retrieve", status) if isinstance(status, str) else str(status)


DRILL_DEFAULTS = {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic"}


def drill_link(h: ExcelHost, conn: str, cell: dict[str, str]) -> dict:
    """Retrieve ONE cell (account on rows, period on columns, the rest as POV) on our hidden grid sheet and return
    {cube_value, url, report}. url is relative to config.EPM_BASE. Read-only: the report is listed, never run."""
    need = {"All Accounts", "All Time Periods", "All Entities"}
    missing = need - set(cell)
    if missing:
        raise InputError(f"sv_drill needs these dimensions in cell: {sorted(missing)}")
    if cell.get("All Balance Types", "Periodic") != "Periodic":
        raise InputError("Drill-through works on Periodic data only; use a period member (quarter/year) instead "
                        "of Accumulated.")
    pov = {**DRILL_DEFAULTS, **{d: m for d, m in cell.items() if d not in ("All Accounts", "All Time Periods")}}
    lay, _, _ = query(h, conn, pov, [Axis("All Accounts", [cell["All Accounts"]])],
                      [Axis("All Time Periods", [cell["All Time Periods"]])], False, False)
    sheet = ensure_grid_sheet(h, conn)
    addr = f"${chr(ord('A') + lay.nr)}${2 + lay.nc}"          # first data cell: row 2+nc, column nr+1
    val = h.wb.Worksheets(sheet).Range(addr).Value
    try:
        r = h.run("B_DrillReports", sheet, addr)
    except HostError as e:
        if "no drill-through reports found" not in str(e).lower():     # rc=61161: data not from Data Integration
            raise
        return {"cube_value": val if isinstance(val, float) else None, "url": None, "report": None}
    if not isinstance(r, tuple) or r[0] != "OK":
        raise HostError(f"Smart View could not list the drill reports of this cell: {r[0] if isinstance(r, tuple) else r}")
    names = r[1] if isinstance(r[1], tuple) else (r[1],)
    urls = r[2] if isinstance(r[2], tuple) else (r[2],)
    if not names or not names[0]:
        return {"cube_value": val if isinstance(val, float) else None, "url": None, "report": None}
    i = next((k for k, n in enumerate(names) if "source" in str(n).lower()), 0)
    return {"cube_value": val if isinstance(val, float) else None, "url": str(urls[i]).lstrip("./"),
            "report": names[i]}
