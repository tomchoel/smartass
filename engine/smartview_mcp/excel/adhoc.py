"""Smart View ad hoc pulls into the user's sheets: write the grid, retrieve, drop rejected members."""
from __future__ import annotations

import time

from . import jobs, login, sheets, target
from ..grid import Axis, full_pov, Layout, parse, verify_labels as _verify_labels
from ..host import ExcelHost, HostError, InputError

def pull(h: ExcelHost, t: target.Target, q: int, connection: str, pov: dict, rows: list[Axis], cols: list[Axis],
         title: str | None = None) -> sheets.Pull:
    """A live Smart View grid in the user's workbook; Smart View's modal message boxes in that Excel are dismissed
    while it runs (SmartViewBoxes), so they cannot block the pull."""
    with login.SmartViewBoxes(t.pid) as boxes:
        return _pull(h, t, q, connection, pov, rows, cols, title, boxes)


def _pull(h: ExcelHost, t: target.Target, q: int, connection: str, pov: dict, rows: list[Axis], cols: list[Axis],
          title: str | None, boxes: "login.SmartViewBoxes") -> sheets.Pull:
    from .. import styling                         # provided by the styling adapter
    from ..grid import validate_axes
    from ..ops import _dims_cache, _grid_labels, _used_values, canon_dim, dims_known, is_rc1000, learn_dims
    from ..timing import stage

    def layout():
        p_ = {canon_dim(connection, d): m for d, m in pov.items()}
        r_ = [Axis(canon_dim(connection, a.dimension), a.members) for a in rows]
        c_ = [Axis(canon_dim(connection, a.dimension), a.members) for a in cols]
        dims = _dims_cache[connection]
        problems = validate_axes(dims, p_, r_, c_)
        if problems:
            raise InputError("Invalid pull: " + "; ".join(problems))
        return Layout(full_pov(dims, p_, r_, c_), r_, c_)

    # dimension labels come from the on-disk cache, else from this sheet's own default grid below - never from the
    # hidden server Excel (a second Excel + login, and a connect there once hung a pull for minutes)
    lay = layout() if dims_known(connection) else None
    name = sheets.unique_sheet_name(t, title or f"Q{q} {connection}")
    t.status(f"fetching {connection} into the sheet '{name}' ...")
    with stage("add sheet"):
        jobs.com(t.wb.Activate)
        ws = sheets.add_sheet(t, name, after_q=q, color=0xA6A6A6)
        jobs.com(ws.Activate)
    with stage("connect"):
        from ..dialogs import not_signed_in
        failed: list[list[str]] = []              # box texts of each failed attempt
        for attempt in range(3):                  # EPM sometimes answers a connect with its extension manifest
            n0 = len(boxes.texts)
            st = t.run("B_Connect", name, connection)
            if st == "OK":
                break
            failed.append(boxes.texts[n0:])
            if not_signed_in(failed):             # retrying cannot fix a session EPM keeps rejecting
                raise HostError(login.NOT_SIGNED_IN)
            if attempt == 2:
                from ..host import _explain
                raise HostError(_explain("Connect", st))
            time.sleep(4 * (attempt + 1))
    # Always let Smart View open its own default grid first (1-2 s). Skipping it was tried on 2026-10-14 and was
    # followed by a 752 error and two Excel crashes (HsAddin.dll access violation) - not worth the time.
    with stage("open default grid"):
        _open_default_grid(t, name)
    if lay is None:                               # first pull of this cube ever: learn its dimensions here
        with stage("learn dimensions"):
            r = t.run("B_GetDimensions", name)
            labels = _grid_labels(jobs.com(lambda: _used_values(ws)), list(r[2])) if isinstance(r, tuple) else None
            if not labels:
                raise HostError(f"Could not read the {connection} dimensions from the default grid on '{name}'.")
            learn_dims(connection, list(r[1]), list(r[2]), labels)
            lay = layout()
    with stage("sheet options"):
        for k, v in sheets.DATA_SHEET_OPTIONS.items():
            t.run("B_SetOption", k, v, name)
    rejected: list[tuple[str, str]] = []
    for attempt in range(2):
        with stage("write layout + retrieve"):
            st, secs = _write_and_retrieve(t, ws, name, lay)
        if st != "OK" and is_rc1000(st) and not rejected:
            with stage("isolate rejected members"):
                lay, rejected = _drop_rejected(t, ws, name, lay, connection)
            if rejected:
                with stage("write layout + retrieve"):
                    st, secs = _write_and_retrieve(t, ws, name, lay)
        if st != "OK":
            raise HostError(f"Retrieve on '{name}' failed: {st}")
        with stage("read + verify"):
            vals = jobs.com(lambda: _used_values(ws))
            parsed = parse(lay, vals, suppress_missing=False, suppress_zeros=False)
            bad = _verify_labels(lay, parsed)
        if not bad or attempt or not _fix_sheet_mode(t, name):
            break                                 # labels off because the sheet opened in alias 'none': fixed, redo
    if bad:
        raise HostError("Smart View returned a different layout than requested, so the numbers cannot be mapped "
                        "safely: " + "; ".join(bad[:5]) + (" ..." if len(bad) > 5 else ""))
    if rejected:
        from ..ops import rejected_note
        parsed.pov_problems.insert(0, rejected_note(connection, rejected))
    fte = sheets.fte_rows(lay, parsed)
    t.status(f"formatting '{name}' ...")
    p = sheets._make_pull(name, connection, lay, parsed, secs)
    with jobs.fast_excel(t.app):
        with stage("style data sheet"):
            styling.apply_data_sheet(ws, header_rows=1 + lay.nc, label_cols=lay.nr, fte_rows=fte)
        with stage("save meta"):
            try:
                sheets.save_meta(t, p)
            except Exception:   # noqa: BLE001 - metadata only helps after restarts
                pass
    return p


def refresh_pull(t: target.Target, p: sheets.Pull) -> sheets.Pull:
    """Retrieve an existing data sheet again IN PLACE (same layout): its name, formulas pointing at it and its
    formatting stay; only the numbers are refreshed. Smart View's message boxes are dismissed meanwhile."""
    with login.SmartViewBoxes(t.pid):
        return _refresh_pull(t, p)


def _refresh_pull(t: target.Target, p: sheets.Pull) -> sheets.Pull:
    from ..ops import _used_values
    from ..timing import stage
    ws = jobs.com(lambda: t.wb.Worksheets(p.sheet))
    t.status(f"refreshing '{p.sheet}' ...")
    with stage("retrieve in place"):
        t0 = time.monotonic()
        st = t.run("B_Retrieve", p.sheet)
        secs = time.monotonic() - t0
    if st != "OK":
        raise HostError(f"Retrieve on '{p.sheet}' failed: {st}")
    with stage("read + verify"):
        parsed = parse(p.layout, jobs.com(lambda: _used_values(ws)), suppress_missing=False, suppress_zeros=False)
        bad = _verify_labels(p.layout, parsed)
    if bad:
        raise HostError(f"'{p.sheet}' no longer has the expected layout: " + "; ".join(bad[:5])
                        + ". Re-run sv_answer for this question.")
    parsed.pov_problems[:0] = [w for w in p.parsed.pov_problems if w.startswith("Smart View rejected")]
    return sheets._make_pull(p.sheet, p.connection, p.layout, parsed, secs)


def _write_and_retrieve(t: target.Target, ws, name: str, lay: Layout) -> tuple[str, float]:
    jobs.com(lambda: ws.Cells.ClearContents())
    m = lay.matrix()
    height, width = len(m), len(m[0])
    jobs.com(lambda: setattr(ws.Range(ws.Cells(1, 1), ws.Cells(1 + lay.nc, width)), "NumberFormat", "@"))
    jobs.com(lambda: setattr(ws.Range(ws.Cells(2 + lay.nc, 1), ws.Cells(height, lay.nr)), "NumberFormat", "@"))
    jobs.com(lambda: setattr(ws.Range(ws.Cells(1, 1), ws.Cells(height, width)), "Value", m))
    t0 = time.monotonic()
    st = t.run("B_Retrieve", name)
    return st, time.monotonic() - t0


def _drop_rejected(t: target.Target, ws, name: str, lay: Layout, connection: str) -> tuple[Layout, list[tuple[str, str]]]:
    """rc=1000: bisect each axis on this sheet to find the members Smart View rejects; return the layout without
    them (unchanged + [] if none could be isolated or an axis would be empty)."""
    from ..ops import MAX_BISECT_RETRIEVES, is_rc1000, isolate_rejected, without
    axes = lay.rows + lay.cols

    def ok(trial):
        lt = Layout(lay.pov, trial[:len(lay.rows)], trial[len(lay.rows):])
        st, _ = _write_and_retrieve(t, ws, name, lt)
        if st != "OK" and not is_rc1000(st):
            raise HostError(f"Retrieve on '{name}' failed: {st}")
        return st == "OK"

    rej = isolate_rejected(axes, ok, f"Retrieve on '{name}' was rejected (rc=1000) and the bad member could not be "
                           f"isolated within {MAX_BISECT_RETRIEVES} test retrieves; split the pull.")
    new = without(axes, rej)
    if not rej or any(not a.members for a in new):
        return lay, []
    return Layout(lay.pov, new[:len(lay.rows)], new[len(lay.rows):]), rej


def _sheet_mode_bad(t: target.Target, name: str) -> bool:
    info = t.run("B_SheetInfo", name)
    meta = dict(zip(info[1], info[2])) if isinstance(info, tuple) and info[0] == "OK" else {}
    return "Standard" not in str(meta.get("Sheet Type", "")) or str(meta.get("Alias Table", "")).lower() == "none"


def _open_default_grid(t: target.Target, name: str) -> None:
    st = t.run("B_Retrieve", name)
    if st != "OK":
        raise HostError(f"Smart View could not open a grid on '{name}': {st}")
    if _sheet_mode_bad(t, name):                  # seen once: FinMax opened as plain 'Ad hoc', alias none
        t.run("B_SetAliasTable", name, "Default")
        t.run("B_Retrieve", name)


def _fix_sheet_mode(t: target.Target, name: str) -> bool:
    """After a label mismatch: if the sheet opened in a non-standard mode / alias 'none', switch to Default."""
    if not _sheet_mode_bad(t, name):
        return False
    t.run("B_SetAliasTable", name, "Default")
    return True
