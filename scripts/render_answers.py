"""Render every visible sheet of a smartass workbook to PNG for a visual review (memory: review-every-answer-visually).

Usage (from engine/):  uv run --with pymupdf python -B ../scripts/render_answers.py "<name fragment>" <out dir>

The workbook may be open in the user's Excel (OneDrive: the file is locked): it is found in the Running Object Table
(its display name is the https URL, so spaces are matched as %20 too) and copied with SaveCopyAs - never saved,
never closed. A closed workbook is copied from disk (give a full path as the fragment). The copy is opened read-only
in a hidden Excel (DispatchEx) and each sheet exported to PDF, then to <tag>_<sheet>_<page>.png (first 2 pages).
"""
import os
import shutil
import sys

import fitz
import pythoncom
import win32com.client as w


def copy_open(fragment: str, out: str) -> bool:
    ctx = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()
    keys = {fragment.lower(), fragment.lower().replace(" ", "%20")}
    for mk in rot:
        name = mk.GetDisplayName(ctx, None).lower()
        if any(k in name for k in keys):
            wb = w.Dispatch(rot.GetObject(mk).QueryInterface(pythoncom.IID_IDispatch))
            wb.SaveCopyAs(out)
            print("copied open workbook", wb.Name)
            return True
    return False


def export(wb, out_dir: str, tag: str) -> None:
    """Each visible sheet of an open workbook to <tag>_<i>.pdf and its first 2 pages to PNG."""
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        if ws.Visible != -1:
            continue
        pdf = os.path.join(out_dir, f"{tag}_{i}.pdf")
        try:
            ws.ExportAsFixedFormat(0, pdf)
        except Exception:   # noqa: BLE001 - an empty sheet has nothing to print
            print(i, ws.Name, 'empty')
            continue
        d = fitz.open(pdf)
        for p in range(min(6 if "Answer" in ws.Name else 2, d.page_count)):   # every page of an answer
            d[p].get_pixmap(dpi=100).save(os.path.join(out_dir, f"{tag}_{i}_{p}.png"))
        print(i, ws.Name, "pages", d.page_count, "print area", ws.PageSetup.PrintArea)
        d.close()


def render(path: str, out_dir: str, tag: str) -> None:
    app = w.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Open(path, 0, True)
        export(wb, out_dir, tag)
        wb.Close(False)
    finally:
        app.Quit()


def main() -> None:
    fragment, out_dir = sys.argv[1], os.path.abspath(sys.argv[2])
    os.makedirs(out_dir, exist_ok=True)
    tag = "answer"
    copy = os.path.join(out_dir, f"{tag}.xlsx")
    pythoncom.CoInitialize()
    if not copy_open(fragment, copy):
        if not os.path.exists(fragment):
            sys.exit(f"not open and not a file: {fragment}")
        shutil.copy(fragment, copy)
        print("copied closed workbook from disk")
    render(copy, out_dir, tag)


if __name__ == "__main__":
    main()
