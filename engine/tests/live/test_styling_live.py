"""Live test of smartview_mcp.styling against its OWN Excel instance (DispatchEx; never attaches to a running Excel).

Run (from engine):  uv run python tests\\live\\test_styling_live.py
Under pytest it skips cleanly when Excel / pywin32 is not available.
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from smartass.core import config  # noqa: E402

WORK = Path(config.ANALYSER_DIR) / "_matrix"      # same scratch folder as matrix.py


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", s.replace(" ", " ").replace(" ", " ")).strip()


def run() -> None:
    import pythoncom
    import win32api
    import win32com.client
    import win32process
    from smartview_mcp import styling
    from smartview_mcp.excel import instances
    from smartview_mcp import styling as st
    tn = st.tn

    WORK.mkdir(parents=True, exist_ok=True)

    def fmt_ok(cell, us):
        """NumberFormat read-back: US syntax or (when Excel returns local syntax) NumberFormatLocal == translated."""
        intl = cell.Application.International
        return cell.NumberFormat == us or cell.NumberFormatLocal == tn.to_local(us, intl[2], intl[3])
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    _, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
    print(f"own Excel PID {pid}")
    try:
        xl.DisplayAlerts = False
        xl.Visible = True          # visible window so freeze panes can be tested
        xl.WindowState = -4137     # xlMaximized: a fresh DispatchEx window has no usable client area for FreezePanes
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Answer"

        # ------------------------------------------------ answer layout (content only)
        put = lambda a, v: setattr(ws.Range(a), "Value2", v)
        put("A1", None)
        put("B1", "Kostnadsbase Wholesale – testrapport")
        put("B2", "Kilde: FinMax, Actual, NOK, YTD Sep 26 · generert av styling-test")
        put("B4", "Hovedpunkter: kostnadene i Wholesale er 12,4 % høyere enn i fjor, drevet av innleide tjenester og "
                  "avskrivninger. Mobil står for størstedelen av økningen, mens fast nett er stabilt. Sum og kontroller "
                  "under er hentet direkte fra regnskapet og skal gå i null mot hovedboken.")
        put("B6", "Kostnad per produktgruppe")
        for a, t in zip("BCDEF", ["Produktgruppe", "Kostnad 2026 (NOK)", "Kostnad 2025 (NOK)", "Endring (NOK)", "Endring %"]):
            put(f"{a}7", t)
        data = [("Mobil", 1234567, 1000000), ("Fast nett", -1234567, -1200000), ("Tv og innhold", 0, 0),
                ("Tjenester (stort tall)", 98765432100, 90000000000)]
        for i, (n, a, b) in enumerate(data, start=8):
            put(f"B{i}", n); put(f"C{i}", a); put(f"D{i}", b)
            put(f"E{i}", f"=C{i}-D{i}"); put(f"F{i}", f"=IF(D{i}=0,0,E{i}/ABS(D{i}))")
        put("B12", "Sum produktgrupper")
        for a in "CDE":
            put(f"{a}12", f"=SUM({a}8:{a}11)")
        put("F12", "=IF(D12=0,0,E12/ABS(D12))")
        put("B13", "Totalt inkl. felleskostnader")
        put("C13", "=C12+500000"); put("D13", "=D12+400000"); put("E13", "=C13-D13"); put("F13", "=IF(D13=0,0,E13/ABS(D13))")
        put("B15", "Note: beløp i NOK; negative tall i parentes; null vises som «–».")
        put("B17", "Kontroller")
        put("B18", "Sum mot hovedbok (differanse)"); put("C18", 0); put("D18", "OK")
        put("B19", "Totalt mot resultatregnskap (differanse)"); put("C19", 12345); put("D19", "AVVIK")

        num4 = {"C": "heltall", "D": "heltall", "E": "heltall", "F": "prosent1"}
        roles = {
            "last_col": "F", "label_cols": ["B"],
            "col_widths": {"A": 2.5, "B": 46, "C:F": 16}, "freeze_rows": 3, "print_header": "Intern",
            "rows": [
                {"row": 1, "kind": "title"}, {"row": 2, "kind": "subtitle"}, {"row": 3, "kind": "blank"},
                {"row": 4, "kind": "text"}, {"row": 5, "kind": "blank"}, {"row": 6, "kind": "section"},
                {"row": 7, "kind": "header", "num_cols": ["C", "D", "E", "F"]},
                *({"row": r, "kind": "line", "num": num4} for r in (8, 9, 10, 11)),
                {"row": 12, "kind": "sum", "num": num4}, {"row": 13, "kind": "grand", "num": num4},
                {"row": 14, "kind": "blank"}, {"row": 15, "kind": "note"}, {"row": 16, "kind": "blank"},
                {"row": 17, "kind": "section"},
                {"row": 18, "kind": "check", "num": {"C": "heltall"}, "status_col": "D"},
                {"row": 19, "kind": "check", "num": {"C": "heltall"}, "status_col": "D"},
            ],
        }
        before = {a: ws.Range(a).Formula for a in ("B1", "C8", "E8", "C12", "D18")}
        styling.ensure_styles(wb)
        styling.ensure_styles(wb)                       # idempotent
        findings = styling.apply_report(ws, roles)
        print("lint findings:", findings)
        assert findings == [], findings
        assert before == {a: ws.Range(a).Formula for a in before}, "contents changed"

        # ------------------------------------------------ assertions on the answer sheet
        assert fmt_ok(ws.Range("C8"), tn.NUM["heltall"])
        assert fmt_ok(ws.Range("F8"), tn.NUM["prosent1"])
        assert ws.Range("C8").Style.Name == "TN Heltall"
        assert ws.Range("C12").Style.Name == "TN Heltall · Sum"
        assert ws.Range("C13").Style.Name == "TN Heltall · Hovedsum"
        assert ws.Range("F13").Style.Name == "TN Prosent 1d · Hovedsum"
        assert ws.Range("B1").Style.Name == "TN Tittel" and ws.Range("B2").Style.Name == "TN Undertittel"
        assert ws.Range("B6").Style.Name == "TN Seksjon" and ws.Range("C7").Style.Name == "TN Kolonne tall"
        assert ws.Range("B7").Style.Name == "TN Kolonne" and ws.Range("B15").Style.Name == "TN Note"
        assert ws.Range("B4").Style.Name == "TN Tekst" and ws.Range("B4").MergeCells
        assert ws.Range("C12").Borders(8).LineStyle == 1 and ws.Range("C12").Font.Bold
        assert ws.Range("C13").Borders(9).LineStyle == -4119            # double under hovedsum
        assert ws.Range("C13").Borders(8).LineStyle == 1
        for edge in (7, 10, 11):
            assert ws.UsedRange.Borders(edge).LineStyle == -4142, f"vertical border edge {edge}"
        assert ws.UsedRange.Font.Name == "Arial"
        assert ws.Range("C8").Text and _norm(ws.Range("C9").Text) == "(1 234 567)", repr(ws.Range("C9").Text)
        assert _norm(ws.Range("C10").Text) == "–", repr(ws.Range("C10").Text)
        assert _norm(ws.Range("C11").Text) == "98 765 432 100", repr(ws.Range("C11").Text)
        print("Text -1234567 =", repr(ws.Range("C9").Text), "| 0 =", repr(ws.Range("C10").Text),
              "| F-pct =", repr(ws.Range("F8").Text))
        fcs = ws.Range("D18").FormatConditions
        assert fcs.Count == 2 and ws.Range("D19").FormatConditions.Count == 2
        assert ws.Range("D18").DisplayFormat.Font.Color == tn.rgb(st.GREEN), "OK should be green"
        assert ws.Range("D19").DisplayFormat.Font.Color == tn.rgb(tn.BREACH_FONT), "AVVIK should be red"
        assert ws.Range("D19").DisplayFormat.Interior.Color == tn.rgb(tn.BREACH_FILL)
        win = wb.Windows(1)
        assert win.DisplayGridlines is False, "gridlines should be off"
        if ":" not in win.VisibleRange.Address:
            # a DispatchEx Excel started from a console can have a window without a laid-out grid (VisibleRange is
            # one cell, even maximised and in the foreground): Excel then refuses FreezePanes, and _freeze leaves no
            # half split behind. Answers in a real Excel window do freeze (checked 2026-10-26 on the answer sheets).
            print("SKIP freeze panes: this Excel's window has no grid area (VisibleRange", win.VisibleRange.Address + ")")
            assert not win.FreezePanes and win.SplitRow == 0, "no half split may be left behind"
        else:
            assert win.FreezePanes and win.SplitRow == 3, (win.FreezePanes, win.SplitRow, win.Visible)

        # ------------------------------------------------ PDF export + save
        pdf = WORK / "styling_test.pdf"
        if pdf.exists():
            pdf.unlink()
        ws.ExportAsFixedFormat(0, str(pdf))
        assert pdf.exists() and pdf.stat().st_size > 1000
        try:
            import pymupdf as fitz
            doc = fitz.open(str(pdf))
            png = WORK / "styling_test_p1.png"
            doc[0].get_pixmap(dpi=130).save(str(png))
            print("PNG:", png)
        except ImportError:
            print("pymupdf missing - PNG skipped")

        # ------------------------------------------------ Smart View style data grid
        ds = wb.Worksheets.Add(After=ws)
        ds.Name = "Grid"
        ds.Range("A1").Value2 = "POV: Actual / NOK"
        for i, p in enumerate(["Jul 26", "Aug 26", "Sep 26", "YTD"], start=2):
            ds.Cells(2, i).Value2 = p
        rows = [("ac.C30 Revenue", 1234567.891, -2500000.4, 0, "#Missing"), ("ac.C40 COGS", -987654, 12.3, 5e9, -1),
                ("  ac.C41 Sub", 100, 200, 300, 600), ("FTE", 123.456, 124.5, 0, 125)]
        for r, (lab, *vals) in enumerate(rows, start=3):
            ds.Cells(r, 1).Value2 = lab
            for c, v in enumerate(vals, start=2):
                ds.Cells(r, c).Value2 = v
        ds.Range("A3:A6").IndentLevel = 0
        ds.Range("A4").IndentLevel = 2
        ds.Range("F3").Formula = "=B3+C3"
        snapshot = ds.UsedRange.Formula
        styling.apply_data_sheet(ds, header_rows=2, label_cols=1, fte_rows=[6])
        assert snapshot == ds.UsedRange.Formula, "contents changed"
        assert ds.UsedRange.Address == "$A$1:$F$6", ds.UsedRange.Address
        assert fmt_ok(ds.Range("B3"), tn.NUM["heltall"])
        assert fmt_ok(ds.Range("D5"), tn.NUM["heltall"]) and fmt_ok(ds.Range("F3"), tn.NUM["heltall"])
        assert fmt_ok(ds.Range("B6"), st.TALL1_THOUSANDS) and fmt_ok(ds.Range("E6"), st.TALL1_THOUSANDS)
        assert ds.Range("A4").IndentLevel == 2 and ds.Range("A4").HorizontalAlignment == -4131
        print("grid texts:", [_norm(ds.Range(a).Text) for a in ("B3", "C3", "D3", "B4", "D4", "B6", "D6", "F3")])
        assert _norm(ds.Range("B3").Text) == "1 234 568" and _norm(ds.Range("C3").Text) == "(2 500 000)"
        assert _norm(ds.Range("D3").Text) == "–" and _norm(ds.Range("B4").Text) == "(987 654)"
        assert _norm(ds.Range("B6").Text) == "123,5" and _norm(ds.Range("D6").Text) == "–"
        assert ds.Range("A1").Font.Bold and ds.Range("B2").Font.Bold and ds.Range("B2").Font.Color == tn.rgb(tn.NAVY)
        assert ds.Range("A2").Borders(9).LineStyle == 1 and not ds.Range("A3").Font.Bold
        assert ds.Range("A1").Font.Name == "Arial" and ds.Range("E6").Font.Name == "Arial"
        assert ds.Columns(1).ColumnWidth > 40 and 14 < ds.Columns(3).ColumnWidth < 18
        assert ds.Tab.Color == tn.rgb(st.GREY_TAB)
        ds.Activate()
        w = wb.Windows(1)
        assert w.FreezePanes and w.SplitRow == 2 and w.SplitColumn == 1, (w.SplitRow, w.SplitColumn)
        for edge in (7, 10, 11):
            assert ds.UsedRange.Borders(edge).LineStyle == -4142
        # hidden window: freeze is skipped silently
        w.Visible = False
        styling.apply_data_sheet(ds, 2, 1)
        w.Visible = True

        # ------------------------------------------------ save (sensitivity label first)
        instances._label(wb)
        out = WORK / "styling_test.xlsx"
        if out.exists():
            out.unlink()
        wb.SaveAs(str(out), 51)
        assert out.exists(), "SaveAs silently cancelled"
        print("saved", out)
        print("ALL STYLING ASSERTIONS PASSED")
        wb.Close(False)
    finally:
        try:
            xl.Quit()
        except Exception:   # noqa: BLE001
            pass
        del xl
        try:
            h = win32api.OpenProcess(0x0400, False, pid)   # PROCESS_QUERY_INFORMATION: still alive?
            win32api.CloseHandle(h)
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
        except Exception:   # noqa: BLE001
            pass


def test_styling_live():
    import pytest
    pytest.importorskip("win32com")
    pytest.importorskip("openpyxl")
    try:
        import win32com.client
        win32com.client.Dispatch  # noqa: B018
    except Exception:   # noqa: BLE001
        pytest.skip("no pywin32")
    if os.environ.get("SV_STYLING_LIVE") != "1":
        pytest.skip("live Excel test; set SV_STYLING_LIVE=1 or run as a script")
    run()


if __name__ == "__main__":
    run()
