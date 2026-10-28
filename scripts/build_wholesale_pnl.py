"""Build Wholesale P&L 2026 vs 2025 headlessly with HsGetValue formulas."""
from __future__ import annotations

import io
import os
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# Add engine to path
ENGINE_DIR = Path(__file__).resolve().parents[1] / "engine"
sys.path.insert(0, str(ENGINE_DIR))

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from smartview_mcp.hsgetvalue import build_hsgetvalue_formula
from smartview_mcp import catalog, tools_fin, atlas, fin_pnl, warehouse
import contextlib

MIP_LABEL_ID = "f604d2c9-1577-460e-b668-57374a0216c3"
MIP_SITE_ID = "1676489c-5c72-46b7-ba63-9ab90c4aad44"

CUSTOM_PROPS_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="2" name="MSIP_Label_{MIP_LABEL_ID}_Enabled">
    <vt:lpwstr>true</vt:lpwstr>
  </property>
  <property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="3" name="MSIP_Label_{MIP_LABEL_ID}_SetDate">
    <vt:lpwstr>2026-10-28T00:00:00Z</vt:lpwstr>
  </property>
  <property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="4" name="MSIP_Label_{MIP_LABEL_ID}_Method">
    <vt:lpwstr>Privileged</vt:lpwstr>
  </property>
  <property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="5" name="MSIP_Label_{MIP_LABEL_ID}_SiteId">
    <vt:lpwstr>{MIP_SITE_ID}</vt:lpwstr>
  </property>
</Properties>
"""

CONTENT_TYPES_OVERRIDE = '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
RELS_CUSTOM_PROP = '<Relationship Id="rIdCustomProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'


def generate_wholesale_pnl() -> Path:
    t0 = time.perf_counter()

    # 1. Fetch from warehouse (or live)
    cat = catalog.load()
    with contextlib.closing(atlas.connect()) as adb:
        entity = tools_fin.resolve_scope(cat, "Wholesale", adb)

    store = warehouse.Warehouse()
    res = fin_pnl.compute(cat, fin_pnl.Scope(entity=entity, period="Sep 26", closed="Sep 26", forecast=None), store)
    table_rows = [r for r in fin_pnl.table(res, cat) if r["comparison"] == "ly"]

    # 2. Build OpenXML Workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Wholesale P&L"
    ws.views.sheetView[0].showGridLines = True

    # Styles
    f_title = Font(name="Segoe UI", size=14, bold=True, color="002060")
    f_sub = Font(name="Segoe UI", size=9, italic=True, color="595959")
    f_hdr = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    f_body = Font(name="Segoe UI", size=10)
    f_bold = Font(name="Segoe UI", size=10, bold=True)
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_subhdr = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    b_thin = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                    top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
    b_total = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))

    # Title block
    ws.cell(row=2, column=2, value="Wholesale P&L 2026 vs 2025").font = f_title
    ws.cell(row=3, column=2, value="Entity: 2WHS Wholesale (un.SY.2WHS) | Source: FinMax Actuals | Live Smart View HsGetValue formulas").font = f_sub

    # Header Row 5
    headers_col = [
        ("B", "Linje / Regnskapspost"),
        ("C", "Sep 26 (M)"),
        ("D", "Sep 25 (M)"),
        ("E", "Diff MNOK"),
        ("F", "Endring %"),
        ("G", "YTD Sep 26"),
        ("H", "YTD Sep 25"),
        ("I", "Diff YTD"),
        ("J", "YTD %"),
    ]
    for col_letter, text in headers_col:
        c = ws[f"{col_letter}5"]
        c.value = text
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal="center" if col_letter != "B" else "left", vertical="center")

    account_map = {
        "rev": "ac.C30",
        "cogs": "ac.C40S",
        "gm": "ac.C48",
        "pers": "ac.C50S",
        "oth": "ac.C60",
        "opex": "ac.C70",
        "boi": "ac.EBITDA_BOI",
        "c73": "ac.C73",
        "ebitda": "ac.EBITDA",
        "da": "ac.C80",
        "fte": "ac.STA20Y",
    }

    cached_injections = {}
    row_idx = 6

    for r in table_rows:
        line_key = r["line"]
        if line_key.startswith("kpi."):
            continue
        acct = account_map.get(line_key)
        is_bold = line_key in ("rev", "gm", "opex", "ebitda")
        font_use = f_bold if is_bold else f_body
        border_use = b_total if line_key == "ebitda" else b_thin

        # Col B: Label
        c_lbl = ws.cell(row=row_idx, column=2, value=r["label"])
        c_lbl.font = font_use
        c_lbl.border = border_use

        # If we have an exact FinMax account, create HsGetValue formulas for Month and YTD
        m_cur = r["month"]["cur"] if r["month"] else 0.0
        m_cmp = r["month"]["cmp"] if r["month"] else 0.0
        y_cur = r["ytd"]["cur"] if r["ytd"] else 0.0
        y_cmp = r["ytd"]["cmp"] if r["ytd"] else 0.0

        if acct:
            f_m26 = build_hsgetvalue_formula("FinMax", {
                "All Accounts": acct,
                "All Entities": "un.SY.2WHS",
                "All Time Periods": "Sep 26",
                "All Balance Types": "Periodic",
            })
            f_m25 = build_hsgetvalue_formula("FinMax", {
                "All Accounts": acct,
                "All Entities": "un.SY.2WHS",
                "All Time Periods": "Sep 25",
                "All Balance Types": "Periodic",
            })
            f_y26 = build_hsgetvalue_formula("FinMax", {
                "All Accounts": acct,
                "All Entities": "un.SY.2WHS",
                "All Time Periods": "Sep 26",
                "All Balance Types": "Accumulated",
            })
            f_y25 = build_hsgetvalue_formula("FinMax", {
                "All Accounts": acct,
                "All Entities": "un.SY.2WHS",
                "All Time Periods": "Sep 25",
                "All Balance Types": "Accumulated",
            })
        else:
            f_m26 = str(m_cur)
            f_m25 = str(m_cmp)
            f_y26 = str(y_cur)
            f_y25 = str(y_cmp)

        # Col C: Sep 26 (M)
        c_m26 = ws.cell(row=row_idx, column=3, value=f_m26)
        c_m26.font = font_use
        c_m26.number_format = "#,##0" if line_key != "fte" else "0.0"
        c_m26.alignment = Alignment(horizontal="right")
        c_m26.border = border_use
        cached_injections[f"C{row_idx}"] = m_cur

        # Col D: Sep 25 (M)
        c_m25 = ws.cell(row=row_idx, column=4, value=f_m25)
        c_m25.font = font_use
        c_m25.number_format = "#,##0" if line_key != "fte" else "0.0"
        c_m25.alignment = Alignment(horizontal="right")
        c_m25.border = border_use
        cached_injections[f"D{row_idx}"] = m_cmp

        # Col E: Diff M
        c_mdiff = ws.cell(row=row_idx, column=5, value=f"=C{row_idx}-D{row_idx}")
        c_mdiff.font = font_use
        c_mdiff.number_format = "#,##0" if line_key != "fte" else "0.0"
        c_mdiff.alignment = Alignment(horizontal="right")
        c_mdiff.border = border_use

        # Col F: Endring % M
        c_mpct = ws.cell(row=row_idx, column=6, value=f'=IF(D{row_idx}<>0, (C{row_idx}-D{row_idx})/ABS(D{row_idx}), "-")')
        c_mpct.font = font_use
        c_mpct.number_format = "0.0%"
        c_mpct.alignment = Alignment(horizontal="right")
        c_mpct.border = border_use

        # Col G: YTD Sep 26
        c_y26 = ws.cell(row=row_idx, column=7, value=f_y26)
        c_y26.font = font_use
        c_y26.number_format = "#,##0" if line_key != "fte" else "0.0"
        c_y26.alignment = Alignment(horizontal="right")
        c_y26.border = border_use
        cached_injections[f"G{row_idx}"] = y_cur

        # Col H: YTD Sep 25
        c_y25 = ws.cell(row=row_idx, column=8, value=f_y25)
        c_y25.font = font_use
        c_y25.number_format = "#,##0" if line_key != "fte" else "0.0"
        c_y25.alignment = Alignment(horizontal="right")
        c_y25.border = border_use
        cached_injections[f"H{row_idx}"] = y_cmp

        # Col I: Diff YTD
        c_ydiff = ws.cell(row=row_idx, column=9, value=f"=G{row_idx}-H{row_idx}")
        c_ydiff.font = font_use
        c_ydiff.number_format = "#,##0" if line_key != "fte" else "0.0"
        c_ydiff.alignment = Alignment(horizontal="right")
        c_ydiff.border = border_use

        # Col J: Endring % YTD
        c_ypct = ws.cell(row=row_idx, column=10, value=f'=IF(H{row_idx}<>0, (G{row_idx}-H{row_idx})/ABS(H{row_idx}), "-")')
        c_ypct.font = font_use
        c_ypct.number_format = "0.0%"
        c_ypct.alignment = Alignment(horizontal="right")
        c_ypct.border = border_use

        row_idx += 1

    # Column widths
    ws.column_dimensions["B"].width = 28
    for col_c in ["C", "D", "E", "G", "H", "I"]:
        ws.column_dimensions[col_c].width = 16
    for col_c in ["F", "J"]:
        ws.column_dimensions[col_c].width = 12

    # Save and inject cached values
    local_app = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
    out_dir = local_app / "smartass" / "analyser"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "Wholesale_PL_2026_vs_2025.xlsx"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    # Post process XML
    with zipfile.ZipFile(buf, "r") as zin:
        sheet_xml = zin.read("xl/worksheets/sheet1.xml")
        root = ET.fromstring(sheet_xml)
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        ET.register_namespace("", ns["main"])

        for c in root.findall(".//main:c", ns):
            r_attr = c.attrib.get("r")
            if r_attr in cached_injections:
                v_elem = c.find("main:v", ns)
                if v_elem is None:
                    v_elem = ET.SubElement(c, f"{{{ns['main']}}}v")
                v_elem.text = str(cached_injections[r_attr])
                c.attrib["t"] = "n"

        modified_sheet_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        try:
            with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "xl/worksheets/sheet1.xml":
                        zout.writestr(item, modified_sheet_xml)
                    elif item.filename == "[Content_Types].xml":
                        data = zin.read(item.filename).replace(b"</Types>", CONTENT_TYPES_OVERRIDE.encode("utf-8") + b"</Types>")
                        zout.writestr(item, data)
                    elif item.filename == "_rels/.rels":
                        data = zin.read(item.filename).replace(b"</Relationships>", RELS_CUSTOM_PROP.encode("utf-8") + b"</Relationships>")
                        zout.writestr(item, data)
                    else:
                        zout.writestr(item, zin.read(item.filename))
                zout.writestr("docProps/custom.xml", CUSTOM_PROPS_XML.encode("utf-8"))
        except PermissionError:
            out_path = out_dir / f"Wholesale_PL_2026_vs_2025_{int(time.time())}.xlsx"
            with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "xl/worksheets/sheet1.xml":
                        zout.writestr(item, modified_sheet_xml)
                    elif item.filename == "[Content_Types].xml":
                        data = zin.read(item.filename).replace(b"</Types>", CONTENT_TYPES_OVERRIDE.encode("utf-8") + b"</Types>")
                        zout.writestr(item, data)
                    elif item.filename == "_rels/.rels":
                        data = zin.read(item.filename).replace(b"</Relationships>", RELS_CUSTOM_PROP.encode("utf-8") + b"</Relationships>")
                        zout.writestr(item, data)
                    else:
                        zout.writestr(item, zin.read(item.filename))
                zout.writestr("docProps/custom.xml", CUSTOM_PROPS_XML.encode("utf-8"))

    elapsed = time.perf_counter() - t0
    print(f"Generated Wholesale P&L in {elapsed:.4f}s -> {out_path}")
    return out_path


if __name__ == "__main__":
    generate_wholesale_pnl()
