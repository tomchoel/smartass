"""Build Mobile Consumer Subscriptions P&L / Direct Contribution headlessly with HsGetValue formulas."""
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

# Data retrieved over HTTP from FinMax
PRODUCTS_DATA = [
    ("pd.SY.M040", "M040 Interconnect", 269368563.00, 277484217.00, 0.0, 0.0, 0.0, 0.0, 269368563.00, 277484217.00),
    ("pd.SY.M112", "M112 Subscription fee voice", 10170031.29, 11376934.49, 0.0, 0.0, 0.0, 0.0, 10170031.29, 11376934.49),
    ("pd.SY.M121", "M121 Voice, outgoing airtime", 79248493.48, 88880773.32, 0.0, 0.0, 0.0, 0.0, 79248493.48, 88880773.32),
    ("pd.SY.M124", "M124 Content services", 45231593.69, 42658116.83, 0.0, 0.0, 0.0, 0.0, 45231593.69, 42658116.83),
    ("pd.SY.M141", "M141 Other mobile rev - own sub", 119864353.86, 103368775.34, 0.0, 0.0, -7.05, -1.06, 119864346.81, 103368774.28),
    ("pd.SY.M210", "M210 Interconnect Mobile costs", 0.0, 0.0, -329858898.50, -334996972.44, 0.0, 0.0, -329858898.50, -334996972.44),
    ("pd.SY.M4115", "M4115 Sub Voice bundled (Hoved)", 5787048544.47, 5747445663.33, 0.0, 0.0, -167685805.77, -157980830.54, 5619362738.70, 5589464832.79),
    ("pd.SY.M4210", "M4210 Data Top-Up", 49481969.53, 51041497.18, 0.0, 0.0, 0.0, 0.0, 49481969.53, 51041497.18),
    ("pd.SY.M4431", "M4431 Messaging SMS IC", 14762076.08, 17573300.09, 0.0, 0.0, 0.0, 0.0, 14762076.08, 17573300.09),
    ("pd.SY.M4514", "M4514 TnShop Other Non Mobile", 8844944.00, 7435802.00, 0.0, 16800.00, -1334394.58, -1568136.50, 7510549.42, 5884465.50),
    ("pd.SY.M5050", "M5050 Roaming-outbound costs", 0.0, 0.0, -107956972.90, -118358309.49, 0.0, 0.0, -107956972.90, -118358309.49),
    ("pd.SY.M5080", "M5080 Other traffic charges", 0.0, 0.0, -8667349.54, -8319442.63, -5039549.37, -4708126.52, -13706898.91, -13027569.15),
    ("pd.SY.M5095", "M5095 Other cost of materials", -118.04, -58.04, -83007711.56, -107324119.44, -363079.85, -789800.89, -83370909.45, -108113978.37),
]


def generate_mobile_consumer_pnl() -> Path:
    t0 = time.perf_counter()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mobil Privat Bidrag"
    ws.views.sheetView[0].showGridLines = True

    # Styles
    f_title = Font(name="Segoe UI", size=14, bold=True, color="002060")
    f_sub = Font(name="Segoe UI", size=9, italic=True, color="595959")
    f_hdr = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    f_body = Font(name="Segoe UI", size=10)
    f_bold = Font(name="Segoe UI", size=10, bold=True)
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_section = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    b_thin = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                    top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
    b_total = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))

    # Title block
    ws.cell(row=2, column=2, value="Mobil Privat (Consumer) - Direkte Bidrag").font = f_title
    ws.cell(row=3, column=2, value="Scope: fam_mobil_own_subscriptions_mobile_consumer | Kilde: FinMax YTD Sep 26 vs Sep 25 (HsGetValue)").font = f_sub

    # Headers
    headers = [
        ("B", "Produkt / Linje"),
        ("C", "Inntekter 26"),
        ("D", "Inntekter 25"),
        ("E", "Varekost 26"),
        ("F", "Varekost 25"),
        ("G", "Opex 26"),
        ("H", "Opex 25"),
        ("I", "Bidrag 26"),
        ("J", "Bidrag 25"),
        ("K", "Endring MNOK"),
        ("L", "Endring %"),
    ]
    for col_letter, text in headers:
        c = ws[f"{col_letter}5"]
        c.value = text
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal="center" if col_letter != "B" else "left", vertical="center")

    cached_injections = {}
    row_idx = 6

    for code, label, r26, r25, c26, c25, o26, o25, e26, e25 in PRODUCTS_DATA:
        # Col B: Label
        ws.cell(row=row_idx, column=2, value=f"{code} {label}").font = f_body
        ws.cell(row=row_idx, column=2).border = b_thin

        # Build HsGetValue formulas
        f_r26 = build_hsgetvalue_formula("FinMax", {"All Accounts": "ac.C30", "All Products": code, "All Time Periods": "Sep 26", "All Entities": "un.SY", "All Balance Types": "Accumulated"})
        f_r25 = build_hsgetvalue_formula("FinMax", {"All Accounts": "ac.C30", "All Products": code, "All Time Periods": "Sep 25", "All Entities": "un.SY", "All Balance Types": "Accumulated"})
        f_c26 = build_hsgetvalue_formula("FinMax", {"All Accounts": "ac.C40S", "All Products": code, "All Time Periods": "Sep 26", "All Entities": "un.SY", "All Balance Types": "Accumulated"})
        f_c25 = build_hsgetvalue_formula("FinMax", {"All Accounts": "ac.C40S", "All Products": code, "All Time Periods": "Sep 25", "All Entities": "un.SY", "All Balance Types": "Accumulated"})
        f_o26 = build_hsgetvalue_formula("FinMax", {"All Accounts": "ac.C70", "All Products": code, "All Time Periods": "Sep 26", "All Entities": "un.SY", "All Balance Types": "Accumulated"})
        f_o25 = build_hsgetvalue_formula("FinMax", {"All Accounts": "ac.C70", "All Products": code, "All Time Periods": "Sep 25", "All Entities": "un.SY", "All Balance Types": "Accumulated"})
        f_e26 = build_hsgetvalue_formula("FinMax", {"All Accounts": "ac.EBITDA_BOI", "All Products": code, "All Time Periods": "Sep 26", "All Entities": "un.SY", "All Balance Types": "Accumulated"})
        f_e25 = build_hsgetvalue_formula("FinMax", {"All Accounts": "ac.EBITDA_BOI", "All Products": code, "All Time Periods": "Sep 25", "All Entities": "un.SY", "All Balance Types": "Accumulated"})

        # Col C: Rev 26
        ws.cell(row=row_idx, column=3, value=f_r26).font = f_body
        ws.cell(row=row_idx, column=3).number_format = "#,##0"
        ws.cell(row=row_idx, column=3).border = b_thin
        cached_injections[f"C{row_idx}"] = r26

        # Col D: Rev 25
        ws.cell(row=row_idx, column=4, value=f_r25).font = f_body
        ws.cell(row=row_idx, column=4).number_format = "#,##0"
        ws.cell(row=row_idx, column=4).border = b_thin
        cached_injections[f"D{row_idx}"] = r25

        # Col E: COGS 26
        ws.cell(row=row_idx, column=5, value=f_c26).font = f_body
        ws.cell(row=row_idx, column=5).number_format = "#,##0"
        ws.cell(row=row_idx, column=5).border = b_thin
        cached_injections[f"E{row_idx}"] = c26

        # Col F: COGS 25
        ws.cell(row=row_idx, column=6, value=f_c25).font = f_body
        ws.cell(row=row_idx, column=6).number_format = "#,##0"
        ws.cell(row=row_idx, column=6).border = b_thin
        cached_injections[f"F{row_idx}"] = c25

        # Col G: Opex 26
        ws.cell(row=row_idx, column=7, value=f_o26).font = f_body
        ws.cell(row=row_idx, column=7).number_format = "#,##0"
        ws.cell(row=row_idx, column=7).border = b_thin
        cached_injections[f"G{row_idx}"] = o26

        # Col H: Opex 25
        ws.cell(row=row_idx, column=8, value=f_o25).font = f_body
        ws.cell(row=row_idx, column=8).number_format = "#,##0"
        ws.cell(row=row_idx, column=8).border = b_thin
        cached_injections[f"H{row_idx}"] = o25

        # Col I: Bidrag 26 (Formula)
        ws.cell(row=row_idx, column=9, value=f"=C{row_idx}+E{row_idx}+G{row_idx}").font = f_bold
        ws.cell(row=row_idx, column=9).number_format = "#,##0"
        ws.cell(row=row_idx, column=9).border = b_thin

        # Col J: Bidrag 25 (Formula)
        ws.cell(row=row_idx, column=10, value=f"=D{row_idx}+F{row_idx}+H{row_idx}").font = f_bold
        ws.cell(row=row_idx, column=10).number_format = "#,##0"
        ws.cell(row=row_idx, column=10).border = b_thin

        # Col K: Endring
        ws.cell(row=row_idx, column=11, value=f"=I{row_idx}-J{row_idx}").font = f_body
        ws.cell(row=row_idx, column=11).number_format = "#,##0"
        ws.cell(row=row_idx, column=11).border = b_thin

        # Col L: Endring %
        ws.cell(row=row_idx, column=12, value=f'=IF(J{row_idx}<>0, (I{row_idx}-J{row_idx})/ABS(J{row_idx}), "-")').font = f_body
        ws.cell(row=row_idx, column=12).number_format = "0.0%"
        ws.cell(row=row_idx, column=12).border = b_thin

        row_idx += 1

    # Total Row
    tot_row = row_idx
    ws.cell(row=tot_row, column=2, value="Total Mobil Privat (Abonnement)").font = f_bold
    ws.cell(row=tot_row, column=2).border = b_total

    for col_c in ["C", "D", "E", "F", "G", "H", "I", "J"]:
        ws.cell(row=tot_row, column=openpyxl.utils.column_index_from_string(col_c), value=f"=SUM({col_c}6:{col_c}{tot_row-1})").font = f_bold
        ws.cell(row=tot_row, column=openpyxl.utils.column_index_from_string(col_c)).number_format = "#,##0"
        ws.cell(row=tot_row, column=openpyxl.utils.column_index_from_string(col_c)).border = b_total

    # Total diff & %
    ws.cell(row=tot_row, column=11, value=f"=I{tot_row}-J{tot_row}").font = f_bold
    ws.cell(row=tot_row, column=11).number_format = "#,##0"
    ws.cell(row=tot_row, column=11).border = b_total

    ws.cell(row=tot_row, column=12, value=f'=IF(J{tot_row}<>0, (I{tot_row}-J{tot_row})/ABS(J{tot_row}), "-")').font = f_bold
    ws.cell(row=tot_row, column=12).number_format = "0.0%"
    ws.cell(row=tot_row, column=12).border = b_total

    # Column widths
    ws.column_dimensions["B"].width = 34
    for c_let in ["C", "D", "E", "F", "G", "H", "I", "J", "K"]:
        ws.column_dimensions[c_let].width = 17
    ws.column_dimensions["L"].width = 13

    # Save to analyser folder
    local_app = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
    out_dir = local_app / "smartass" / "analyser"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "Mobile_Consumer_2026_vs_2025.xlsx"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

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
            out_path = out_dir / f"Mobile_Consumer_2026_vs_2025_{int(time.time())}.xlsx"
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
    print(f"Generated Mobile Consumer P&L in {elapsed:.4f}s -> {out_path}")
    return out_path


if __name__ == "__main__":
    generate_mobile_consumer_pnl()
