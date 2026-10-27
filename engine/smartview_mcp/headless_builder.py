"""Headless Excel workbook generator with Smart View formulas and cached values."""

from __future__ import annotations

import io
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET
import zipfile

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

MIP_LABEL_ID = "f604d2c9-1577-460e-b668-57374a0216c3"
MIP_SITE_ID = "1676489c-5c72-46b7-ba63-9ab90c4aad44"

CUSTOM_PROPS_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="2" name="MSIP_Label_{MIP_LABEL_ID}_Enabled">
    <vt:lpwstr>true</vt:lpwstr>
  </property>
  <property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="3" name="MSIP_Label_{MIP_LABEL_ID}_SetDate">
    <vt:lpwstr>2026-10-27T00:00:00Z</vt:lpwstr>
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


class HeadlessWorkbookBuilder:
    """Builds a formatted Excel workbook with formulas and pre-populated values."""

    def __init__(self, title: str):
        self.title = title
        self.rows: list[dict[str, Any]] = []

    def add_data_row(
        self,
        label: str,
        aug_val: float,
        aug_formula: str,
        sep_val: float,
        sep_formula: str,
    ) -> None:
        self.rows.append(
            {
                "label": label,
                "aug_val": aug_val,
                "aug_formula": aug_formula,
                "sep_val": sep_val,
                "sep_formula": sep_formula,
            }
        )

    def build_workbook(self) -> openpyxl.Workbook:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "HsGetValue Svar"

        # Styles
        font_title = Font(name="Segoe UI", size=14, bold=True, color="002060")
        font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        font_body = Font(name="Segoe UI", size=10)
        font_bold = Font(name="Segoe UI", size=10, bold=True)
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )
        double_bottom = Border(
            top=Side(style="thin", color="000000"),
            bottom=Side(style="double", color="000000"),
        )

        # Title
        ws.cell(row=2, column=2, value=self.title).font = font_title
        ws.cell(
            row=3,
            column=2,
            value="FinMax Actuals via headless =HsGetValue formulas (Smart View refreshable)",
        ).font = Font(name="Segoe UI", size=9, italic=True)

        # Table Headers
        headers = ["Linje / Enhet", "August 2026", "September 2026", "YTD Sum (Excel)"]
        header_row = 5
        for col_idx, text in enumerate(headers, start=2):
            cell = ws.cell(row=header_row, column=col_idx, value=text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center" if col_idx > 2 else "left", vertical="center")

        current_row = header_row + 1
        for row_data in self.rows:
            # Col B: Label
            c_lbl = ws.cell(row=current_row, column=2, value=row_data["label"])
            c_lbl.font = font_body
            c_lbl.border = thin_border

            # Col C: Aug (Formula + Value)
            c_aug = ws.cell(row=current_row, column=3, value=row_data["aug_formula"])
            c_aug.font = font_body
            c_aug.number_format = "#,##0"
            c_aug.alignment = Alignment(horizontal="right")
            c_aug.border = thin_border

            # Col D: Sep (Formula + Value)
            c_sep = ws.cell(row=current_row, column=4, value=row_data["sep_formula"])
            c_sep.font = font_body
            c_sep.number_format = "#,##0"
            c_sep.alignment = Alignment(horizontal="right")
            c_sep.border = thin_border

            # Col E: Native Excel SUM formula
            c_sum = ws.cell(
                row=current_row,
                column=5,
                value=f"=SUM(C{current_row}:D{current_row})",
            )
            c_sum.font = font_bold
            c_sum.number_format = "#,##0"
            c_sum.alignment = Alignment(horizontal="right")
            c_sum.border = thin_border

            current_row += 1

        # Summary Row (Total)
        if self.rows:
            ws.cell(row=current_row, column=2, value="Total EBITDA").font = font_bold
            ws.cell(row=current_row, column=2).border = double_bottom

            c_tot_aug = ws.cell(row=current_row, column=3, value=f"=SUM(C6:C{current_row - 1})")
            c_tot_aug.font = font_bold
            c_tot_aug.number_format = "#,##0"
            c_tot_aug.alignment = Alignment(horizontal="right")
            c_tot_aug.border = double_bottom

            c_tot_sep = ws.cell(row=current_row, column=4, value=f"=SUM(D6:D{current_row - 1})")
            c_tot_sep.font = font_bold
            c_tot_sep.number_format = "#,##0"
            c_tot_sep.alignment = Alignment(horizontal="right")
            c_tot_sep.border = double_bottom

            c_tot_ytd = ws.cell(row=current_row, column=5, value=f"=SUM(E6:E{current_row - 1})")
            c_tot_ytd.font = font_bold
            c_tot_ytd.number_format = "#,##0"
            c_tot_ytd.alignment = Alignment(horizontal="right")
            c_tot_ytd.border = double_bottom

        # Column widths
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 20

        return wb

    def save(self, target_path: Path | str) -> None:
        target = Path(target_path)
        wb = self.build_workbook()
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        # Post-process the zip to inject <v> cached values and MIP sensitivity label
        target.parent.mkdir(parents=True, exist_ok=True)
        with (
            zipfile.ZipFile(buf, "r") as zin,
            zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout,
        ):
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = self._inject_cached_values(data)
                elif item.filename == "[Content_Types].xml":
                    data = data.replace(
                        b"</Types>",
                        CONTENT_TYPES_OVERRIDE.encode("utf-8") + b"</Types>",
                    )
                elif item.filename == "_rels/.rels":
                    data = data.replace(
                        b"</Relationships>",
                        RELS_CUSTOM_PROP.encode("utf-8") + b"</Relationships>",
                    )
                zout.writestr(item, data)

            # Write custom.xml
            zout.writestr("docProps/custom.xml", CUSTOM_PROPS_XML.encode("utf-8"))

    def _inject_cached_values(self, sheet_xml_bytes: bytes) -> bytes:
        root = ET.fromstring(sheet_xml_bytes)
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        ET.register_namespace("", ns["main"])

        for row_idx, row_data in enumerate(self.rows, start=6):
            # Target C<row> and D<row>
            aug_ref = f"C{row_idx}"
            sep_ref = f"D{row_idx}"

            for c in root.findall(".//main:c", ns):
                r_attr = c.attrib.get("r")
                if r_attr == aug_ref:
                    v_elem = c.find("main:v", ns)
                    if v_elem is None:
                        v_elem = ET.SubElement(c, f"{{{ns['main']}}}v")
                    v_elem.text = str(row_data["aug_val"])
                    c.attrib["t"] = "n"
                elif r_attr == sep_ref:
                    v_elem = c.find("main:v", ns)
                    if v_elem is None:
                        v_elem = ET.SubElement(c, f"{{{ns['main']}}}v")
                    v_elem.text = str(row_data["sep_val"])
                    c.attrib["t"] = "n"

        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
