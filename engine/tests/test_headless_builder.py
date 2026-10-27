import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from smartview_mcp.headless_builder import HeadlessWorkbookBuilder


def test_headless_builder_generates_xml_with_formula_and_cached_value(tmp_path: Path):
    target = tmp_path / "test_out.xlsx"
    builder = HeadlessWorkbookBuilder(title="Test Report")

    # Add one row of data
    builder.add_data_row(
        label="Wholesale EBITDA",
        aug_val=123456.78,
        aug_formula='=HsGetValue("FinMax", "Accounts#ac.EBITDA", "Period#Aug")',
        sep_val=234567.89,
        sep_formula='=HsGetValue("FinMax", "Accounts#ac.EBITDA", "Period#Sep")',
    )
    builder.save(target)

    assert target.exists()

    # Inspect the raw openxml to verify <f> and <v> presence
    with zipfile.ZipFile(target, "r") as z:
        sheet_xml = z.read("xl/worksheets/sheet1.xml")
        root = ET.fromstring(sheet_xml)
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        cells = root.findall(".//main:c", ns)
        formula_cells = [c for c in cells if c.find("main:f", ns) is not None]
        assert len(formula_cells) >= 2

        # Verify first formula cell has both <f> and <v>
        f_elem = formula_cells[0].find("main:f", ns)
        v_elem = formula_cells[0].find("main:v", ns)
        assert f_elem is not None
        assert "HsGetValue" in f_elem.text
        assert v_elem is not None
        assert float(v_elem.text) == 123456.78


def test_headless_builder_sensitivity_label(tmp_path: Path):
    target = tmp_path / "test_label.xlsx"
    builder = HeadlessWorkbookBuilder(title="Labeled Report")
    builder.save(target)

    with zipfile.ZipFile(target, "r") as z:
        assert "docProps/custom.xml" in z.namelist()
        custom_xml = z.read("docProps/custom.xml").decode("utf-8")
        assert "f604d2c9-1577-460e-b668-57374a0216c3" in custom_xml
