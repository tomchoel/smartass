"""Export a sheet of an open workbook (user's Excel) to PDF and render page 1 to PNG for visual review.
Usage: python render_sheet.py <workbook name part> <sheet name> <out.png>"""
import os, sys
sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
import pythoncom
from smartview_mcp import config, userexcel as u

pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
part, sheet, png = sys.argv[1], sys.argv[2], sys.argv[3]
pdf = os.path.splitext(png)[0] + ".pdf"
for pid, hwnd, app in u.interactive_excels():
    for wb in app.Workbooks:
        if part in wb.Name:
            ws = wb.Worksheets(sheet)
            ws.ExportAsFixedFormat(0, pdf)            # xlTypePDF
            import fitz
            doc = fitz.open(pdf)
            doc[0].get_pixmap(dpi=110).save(png)
            print("rendered", png, "pages:", len(doc))
            sys.exit(0)
print("workbook not found")
