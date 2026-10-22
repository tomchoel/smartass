"""Read-only: show a formula from 'Q1 Svar' and the header/label cells of 'Q1 FinMax' in the open analysis workbook."""
import sys
sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
import pythoncom
from smartview_mcp import userexcel as u

pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
name_part = sys.argv[1] if len(sys.argv) > 1 else "un.SY.2TEC"
for pid, hwnd, app in u.interactive_excels():
    for wb in app.Workbooks:
        if name_part in wb.Name:
            print("workbook:", wb.Name, "| FullName:", wb.FullName)
            ans = wb.Worksheets("Q1 Svar")
            for r in range(1, 14):
                f = ans.Cells(r, 3).Formula2
                if f:
                    print(f"  Svar C{r}: {str(f)[:400]}  => {ans.Cells(r, 3).Text}")
            d = wb.Worksheets("Q1 FinMax")
            v = d.UsedRange.Value2
            print("  data used:", d.UsedRange.Address)
            for row in v[:6]:
                print("   ", [repr(x)[:40] for x in row[:5]])
