"""Close (without saving) and delete analysis workbooks created by tests: names containing argv[1] in the analyser folder."""
import os, sys, time
sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
import pythoncom
from smartview_mcp import config, userexcel as u

pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
part = sys.argv[1]
closed = []
for pid, hwnd, app in u.interactive_excels():
    for wb in list(app.Workbooks):
        lp = u.local_path(wb.FullName)
        if part in wb.Name and os.path.normcase(str(config.ANALYSER_DIR)) in os.path.normcase(lp):
            wb.Close(False)
            closed.append(lp)
time.sleep(1)
for lp in closed:
    try:
        os.remove(lp)
        print("closed + deleted", lp)
    except OSError as e:
        print("closed, could not delete", lp, e)
