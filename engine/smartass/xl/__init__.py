"""Working inside the USER's interactive Excel (any open workbook), on a job's own STA thread (xl/userjobs.py).

Proven on 2026-10-13 (spikes/spike_hwnd.py):
- attach through the Excel main window (AccessibleObjectFromWindow on XLMAIN>XLDESK>EXCEL7), never GetObject(path) -
  that spawns a hidden '-Embedding' Excel when the workbook is not in the ROT yet;
- never send COM calls while Smart View's 'Login' dialog is open in that Excel (it blocks): such an Excel is refused
  at once (xl/dialogs.py) - smartass uses no Smart View, but it stays installed in the user's Excel.
The user's Excel is NEVER killed or saved by this package; the only Excel it starts is the user's own, on a workbook
(target._launch_excel).
"""
