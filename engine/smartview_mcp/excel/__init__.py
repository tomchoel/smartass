"""Working inside the USER's interactive Excel (any open workbook), on the server's COM thread.

Proven on 2026-10-13 (spikes/spike_hwnd.py):
- attach through the Excel main window (AccessibleObjectFromWindow on XLMAIN>XLDESK>EXCEL7), never GetObject(path) -
  that spawns a hidden '-Embedding' Excel when the workbook is not in the ROT yet;
- load the VBA bridge as a hidden add-in (svbridge.xlam) into that Excel; add-ins resolve by name, they are not enumerated;
- never send COM calls while Smart View's 'Login' dialog is open in that Excel (deadlock) - wait for the SSO watcher;
- Smart View acts on the ACTIVE workbook/sheet, so the target sheet is activated (the user's view jumps to it).
The user's Excel is NEVER killed or saved by this package.
"""
