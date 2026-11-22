"""xl_* tools: work LIVE in the user's open workbooks (any Excel window), without Smart View.

Agreed with the user 2026-10-15:
- the ACTIVE workbook by default, another open one by name;
- a snapshot before EVERY change (values, formulas, formats via Excel's XML-spreadsheet export) + xl_revert,
  because Ctrl+Z cannot undo COM writes. Snapshots: %LOCALAPPDATA%\\smartass\\snapshots (config.HOME), kept 7 days, never OneDrive;
- the screen is frozen while writing, then Excel jumps to the changed area;
- the user's workbook is NEVER saved here (no Save / SaveAs anywhere in this package);
- every written formula goes through the tn-formelstil checker; 'block' findings must be fixed.
All functions run on a user-Excel job thread (xl/userjobs.py) and take that Excel's turn.
"""
