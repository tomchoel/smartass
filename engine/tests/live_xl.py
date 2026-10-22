"""Live: every xl_* operation against an OPEN test workbook in the user's Excel (default 'formelstil_test'),
including snapshot/revert round trips. Never saves. Usage: python tests/live_xl.py [workbook]"""
import json, os, sys
os.environ["SV_XL_JUMP"] = "0"
import pythoncom
pythoncom.CoInitialize()
from smartview_mcp import livexl as x, userexcel
from smartview_mcp.host import HostError
WB = sys.argv[1] if len(sys.argv) > 1 else "formelstil_test"
fails = []
def ok(cond, what):
    print(("PASS " if cond else "FAIL ") + what); (None if cond else fails.append(what))
def show(o): print("   ", json.dumps(o, ensure_ascii=False, default=str)[:400])
def cleanup():
    for c in x.changes(WB, limit=200):
        if not c["reverted"]:
            try:
                x.revert(c["id"], WB, force=True)
            except HostError as e:
                print("   cleanup:", e)
try:
    cleanup()
    wbs = x.workbooks(); ok(any(w["name"].startswith(WB) for w in wbs), "xl_workbooks lists the test workbook")
    a = x.add_sheet(WB, "XL test"); ok(a["sheet"] == "XL test", "add_sheet")
    w1 = x.write(WB, "XL test", "B2", values=[["Line", "2026", "2025", "Var", "Var %"],
                                              ["Revenue", 100, 80, None, None], ["COGS", -30, -20, None, None],
                                              ["Opex", -20, -25, None, None], ["Total", None, None, None, None],
                                              ["Sep 26", "1/2", "=1+1", None, None]])
    r = x.read(WB, "XL test", "B7:D7"); ok(r["values"][0][0] == "Sep 26" and r["values"][0][1] == "1/2", "text stays text (no date / fraction)")
    ok(r["values"][0][2] == 2, "'=..' in values is a formula")
    w2 = x.write(WB, "XL test", "C6:D6", formula="=SUM(C3:C5)"); r = x.read(WB, "XL test", "C6:D6", "both")
    ok(r["values"][0] == [50, 35] and r["formulas"]["D6"] == "=SUM(D3:D5)", "formula filled across with relative refs")
    ok(w2["formula_style"]["ok"] is True, "clean formulas pass the checker")
    w3 = x.write(WB, "XL test", "E3:E6", formula="=C3-D3"); w4 = x.write(WB, "XL test", "F3:F6", formula='=IF(D3=0,"",(C3-D3)/ABS(D3))')
    bad = x.write(WB, "XL test", "C8", formulas=[["=C3+C4+C5"]])
    ok(bad["formula_style"]["ok"] is False and bad["formula_style"]["block"][0]["rule"] == "PLUS_CHAIN", "a + chain is blocked")
    show(bad["formula_style"]["block"][0])
    rv = x.revert(bad["change_id"], WB); r = x.read(WB, "XL test", "C8", "both"); ok(r["values"][0][0] is None, "revert restores the empty cell")
    t = x.trace(WB, "XL test", "C6"); ok([p["address"] for p in t["precedents"]] == ["C3:C5"] and t["precedents"][0]["sum"] == 50, "trace precedents")
    show(t)
    d = x.trace(WB, "XL test", "C3", "dependents"); from openpyxl.utils.cell import range_boundaries as rb
    cells = set()
    for i in d["dependents"]:
        c1, r1, c2, r2 = rb(i["address"])
        cells |= {(c, r) for c in range(c1, c2 + 1) for r in range(r1, r2 + 1)}
    ok({(3, 6), (5, 3), (6, 3)} <= cells, "trace dependents"); show(d)
    pnl = x.trace(WB, "Q1 Svar", "C13"); ok(pnl["precedents"] and pnl["precedents"][0]["sheet"] == "Q1 FinMax", "trace across sheets (smartview answer)")
    show(pnl)
    f = x.format_range(WB, "XL test", "C3:E6", style="heltall"); r = x.read(WB, "XL test", "C3", "formats"); ok("TN Heltall" in r["styles"], "format with TN style")
    fs = x.format_sheet(WB, "XL test", "F", {"2": "header", "6": "sum"}, label_cols=["B"], num={"F": "prosent1"})
    show(fs); ok(isinstance(fs["style_lint"], list), "format_sheet runs tn-rapportstil")
    rf = x.revert(fs["change_id"], WB); r = x.read(WB, "XL test", "C3", "formats"); ok("TN Heltall" in r["styles"], "revert format_sheet back to previous look")
    ins = x.rows_cols(WB, "XL test", "insert_rows", "4"); r = x.read(WB, "XL test", "C7", "formulas"); ok(r["formulas"].get("C7") == "=SUM(C3:C6)", "insert row expands the SUM")
    x.revert(ins["change_id"], WB); r = x.read(WB, "XL test", "C6", "formulas"); ok(r["formulas"].get("C6") == "=SUM(C3:C5)", "revert insert")
    x.write(WB, "XL test", "H10", formulas=[["=C3"]])          # a single-cell reference: deleting row 3 breaks it
    try:
        x.rows_cols(WB, "XL test", "delete_rows", "3"); ok(False, "delete with dependents refused")
    except HostError as e:
        ok("#REF!" in str(e) or "depend" in str(e), "delete with dependents refused")
    x.revert(x.changes(WB)[0]["id"], WB)                        # take H10 out again
    x.write(WB, "XL test", "H2", values=[["scratch"]]); dl = x.rows_cols(WB, "XL test", "delete_cols", "H")
    x.revert(dl["change_id"], WB); r = x.read(WB, "XL test", "H2"); ok(r["values"][0][0] == "scratch", "revert delete column restores content")
    # conflict: user edits after my change
    w5 = x.write(WB, "XL test", "B10", values=[["mine"]])
    x.write(WB, "XL test", "B10", values=[["later edit"]])
    try:
        x.revert(w5["change_id"], WB); ok(False, "revert refuses when later changes exist")
    except HostError as e:
        ok("Later changes" in str(e), "revert refuses when later changes exist")
    # review fixes
    ins2 = x.rows_cols(WB, "XL test", "insert_rows", "12")
    x.write(WB, "XL test", "B12", values=[["typed by the user"]])
    try:
        x.revert(ins2["change_id"], WB); ok(False, "insert revert refuses when the new row has content")
    except HostError as e:
        ok("content" in str(e) or "Later changes" in str(e), "insert revert refuses when the new row has content")
    for bad_call, what in ((lambda: x.write(WB, "XL test", "B2:B3;D2:D3", values=[[1]]), "multi-area range refused"),
                           (lambda: x.add_sheet(WB, "bad/name"), "invalid sheet name refused"),
                           (lambda: x.rows_cols(WB, "Q1 FinMax x", "delete_rows", "2"), "cross-sheet delete guard (header row)")):
        try:
            bad_call(); ok(False, what)
        except HostError as e:
            ok(True, what); print("    ", str(e)[:140])
    chk = x.check(WB, "Q1 Svar"); ok(chk["ok"] is True, "xl_check on the smartview answer sheet is clean"); show({k: chk[k] for k in ("ok", "checked")})
    ov = x.overview(WB, None); ok(any(s["name"] == "XL test" for s in ov["sheets"]), "overview"); 
    ch = x.changes(WB); ok(len(ch) >= 10, f"changes listed ({len(ch)})")
    # clean up: revert everything newest first, sheet last
    for c in x.changes(WB, limit=100):
        if not c["reverted"]:
            try:
                x.revert(c["id"], WB, force=True)
            except HostError as e:
                print("   cleanup:", e)
    ok("XL test" not in x.workbooks()[[w["name"].startswith(WB) for w in x.workbooks()].index(True)]["sheets"], "full revert removed the test sheet")
finally:
    userexcel.release_turns()
print("FAILS:", fails or "none")
