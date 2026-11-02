"""Unit tests for sv_drill's pure parts (fixtures from the 15.10.2026 spike)."""
import json
from pathlib import Path

import pytest

from smartview_mcp import drill_lines as dl

FIX = Path(__file__).parent / "fixtures" / "drill"
FIXES = {"H\ufffd\ufffdyhastighet": "Høyhastighet"}


def _items(name):
    return json.loads((FIX / name).read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def cols():
    return dl.columns_from_view(json.loads((FIX / "view_definition.json").read_text(encoding="utf-8")))


@pytest.fixture(scope="module")
def whs(cols):
    return dl.normalize(_items("whs_c40s_q326_items.json"), cols, FIXES)


def test_columns_follow_the_drill_page(cols):
    labels = [c[0] for c in cols]
    assert len(cols) == 45   # plan said 46; the DrillFilter view in the fixture has 45
    assert labels[:4] == ["Scenario", "Period", "GL_POSTED_DATE", "GL_VOUCHER_NUMBER"]
    assert ("AMOUNT", "AMOUNTX") in cols and ("AP_CUSTOMER_NAME", "ATTR13") in cols


def test_values_are_typed(whs):
    rows, _ = whs
    r = rows[0]
    assert isinstance(r["AMOUNT"], float)
    assert r["GL_POSTED_DATE"].year == 2026
    assert isinstance(r["ACCOUNT"], str)            # account stays text ("43200"), never a number


def test_reconcile_exact(whs):
    rows, _ = whs
    rec = dl.reconcile(-24706615.55, rows)
    assert rec["lines"] == 1231 and rec["status"] == "exact" and rec["gap"] == 0.0
    assert rec["debit"] == -369759826.95 and rec["credit"] == 345053211.4


def test_reconcile_reports_gap():
    rec = dl.reconcile(-100.0, [{"AMOUNT": -60.0}])
    assert rec["status"] == "gap" and rec["gap"] == -40.0


def test_replacement_chars_fixed_by_list(whs):
    rows, stats = whs
    assert stats["fixed_by_list"] == 128 and stats["unfixed"] == []
    assert any("Høyhastighet" in (r["PROJECT_DESC"] or "") for r in rows)
    assert not any("\ufffd" in str(v) for r in rows for v in r.values())


def test_unfixed_replacement_chars_are_reported(cols):
    rows, stats = dl.normalize(_items("whs_c40s_q326_items.json"), cols, {})
    assert stats["fixed_by_list"] == 0
    assert stats["unfixed"] == ["SYPR68 Wholesale H\ufffd\ufffdyhastighet", "SYPR68 Wholesale H\ufffd\ufffdyhastighet."]


def test_mojibake_repaired(cols):
    rows, stats = dl.normalize(_items("tec_c40s_sep26_items.json"), cols, FIXES)
    assert stats["fixed_mojibake"] == 2
    texts = [r["GL_DESCRIPTION_LINE"] for r in rows if "Fjordb" in str(r["GL_DESCRIPTION_LINE"])]
    assert texts and all("Fjordbåndet" in t for t in texts)


def test_repair_text_leaves_good_norwegian_alone():
    stats = {"fixed_mojibake": 0, "fixed_by_list": 0, "unfixed": set()}
    assert dl.repair_text("Strøm til produksjon", {}, stats) == "Strøm til produksjon"
    assert stats["fixed_mojibake"] == 0


def test_user_text_fixes_merge_over_the_shipped_list(tmp_path):
    shipped, user = tmp_path / "shipped.json", tmp_path / "user.json"
    shipped.write_text(json.dumps({"fixes": {"H\ufffd\ufffdy": "Høy", "Fjordb\ufffdnd": "Fjordbond"}}), encoding="utf-8")
    user.write_text(json.dumps({"fixes": {"Fjordb\ufffdnd": "Fjordbånd", "Str\ufffdm": "Strøm"}}), encoding="utf-8")
    assert dl.load_text_fixes(shipped, user) == {"H\ufffd\ufffdy": "Høy", "Fjordb\ufffdnd": "Fjordbånd",
                                                 "Str\ufffdm": "Strøm"}
    assert dl.load_text_fixes(shipped, tmp_path / "absent.json") == dl.load_text_fixes(shipped)

def test_profile_names_the_drivers(whs):
    rows, _ = whs
    p = dl.profile(rows, top=5)
    sup = p["supplier"]["top"]
    assert sup[0]["key"] == "" and sup[1]["key"] == "SUPPLIER 037 AS" and sup[1]["amount"] == -60646654.33
    assert p["no_supplier"]["lines"] == 872
    assert p["signs_mixed"] is True
    assert any("allokering" in g["key"] for g in p["journals_without_supplier"]["top"])
    assert {g["key"] for g in p["month"]["top"]} == {"Jul-26", "Aug-26", "Sep-26"}

from smartview_mcp import drill_split as ds

TREE = {("All Time Periods", "Yr 26"): ["Q1 26", "Q2 26", "Q3 26", "Q4 26"],
        ("All Time Periods", "Q3 26"): ["Jul 26", "Aug 26", "Sep 26"],
        ("All Entities", "un.SY.2WHS"): ["un.SY.3WHS_W"],
        ("All Entities", "un.SY.3WHS_W"): ["un.SY.NS74", "un.SY.A406"]}


def kids(dim, member):
    return TREE.get((dim, member), [])


CELL = {"All Accounts": "ac.C40S", "All Entities": "un.SY.2WHS", "All Time Periods": "Q3 26"}


def test_split_by_period_first():
    parts = ds.plan_split(CELL, kids)
    assert [p["All Time Periods"] for p in parts] == ["Jul 26", "Aug 26", "Sep 26"]
    assert all(p["All Entities"] == "un.SY.2WHS" for p in parts)


def test_split_periods_in_calendar_order():
    # the FinLite outline (atlas ord) lists months newest first: pieces and the grid must still run Jul -> Sep
    rev = lambda dim, member: list(reversed(kids(dim, member)))
    assert [p["All Time Periods"] for p in ds.plan_split(CELL, rev)] == ["Jul 26", "Aug 26", "Sep 26"]
    assert [p["All Time Periods"] for p in ds.plan_split({**CELL, "All Time Periods": "Yr 26"}, rev)] ==         ["Q1 26", "Q2 26", "Q3 26", "Q4 26"]
    assert ds.chronological(["Dec 25", "Jan 26"]) == ["Dec 25", "Jan 26"]
    assert ds.chronological(["b", "Jul 26"]) == ["b", "Jul 26"]            # not all periods: order kept


def test_split_by_entity_when_period_is_a_leaf():
    parts = ds.plan_split({**CELL, "All Time Periods": "Sep 26"}, kids)
    assert [p["All Entities"] for p in parts] == ["un.SY.3WHS_W"]        # single child: still progress down


def test_no_split_possible():
    leaf = {"All Accounts": "ac.40100", "All Entities": "un.SY.NS74", "All Time Periods": "Sep 26"}
    assert ds.plan_split(leaf, kids) is None


import datetime as _dt

from smartview_mcp import drill_store as st


def test_store_roundtrip(tmp_path, whs):
    rows, _ = whs
    db = st.connect(tmp_path / "d.db")
    did = st.save(db, connection="FinLite", cell=CELL, rec={"lines": len(rows), "gap": 0.0},
                  columns=list(rows[0]), rows=rows, workbook="C:/x.xlsx", q=3, alias="d")
    meta, back = st.load(db, did)
    assert meta["q"] == 3 and meta["cell"] == CELL and len(back) == 1231
    assert back[0]["GL_POSTED_DATE"] == rows[0]["GL_POSTED_DATE"]           # datetimes survive
    assert st.find(db, "C:/x.xlsx", 3) == {"d": did}


def test_store_prunes_old(tmp_path):
    db = st.connect(tmp_path / "d.db")
    old = st.save(db, connection="FinLite", cell=CELL, rec={}, columns=["AMOUNT"], rows=[{"AMOUNT": 1.0}],
                  workbook="w", q=1, alias="d", created=_dt.datetime.now() - _dt.timedelta(days=91))
    assert st.prune(db, keep_days=90) == 1
    with pytest.raises(KeyError):
        st.load(db, old)


def test_reconcile_counts_non_numeric_amounts(whs):
    rows, _ = whs
    assert dl.reconcile(-24706615.55, rows)["non_numeric"] == 0
    rec = dl.reconcile(None, [{"AMOUNT": -60.0}, {"AMOUNT": "n/a"}, {"AMOUNT": None}])
    assert rec["non_numeric"] == 2 and rec["lines"] == 3 and rec["lines_sum"] == -60.0


def test_profile_groups_carry_debit_and_credit(whs):
    rows, _ = whs
    pay = next(g for g in dl.profile(rows)["je_source"]["top"] if g["key"] == "Payables")
    assert pay["amount"] == -123132549.19 and pay["lines"] == 359
    assert pay["debit"] == -123802585.86 and pay["credit"] == 670036.67


from smartview_mcp import drill_xl as dx


def test_table_name():
    assert dx.table_name(1) == "Lines_Q1" and dx.table_name(4, "b") == "Lines_Q4_b"


def test_formula_sumifs_for_safe_keys():
    f = dx.drill_formula("Lines_Q1", {"JE_SOURCE": "Payables"})
    assert f == '=SUMIFS(Lines_Q1[AMOUNT],Lines_Q1[JE_SOURCE],"Payables")'
    assert dx.drill_formula("Lines_Q1", {"JE_SOURCE": "Payables"}, "count") == \
        '=COUNTIFS(Lines_Q1[JE_SOURCE],"Payables")'
    assert dx.drill_formula("T", {"JE_SOURCE": "Payables"}, "credit") == \
        '=SUMIFS(T[AMOUNT],T[AMOUNT],">0",T[JE_SOURCE],"Payables")'


def test_formula_exact_match_for_date_like_and_long_digits():
    assert dx.drill_formula("T", {"Period": "Jul-26"}) == '=SUM(FILTER(T[AMOUNT],T[Period]="Jul-26",0))'
    assert dx.drill_formula("T", {"AP_INVOICE_NUMBER": "1234567890123456"}, "count") == \
        '=SUM(--(T[AP_INVOICE_NUMBER]="1234567890123456"))'
    two = dx.drill_formula("T", {"Period": "Jul-26", "JE_SOURCE": "Payables"})
    assert two == '=SUM(FILTER(T[AMOUNT],(T[Period]="Jul-26")*(T[JE_SOURCE]="Payables"),0))'
    assert dx.drill_formula("T", {"Period": "Jul-26"}, "debit") == \
        '=SUM(FILTER(T[AMOUNT],(T[Period]="Jul-26")*(T[AMOUNT]<0),0))'


def test_formula_exact_match_where_sumifs_would_read_the_key_as_a_pattern():
    # SUMIFS treats * ? ~ as wildcards and a leading < > = as a comparison; leading zeros are lost to a number
    for v in ["AB*", "what?", "a~b", "<none>", "=x", "0123", "x" * 256, "12:30", "7:05:00", "50%", "TRUE",
              "false", "1,5", "-12,75"]:
        assert not dx.safe_for_sumifs(v), v
        assert dx.drill_formula("T", {"X": v}).startswith("=SUM(FILTER(")
    for v in ["Payables", "43200", "SUPPLIER 037 AS", "", "Tele2 Norge AS"]:
        assert dx.safe_for_sumifs(v), v


def test_formula_totals_debit_credit_and_blank():
    assert dx.drill_formula("T", {}) == "=SUM(T[AMOUNT])"
    assert dx.drill_formula("T", {}, "count") == "=ROWS(T[AMOUNT])"         # all lines, like COUNTIFS
    assert dx.drill_formula("T", {}, "debit") == '=SUMIFS(T[AMOUNT],T[AMOUNT],"<0")'
    assert dx.drill_formula("T", {"AP_CUSTOMER_NAME": ""}) == '=SUMIFS(T[AMOUNT],T[AP_CUSTOMER_NAME],"")'
    assert dx.drill_formula("T", {"X": 'a"b'}) == '=SUMIFS(T[AMOUNT],T[X],"a""b")'


def test_formulas_pass_the_style_checker():
    from smartview_mcp import formulas
    t = "Lines_Q1"
    fs = [dx.drill_formula(t, {"JE_SOURCE": "Payables"}),
          dx.drill_formula(t, {"JE_SOURCE": "Payables", "AP_CUSTOMER_NAME": "SUPPLIER 037 AS"}, "count"),
          dx.drill_formula(t, {"Period": "Jul-26"}),
          dx.drill_formula(t, {"Period": "Jul-26", "JE_SOURCE": "Payables"}, "debit"),
          dx.drill_formula(t, {"AP_INVOICE_NUMBER": "1234567890123456"}, "count"),
          dx.drill_formula(t, {"Period": "Jul-26"}, "count"),
          dx.drill_formula(t, {}, "credit"),
          dx.drill_formula(t, {}),
          dx.drill_formula(t, {}, "count")]
    res = formulas.check([("Q1 Answer", f"C{i}", f) for i, f in enumerate(fs, start=10)])
    assert res["ok"] is True and res["block"] == [], res


from smartview_mcp import specs

DRILL_SPEC = {"columns": [{"header": "NOK"}, {"header": "Lines"}],
              "sections": [{"title": "By source", "lines": [
                  {"key": "pay", "label": "Payables", "cells": [
                      {"drill": "d", "where": {"JE_SOURCE": "Payables"}},
                      {"drill": "d", "where": {"JE_SOURCE": "Payables"}, "agg": "count"}]},
                  {"key": "jul", "label": "July", "cells": [
                      {"drill": "d", "where": {"Period": "Jul-26"}, "agg": "debit"}, None]},
                  {"key": "tot", "label": "Total", "kind": "sum", "cells": [
                      {"drill": "d", "where": {}}, {"drill": "d", "where": {}, "agg": "count"}]}]}]}


def test_evaluate_drill_cells(whs):
    rows, _ = whs
    vals = specs.evaluate(DRILL_SPEC, {}, drills={"d": rows})
    assert round(vals["pay"][0], 2) == -123132549.19 and vals["pay"][1] == 359
    assert round(vals["tot"][0], 2) == -24706615.55 and vals["tot"][1] == 1231
    jul = sum(r["AMOUNT"] for r in rows if r["Period"] == "Jul-26" and r["AMOUNT"] < 0)
    assert vals["jul"][0] == pytest.approx(jul) and jul < 0


def test_evaluate_drill_cells_without_rows_are_empty():
    vals = specs.evaluate(DRILL_SPEC, {})
    assert vals["pay"] == [None, None] and vals["tot"] == [None, None]
    assert specs.prepare(DRILL_SPEC, {}, {})["sections"]        # drills stays optional


def test_prepare_passes_drill_rows(whs, monkeypatch):
    rows, _ = whs
    seen = []
    real = specs.evaluate

    def spy(s, p, drills=None):
        seen.append(drills)
        return real(s, p, drills)
    monkeypatch.setattr(specs, "evaluate", spy)
    specs.prepare(DRILL_SPEC, {}, {}, drills={"d": rows})
    assert seen and seen[0]["d"] is rows


class _Cell:
    def __init__(self):
        self.Value = self.Formula2 = None


class _Sheet:
    def __init__(self):
        self.cells = {}

    def Cells(self, r, c):
        cell = self.cells.setdefault((r, c), _Cell())
        cell.rc = (r, c)
        return cell

    def Range(self, a, b):
        return _Rect(self, a.rc, b.rc)


class _Rect:
    """A rectangle written in one go (answer_write._Cells): spread over the cells, formulas as Formula2, the rest Value."""
    def __init__(self, ws, a, b):
        self.__dict__.update(ws=ws, a=a)

    def __setattr__(self, k, grid):
        for i, row in enumerate(grid):
            for j, v in enumerate(row):
                if v != "":
                    cell = self.ws.Cells(self.a[0] + i, self.a[1] + j)
                    setattr(cell, "Formula2" if isinstance(v, str) and v.startswith("=") else "Value", v)


class _Target:
    def status(self, _msg):
        pass


def test_answer_writes_drill_formulas(monkeypatch):
    from smartview_mcp import answer_write
    monkeypatch.setattr(answer_write, "com", lambda call, *a, **k: call())
    ws = _Sheet()
    out = answer_write._write(_Target(), 2, DRILL_SPEC, {}, ws)
    rows_by_key, warnings = out[2], out[-1]
    assert warnings == []
    c = answer_write.FIRST_COL
    assert ws.cells[(rows_by_key["pay"], c)].Formula2 == '=SUMIFS(Lines_Q2[AMOUNT],Lines_Q2[JE_SOURCE],"Payables")'
    assert ws.cells[(rows_by_key["pay"], c + 1)].Formula2 == '=COUNTIFS(Lines_Q2[JE_SOURCE],"Payables")'
    assert ws.cells[(rows_by_key["jul"], c)].Formula2 == \
        '=SUM(FILTER(Lines_Q2[AMOUNT],(Lines_Q2[Period]="Jul-26")*(Lines_Q2[AMOUNT]<0),0))'
    assert ws.cells[(rows_by_key["tot"], c)].Formula2 == "=SUM(Lines_Q2[AMOUNT])"
    other = {**DRILL_SPEC["sections"][0]["lines"][2], "key": "b", "cells": [{"drill": "b", "where": {}}]}
    ws2 = _Sheet()
    out2 = answer_write._write(_Target(), 2, {**DRILL_SPEC, "sections": [{"title": "x", "lines": [other]}]}, {}, ws2)
    assert ws2.cells[(out2[2]["b"], c)].Formula2 == "=SUM(Lines_Q2_b[AMOUNT])"


def test_drill_value_matches_like_excel_equals():
    rows = [{"JE_SOURCE": "Payables", "ACCOUNT": "43200", "AMOUNT": -10.0},
            {"JE_SOURCE": "PAYABLES", "ACCOUNT": "43200", "AMOUNT": -5.0},
            {"JE_SOURCE": None, "ACCOUNT": "43201", "AMOUNT": 2.0}]
    assert specs._drill_value(rows, {"JE_SOURCE": "payables"}, "sum") == -15.0      # case-insensitive
    assert specs._drill_value(rows, {"ACCOUNT": 43200}, "count") == 2               # non-string where value
    assert specs._drill_value(rows, {"JE_SOURCE": ""}, "sum") == 2.0                # None = blank


def test_unknown_agg_is_rejected():
    with pytest.raises(ValueError):
        dx.drill_formula("T", {"JE_SOURCE": "Payables"}, "avg")
    with pytest.raises(ValueError):
        specs._drill_value([{"AMOUNT": 1.0}], {}, "avg")
    spec = {"columns": [{"header": "NOK"}], "sections": [{"title": "x", "lines": [
        {"key": "a", "label": "A", "cells": [{"drill": "d", "where": {}, "agg": "avg"}]}]}]}
    assert specs.evaluate(spec, {}, drills={"d": [{"AMOUNT": 1.0}]})["a"] == [None]


def test_answer_reports_unknown_agg_as_warning(monkeypatch):
    from smartview_mcp import answer_write
    monkeypatch.setattr(answer_write, "com", lambda call, *a, **k: call())
    spec = {"columns": [{"header": "NOK"}], "sections": [{"title": "x", "lines": [
        {"key": "a", "label": "A", "cells": [{"drill": "d", "where": {}, "agg": "avg"}]}]}]}
    warnings = answer_write._write(_Target(), 1, spec, {}, _Sheet())[-1]
    assert len(warnings) == 1 and "avg" in warnings[0]


def test_answer_drill_wins_over_pull_like_specs(monkeypatch):
    from smartview_mcp import answer_write
    monkeypatch.setattr(answer_write, "com", lambda call, *a, **k: call())
    spec = {"columns": [{"header": "NOK"}], "sections": [{"title": "x", "lines": [
        {"key": "a", "label": "A", "cells": [{"drill": "d", "where": {}, "pull": "nope"}]}]}]}
    ws = _Sheet()
    out = answer_write._write(_Target(), 1, spec, {}, ws)
    assert out[-1] == [] and ws.cells[(out[2]["a"], answer_write.FIRST_COL)].Formula2 == "=SUM(Lines_Q1[AMOUNT])"


# ------------------------------------------------------------------ drill_browser (no browser started)
def test_kill_our_chrome_matches_exact_profile_only(monkeypatch):
    from smartview_mcp import drill_browser as db

    class P:
        def __init__(self, pid, name, cmd):
            self.pid, self.info, self.killed = pid, {"name": name, "cmdline": cmd}, False

        def kill(self):
            self.killed = True
    prof = Path(r"C:\Users\x\AppData\Local\svkit\drill_chrome")
    procs = [P(1, "chrome.exe", ["chrome.exe", r"--user-data-dir=C:\Users\x\AppData\Local\svkit\drill_chrome"]),
             P(2, "chrome.exe", ["chrome.exe", r"--user-data-dir=C:\Users\x\AppData\Local\Google\Chrome\User Data"]),
             P(3, "msedge.exe", ["msedge.exe", r"--user-data-dir=C:\Users\x\AppData\Local\svkit\drill_chrome"]),
             P(4, "chrome.exe", None),
             P(5, "chrome.exe", ["chrome.exe", r"--user-data-dir=C:\Users\x\AppData\Local\svkit\drill_chrome_old"]),
             P(6, "chrome.exe", ["chrome.exe", '--user-data-dir="c:\\users\\x\\appdata\\local\\svkit\\drill_chrome"',
                                 "--type=renderer"]),
             P(7, "chrome.exe", ["chrome.exe", r"--load-extension=C:\Users\x\AppData\Local\svkit\drill_chrome\ext"]),
             P(8, "msedge.exe", ["msedge.exe", r"--user-data-dir=C:\Users\x\AppData\Local\Microsoft\Edge\User Data"])]
    monkeypatch.setattr(db.psutil, "process_iter", lambda attrs=None: iter(procs))
    assert db.kill_our_chrome(prof) == [1, 3, 6]          # the drill browser is Edge since 22.10.2026
    assert [p.pid for p in procs if p.killed] == [1, 3, 6]


def test_redact_strips_queries_and_tokens():
    from smartview_mcp.drill_browser import redact
    s = "Could not open https://x.oraclecloud.com/aif/FdmeeMain.jsp?module=a&sso_token=SECRET&attribute=b: timeout"
    out = redact(s)
    assert "SECRET" not in out and "module=" not in out and out.startswith("Could not open https://x.oraclecloud")
    assert "SECRET" not in redact("token sso_token=SECRET&x=1 in a console line")


def _pages(total, sizes, has_more_always=False, include_total=True):
    """Fake getData: answers per offset from a list of page sizes."""
    pages, off = {}, 0
    for i, n in enumerate(sizes):
        pay = {"items": [{"AMOUNTX": 1.0}] * n, "hasMore": has_more_always or i < len(sizes) - 1}
        if include_total:
            pay["totalResults"] = total
        pages[off] = json.dumps({"responsePayload": pay})
        off += n
    calls = []
    extra = {"items": [{"AMOUNTX": 1.0}], "hasMore": True}
    if include_total:
        extra["totalResults"] = total

    def post(o):
        calls.append(o)
        return 200, pages.get(o, json.dumps({"responsePayload": extra}))
    return post, calls


def test_page_through_pages_until_total():
    from smartview_mcp.drill_browser import page_through
    post, calls = _pages(1231, [500, 500, 231])
    seen = []
    items, total = page_through(post, 300000, lambda n, t: seen.append((n, t)))
    assert (len(items), total, calls) == (1231, 1231, [0, 500, 1000])
    assert seen[-1] == (1231, 1231)


def test_page_through_follows_has_more_past_a_short_total(caplog):
    from smartview_mcp.drill_browser import page_through
    post, calls = _pages(81, [81, 1])             # spec §9.2: totalResults may be one short
    items, total = page_through(post, 300000)
    assert len(items) == 82 and total == 81 and calls == [0, 81]
    assert "totalResults was 81" in caplog.text


def test_page_through_safety_cap_when_has_more_never_ends(monkeypatch):
    from smartview_mcp import config
    from smartview_mcp.drill_browser import page_through
    monkeypatch.setattr(config, "DRILL_PAGE", 10)
    post, calls = _pages(81, [81], has_more_always=True)   # then 1-line pages with hasMore forever
    items, total = page_through(post, 300000)
    assert len(items) == 81 + 10 and len(calls) == 11


def test_page_through_limit_holds_without_total():
    from smartview_mcp.drill_browser import TooManyLines, page_through
    post, _ = _pages(None, [500] * 10, has_more_always=True, include_total=False)
    with pytest.raises(TooManyLines) as e:
        page_through(post, 1200)
    assert e.value.total >= 1200


def test_page_through_limit_from_total():
    from smartview_mcp.drill_browser import TooManyLines, page_through
    post, calls = _pages(5000, [500], has_more_always=True)
    with pytest.raises(TooManyLines) as e:
        page_through(post, 1000)
    assert e.value.total == 5000 and calls == [0]


def test_page_through_http_error_with_json_fails():
    from smartview_mcp.drill_browser import DrillError, page_through
    body = json.dumps({"responsePayload": {"items": [], "totalResults": 0}})
    with pytest.raises(DrillError, match="HTTP 500"):
        page_through(lambda o: (500, body), 1000, recover=lambda: None)


def test_page_through_relogin_once_and_resume_at_offset():
    from smartview_mcp.drill_browser import page_through
    post, calls = _pages(1231, [500, 500, 231])
    state = {"expired": True, "recovered": 0}

    def flaky(o):
        if o == 500 and state["expired"]:
            calls.append(o)
            return 0, "TypeError: Failed to fetch"
        return post(o)

    def recover():
        state["expired"] = False
        state["recovered"] += 1
    items, total = page_through(flaky, 300000, recover=recover)
    assert len(items) == 1231 and state["recovered"] == 1 and calls == [0, 500, 500, 1000]


def test_page_through_second_expiry_fails():
    from smartview_mcp.drill_browser import DrillError, page_through
    n = {"r": 0}
    with pytest.raises(DrillError, match="HTTP 401"):
        page_through(lambda o: (401, "<html>sign in</html>"), 1000, recover=lambda: n.__setitem__("r", n["r"] + 1))
    assert n["r"] == 1


class _FakeClock:
    def __init__(self):
        self.t = 0.0

    def __call__(self):
        return self.t


class _FakePage:
    """Drill page that redirects to sign-in first; after the re-login the getData request comes 40 s later."""
    def __init__(self, clock, base):
        self.clock, self.base, self.url, self.handlers, self.gotos = clock, base, "", {}, []
        self.armed_at = None

    def on(self, ev, fn):
        self.handlers[ev] = fn

    def remove_listener(self, ev, fn):
        self.handlers.pop(ev, None)

    def goto(self, url, **kw):
        self.gotos.append(url)
        self.url = self.base + "/ui/v1/signin?x" if len(self.gotos) == 1 else url
        if len(self.gotos) > 1:
            self.armed_at = self.clock.t + 40

    def wait_for_timeout(self, ms):
        self.clock.t += ms / 1000
        if self.armed_at is not None and self.clock.t >= self.armed_at and "request" in self.handlers:
            base = self.base

            class R:
                url = base + "/aif/ui/model/drilldown/getData?limit=500&offset=0"
                post_data, headers = "{}", {}

            class V:
                url = base + "/aif/ui/model/view?page=DRILL_THROUGH&appId=52"

                @staticmethod
                def json():
                    return {"views": []}
            self.handlers["request"](R)
            self.handlers["response"](V)


def test_capture_resets_timer_after_relogin(monkeypatch):
    from smartview_mcp import drill_browser as db
    clock = _FakeClock()
    monkeypatch.setattr(db.time, "monotonic", clock)
    b = db.DrillBrowser.__new__(db.DrillBrowser)
    b.base = "https://epm"
    s = db._Session()
    s.page = _FakePage(clock, b.base)
    logins = []

    def fake_login(sess):
        clock.t += 30                     # a slow SSO round trip eats most of the old 45 s window
        logins.append(1)
        sess.ready = True
    b._login = fake_login
    got = b._capture(s, "./aif/x.jsp?sso_token=T")
    assert "req" in got and "view" in got and logins == [1] and len(s.page.gotos) == 2


def test_capture_timeout_reports_console_errors_without_token(monkeypatch):
    from smartview_mcp import drill_browser as db
    clock = _FakeClock()
    monkeypatch.setattr(db.time, "monotonic", clock)

    class Blank(_FakePage):
        def goto(self, url, **kw):
            self.gotos.append(url)
            self.url = url

        def reload(self, **kw):
            self.reloads = getattr(self, "reloads", 0) + 1

        def wait_for_timeout(self, ms):
            self.clock.t += ms / 1000
            if "console" in self.handlers and self.clock.t < 1:
                class M:
                    type, text = "error", "CDN_HOST null at https://epm/x.js?sso_token=T"
                self.handlers["console"](M)
    b = db.DrillBrowser.__new__(db.DrillBrowser)
    b.base = "https://epm"
    s = db._Session()
    s.page = Blank(clock, b.base)
    with pytest.raises(db.DrillError) as e:
        b._capture(s, "aif/x.jsp?sso_token=T")
    assert "CDN_HOST null" in str(e.value) and "sso_token=T" not in str(e.value)
    assert s.page.reloads == 1 and "also after one reload" in str(e.value)          # blank page: reloaded once


def test_login_retries_once_in_fresh_tab():
    from smartview_mcp import drill_browser as db

    class Tab:
        closed = False

        def close(self):
            self.closed = True

    class Ctx:
        def new_page(self):
            return Tab()
    b = db.DrillBrowser.__new__(db.DrillBrowser)
    tries = []

    def once(p, deadline):
        tries.append(p)
        if len(tries) == 1:
            raise db.DrillError("last page: signin")
    b._login_once = once
    s = db._Session()
    s.ctx, s.page = Ctx(), Tab()
    first = s.page
    b._login(s)
    assert s.ready and len(tries) == 2 and tries[1] is s.page and s.page is not first and first.closed

    def never(p, deadline):
        raise db.DrillError("last page: signin")
    b._login_once = never
    with pytest.raises(db.DrillError, match="Sign-in to EPM did not complete"):
        b._login(s)
    assert not s.ready


def test_close_always_stops_driver(monkeypatch):
    from smartview_mcp import drill_browser as db
    stopped = []

    class Ctx:
        def close(self):
            raise RuntimeError("browser already gone")

    class Pw:
        def stop(self):
            stopped.append(1)
    monkeypatch.setattr(db, "_kill_driver", lambda d: stopped.append(("kill", tuple(d))))
    s = db._Session()
    s.ctx, s.pw, s.page, s.ready, s.driver = Ctx(), Pw(), object(), True, [(9, 1.0)]
    db.DrillBrowser._close(s)
    assert stopped == [1, ("kill", ((9, 1.0),))] and s.ctx is None and s.pw is None and not s.ready


def test_kill_replaces_hung_worker_and_new_worker_runs(monkeypatch):
    import threading
    from smartview_mcp import drill_browser as db
    monkeypatch.setattr(db, "kill_our_chrome", lambda prof: [])
    monkeypatch.setattr(db, "_kill_driver", lambda d: [])
    monkeypatch.setattr(db, "KILL_WAIT_S", 0.3)
    b = db.DrillBrowser(base="https://epm", profile=Path("x"))
    gate = threading.Event()
    hung = b.submit(lambda s: gate.wait(30))
    queued = b.submit(lambda s: "never")
    old_thread = b.thread
    b.kill()
    assert b.thread is not old_thread and queued.cancelled()
    assert b.submit(lambda s: "fresh").result(timeout=5) == "fresh"
    gate.set()
    hung.result(timeout=5)
    old_thread.join(timeout=5)
    assert not old_thread.is_alive()
    b.shutdown()


def test_cell_keeps_keys_as_text():
    assert dx._cell("ACCOUNT", "43200") == "'43200"
    assert dx._cell("Period", "Jul-26") == "'Jul-26"
    assert dx._cell("GL_DESCRIPTION_LINE", "=HYPERLINK(1)") == "'=HYPERLINK(1)"
    assert dx._cell("AMOUNT", -5.0) == -5.0 and dx._cell("X", None) is None
    # any text Excel would otherwise coerce (logical, time, percent, Norwegian decimal) stays text too
    for v in ("TRUE", "12:30", "50%", "1,5", "Mar-26"):
        assert dx._cell("JOURNAL_NAME", v) == "'" + v
    assert dx._cell("ACCOUNT", 43200) == "'43200"


def test_cell_writes_dates_as_excel_serials():
    # pywin32 shifts naive datetimes by the UTC offset (10:25 became 08:25); a serial number is exact
    import datetime as dt_
    assert dx._cell("GL_POSTED_DATE", dt_.datetime(2026, 7, 6, 10, 25, 27)) == pytest.approx(46209 + (10 * 3600 + 25 * 60 + 27) / 86400)
    assert dx._cell("BOOKING_DATE", dt_.datetime(1900, 3, 1)) == 61.0


def test_gl_posted_date_is_oslo_wall_clock():
    import datetime as dt_
    assert dl._value("ATTR17", "2026-07-06T10:25:27.000+00:00") == dt_.datetime(2026, 7, 6, 12, 25, 27)   # CEST
    assert dl._value("ATTR17", "2026-01-15T23:30:00.000+00:00") == dt_.datetime(2026, 1, 16, 0, 30)       # CET
    assert dl._value("ATTR4", "2026-07-06T00:00:00.000+00:00") == dt_.datetime(2026, 7, 6)              # unchanged


def test_oslo_without_tzdata_ignores_pc_time_zone(monkeypatch):
    # Windows has no tzdata: the fallback must apply EU summer time, not this PC's zone (CI runs in UTC)
    import builtins
    import datetime as dt_
    real_import = builtins.__import__

    def no_zoneinfo(name, *a, **k):
        if name == "zoneinfo":
            raise ImportError
        return real_import(name, *a, **k)
    monkeypatch.setattr(builtins, "__import__", no_zoneinfo)
    U = dt_.datetime
    assert dl._oslo(U(2026, 7, 6, 10, 25, 27)) == U(2026, 7, 6, 12, 25, 27)       # CEST
    assert dl._oslo(U(2026, 1, 15, 23, 30)) == U(2026, 1, 16, 0, 30)              # CET
    assert dl._oslo(U(2026, 3, 29, 0, 59)) == U(2026, 3, 29, 1, 59)               # last minute of CET
    assert dl._oslo(U(2026, 3, 29, 1, 0)) == U(2026, 3, 29, 3, 0)                 # summer time starts 01:00 UTC
    assert dl._oslo(U(2026, 10, 25, 0, 59)) == U(2026, 10, 25, 2, 59)             # last minute of CEST
    assert dl._oslo(U(2026, 10, 25, 1, 0)) == U(2026, 10, 25, 2, 0)               # back to CET


# ------------------------------------------------------------------ sheet placement and re-drill (fake COM)
class _Sh:
    def __init__(self, name):
        self.Name = name
        self.Tab = type("Tab", (), {"Color": None})()


class _Sheets(list):
    def __call__(self, name):                          # Worksheets("name")
        return next(s for s in self if s.Name == name)

    def Add(self, before=None, after=None):          # positional like Excel's Sheets.Add(Before, After)
        new = _Sh("new")
        self.insert(self.index(after) + 1 if after is not None else self.index(before), new)
        return new


def _target(names):
    from types import SimpleNamespace
    return SimpleNamespace(wb=SimpleNamespace(Worksheets=_Sheets(_Sh(n) for n in names)), sheets_added=[])


@pytest.mark.parametrize("names, after_q, expect", [
    (["Sources"], 1, ["X", "Sources"]),                                              # new workbook: before Sources
    (["Sources", "_svmeta"], None, ["X", "Sources", "_svmeta"]),
    (["Q1 FinLite", "Q1 Lines", "Sources", "_svmeta"], 1, ["Q1 FinLite", "Q1 Lines", "X", "Sources", "_svmeta"]),
    (["Q1 Answer", "Q1 FinMax", "Q2 FinLite", "Sources"], 1, ["Q1 Answer", "Q1 FinMax", "X", "Q2 FinLite", "Sources"]),
    (["Q1 Answer", "Q1 FinMax", "Sources"], 2, ["Q1 Answer", "Q1 FinMax", "X", "Sources"]),
])
def test_add_sheet_anchor(names, after_q, expect):
    from smartview_mcp.excel import sheets
    t = _target(names)
    sheets.add_sheet(t, "X", after_q=after_q)
    assert [s.Name for s in t.wb.Worksheets] == expect and t.sheets_added == ["X"]


def test_question_grouping_includes_lines(monkeypatch):
    from smartview_mcp.excel import sheets
    assert sheets._Q.match("Q3 Lines b") and not sheets._Q.match("Q3 Forecast")
    t = _target(["Q2 Answer", "Q2 FinLite", "Q2 Lines", "Sources"])
    assert sheets.next_question(t) == 3


def test_redrill_clear_ranges():
    # old table A4:AS1235 (1231 rows, 45 cols) -> new 1000 rows, 20 cols: clear the rows below and the cols right
    assert dx._stale_ranges(4, 1235, 45, 1004, 20) == ["A1005:AS1235", "U4:AS1004"]
    assert dx._stale_ranges(4, 1235, 45, 1300, 45) == []                       # grows: nothing stale


def test_where_columns_checked_against_table():
    assert dx.missing_columns(["Period", "AMOUNT"], {"Period": "Jul-26", "GL_DESCRIPTION_LINE": "x"}) == \
        ["GL_DESCRIPTION_LINE"]
    assert dx.missing_columns(None, {"X": "1"}) == []                          # table unknown: no check


def test_answer_warns_on_where_column_missing_from_table(monkeypatch):
    from smartview_mcp import answer_write
    monkeypatch.setattr(answer_write, "com", lambda call, *a, **k: call())
    monkeypatch.setattr(answer_write, "drill_table_columns", lambda wb, tn: ["Period", "AMOUNT", "JE_SOURCE"])
    spec = {"columns": [{"header": "NOK"}], "sections": [{"title": "x", "lines": [
        {"key": "a", "label": "A", "cells": [{"drill": "d", "where": {"BATCH_ID": "7"}}]},
        {"key": "b", "label": "B", "cells": [{"drill": "d", "where": {"JE_SOURCE": "Payables"}}]}]}]}
    ws, t = _Sheet(), _Target()
    t.wb = None
    out = answer_write._write(t, 1, spec, {}, ws)
    assert len(out[-1]) == 1 and "BATCH_ID" in out[-1][0]
    assert (out[2]["a"], answer_write.FIRST_COL) not in ws.cells
    assert ws.cells[(out[2]["b"], answer_write.FIRST_COL)].Formula2.startswith("=SUMIFS(")


# ------------------------------------------------------------------ missing Lines table / wrong q (0x800A03EC)
def _drill_cell(alias, where):
    return {"drill": alias, "where": where}


def _spec(cell):
    from smartview_mcp.core import contract
    return contract.validate({"version": 1, "question": "Top suppliers", "gate": {"outcome": "clear"},
                              "layout": {"title": "t", "columns": [{"header": "NOK"}], "sections": [
                                  {"title": "s", "lines": [{"key": "a", "label": "KPMG AS", "cells": [cell]}]}]}})


def _spec_without_drill():
    return _spec({"value": 1})


def _lo_sheet(name, tables=()):
    from types import SimpleNamespace
    return SimpleNamespace(Name=name, ListObjects=[SimpleNamespace(Name=tn) for tn in tables])


def _target_with_tables(tables):
    from types import SimpleNamespace
    return SimpleNamespace(wb=SimpleNamespace(Worksheets=[_lo_sheet(f"S{i}", [tn]) for i, tn in enumerate(tables)]))


def _target_with_sheets(names):
    from types import SimpleNamespace
    return SimpleNamespace(wb=SimpleNamespace(Worksheets=[_lo_sheet(n) for n in names]))


RAISES = object()


def _render_drill_cell(monkeypatch, tables):
    from smartview_mcp import answer_write

    def cols(wb, tn):
        if tables is RAISES:
            raise RuntimeError("COM error")
        return tables.get(tn)
    monkeypatch.setattr(answer_write, "com", lambda call, *a, **k: call())
    monkeypatch.setattr(answer_write, "drill_table_columns", cols)
    spec = {"columns": [{"header": "NOK"}], "sections": [{"title": "x", "lines": [
        {"key": "a", "label": "KPMG AS", "cells": [_drill_cell("d2", {"AP_CUSTOMER_NAME": "KPMG AS"})]}]}]}
    ws, t = _Sheet(), _Target()
    t.wb = None
    warnings = answer_write._write(t, 2, spec, {}, ws)[-1]
    return [c.Formula2 or c.Value for c in ws.cells.values()], warnings


def test_answer_refuses_drill_cell_whose_table_is_missing_and_names_existing_tables():
    from smartview_mcp.core import pipeline
    from smartview_mcp.host import InputError
    t = _target_with_tables(["Lines_Q1_d1", "Lines_Q1_d2"])
    with pytest.raises(InputError) as e:
        pipeline.check_drill_tables(t, 2, _spec(_drill_cell("d2", {"AP_CUSTOMER_NAME": "KPMG AS"})))
    assert all(s in str(e.value) for s in ("Lines_Q2_d2", "Lines_Q1_d2", "sv_drill", "q=2"))
    pipeline.check_drill_tables(t, 1, _spec(_drill_cell("d2", {})))           # the table exists: no error


def test_answer_write_marks_missing_table_as_cell_error_not_formula(monkeypatch):
    written, warnings = _render_drill_cell(monkeypatch, tables={})            # find_table -> None
    assert not any("SUMIFS(" in str(v) for v in written)
    assert any("Lines_Q2_d2 not found" in w for w in warnings)
    assert any("Lines_Q2_d2 not found" in str(v) for v in written)


def test_answer_write_keeps_cannot_check_path_when_tables_unreadable(monkeypatch):
    written, _ = _render_drill_cell(monkeypatch, tables=RAISES)             # today's behaviour: formula written
    assert any("SUMIFS(" in str(v) for v in written)


def test_new_spec_with_drill_cells_takes_the_pending_drill_q():
    from smartview_mcp.core import pipeline
    from smartview_mcp.excel import sheets
    t = _target_with_sheets(["Q1 Lines d1", "Q1 Lines d2"])               # the live case
    assert pipeline.question_for(t, _spec(_drill_cell("d2", {}))) == 1
    assert pipeline.question_for(t, _spec_without_drill()) == 2
    assert sheets.pending_drill_question(_target_with_sheets(["Q1 Answer", "Q1 Lines d1"])) is None


# ------------------------------------------------------------------ cross-process profile lock
def test_profile_lock_is_exclusive_and_reusable(tmp_path):
    from smartview_mcp import drill_browser as db
    a, b = db.ProfileLock(tmp_path / "prof"), db.ProfileLock(tmp_path / "prof")
    assert a.acquire() and a.acquire() and a.held          # idempotent for its owner
    assert not b.acquire() and not b.held                  # another handle (another server) is refused
    a.release()
    assert not a.held and b.acquire()
    b.release()


def test_kill_never_touches_chrome_without_the_profile_lock(monkeypatch, tmp_path):
    from smartview_mcp import drill_browser as db
    killed = []
    monkeypatch.setattr(db, "kill_our_chrome", lambda prof: killed.append(prof) or [1])
    b = db.DrillBrowser(base="https://epm", profile=tmp_path / "prof")
    other = db.ProfileLock(tmp_path / "prof")
    assert other.acquire()                                  # a different server owns the profile
    assert b.kill() == [] and killed == []
    with pytest.raises(db.DrillBusy, match="different session"):
        b._start(db._Session())
    b.shutdown()
    assert killed == []                                     # shutdown fallback is guarded too
    other.release()


def test_idle_worker_closes_browser_and_frees_lock(monkeypatch, tmp_path):
    import threading
    from smartview_mcp import drill_browser as db
    lock = db.ProfileLock(tmp_path / "prof")
    assert lock.acquire()
    closed = threading.Event()
    real = db.DrillBrowser._close

    def spy(s):
        real(s)
        closed.set()
    monkeypatch.setattr(db.DrillBrowser, "_close", staticmethod(spy))
    monkeypatch.setattr(db, "_kill_driver", lambda d: [])
    import queue
    s, q = db._Session(), queue.Queue()
    s.plock = lock
    t = threading.Thread(target=db.DrillBrowser._loop, args=(q, s, 0.05), daemon=True)
    t.start()
    assert closed.wait(5) and not lock.held and s.plock is None
    q.put((None, None))
    t.join(5)


# ------------------------------------------------------------------ server helpers (no Excel, no browser)
def test_server_lines_of_normalises_each_piece_once(monkeypatch):
    from smartview_mcp.core import drillstep as server
    items = _items("whs_c40s_q326_items.json")
    view = json.loads((FIX / "view_definition.json").read_text(encoding="utf-8"))
    calls = []
    real = dl.normalize
    monkeypatch.setattr(dl, "normalize", lambda it, c, f: calls.append(len(it)) or real(it, c, f))
    half = len(items) // 2
    pieces = [{"cell": {"p": 1}, "cube_value": None, "items": items[:half]},
              {"cell": {"p": 2}, "cube_value": None, "items": items[half:]}]
    L = server._lines_of(pieces, view)
    assert calls == [half, len(items) - half]                    # each piece once, never the whole again
    assert L["rec"]["lines"] == 1231 and L["rec"]["lines_sum"] == -24706615.55
    assert [p["lines"] for p in L["piece_recs"]] == [half, len(items) - half]
    assert isinstance(L["stats"]["unfixed"], list) and L["cols"][0] == "Scenario"


def test_drill_alias_syntax():
    from smartview_mcp import tools_drill as server
    assert server._ALIAS.match("d") and server._ALIAS.match("b2_x")
    for bad in ("", "2b", "a-b", "a b", "a]"):
        assert not server._ALIAS.match(bad)


# ------------------------------------------------------------------ re-drill decisions (last review round)
def test_same_header_decides_reuse():
    assert dx.same_header(("Period", "AMOUNT"), ["Period", "AMOUNT"])
    assert not dx.same_header(("Period", "AMOUNT"), ["Period", "AMOUNT", "ACCOUNT"])
    assert not dx.same_header(("AMOUNT", "Period"), ["Period", "AMOUNT"]) and not dx.same_header(None, ["A"])


def test_grid_key_identifies_the_same_grid():
    from smartview_mcp import tools_drill as server
    a = server._grid_key("FinLite", {"All Entities": "un.SY.2WHS", "All Scenarios": "Actual"}, "ac.C40S", ["Q3 26"])
    b = server._grid_key("FinLite", {"All Scenarios": "Actual", "All Entities": "un.SY.2WHS"}, "ac.C40S", ["Q3 26"])
    assert a == b
    assert a != server._grid_key("FinLite", {"All Entities": "un.SY.2WHS", "All Scenarios": "Actual"}, "ac.C40S",
                                 ["Sep 26"])


def test_failed_redrill_restores_previous_lines_or_marks_incomplete(monkeypatch, tmp_path):
    from smartview_mcp import drill_store, tools_drill as server
    db = drill_store.connect(tmp_path / "d.db")
    did = drill_store.save(db, connection="FinLite", cell={"All Accounts": "ac.C40S"}, rec={}, columns=["AMOUNT"],
                           rows=[{"AMOUNT": -1.0}], workbook="w.xlsx", q=1, alias="d")
    db.close()
    monkeypatch.setattr(drill_store, "connect", lambda path=None: drill_store.sqlite3.connect(tmp_path / "d.db"))
    prev = server._prev_drill("w.xlsx", 1, "d")
    assert prev["id"] == did and prev["columns"] == ["AMOUNT"]
    calls = []
    monkeypatch.setattr(dx, "write_lines", lambda t, q, rows, cols, **kw: calls.append(("write", rows, cols, kw)))
    monkeypatch.setattr(dx, "mark_incomplete", lambda t, q, alias="d": calls.append(("incomplete", q, alias)))
    assert "restored" in server._restore_lines(None, 1, "d", prev, "src")
    assert calls[0][:3] == ("write", [{"AMOUNT": -1.0}], ["AMOUNT"]) and "forrige uttrekk" in calls[0][3]["title"]
    calls.clear()
    assert "INCOMPLETE" in server._restore_lines(None, 1, "d", None, "src") and calls == [("incomplete", 1, "d")]


def test_chrome_start_failure_releases_the_profile_lock(monkeypatch, tmp_path):
    from smartview_mcp import drill_browser as db
    stopped = []

    class Chromium:
        def launch_persistent_context(self, *a, **k):
            raise RuntimeError("chrome.exe not found")

    class Pw:
        chromium = Chromium()

        def stop(self):
            stopped.append(1)

    class SP:
        def start(self):
            return Pw()
    import playwright.sync_api
    monkeypatch.setattr(playwright.sync_api, "sync_playwright", lambda: SP())
    monkeypatch.setattr(db, "kill_our_chrome", lambda prof: [])
    monkeypatch.setattr(db, "_kill_driver", lambda d: [])
    monkeypatch.setattr(db, "user_data_dir_policy", lambda ch: None)
    b = db.DrillBrowser.__new__(db.DrillBrowser)
    b.base, b.profile, b.plock = "https://epm", tmp_path / "prof", db.ProfileLock(tmp_path / "prof")
    s = db._Session()
    with pytest.raises(db.DrillError, match="not found"):
        b._start(s)
    assert stopped == [1] and not b.plock.held and s.plock is None
    other = db.ProfileLock(tmp_path / "prof")
    assert other.acquire()                                  # another session can take the profile right away
    other.release()


def test_launch_skips_policy_blocked_channel_retries_and_falls_back(monkeypatch, tmp_path):
    from smartview_mcp import config
    from smartview_mcp import drill_browser as db
    tried = []

    class Chromium:
        def launch_persistent_context(self, prof, channel, **k):
            tried.append((channel, Path(prof).name))
            if channel == "msedge":
                raise RuntimeError("Target page, context or browser has been closed")
            return "ctx"
    monkeypatch.setattr(config, "DRILL_CHANNELS", ("msedge", "chrome", "chromium"))
    monkeypatch.setattr(config, "DRILL_LAUNCH_TRIES", 2)
    monkeypatch.setattr(db.time, "sleep", lambda s: None)
    monkeypatch.setattr(db, "kill_our_chrome", lambda prof: [])
    monkeypatch.setattr(db, "user_data_dir_policy", lambda ch: r"C:\x" if ch == "chrome" else None)
    b = db.DrillBrowser.__new__(db.DrillBrowser)
    b.profile = tmp_path / "drill_edge"

    class Pw:
        chromium = Chromium()
    assert b._launch(Pw()) == ("ctx", "chromium")
    assert tried == [("msedge", "drill_edge")] * 2 + [("chromium", "drill_edge_chromium")]   # chrome never launched

    monkeypatch.setattr(config, "DRILL_CHANNELS", ("msedge", "chrome"))
    with pytest.raises(db.DrillError, match=r"msedge: Target page.*chrome: skipped, policy UserDataDir"):
        b._launch(Pw())


def test_dead_browser_is_restarted_and_the_drill_retried_once(monkeypatch):
    from smartview_mcp import drill_browser as db
    b = db.DrillBrowser.__new__(db.DrillBrowser)
    s, calls, closed = db._Session(), [], []
    s.ctx = "ctx"
    monkeypatch.setattr(db.DrillBrowser, "_alive", staticmethod(lambda s: False))
    monkeypatch.setattr(db.DrillBrowser, "_close", staticmethod(lambda s: closed.append(1)))

    def once(s_, url, n, prog):
        calls.append(url)
        if len(calls) == 1:
            raise db.DrillError("Target closed")
        return {"items": [], "total": 0}
    monkeypatch.setattr(b, "_fetch_once", once)
    assert b._fetch(s, "u", 10, None) == {"items": [], "total": 0} and calls == ["u", "u"] and closed == [1]

    calls.clear()
    monkeypatch.setattr(db.DrillBrowser, "_alive", staticmethod(lambda s: True))    # live page: a real answer
    monkeypatch.setattr(b, "_fetch_once", lambda *a: calls.append(1) or (_ for _ in ()).throw(db.DrillError("x")))
    with pytest.raises(db.DrillError):
        b._fetch(s, "u", 10, None)
    assert calls == [1]


# ------------------------------------------------------------------ final review round
def _drill_fakes(monkeypatch, links: dict, totals: dict, kids: dict, split_at: int = 600):
    """Fake Smart View links, a fake browser and a fake atlas for server._drill_cell."""
    from smartview_mcp import config, drill_split, ops
    from smartview_mcp.core import drillstep as server
    from smartview_mcp.drill_browser import TooManyLines
    asked = []

    async def fake_call(fn, timeout=None, what="call"):
        return fn(None)

    def fake_link(h, conn, cell):
        asked.append(cell["All Time Periods"])
        return links[cell["All Time Periods"]]

    async def fake_fetch(url, max_lines):
        n = totals[url]
        if n > max_lines:
            raise TooManyLines(n)
        return {"items": [{"AMOUNTX": "-1"}] * n, "total": n, "view": {"v": 1}, "secs": 0}
    monkeypatch.setattr(server, "_call", fake_call)
    monkeypatch.setattr(ops, "drill_link", fake_link)
    monkeypatch.setattr(server, "_drill_fetch", fake_fetch)
    monkeypatch.setattr(config, "DRILL_SPLIT_AT", split_at)
    monkeypatch.setattr(drill_split, "atlas_children", lambda conn: lambda dim, m: kids.get(m, []))
    return asked


_LINKS = {"Q3 26": {"url": "q3", "cube_value": -1231.0}, "Jul 26": {"url": "jul", "cube_value": -400.0},
          "Aug 26": {"url": "aug", "cube_value": -400.0}, "Sep 26": {"url": "sep", "cube_value": -431.0}}
_KIDS = {"Q3 26": ["Sep 26", "Aug 26", "Jul 26"]}
_CELL = {"All Accounts": "ac.C40S", "All Entities": "un.SY.2WHS", "All Time Periods": "Q3 26"}


def test_drill_cell_splits_and_keeps_the_parent_cube(monkeypatch):
    import asyncio
    from smartview_mcp.core import drillstep as server
    _drill_fakes(monkeypatch, _LINKS, {"q3": 1231, "jul": 400, "aug": 400, "sep": 431}, _KIDS)
    pieces, parent = asyncio.run(server._drill_cell("FinLite", _CELL, 300000, "auto"))
    assert [p["cell"]["All Time Periods"] for p in pieces] == ["Jul 26", "Aug 26", "Sep 26"]
    assert [len(p["items"]) for p in pieces] == [400, 400, 431] and parent == -1231.0


def test_drill_cell_max_lines_guard_before_splitting(monkeypatch):
    import asyncio
    from smartview_mcp.core import drillstep as server
    asked = _drill_fakes(monkeypatch, _LINKS, {"q3": 1231, "jul": 400, "aug": 400, "sep": 431}, _KIDS)
    with pytest.raises(server.ToolError, match="1231 lines"):
        asyncio.run(server._drill_cell("FinLite", _CELL, 1000, "auto"))
    assert asked == ["Q3 26"]                                   # no child was drilled


def test_drill_cell_passes_the_remaining_budget_down(monkeypatch):
    import asyncio
    from smartview_mcp.core import drillstep as server
    from smartview_mcp.drill_browser import TooManyLines
    _drill_fakes(monkeypatch, _LINKS, {"q3": 1231, "jul": 400, "aug": 400, "sep": 431}, _KIDS)
    real = server._drill_fetch

    async def unknown_total(url, max_lines):                    # totalResults missing: only "more than the limit"
        if url == "q3":
            raise TooManyLines(max_lines + 1)
        return await real(url, max_lines)
    monkeypatch.setattr(server, "_drill_fetch", unknown_total)
    with pytest.raises(server.ToolError, match="max_lines=1000"):
        asyncio.run(server._drill_cell("FinLite", _CELL, 1000, "auto"))     # 400 + 400 + 431 > 1000


def test_drill_cell_mixed_piece_without_drill_report(monkeypatch):
    import asyncio
    from smartview_mcp.core import drillstep as server
    links = {**_LINKS, "Sep 26": {"url": None, "cube_value": -431.0}}
    _drill_fakes(monkeypatch, links, {"q3": 1231, "jul": 400, "aug": 400}, _KIDS)
    pieces, parent = asyncio.run(server._drill_cell("FinLite", _CELL, 300000, "auto"))
    assert [len(p["items"]) for p in pieces] == [400, 400, 0] and "no drill-through report" in pieces[2]["note"]


def test_drill_cell_too_wide_splits_by_entity_and_skips_empty(monkeypatch):
    """EPMFDM-140373 on a top entity: drill its child entities (not periods); #Missing children are skipped."""
    import asyncio
    from smartview_mcp import drill_split, ops
    from smartview_mcp.core import drillstep as server
    from smartview_mcp.drill_browser import CellTooWide
    links = {"un.SY": {"url": "sy", "cube_value": -161.0}, "un.SY.2SSM": {"url": "ssm", "cube_value": -98.0},
             "un.SY.2I15": {"url": "i15", "cube_value": -63.0}, "un.SY.2CON": {"url": "con", "cube_value": None}}
    fetched = []

    async def fake_call(fn, timeout=None, what="call"):
        return fn(None)

    async def fake_fetch(url, max_lines):
        fetched.append(url)
        if url == "sy":
            raise CellTooWide("EPMFDM-140373: more than 3 dimensions have more than 1000 child members")
        return {"items": [{"AMOUNTX": "-1"}], "total": 1, "view": {"v": 1}, "secs": 0}
    monkeypatch.setattr(server, "_call", fake_call)
    monkeypatch.setattr(ops, "drill_link", lambda h, conn, cell: links[cell["All Entities"]])
    monkeypatch.setattr(server, "_drill_fetch", fake_fetch)
    monkeypatch.setattr(server, "_too_wide", set())
    kids = {"Q3 26": ["Sep 26"], "un.SY": ["un.SY.2SSM", "un.SY.2I15", "un.SY.2CON"]}
    monkeypatch.setattr(drill_split, "atlas_children", lambda conn: lambda dim, m: kids.get(m, []))
    cell = {"All Accounts": "ac.12881", "All Entities": "un.SY", "All Time Periods": "Q3 26"}
    pieces, parent = asyncio.run(server._drill_cell("FinLite", cell, 300000, "auto"))
    assert [p["cell"]["All Entities"] for p in pieces] == ["un.SY.2SSM", "un.SY.2I15", "un.SY.2CON"]
    assert all(p["cell"]["All Time Periods"] == "Q3 26" for p in pieces) and parent == -161.0
    assert fetched == ["sy", "ssm", "i15"] and pieces[2]["note"] == "empty cell (skipped)"
    fetched.clear()                                             # a re-drill splits at once, no second refusal
    pieces2, _ = asyncio.run(server._drill_cell("FinLite", cell, 300000, "auto"))
    assert fetched == ["ssm", "i15"] and len(pieces2) == 3


def test_drill_cell_too_wide_on_a_leaf_is_a_readable_error(monkeypatch):
    import asyncio
    from smartview_mcp import drill_split, ops
    from smartview_mcp.core import drillstep as server
    from smartview_mcp.drill_browser import CellTooWide

    async def fake_call(fn, timeout=None, what="call"):
        return fn(None)

    async def fake_fetch(url, max_lines):
        raise CellTooWide("EPMFDM-140373")
    monkeypatch.setattr(server, "_call", fake_call)
    monkeypatch.setattr(ops, "drill_link", lambda h, conn, cell: {"url": "x", "cube_value": -1.0})
    monkeypatch.setattr(server, "_drill_fetch", fake_fetch)
    monkeypatch.setattr(drill_split, "atlas_children", lambda conn: lambda dim, m: [])
    with pytest.raises(server.ToolError, match="too wide"):
        asyncio.run(server._drill_cell("FinLite", {"All Entities": "un.SY.A101", "All Accounts": "ac.12881",
                                                   "All Time Periods": "Sep 26"}, 300000, "auto"))


def test_lines_of_reconciles_against_the_parent_and_checks_the_split():
    from smartview_mcp.core import drillstep as server
    items = _items("whs_c40s_q326_items.json")
    view = json.loads((FIX / "view_definition.json").read_text(encoding="utf-8"))
    half = len(items) // 2
    pieces = [{"cell": {"p": 1}, "cube_value": -100.0, "items": items[:half]},
              {"cell": {"p": 2}, "cube_value": -200.0, "items": items[half:]}]
    L = server._lines_of(pieces, view, parent_cube=-24706615.55)
    assert L["rec"]["status"] == "exact" and L["rec"]["cube_value"] == -24706615.55
    assert L["split_check"] == round(-300.0 + 24706615.55, 2)
    assert server._lines_of(pieces[:1], view, parent_cube=-100.0)["split_check"] is None   # no split


def test_gap_warning_shows_share_of_cell():
    from smartview_mcp.core import drillstep as server
    w = server._gap_warning({"gap": -250.0, "cube_value": -1000.0})
    assert "-250.00" in w and "25.0 %" in w


def test_drill_retries_once_on_empty_provider_response(monkeypatch):
    import asyncio
    from smartview_mcp.core import drillstep
    tries, notes = [], []

    async def flaky(conn, cell, max_lines, split):
        tries.append(1)
        if len(tries) == 1:
            raise drillstep.ToolError("Smart View error: Empty response from provider")
        return [{"cell": cell, "cube_value": -5.0, "items": [], "view": None, "secs": 0, "note": "no report"}], -5.0
    monkeypatch.setattr(drillstep, "_drill_cell", flaky)
    res = asyncio.run(drillstep.drill(_CELL, say=notes.append))
    assert len(tries) == 2 and len(notes) == 1 and "retrying" in notes[0]
    assert res.cell_value == -5.0 and res.rows == []

    tries.clear()

    async def always(conn, cell, max_lines, split):
        tries.append(1)
        raise drillstep.ToolError("Empty response from provider")
    monkeypatch.setattr(drillstep, "_drill_cell", always)
    with pytest.raises(drillstep.ToolError):
        asyncio.run(drillstep.drill(_CELL, say=notes.append))
    assert len(tries) == 2                                      # one retry only

    async def other(conn, cell, max_lines, split):
        tries.append(1)
        raise drillstep.ToolError("EPM refused")
    tries.clear()
    monkeypatch.setattr(drillstep, "_drill_cell", other)
    with pytest.raises(drillstep.ToolError, match="refused"):
        asyncio.run(drillstep.drill(_CELL, say=notes.append))
    assert len(tries) == 1


def test_fetch_never_leaks_the_sso_token():
    from smartview_mcp import drill_browser as db
    b = db.DrillBrowser.__new__(db.DrillBrowser)
    b._ensure = lambda s: None

    def boom(s, url):
        raise db.PwError("Page.goto: net::ERR at https://epm/aif/x.jsp?sso_token=abc&x=1")
    b._capture = boom
    with pytest.raises(db.DrillError) as e:
        b._fetch(db._Session(), "aif/x.jsp?sso_token=abc", 10, None)
    assert "abc" not in str(e.value)


def test_sv_drill_error_text_is_redacted(monkeypatch):
    import asyncio
    from smartview_mcp import tools_drill as server
    logged = []

    async def boom(*a, **k):
        raise RuntimeError("failed at https://epm/aif/x.jsp?sso_token=abc")
    monkeypatch.setattr(server, "_sv_drill", boom)
    monkeypatch.setattr(server.qlog, "write", logged.append)
    with pytest.raises(Exception) as e:
        asyncio.run(server.sv_drill("q", _CELL))
    assert "abc" not in str(e.value) and "abc" not in json.dumps(logged) and logged


def test_page_url_replaces_only_limit_and_offset():
    from urllib.parse import parse_qs, urlsplit
    from smartview_mcp import drill_browser as db
    u = db._page_url("https://epm/aif/ui/model/drilldown/getData?limit=25&offset=0&locale=nb&x=a%20b", 1500)
    assert u.startswith("https://epm/aif/ui/model/drilldown/getData?")
    assert parse_qs(urlsplit(u).query) == {"limit": ["500"], "offset": ["1500"], "locale": ["nb"], "x": ["a b"]}


def test_drill_link_reports_a_failed_status(monkeypatch):
    from smartview_mcp import ops
    from smartview_mcp.host import HostError

    class Lay:
        nr, nc = 1, 1
    monkeypatch.setattr(ops, "query", lambda *a, **k: (Lay(), None, None))
    monkeypatch.setattr(ops, "ensure_grid_sheet", lambda h, conn: "grid")

    class Rng:
        Value = -5.0

    class WS:
        def Range(self, a):
            return Rng()

    class WB:
        def Worksheets(self, name):
            return WS()

    class H:
        wb = WB()
        reply = None

        def run(self, fn, *a):
            return self.reply
    h = H()
    h.reply = ("ERR|-17|no connection|", None, None, None)
    with pytest.raises(HostError, match="ERR"):
        ops.drill_link(h, "FinLite", _CELL)
    h.reply = ("OK", None, None, None)
    assert ops.drill_link(h, "FinLite", _CELL)["url"] is None


def test_drill_formula_treats_none_as_blank():
    assert dx.drill_formula("T", {"AP_CUSTOMER_NAME": None}) == dx.drill_formula("T", {"AP_CUSTOMER_NAME": ""})


def test_evaluate_drill_cell_on_a_column_not_in_excel_is_empty():
    spec = {"columns": [{"header": "NOK"}], "sections": [{"title": "x", "lines": [
        {"key": "a", "label": "A", "cells": [{"drill": "d", "where": {"BATCH_ID": "7"}}]},
        {"key": "b", "label": "B", "cells": [{"drill": "d", "where": {"je_source": "Payables"}}]}]}]}
    rows = [{"JE_SOURCE": "Payables", "AMOUNT": -2.0}]
    v = specs.evaluate(spec, {}, drills={"d": rows})
    assert v["a"] == [None] and v["b"] == [-2.0]


def test_profile_shares_for_no_supplier_and_untagged(whs):
    rows, _ = whs
    p = dl.profile(rows)
    for k in ("no_supplier", "untagged_product"):
        assert 0 <= p[k]["share_lines"] <= 1 and "share_amount" in p[k]
    assert p["no_supplier"]["share_lines"] == round(p["no_supplier"]["lines"] / len(rows), 4)


# ------------------------------------------------------------------ reliability audit (Task 11)
def test_drill_cell_defaults_entity_and_needs_account_and_period():
    from smartview_mcp.host import InputError
    from smartview_mcp.ops import drill_cell_defaults
    cell, w = drill_cell_defaults({"All Accounts": "ac.12881", "All Time Periods": "Sep 26"})
    assert cell["All Entities"] == "All Entities" and "All Entities not given" in w[0]
    cell, w = drill_cell_defaults({"All Accounts": "a", "All Time Periods": "p", "All Entities": "un.SY"})
    assert cell["All Entities"] == "un.SY" and w == []
    for bad in ({"All Time Periods": "Sep 26"}, {"All Accounts": "ac.12881"}):
        with pytest.raises(InputError):
            drill_cell_defaults(bad)


def test_classify_stall():
    from smartview_mcp.drill_browser import classify_stall
    assert classify_stall("https://login.microsoftonline.com/x", "Sign in") == "sso"
    assert classify_stall("https://epm/cloudgate/v1/oauth2/callback", "") == "sso"
    assert classify_stall("https://epm/aif/drill", "  ") == "blank"
    assert classify_stall("https://epm/aif/drill", "Drill Through") == "page"


def test_login_action_nudges_a_stalled_redirect_once():
    from smartview_mcp.drill_browser import SIGNIN, login_action
    base = "https://epm"
    assert login_action(base + SIGNIN, base, 1, 2, clicked=False, nudged=False) == "click"
    assert login_action(base + "/HyperionPlanning/", base, 5, 5, clicked=True, nudged=False) == "done"
    assert login_action(base + "/HyperionPlanning/", base, 2, 2, clicked=True, nudged=False) == "wait"
    cb = base + "/cloudgate/v1/oauth2/callback"
    assert login_action(cb, base, 10, 5, clicked=True, nudged=False) == "wait"
    assert login_action(cb, base, 30, 20, clicked=True, nudged=False) == "home"     # stuck 20 s on one page
    assert login_action(cb, base, 50, 20, clicked=True, nudged=True) == "wait"      # only once


def test_login_skips_the_retry_when_the_drill_deadline_is_near():
    from smartview_mcp import drill_browser as db
    b = db.DrillBrowser.__new__(db.DrillBrowser)
    tries = []

    def never(p, deadline):
        tries.append(p)
        raise db.DrillError("last page: oauth2/callback")
    b._login_once = never
    s = db._Session()
    s.page, s.deadline = object(), db.time.monotonic() + 5
    with pytest.raises(db.DrillError, match="open EPM once in Edge"):
        b._login(s)
    assert len(tries) == 1


def test_page_loads_are_bounded_by_the_drill_deadline():
    from smartview_mcp.drill_browser import load_timeout_ms
    assert load_timeout_ms(float("inf"), 0) == 90_000
    assert load_timeout_ms(100.0, 70.0) == 30_000
    assert load_timeout_ms(100.0, 120.0) == 1_000


def test_no_nudge_without_time_left():
    from smartview_mcp.drill_browser import login_action
    cb = "https://epm/cloudgate/v1/oauth2/callback"
    assert login_action(cb, "https://epm", 30, 20, clicked=True, nudged=False, left=10) == "wait"


def test_crash_retry_needs_time_left(monkeypatch):
    from smartview_mcp import drill_browser as db
    b = db.DrillBrowser.__new__(db.DrillBrowser)
    calls = []

    def once(s, url, n, progress):
        calls.append(1)
        s.ctx = object()                           # a browser that is dead now
        s.deadline = db.time.monotonic() + 5       # most of the drill's time is used up
        raise db.DrillError("Page crashed")
    b._fetch_once, b._alive, b._close = once, lambda s: False, lambda s: None
    with pytest.raises(db.DrillError, match="browser crashed"):
        b._fetch(db._Session(), "u", 10, None)
    assert len(calls) == 1


def test_crash_while_paging_retries_with_a_fresh_deadline():
    from smartview_mcp import drill_browser as db
    b = db.DrillBrowser.__new__(db.DrillBrowser)
    calls = []

    def once(s, url, n, progress):
        calls.append(s.deadline - db.time.monotonic())
        if len(calls) == 1:
            s.ctx, s.paging = object(), True        # captured, then crashed while paging
            s.deadline = db.time.monotonic() - 50   # long past the setup deadline
            raise db.DrillError("Page crashed")
        return {"items": []}
    b._fetch_once, b._alive, b._close = once, lambda s: False, lambda s: None
    assert b._fetch(db._Session(), "u", 10, None) == {"items": []}
    assert calls[1] > 60


class _FakeLoginPage:
    def __init__(self, clock):
        self.url = "about:blank"
        self.clock = clock

    def goto(self, url, wait_until=None, timeout=None):
        self.url = url

    def wait_for_load_state(self, *a, **k):
        pass

    def wait_for_timeout(self, ms):
        self.clock[0] += ms / 1000


def test_login_once_signs_in_on_another_pod(tmp_path, monkeypatch):
    from smartview_mcp import drill_browser as db
    clock = [0.0]                                   # fake clock: waiting advances it, so login_action's 3 s passes at once
    monkeypatch.setattr(db.time, "monotonic", lambda: clock[0])
    b = db.DrillBrowser(base="https://plan1.example", profile=tmp_path / "p")
    p = _FakeLoginPage(clock)
    b._login_once(p, float("inf"), base="https://strat.example/")
    assert p.url == "https://strat.example/HyperionPlanning/"
    b._login_once(p, float("inf"))
    assert p.url == "https://plan1.example/HyperionPlanning/"
    b.shutdown()


def test_shutdown_shared_is_idempotent(monkeypatch):
    from smartview_mcp import drill_browser as db
    monkeypatch.setattr(db, "_shared", None)
    class Fake:
        shutdowns = 0
        def shutdown(self):
            Fake.shutdowns += 1
    monkeypatch.setattr(db, "DrillBrowser", Fake)
    first = db.shared_browser()
    db.shutdown_shared()
    db.shutdown_shared()
    assert Fake.shutdowns == 1
    assert db.shared_browser() is not first


def test_shared_browser_is_one_instance(monkeypatch):
    from smartview_mcp import drill_browser as db
    from smartview_mcp.core import drillstep
    monkeypatch.setattr(db, "_shared", None)
    made = []
    monkeypatch.setattr(db, "DrillBrowser", lambda: made.append(1) or object())
    first = db.shared_browser()
    assert db.shared_browser() is first and drillstep.browser() is first and made == [1]
