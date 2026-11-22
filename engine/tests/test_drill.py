"""Unit tests for fin_drill's pure parts (fixtures from the 15.10.2026 spike)."""
import json
from pathlib import Path

import pytest

from smartass.epm import drill_lines as dl

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

from smartass.epm import drill_split as ds

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


from smartass.xl import drill_xl as dx


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
    from smartass.xl import formulas
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


DRILL_SPEC = {"columns": [{"header": "NOK"}, {"header": "Lines"}],
              "sections": [{"title": "By source", "lines": [
                  {"key": "pay", "label": "Payables", "cells": [
                      {"drill": "d", "where": {"JE_SOURCE": "Payables"}},
                      {"drill": "d", "where": {"JE_SOURCE": "Payables"}, "agg": "count"}]},
                  {"key": "jul", "label": "July", "cells": [
                      {"drill": "d", "where": {"Period": "Jul-26"}, "agg": "debit"}, None]},
                  {"key": "tot", "label": "Total", "kind": "sum", "cells": [
                      {"drill": "d", "where": {}}, {"drill": "d", "where": {}, "agg": "count"}]}]}]}


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
    from smartass.xl import sheets
    t = _target(names)
    sheets.add_sheet(t, "X", after_q=after_q)
    assert [s.Name for s in t.wb.Worksheets] == expect and t.sheets_added == ["X"]


def test_question_grouping_includes_lines(monkeypatch):
    from smartass.xl import sheets
    assert sheets._Q.match("Q3 Lines b") and not sheets._Q.match("Q3 Forecast")
    t = _target(["Q2 Answer", "Q2 FinLite", "Q2 Lines", "Sources"])
    assert sheets.next_question(t) == 3


# ------------------------------------------------------------------ missing Lines table / wrong q (0x800A03EC)
def _drill_cell(alias, where):
    return {"drill": alias, "where": where}


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


# ------------------------------------------------------------------ server helpers (no Excel, no EPM)
def test_server_lines_of_normalises_each_piece_once(monkeypatch):
    from smartass.epm import drillstep as server
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


# ------------------------------------------------------------------ final review round
def _drill_fakes(monkeypatch, links: dict, totals: dict, kids: dict, split_at: int = 600):
    """Fake drill links (the HTTP route), a fake line fetch and a fake atlas for server._drill_cell; the split
    children's values are not read (None: every child is linked)."""
    from smartass.core import config
    from smartass.epm import drill_split
    from smartass.epm import drillstep as server
    from smartass.epm.drillcore import TooManyLines
    asked = []

    def fake_link(conn, cell):
        asked.append(cell["All Time Periods"])
        return links[cell["All Time Periods"]]

    async def no_values(connection, parts):
        return None

    async def fake_fetch(url, max_lines):
        n = totals[url]
        if n > max_lines:
            raise TooManyLines(n)
        return {"items": [{"AMOUNTX": "-1"}] * n, "total": n, "view": {"v": 1}, "secs": 0}
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")
    monkeypatch.setattr(server, "_http_link", fake_link)
    monkeypatch.setattr(server, "_child_values", no_values)
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
    from smartass.epm import drillstep as server
    _drill_fakes(monkeypatch, _LINKS, {"q3": 1231, "jul": 400, "aug": 400, "sep": 431}, _KIDS)
    pieces, parent = asyncio.run(server._drill_cell("FinLite", _CELL, 300000, "auto"))
    assert [p["cell"]["All Time Periods"] for p in pieces] == ["Jul 26", "Aug 26", "Sep 26"]
    assert [len(p["items"]) for p in pieces] == [400, 400, 431] and parent == -1231.0


def test_drill_cell_max_lines_guard_before_splitting(monkeypatch):
    import asyncio
    from smartass.epm import drillstep as server
    asked = _drill_fakes(monkeypatch, _LINKS, {"q3": 1231, "jul": 400, "aug": 400, "sep": 431}, _KIDS)
    with pytest.raises(server.ToolError, match="1231 lines"):
        asyncio.run(server._drill_cell("FinLite", _CELL, 1000, "auto"))
    assert asked == ["Q3 26"]                                   # no child was drilled


def test_drill_cell_passes_the_remaining_budget_down(monkeypatch):
    import asyncio
    from smartass.epm import drillstep as server
    from smartass.epm.drillcore import TooManyLines
    _drill_fakes(monkeypatch, _LINKS, {"q3": 1231, "jul": 400, "aug": 400, "sep": 431}, _KIDS)
    real = server._drill_fetch

    async def unknown_total(url, max_lines):                    # totalResults missing: only "more than the limit"
        if url == "q3":
            raise TooManyLines(max_lines + 1)
        return await real(url, max_lines)
    monkeypatch.setattr(server, "_drill_fetch", unknown_total)
    with pytest.raises(server.ToolError, match="max_lines=1000"):
        asyncio.run(server._drill_cell("FinLite", _CELL, 1000, "auto"))     # 400 + 400 + 431 > 1000


def test_drill_cell_fetches_the_children_at_once_in_order(monkeypatch):
    """The split children are drilled concurrently; the pieces keep the plan's order."""
    import asyncio
    from smartass.epm import drillstep as server
    _drill_fakes(monkeypatch, _LINKS, {"q3": 1231, "jul": 400, "aug": 400, "sep": 431}, _KIDS)
    real, live = server._drill_fetch, {"now": 0, "peak": 0}

    async def slow(url, max_lines):
        live["now"] += 1
        live["peak"] = max(live["peak"], live["now"])
        await asyncio.sleep(0.01 if url != "jul" else 0.03)      # the first child finishes last
        live["now"] -= 1
        return await real(url, max_lines)
    monkeypatch.setattr(server, "_drill_fetch", slow)
    pieces, _ = asyncio.run(server._drill_cell("FinLite", _CELL, 300000, "auto"))
    assert live["peak"] == 3 and [p["cell"]["All Time Periods"] for p in pieces] == ["Jul 26", "Aug 26", "Sep 26"]


def test_drill_cell_mixed_piece_without_drill_report(monkeypatch):
    import asyncio
    from smartass.epm import drillstep as server
    links = {**_LINKS, "Sep 26": {"url": None, "cube_value": -431.0}}
    _drill_fakes(monkeypatch, links, {"q3": 1231, "jul": 400, "aug": 400}, _KIDS)
    pieces, parent = asyncio.run(server._drill_cell("FinLite", _CELL, 300000, "auto"))
    assert [len(p["items"]) for p in pieces] == [400, 400, 0] and "no drill-through report" in pieces[2]["note"]


def test_drill_cell_too_wide_splits_by_entity_and_skips_empty(monkeypatch):
    """EPMFDM-140373 on a top entity: drill its child entities (not periods); #Missing children are skipped."""
    import asyncio
    from smartass.epm import drill_split
    from smartass.epm import drillstep as server
    from smartass.epm.drillcore import CellTooWide
    links = {"un.SY": {"url": "sy", "cube_value": -161.0}, "un.SY.2SSM": {"url": "ssm", "cube_value": -98.0},
             "un.SY.2I15": {"url": "i15", "cube_value": -63.0}, "un.SY.2CON": {"url": "con", "cube_value": None}}
    fetched = []

    async def no_values(connection, parts):
        return None

    async def fake_fetch(url, max_lines):
        fetched.append(url)
        if url == "sy":
            raise CellTooWide("EPMFDM-140373: more than 3 dimensions have more than 1000 child members")
        return {"items": [{"AMOUNTX": "-1"}], "total": 1, "view": {"v": 1}, "secs": 0}
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")
    monkeypatch.setattr(server, "_http_link", lambda conn, cell: links[cell["All Entities"]])
    monkeypatch.setattr(server, "_child_values", no_values)
    monkeypatch.setattr(server, "_drill_fetch", fake_fetch)
    monkeypatch.setattr(drill_split, "_too_wide", set())
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


def test_too_wide_memo_survives_a_server_restart(monkeypatch, tmp_path):
    from smartass.core import config
    from smartass.epm import drill_split
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    monkeypatch.setattr(drill_split, "_too_wide", None)
    key = drill_split.too_wide_key("FinLite", {"All Entities": "un.SY", "All Time Periods": "Yr 26"})
    assert key not in drill_split.known_too_wide()
    drill_split.remember_too_wide(key)
    monkeypatch.setattr(drill_split, "_too_wide", None)              # a new server process
    assert key in drill_split.known_too_wide()


def test_split_links_only_the_children_with_a_value(monkeypatch):
    """One retrieve of all the children's values first: an empty (#Missing) child gets no Smart View drill link."""
    import asyncio
    from smartass.epm import drillstep as server
    asked = _drill_fakes(monkeypatch, _LINKS, {"q3": 1231, "jul": 400, "aug": 400, "sep": 431}, _KIDS)
    got = []

    async def values(connection, parts):
        got.append([p["All Time Periods"] for p in parts])
        return [-400.0, None, -431.0]
    monkeypatch.setattr(server, "_child_values", values)
    pieces, _ = asyncio.run(server._drill_cell("FinLite", _CELL, 300000, "auto"))
    assert got == [["Jul 26", "Aug 26", "Sep 26"]] and asked == ["Q3 26", "Jul 26", "Sep 26"]
    assert [p.get("note") for p in pieces] == [None, "empty cell (skipped)", None]


def test_drill_cell_too_wide_on_a_leaf_is_a_readable_error(monkeypatch):
    import asyncio
    from smartass.epm import drill_split
    from smartass.epm import drillstep as server
    from smartass.epm.drillcore import CellTooWide

    async def fake_fetch(url, max_lines):
        raise CellTooWide("EPMFDM-140373")
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")
    monkeypatch.setattr(server, "_http_link", lambda conn, cell: {"url": "x", "cube_value": -1.0})
    monkeypatch.setattr(server, "_drill_fetch", fake_fetch)
    monkeypatch.setattr(drill_split, "atlas_children", lambda conn: lambda dim, m: [])
    with pytest.raises(server.ToolError, match="too wide"):
        asyncio.run(server._drill_cell("FinLite", {"All Entities": "un.SY.A101", "All Accounts": "ac.12881",
                                                   "All Time Periods": "Sep 26"}, 300000, "auto"))


def test_lines_of_reconciles_against_the_parent_and_checks_the_split():
    from smartass.epm import drillstep as server
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
    from smartass.epm import drillstep as server
    w = server._gap_warning({"gap": -250.0, "cube_value": -1000.0})
    assert "-250.00" in w and "25.0 %" in w


def test_drill_retries_once_on_empty_provider_response(monkeypatch):
    import asyncio
    from smartass.epm import drillstep
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


def test_drill_link_reports_a_failed_status(monkeypatch):
    from smartass.epm import ops
    from smartass.xl.host import HostError

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


def test_profile_shares_for_no_supplier_and_untagged(whs):
    rows, _ = whs
    p = dl.profile(rows)
    for k in ("no_supplier", "untagged_product"):
        assert 0 <= p[k]["share_lines"] <= 1 and "share_amount" in p[k]
    assert p["no_supplier"]["share_lines"] == round(p["no_supplier"]["lines"] / len(rows), 4)


# ------------------------------------------------------------------ reliability audit (Task 11)
def test_drill_cell_defaults_entity_and_needs_account_and_period():
    from smartass.xl.host import InputError
    from smartass.epm.ops import drill_cell_defaults
    cell, w = drill_cell_defaults({"All Accounts": "ac.12881", "All Time Periods": "Sep 26"})
    assert cell["All Entities"] == "All Entities" and "All Entities not given" in w[0]
    cell, w = drill_cell_defaults({"All Accounts": "a", "All Time Periods": "p", "All Entities": "un.SY"})
    assert cell["All Entities"] == "un.SY" and w == []
    for bad in ({"All Time Periods": "Sep 26"}, {"All Accounts": "ac.12881"}):
        with pytest.raises(InputError):
            drill_cell_defaults(bad)


def test_drill_link_uses_http_and_never_excel(monkeypatch):
    import asyncio
    from smartass.epm import drillstep as server
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")
    monkeypatch.setattr(server, "_http_link", lambda conn, cell: {"url": "u", "cube_value": -1.0, "report": "r"})

    async def no_excel(fn, what="call"):
        raise AssertionError("Excel was used")
    monkeypatch.setattr(server, "_call", no_excel)
    assert asyncio.run(server._link("FinLite", _CELL)) == {"url": "u", "cube_value": -1.0, "report": "r"}


def test_a_drill_link_the_http_route_cannot_fetch_is_a_clear_error_never_excel(monkeypatch):
    """Planning REST has no drill-report request: there is no second route for the drill link."""
    import asyncio
    from smartass.epm import svhttp
    from smartass.epm import drillstep as server
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")

    def broken(conn, cell):
        raise svhttp.SvHttpError("Smart View provider: session gone")
    monkeypatch.setattr(server, "_http_link", broken)

    async def no_excel(fn, what="call"):
        raise AssertionError("Excel was used")
    monkeypatch.setattr(server, "_call", no_excel)
    with pytest.raises(server.ToolError, match="drill link .* could not be fetched.*session gone.*no other route"):
        asyncio.run(server._link("FinLite", _CELL))


@pytest.mark.parametrize("error, advice", [
    ("NoScope", "Add the scope of the pod named above under [oauth.scopes]"),
    ("AtlasStale", "Build the atlas (`smartass atlas build`), then drill again."),
    ("TokenRefused", "Run `smartass login --force` to sign in again (without --force a pod that still looks signed in is"),
    ("NoFallback", "Try again in a minute."),                   # the data-source lookup busy: transient
    ("NoToken", "Run `smartass login --status`"),             # no access token: the sign-in, not 'try again'
    ("GridRefused", "Look the cell's members up with fin_lookup"),   # an alias clash: deterministic
    ("SvHttpError", "Try again in a minute."),
    ("RequestLost", "Try again in a minute."),
])
def test_a_drill_link_that_fails_gives_the_next_step_of_its_cause(monkeypatch, error, advice):
    """'Try again in a minute' only for a transient failure; a missing scope, a stale atlas and a refused token each
    get their own next step - one only: the tool's guard adds none to a message that has one."""
    import asyncio
    from smartass.core import serverkit
    from smartass.epm import svhttp
    from smartass.epm import drillstep as server
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")

    def broken(conn, cell):
        raise getattr(svhttp, error)("Smart View provider: it failed")
    monkeypatch.setattr(server, "_http_link", broken)
    with pytest.raises(server.ToolError) as e:
        asyncio.run(server._link("FinLite", _CELL))
    text = str(e.value)
    assert "it failed" in text and "no other route" in text and f"Next step: {advice}" in text
    assert text.count("Next step:") == 1 and serverkit.with_next_step(text) == text
    assert ("Try again in a minute" in text) == (advice == "Try again in a minute.")


def test_a_refused_token_gets_login_force_and_names_a_missing_role_or_scope(monkeypatch):
    """Carry-in (Task 3c review, Important): login without --force skips a pod that still looks signed in, so the user
    looped; and EPM's gate refuses a valid token for a missing role or scope too."""
    import asyncio
    from smartass.epm import svhttp
    from smartass.epm import drillstep as server
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")

    def refused(conn, cell):
        raise svhttp.TokenRefused("Smart View provider at https://p: EPM refused the access token twice (HTTP 403)")
    monkeypatch.setattr(server, "_http_link", refused)
    with pytest.raises(server.ToolError) as e:
        asyncio.run(server._link("FinLite", _CELL))
    step = str(e.value).split("Next step: ", 1)[1]
    assert "`smartass login --force`" in step and "role" in step and "scope" in step


def test_a_drill_link_without_an_access_token_keeps_the_sign_in_text_and_one_sign_in_step(monkeypatch):
    """Carry-in (Task 3c review, Minor): the oauth error's own text stays in the reply, and the one next step is about
    the sign-in - not a generic 'try again in a minute' after it."""
    import asyncio
    from smartass.epm import svhttp
    from smartass.epm import drillstep as server
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")
    why = "no EPM access token for https://plan1.example: the identity domain could not renew the sign-in (400)"

    def no_token(conn, cell):
        raise svhttp.NoToken(why)
    monkeypatch.setattr(server, "_http_link", no_token)
    with pytest.raises(server.ToolError) as e:
        asyncio.run(server._link("FinLite", _CELL))
    text = str(e.value)
    assert why in text and text.count("Next step:") == 1 and "Try again in a minute" not in text
    assert "`smartass login --force`" in text.split("Next step: ", 1)[1]


def test_with_the_http_route_switched_off_the_drill_link_says_so(monkeypatch):
    import asyncio
    from smartass.epm import drillstep as server

    async def no_excel(fn, what="call"):
        raise AssertionError("Excel was used")
    monkeypatch.setattr(server, "_call", no_excel)
    monkeypatch.setattr(server, "_http_link", lambda conn, cell: pytest.fail("the HTTP route is off"))
    with pytest.raises(server.ToolError, match="switched off.*SV_TEST_NO_HTTP_DRILL"):
        asyncio.run(server._link("FinLite", _CELL))


def test_split_values_come_over_rest_when_the_provider_route_cannot_read_them(monkeypatch):
    import asyncio
    from smartass.epm import svhttp
    from smartass.epm import drillstep as server
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")
    asked = []

    class Routes:
        def child_values(self, conn, parts, rest=False):
            asked.append(rest)
            if not rest:
                raise svhttp.SvHttpError("Smart View provider: session gone")
            return [-1.0, None]
    monkeypatch.setattr(svhttp, "shared", Routes)

    async def no_excel(fn, what="call"):
        raise AssertionError("Excel was used")
    monkeypatch.setattr(server, "_call", no_excel)
    parts = [_CELL, {**_CELL, "All Time Periods": "Aug 26"}]
    assert asyncio.run(server._child_values("FinLite", parts)) == [-1.0, None] and asked == [False, True]


def test_split_values_are_not_asked_of_rest_after_an_error_rest_would_meet_too(monkeypatch):
    """A NoFallback error (the same token, scope or atlas; the data-source lookup busy) is met by REST as well: the
    values are only an optimisation, so the children are linked without them and the wait is not doubled."""
    import asyncio
    from smartass.epm import svhttp
    from smartass.epm import drillstep as server
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")
    for error in (svhttp.NoFallback("the HTTP route is busy looking up the EPM data sources"),
                  svhttp.TokenRefused("EPM refused the access token twice (HTTP 401)"),
                  svhttp.NoScope("no OAuth scope is configured"), svhttp.AtlasStale("the atlas is stale")):
        asked = []

        class Routes:
            def child_values(self, conn, parts, rest=False, error=error):
                asked.append(rest)
                raise error
        monkeypatch.setattr(svhttp, "shared", Routes)
        parts = [_CELL, {**_CELL, "All Time Periods": "Aug 26"}]
        assert asyncio.run(server._child_values("FinLite", parts)) is None and asked == [False]


def test_split_values_neither_route_can_read_are_none_and_every_child_is_linked(monkeypatch):
    import asyncio
    from smartass.epm import svhttp
    from smartass.epm import drillstep as server
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")

    class Broken:
        def child_values(self, conn, parts, rest=False):
            raise svhttp.SvHttpError("down")
    monkeypatch.setattr(svhttp, "shared", Broken)

    async def no_excel(fn, what="call"):
        raise AssertionError("Excel was used")
    monkeypatch.setattr(server, "_call", no_excel)
    assert asyncio.run(server._child_values("FinLite", [_CELL, {**_CELL, "All Time Periods": "Aug 26"}])) is None
    monkeypatch.setenv("SV_TEST_NO_HTTP_DRILL", "1")                     # the HTTP route off: not read at all
    assert asyncio.run(server._child_values("FinLite", [_CELL, {**_CELL, "All Time Periods": "Aug 26"}])) is None


def test_a_drill_link_without_a_sign_in_is_raised_not_sent_to_excel(monkeypatch):     # Review Focus 4
    import asyncio
    from smartass.epm import oauth
    from smartass.epm import drillstep as server
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")

    def not_signed_in(*a):
        raise oauth.NotSignedIn("not signed in to https://plan1.example: run `smartass login`")
    monkeypatch.setattr(server, "_http_link", not_signed_in)

    async def no_excel(fn, what="call"):
        raise AssertionError("Excel was used")
    monkeypatch.setattr(server, "_call", no_excel)
    with pytest.raises(oauth.NotSignedIn):
        asyncio.run(server._link("FinLite", _CELL))


def test_split_values_without_a_sign_in_are_raised_not_sent_to_excel(monkeypatch):
    import asyncio
    from smartass.epm import oauth, svhttp
    from smartass.epm import drillstep as server
    monkeypatch.delenv("SV_TEST_NO_HTTP_DRILL")

    class NoSignIn:
        def child_values(self, *a, **k):
            raise oauth.NotSignedIn("not signed in")
    monkeypatch.setattr(svhttp, "shared", NoSignIn)

    async def no_excel(fn, what="call"):
        raise AssertionError("Excel was used")
    monkeypatch.setattr(server, "_call", no_excel)
    with pytest.raises(oauth.NotSignedIn):
        asyncio.run(server._child_values("FinLite", [_CELL, {**_CELL, "All Time Periods": "Aug 26"}]))


# ------------------------------------------------------------------ split siblings from one EPM query (step 1b)
def _derive_setup(monkeypatch, members, cubes, lines_for, template_query=True):
    """Children of un.SY by entity; members[] = atlas lists; cubes[] = child values; lines_for(entity) = the
    AMOUNTX values a built query returns. Records which children were drilled the usual way."""
    import json as _json
    from smartass.core import config
    from smartass.epm import drill_derive, drill_token
    from smartass.epm import drillstep as server
    from smartass.epm.drillcore import SuiQuery
    monkeypatch.setattr(config, "DRILL_DERIVE", True)
    monkeypatch.setattr(drill_derive, "members", lambda conn, m, db=None: members.get(m))
    usual = []

    def body_of(ents):
        q = f"SELECT 1 WHERE Entityx IN ({','.join('?' * len(ents))}) AND Accountx = ?"
        return _json.dumps({"QUERY": q, "PARAMS": list(ents) + ["ac.C60090"],
                            "PARAMSTYPE": ["String"] * (len(ents) + 1)})

    async def piece(conn, cell, max_lines, split, state, depth=0):
        e = cell["All Entities"]
        usual.append(e)
        q = SuiQuery(body_of(members.get(e, [e])), 52, 39) if template_query else None
        return [{"cell": cell, "cube_value": cubes[e], "drill_url": "u", "items": [{"AMOUNTX": str(cubes[e])}],
                 "view": {"v": 1}, "secs": 0, "query": q}]
    monkeypatch.setattr(server, "_drill_piece", piece)

    class Token:
        async def lines(self, q, warm_url, max_lines, progress=None):
            ents = _json.loads(q.body)["PARAMS"][:-1]
            return {"items": [{"AMOUNTX": str(a)} for a in lines_for(ents[0])], "total": None, "view": {"v": 1},
                    "secs": 0, "query": q}
    monkeypatch.setattr(drill_token, "shared", lambda: Token())
    return server, usual


def test_split_asks_epm_once_and_builds_the_siblings(monkeypatch):
    import asyncio
    members = {"un.SY.2TEC": ["un.SY.2TEC", "a", "b"], "un.SY.2BUS": ["un.SY.2BUS", "c"],
               "un.SY.2EI": ["un.SY.2EI"], "un.SY.2SSM": ["un.SY.2SSM", "d"]}
    cubes = {"un.SY.2TEC": -10.0, "un.SY.2BUS": -5.0, "un.SY.2EI": -2.0, "un.SY.2SSM": -7.0, "un.SY.2WHS": None}
    lines = {"un.SY.2TEC": [-4, -6], "un.SY.2BUS": [-5], "un.SY.2SSM": [-1]}      # 2SSM does not tie: ask EPM
    server, usual = _derive_setup(monkeypatch, members, cubes, lambda e: lines.get(e, []))
    parts = [{"All Accounts": "ac.C60090", "All Entities": e, "All Time Periods": "Sep 26"} for e in cubes]
    state = {"left": 1000, "cube": None}
    out = asyncio.run(server._children_derived("FinLite", parts, [cubes[e] for e in cubes], 1000, "auto", state, 0))
    assert usual == ["un.SY.2EI", "un.SY.2SSM"]           # template = fewest members; then the one that did not tie
    assert [ps[0]["cell"]["All Entities"] for ps in out] == list(cubes)                   # plan order kept
    assert out[0][0]["derived"] and out[1][0]["derived"] and "derived" not in out[3][0]
    assert out[4][0]["note"] == "empty cell (skipped)"
    assert state["left"] == 1000 - 3                      # only the built pieces' lines here (_drill_piece is faked)


def test_split_without_a_template_drills_every_child_the_usual_way(monkeypatch):
    import asyncio
    members = {"un.SY.2TEC": ["un.SY.2TEC"], "un.SY.2BUS": ["un.SY.2BUS"], "un.SY.2EI": ["un.SY.2EI"]}
    cubes = {"un.SY.2TEC": -1.0, "un.SY.2BUS": -1.0, "un.SY.2EI": -1.0}
    server, usual = _derive_setup(monkeypatch, members, cubes, lambda e: [-1], template_query=False)
    parts = [{"All Accounts": "ac.C60090", "All Entities": e, "All Time Periods": "Sep 26"} for e in cubes]
    asyncio.run(server._children_derived("FinLite", parts, [-1.0] * 3, 1000, "auto", {"left": 1000}, 0))
    assert sorted(usual) == sorted(cubes)                 # 3 template tries (no query), nothing built


def test_period_splits_are_not_derived(monkeypatch):
    import asyncio
    server, usual = _derive_setup(monkeypatch, {}, {"un.SY": -1.0}, lambda e: [-1])
    parts = [{"All Accounts": "ac.C60090", "All Entities": "un.SY", "All Time Periods": p} for p in ("Jul 26", "Aug 26")]
    asyncio.run(server._children_derived("FinLite", parts, [-1.0, -1.0], 1000, "auto", {"left": 1000}, 0))
    assert usual == ["un.SY", "un.SY"]


def test_a_built_piece_without_lines_or_with_an_error_asks_epm(monkeypatch):
    import asyncio
    members = {"un.SY.2TEC": ["un.SY.2TEC", "a"], "un.SY.2BUS": ["un.SY.2BUS"], "un.SY.2EI": ["un.SY.2EI", "b", "c"]}
    cubes = {"un.SY.2TEC": 0.0, "un.SY.2BUS": -1.0, "un.SY.2EI": -3.0}

    def lines_for(e):
        if e == "un.SY.2EI":
            raise RuntimeError("connection reset")             # any error, not only a DrillError
        return []                                          # 2TEC: no lines "tie" to its zero cube - not trusted
    server, usual = _derive_setup(monkeypatch, members, cubes, lines_for)
    parts = [{"All Accounts": "ac.C60090", "All Entities": e, "All Time Periods": "Sep 26"} for e in cubes]
    out = asyncio.run(server._children_derived("FinLite", parts, list(cubes.values()), 1000, "auto",
                                               {"left": 1000}, 0))
    assert usual[0] == "un.SY.2BUS" and sorted(usual) == sorted(cubes)   # template (fewest members), then both
    assert all("derived" not in ps[0] for ps in out)       # every child came from EPM's own query


@pytest.mark.parametrize("derive", [True, False])
def test_split_siblings_are_built_from_one_epm_query_only_with_drill_derive(monkeypatch, derive):
    """The derive gate on the (only) token route: _children builds the siblings' queries only when
    config.DRILL_DERIVE is on; off, every child with a value is drilled the usual way."""
    import asyncio

    from smartass.core import config
    members = {"un.SY.2TEC": ["un.SY.2TEC", "a", "b"], "un.SY.2BUS": ["un.SY.2BUS", "c"], "un.SY.2EI": ["un.SY.2EI"]}
    cubes = {"un.SY.2TEC": -10.0, "un.SY.2BUS": -5.0, "un.SY.2EI": -2.0, "un.SY.2WHS": None}
    lines = {"un.SY.2TEC": [-4, -6], "un.SY.2BUS": [-5]}
    server, usual = _derive_setup(monkeypatch, members, cubes, lambda e: lines.get(e, []))
    monkeypatch.setattr(config, "DRILL_DERIVE", derive)

    async def values(conn, parts):
        return [cubes[p["All Entities"]] for p in parts]
    monkeypatch.setattr(server, "_child_values", values)
    parts = [{"All Accounts": "ac.C60090", "All Entities": e, "All Time Periods": "Sep 26"} for e in cubes]
    out = asyncio.run(server._children("FinLite", parts, 1000, "auto", {"left": 1000, "cube": None}, 0))
    assert [p["cell"]["All Entities"] for p in out] == list(cubes)                         # plan order kept
    assert [bool(p.get("derived")) for p in out] == ([True, True, False, False] if derive else [False] * 4)
    assert usual == (["un.SY.2EI"] if derive else ["un.SY.2TEC", "un.SY.2BUS", "un.SY.2EI"])
    assert out[3]["note"] == "empty cell (skipped)"


def test_the_lines_come_over_the_token_route(monkeypatch):
    import asyncio

    from smartass.epm import drill_token, drillstep

    class Fake:
        async def fetch(self, url, max_lines, progress=None):
            return {"items": [{"AMOUNTX": "1"}], "total": 1, "view": {"items": []}, "secs": 0.1, "query": None}
    monkeypatch.setattr(drill_token, "shared", lambda: Fake())
    assert asyncio.run(drillstep._drill_fetch("u", 10))["total"] == 1


def test_an_error_on_the_token_route_is_raised_as_it_is(monkeypatch):
    import asyncio

    from smartass.epm import drill_token, drillcore, drillstep

    class Fails:
        async def fetch(self, url, max_lines, progress=None):
            raise drillcore.DrillError("getData returned HTTP 500 with an error at offset 0.")
    monkeypatch.setattr(drill_token, "shared", lambda: Fails())
    with pytest.raises(drillcore.DrillError, match="offset 0"):
        asyncio.run(drillstep._drill_fetch("u", 10))


def test_the_token_route_gives_an_empty_piece_for_no_drill_data(monkeypatch):
    import asyncio

    from smartass.epm import drill_token, drillcore, drillstep

    class Empty:
        async def fetch(self, url, max_lines, progress=None):
            raise drillcore.NoDrillData("no drill data here")
    monkeypatch.setattr(drill_token, "shared", lambda: Empty())
    res = asyncio.run(drillstep._drill_fetch("u", 10))
    assert res["items"] == [] and res["note"] == "no drill data here"


# ---------------------------------------------------------------- a piece's time budget (final review I2)
def test_a_pieces_budget_grows_with_its_pages_from_a_floor_to_a_cap(monkeypatch):
    from smartass.core import config
    from smartass.epm import drillstep
    monkeypatch.setattr(config, "DRILL_PAGE", 5000)
    monkeypatch.setattr(config, "DRILL_MAX_LINES", 300_000)
    assert drillstep.RENEWAL_S == 360                       # the lost request, the drain, the key request: 120 s each
    assert drillstep.piece_budget(None) == 660               # until the first page tells the total: 300 s + 360 s
    assert drillstep.piece_budget(1) == 660                  # max(300 s, 1 page * 15 s) + 360 s
    assert drillstep.piece_budget(150_000) == 810            # 30 pages: where a flat 300 s ran out (about 11 s a page)
    assert drillstep.piece_budget(300_000) == drillstep.piece_cap() == 1260     # 60 pages: the line limit needs more
    assert drillstep.piece_budget(10_000_000) == 1260
    monkeypatch.setattr(config, "DRILL_MAX_LINES", 50_000)
    assert drillstep.piece_cap() == 900 and drillstep.piece_budget(300_000) == 900   # the old page route's cap


def _slow_source(monkeypatch, pages: list, total, view=None):
    """drill_token.shared() -> a fetch that tells progress after each page: pages = seconds per page."""
    import asyncio

    from smartass.epm import drill_token

    class Slow:
        async def fetch(self, url, max_lines, progress=None):
            n = 0
            for secs in pages:
                await asyncio.sleep(secs)
                n += 5000
                if progress:
                    progress(n, total)
            return {"items": [{"AMOUNTX": "1"}] * 3, "total": 3, "view": view or {"items": []}, "secs": 0,
                    "query": None}
    monkeypatch.setattr(drill_token, "shared", lambda: Slow())


def _small_budget(monkeypatch):
    from smartass.core import config
    from smartass.epm import drillstep
    monkeypatch.setattr(config, "DRILL_PAGE", 5000)
    monkeypatch.setattr(drillstep, "PIECE_FLOOR_S", 0.3)
    monkeypatch.setattr(drillstep, "PAGE_S", 0.2)
    monkeypatch.setattr(drillstep, "RENEWAL_S", 0.0)


def test_a_big_piece_gets_the_time_its_pages_need_once_its_total_is_known(monkeypatch):
    import asyncio

    from smartass.epm import drillstep
    _small_budget(monkeypatch)
    _slow_source(monkeypatch, [0.05, 0.2, 0.2, 0.2], 20_000)         # 0.65 s: over the floor, within 4 * 0.2 s
    assert asyncio.run(drillstep._drill_fetch("u", 100_000))["total"] == 3


def test_a_piece_over_its_budget_is_a_readable_error_that_names_the_cell(monkeypatch):
    import asyncio

    from mcp.server.mcpserver.exceptions import ToolError

    from smartass.core import serverkit
    from smartass.epm import drillstep
    _small_budget(monkeypatch)
    _slow_source(monkeypatch, [0.05, 0.6], 10_000)                   # 2 pages: 0.4 s, the second one takes longer

    async def link(connection, cell):
        return {"url": "u", "cube_value": -5.0}
    monkeypatch.setattr(drillstep, "_link", link)

    async def fin_drill(cell=None):
        return await drillstep.drill(cell, split="off")
    with pytest.raises(ToolError) as e:
        asyncio.run(serverkit.guard(fin_drill)(cell={"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC",
                                                     "All Time Periods": "Sep 26"}))
    msg = str(e.value)
    assert "TimeoutError" not in msg and "ac.C60 × un.SY.2TEC × Sep 26" in msg
    assert "5,000 of 10,000 lines" in msg and "Narrow the cell" in msg and "Retry" not in msg


def test_a_timeout_inside_the_fetch_is_not_taken_for_the_budget(monkeypatch):
    import asyncio

    from smartass.epm import drill_token, drillstep

    class Raises:
        async def fetch(self, url, max_lines, progress=None):
            raise TimeoutError("EPM's answer did not end in time")
    monkeypatch.setattr(drill_token, "shared", lambda: Raises())
    with pytest.raises(TimeoutError, match="did not end in time"):
        asyncio.run(drillstep._drill_fetch("u", 10))

# ---------------------------------------------------------------- a sign of life during a drill (final review M1)
def test_drill_notes_say_lines_and_pieces_at_most_once_per_ten_seconds():
    from smartass.epm import drillstep
    now, said = [0.0], []
    notes = drillstep._Notes(said.append, clock=lambda: now[0])
    for t, event in [(1, ("page", "a", 5000, 20000)), (10.5, ("page", "a", 10000, 20000)), (12, ("page", "b", 700, 700)),
                     (15, ("done",)), (21, ("page", "a", 15000, 20000)), (25, ("done",)), (31.5, ("done",))]:
        now[0] = t
        notes.page(*event[1:]) if event[0] == "page" else notes.piece_done()
    assert said == ["still working: drill (10,000 of 20,000 lines so far, 0 pieces done) (10s)",
                    "still working: drill (15,700 of 20,700 lines so far, 1 piece done) (21s)",
                    "still working: drill (15,700 of 20,700 lines so far, 3 pieces done) (32s)"]


def test_a_drill_tells_its_progress_with_counts_only(monkeypatch):
    import asyncio

    from smartass.core import timing
    from smartass.epm import drillstep
    monkeypatch.setattr(timing, "PROGRESS_AFTER_S", 0)              # every page and piece may speak
    _slow_source(monkeypatch, [0.0, 0.0], 10_000, json.loads((FIX / "view_definition.json").read_text(encoding="utf-8")))

    async def link(connection, cell):
        return {"url": "u", "cube_value": -1234.5}
    monkeypatch.setattr(drillstep, "_link", link)
    said = []
    asyncio.run(drillstep.drill({"All Accounts": "ac.C60", "All Entities": "un.SY.2TEC", "All Time Periods": "Sep 26"},
                                say=said.append, split="off"))
    assert [s.split(" (", 1)[1].rsplit(") (", 1)[0] for s in said] == [
        "5,000 of 10,000 lines so far, 0 pieces done", "10,000 of 10,000 lines so far, 0 pieces done",
        "10,000 of 10,000 lines so far, 1 piece done"]
    assert all(s.startswith("still working: drill (") and "1234" not in s and "1,234" not in s for s in said)

def test_the_unused_drill_settings_and_the_uncalled_drill_all_are_gone():
    from smartass.core import config
    from smartass.epm import drillstep
    assert not hasattr(config, "DRILL_DB") and not hasattr(config, "DRILL_KEEP_DAYS")
    assert not hasattr(drillstep, "drill_all") and not hasattr(drillstep, "combine")

def test_a_renewal_before_the_first_page_has_its_allowance_too(monkeypatch):
    import asyncio

    from smartass.core import config
    from smartass.epm import drillstep
    monkeypatch.setattr(config, "DRILL_PAGE", 5000)
    monkeypatch.setattr(drillstep, "PIECE_FLOOR_S", 0.2)
    monkeypatch.setattr(drillstep, "PAGE_S", 0.05)
    monkeypatch.setattr(drillstep, "RENEWAL_S", 0.3)
    _slow_source(monkeypatch, [0.35], 5000)          # a lost session on the drill query: the first page after the floor
    assert asyncio.run(drillstep._drill_fetch("u", 100_000))["total"] == 3
