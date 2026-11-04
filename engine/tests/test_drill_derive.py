"""drill_derive: a sibling's drill query from EPM's query for one cell (shapes from the live probe 2026-11-03)."""
import json

from smartview_mcp import drill_derive as dv

Q = ("SELECT ENTITY, AMOUNTX FROM TDATASEG abdv WHERE ( tpovpartition.partitionkey in ( ?,? ) ) AND abdv.catKey IN"
     " ( ? ) AND abdv.periodKey IN ( ? ) AND Entityx IN (?,?,? ) AND Accountx IN (?,? ) AND Ud1x = ?")
PARAMS = ["39", "39", "57", "2026-09-30", "un.SY.2TEC", "un.SY.A", "un.SY.B", "ac.C60090", "ac.67100", "gl.x"]
TYPES = ["BigDecimal", "BigDecimal", "BigDecimal", "Date", "String", "String", "String", "String", "String", "String"]
BODY = json.dumps({"QUERY": Q, "PARAMS": PARAMS, "PARAMSTYPE": TYPES})
TEC = ["un.SY.2TEC", "un.SY.B", "un.SY.A"]                       # the atlas order may differ from EPM's


def test_entity_list_is_swapped_and_the_rest_kept():
    out = json.loads(dv.derive(BODY, "All Entities", TEC, ["un.SY.2BUS", "un.SY.C"]))
    assert "Entityx IN (?,?) AND Accountx IN (?,? )" in out["QUERY"]
    assert out["PARAMS"] == ["39", "39", "57", "2026-09-30", "un.SY.2BUS", "un.SY.C", "ac.C60090", "ac.67100", "gl.x"]
    assert out["PARAMSTYPE"] == TYPES[:4] + ["String"] * 2 + TYPES[7:]
    assert out["QUERY"].count("?") == len(out["PARAMS"])


def test_account_list_and_a_single_member():
    out = json.loads(dv.derive(BODY, "All Accounts", ["ac.67100", "ac.C60090"], ["ac.C60080"]))
    assert "Accountx IN (?) AND Ud1x" in out["QUERY"] and out["PARAMS"][7] == "ac.C60080"


def test_an_equals_clause_is_widened():
    body = json.dumps({"QUERY": "SELECT 1 WHERE Entityx = ? AND x = ?", "PARAMS": ["un.SY.A", "1"],
                       "PARAMSTYPE": ["String", "BigDecimal"]})
    out = json.loads(dv.derive(body, "All Entities", ["un.SY.A"], ["un.SY.B", "un.SY.C"]))
    assert out["QUERY"] == "SELECT 1 WHERE Entityx IN (?,?) AND x = ?" and out["PARAMS"] == ["un.SY.B", "un.SY.C", "1"]


def test_guards_fall_back_to_epm():
    assert dv.derive(BODY, "All Entities", TEC[:2], ["x"]) is None              # atlas != EPM's list
    assert dv.derive(BODY, "All Entities", ["un.SY.2TEC", "un.SY.A", "un.SY.X"], ["x"]) is None
    assert dv.derive(BODY, "All Time Periods", TEC, ["x"]) is None             # periods are not derived
    assert dv.derive(BODY, "All Entities", TEC, []) is None
    assert dv.derive(BODY, "All Entities", TEC, ["m"] * (dv.MAX_MEMBERS + 1)) is None
    assert dv.derive("not json", "All Entities", TEC, ["x"]) is None
    short = json.dumps({"QUERY": Q, "PARAMS": PARAMS[:-1], "PARAMSTYPE": TYPES[:-1]})
    assert dv.derive(short, "All Entities", TEC, ["x"]) is None                 # placeholders != params
    twice = json.dumps({"QUERY": Q + " OR Entityx IN (?)", "PARAMS": PARAMS + ["y"], "PARAMSTYPE": TYPES + ["String"]})
    assert dv.derive(twice, "All Entities", TEC, ["x"]) is None                 # two clauses: ambiguous
    typed = json.dumps({"QUERY": Q, "PARAMS": PARAMS, "PARAMSTYPE": TYPES[:4] + ["BigDecimal"] + TYPES[5:]})
    assert dv.derive(typed, "All Entities", TEC, ["x"]) is None


def test_members_are_the_member_and_its_descendants(tmp_path, monkeypatch):
    from smartass.knowledge import atlas
    monkeypatch.setattr(atlas, "descendants", lambda db, c, m, limit=5000: [{"name": "un.SY.A"}, {"name": "un.SY.B"}])

    class Db:
        def close(self):
            pass
    assert dv.members("FinLite", "un.SY.2TEC", Db()) == ["un.SY.2TEC", "un.SY.A", "un.SY.B"]
    monkeypatch.setattr(atlas, "descendants", lambda db, c, m, limit=5000: [{"name": f"m{i}"} for i in range(limit)])
    assert dv.members("FinLite", "un.SY", Db()) is None                         # too many: EPM refuses it anyway
