"""epm/rest.py: one grid over Planning REST exportdataslice, returned as the provider route's Fetched. EPM is faked:
an opener under a real EpmSession that answers the plan type's dimension list and exportdataslice in Oracle's
documented shapes (confirmed live 2026-11-21 by key paths and types only). Every member, label and figure here is
invented; the dimension names are the cubes' (as in svxml_util)."""
import http.client
import io
import json
import urllib.error
from itertools import product
from urllib.parse import urlsplit

import pytest

from smartass.epm import epmhttp, oauth, rest, svhttp
from smartass.epm.grid import Axis
from smartass.knowledge import atlas
from smartass.xl.host import InputError

POD = "https://plan1.example"

# (connection, dimension, [(name, label), ...]): the atlas as Smart View's zoom stored it. A member whose display has
# no "code alias (prefix)" form is stored under its label (Sep 26, the FinPlan members): REST answers its real name.
ATLAS = [
    ("FinMax", "All Accounts", [("All Accounts", "All Accounts"), ("ac.C30", "C30 REVENUES (ac)"),
                                ("ac.C70", "C70 COSTS (ac)"), ("ac.C90", "C90 OTHER (ac)")]),
    ("FinMax", "All Balance Types", [("All Balance Types", "All Balance Types"), ("Periodic", "Periodic")]),
    ("FinMax", "All Currencies", [("All Currencies", "All Currencies"), ("NOK", "NOK")]),
    ("FinMax", "All Entities", [("All Entities", "All Entities"), ("un.SY.E1", "E1 Entity one (un.SY)"),
                                ("un.SY.E2", "E2 Entity two (un.SY)")]),
    ("FinMax", "All Scenarios", [("All Scenarios", "All Scenarios"), ("Actual", "Actual")]),
    ("FinMax", "All Time Periods", [("All Time Periods", "All Time Periods"), ("Aug 26", "Aug 26"), ("Sep 26", "Sep 26")]),
    ("FinPlan", "Account", [("Account", "Account"), ("EBX Ebitda X", "EBX Ebitda X")]),
    ("FinPlan", "All Balance Types", [("All Balance Types", "All Balance Types"), ("Periodic", "Periodic")]),
    ("FinPlan", "Entity", [("Entity", "Entity"), ("1TEC Tech", "1TEC Tech"), ("2BUS Biz", "2BUS Biz")]),
    ("FinPlan", "Period", [("Period", "Period"), ("Dec", "Dec"), ("Jan", "Jan")]),
    ("FinPlan", "Scenario", [("Scenario", "Scenario"), ("Target", "Target")]),
    ("FinPlan", "Version", [("Version", "Version"), ("Final", "Final")]),
    ("FinPlan", "Years", [("Years", "Years"), ("FY25", "FY25"), ("FY26", "FY26")]),
]
REAL = {"Sep 26": "Y26M09", "Aug 26": "Y26M08", "EBX Ebitda X": "EBX", "1TEC Tech": "1TEC", "2BUS Biz": "2BUS",
        "C30 REVENUES (ac)": "ac.C30"}
# The plan types' dimension lists as GET .../plantypes/{cube}/dimensions?aliasTableName=Default answers them (live:
# FinMax names without "All " and spaces, the atlas's names as their Default alias; FinPlan's mostly without alias)
LISTED = {
    "FinMax": [("Accounts", "All Accounts"), ("BalanceTypes", "All Balance Types"), ("Currencies", "All Currencies"),
               ("Entities", "All Entities"), ("Scenarios", "All Scenarios"), ("TimePeriods", "All Time Periods")],
    "FinLite": [("Accounts", "All Accounts"), ("BalanceTypes", "All Balance Types"), ("Currencies", "All Currencies"),
                ("Entities", "All Entities"), ("Scenarios", "All Scenarios"), ("TimePeriods", "All Time Periods")],
    "FinPlan": [("Account", None), ("Period", None), ("Years", None), ("Scenario", None), ("Version", None),
                ("Entity", None), ("BalanceTypes", "All Balance Types")],
}


@pytest.fixture
def db(tmp_path):
    d = atlas.connect(tmp_path / "atlas.db")
    for conn, dim, members in ATLAS:
        atlas.store(d, conn, dim, [{"name": n, "label": lab, "parent": None if i == 0 else dim, "depth": int(i > 0),
                                    "order": i} for i, (n, lab) in enumerate(members)], 1.0)
    yield d
    d.close()


@pytest.fixture(autouse=True)
def _no_cached_lists(monkeypatch):
    monkeypatch.setattr(rest, "_listed", {})


class _Tokens:
    def __init__(self, *tokens):
        self.tokens, self.asked = list(tokens or ["at-1"]), []

    def scope_for(self, base):
        return "urn:scope"

    def access_token(self, base, reject=None):
        self.asked.append(reject)
        t = self.tokens.pop(0) if len(self.tokens) > 1 else self.tokens[0]
        if isinstance(t, Exception):
            raise t
        return t


class _Answer(io.BytesIO):
    def __init__(self, status, text):
        super().__init__(text.encode("utf-8"))
        self.status, self.headers = status, http.client.HTTPMessage()

    def info(self):
        return self.headers


def value(row, col, pov):
    """Invented figures by member: ac.C30 has numbers, ac.C70 is empty, ac.C90 has text forms of numbers."""
    a, *rest_ = row
    if a == "ac.C70":
        return ""
    if a == "ac.C90":
        return {"Y26M08": "1234.5", "Y26M09": " "}[col[0]]
    return float(len("".join(row)) * 10 + (1 if col[0].endswith("9") else 0)) + 0.25


class _Epm:
    """Planning REST as the session's opener sees it: the dimension list of each plan type, and exportdataslice -
    the requested cross product in order, each sent member answered under its real name (REAL), cells from
    cell(row names, column names, pov names). answer(path, body) may replace any answer: (status, text)."""

    def __init__(self, cell=value, listed=LISTED, answer=None):
        self.cell, self.listed, self.answer, self.sent, self.bodies = cell, listed, answer, [], []

    def open(self, req, timeout):
        path = urlsplit(req.full_url).path
        self.sent.append((req.get_method(), req.full_url))
        body = json.loads(req.data) if req.data else None
        self.bodies.append(body)
        got = self.answer(path, body) if self.answer else None
        if got is None:
            got = self._dimensions(path) if path.endswith("/dimensions") else self._slice(body)
        status, text = got
        if isinstance(status, Exception):
            raise status
        if 200 <= status < 300:
            return _Answer(status, text)
        raise urllib.error.HTTPError(req.full_url, status, "no", http.client.HTTPMessage(), io.BytesIO(text.encode()))

    def _dimensions(self, path):
        cube = path.split("/")[-2]
        items = [{"name": n, "alias": a, "dimType": "Custom", "id": "x"} for n, a in self.listed[cube]]
        return 200, json.dumps({"items": items, "totalResults": len(items), "hasMore": False, "links": []})

    def _slice(self, body):
        g = body["gridDefinition"]
        name = lambda m: REAL.get(m, m)                                     # noqa: E731
        pov = [name(m[0]) for m in g["pov"]["members"]]
        (cseg,), rows = g["columns"], []
        cols = list(product(*[[name(m) for m in ms] for ms in cseg["members"]]))
        for seg in g["rows"]:
            for combo in product(*[[name(m) for m in ms] for ms in seg["members"]]):
                rows.append({"headers": list(combo), "data": [self.cell(combo, c, pov) for c in cols]})
        return 200, json.dumps({"pov": pov, "columns": [list(x) for x in zip(*cols)], "rows": rows})


def _session(epm, *tokens):
    return epmhttp.EpmSession(POD, tokens=_Tokens(*tokens), opener=epm)


def _grid(epm, db, accounts=("ac.C30", "ac.C70"), entities=("un.SY.E1",), periods=("Aug 26", "Sep 26"), pov=None,
          conn="FinMax", s=None):
    pov = {"All Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic"} if pov is None else pov
    return rest.fetch_grid(s or _session(epm), conn, pov, [Axis("All Accounts", list(accounts)),
                                                           Axis("All Entities", list(entities))],
                           [Axis("All Time Periods", list(periods))], db)


def _slices(epm):
    return [b for (m, u), b in zip(epm.sent, epm.bodies) if u.endswith("/exportdataslice")]


# ---------------------------------------------------------------- the grid, as the provider route returns it
def test_a_grid_comes_back_in_the_provider_routes_shape_with_its_labels(db):
    """The provider's Fetched from A1: the POV row, the column headers, the row labels (repeated) and the cells -
    labels as Smart View shows them (the atlas's), None for an empty cell (as the provider's), floats for numbers."""
    f = _grid(_Epm(), db, accounts=("ac.C30", "ac.C70", "ac.C90"))
    assert f.route == "rest" and f.warnings == [] and f.seconds >= 0
    assert [(d, m) for d, m in f.layout.pov] == [("All Balance Types", "Periodic"), ("All Currencies", "NOK"),
                                                   ("All Scenarios", "Actual")]
    c30 = [value(("ac.C30", "un.SY.E1"), (p,), None) for p in ("Y26M08", "Y26M09")]
    assert f.values == [
        [None, None, "Periodic", "NOK", "Actual"],
        [None, None, "Aug 26", "Sep 26", None],
        ["C30 REVENUES (ac)", "E1 Entity one (un.SY)", *c30, None],
        ["C70 COSTS (ac)", "E1 Entity one (un.SY)", None, None, None],
        ["C90 OTHER (ac)", "E1 Entity one (un.SY)", 1234.5, None, None]]


def test_the_request_is_oracles_grid_definition_with_rests_dimension_names(db):
    """No suppression (the provider route suppresses nothing either: grid.parse does), no planning data; the
    dimensions as the plan type lists them; a POV left out is the dimension's top member, sent as REST names it."""
    epm = _Epm()
    _grid(epm, db, pov={"All Scenarios": "Actual"})
    assert epm.sent[0] == ("GET", POD + "/HyperionPlanning/rest/v3/applications/FinMax/plantypes/FinMax/dimensions"
                                        "?aliasTableName=Default")
    assert epm.sent[1] == ("POST", POD + "/HyperionPlanning/rest/v3/applications/FinMax/plantypes/FinMax/exportdataslice")
    assert epm.bodies[1] == {"exportPlanningData": False, "gridDefinition": {
        "suppressMissingBlocks": False, "suppressMissingRows": False, "suppressMissingColumns": False,
        "pov": {"dimensions": ["BalanceTypes", "Currencies", "Scenarios"],
                "members": [["BalanceTypes"], ["Currencies"], ["Actual"]]},
        "columns": [{"dimensions": ["TimePeriods"], "members": [["Aug 26", "Sep 26"]]}],
        "rows": [{"dimensions": ["Accounts", "Entities"], "members": [["ac.C30", "ac.C70"], ["un.SY.E1"]]}]}}


def test_the_top_member_rest_answers_under_its_own_name_shows_the_dimensions_label(db):
    f = _grid(_Epm(), db, pov={"All Scenarios": "Actual"})
    assert f.values[0][2:] == ["All Balance Types", "All Currencies", "Actual"] and f.warnings == []


def test_finlite_is_its_own_plan_type_in_the_finmax_application(db):
    for conn, dim, members in ATLAS[:6]:
        atlas.store(db, "FinLite", dim, [{"name": n, "label": lab, "parent": None, "depth": 0, "order": i}
                                         for i, (n, lab) in enumerate(members)], 1.0)
    epm = _Epm()
    f = _grid(epm, db, conn="FinLite")
    assert [u for _, u in epm.sent] == [
        POD + "/HyperionPlanning/rest/v3/applications/FinMax/plantypes/FinLite/dimensions?aliasTableName=Default",
        POD + "/HyperionPlanning/rest/v3/applications/FinMax/plantypes/FinLite/exportdataslice"]
    assert f.values[2][:2] == ["C30 REVENUES (ac)", "E1 Entity one (un.SY)"]


def test_the_plan_types_dimension_list_is_read_once_per_pod(db):
    epm = _Epm()
    s = _session(epm)
    _grid(epm, db, s=s)
    _grid(epm, db, s=s)
    assert [m for m, _ in epm.sent] == ["GET", "POST", "POST"]


def test_the_dimension_list_is_read_again_after_it_expires(db, monkeypatch):
    epm = _Epm()
    s = _session(epm)
    now = [1000.0]
    monkeypatch.setattr(rest, "_clock", lambda: now[0])
    _grid(epm, db, s=s)
    now[0] += rest.LIST_TTL_S - 1
    _grid(epm, db, s=s)
    assert sum(u.endswith("Default") for _, u in epm.sent) == 1
    now[0] += 2
    _grid(epm, db, s=s)
    assert sum(u.endswith("Default") for _, u in epm.sent) == 2


def test_a_list_that_no_longer_maps_is_read_again_once_not_kept(db):
    """A dimension renamed in EPM and the atlas rebuilt: the cached list no longer maps, so it is read again."""
    epm = _Epm()
    s = _session(epm)
    _grid(epm, db, s=s)
    with db:
        for table in ("member", "build"):
            db.execute(f"UPDATE {table} SET dimension='Plan Scenarios' WHERE dimension='All Scenarios'")
    epm.listed = {**LISTED, "FinMax": [(n, "Plan Scenarios" if a == "All Scenarios" else a)
                                       for n, a in LISTED["FinMax"]]}
    f = _grid(epm, db, s=s, pov={"Plan Scenarios": "Actual", "All Currencies": "NOK", "All Balance Types": "Periodic"})
    assert f.values[2][0] == "C30 REVENUES (ac)"
    assert sum(u.endswith("Default") for _, u in epm.sent) == 2


# ---------------------------------------------------------------- dimension names: from the plan type's own list
@pytest.mark.parametrize("conn, want", [
    ("FinMax", {"All Accounts": "Accounts", "All Balance Types": "BalanceTypes", "All Currencies": "Currencies",
                "All Entities": "Entities", "All Scenarios": "Scenarios", "All Time Periods": "TimePeriods"}),
    ("FinLite", {"All Accounts": "Accounts", "All Balance Types": "BalanceTypes", "All Currencies": "Currencies",
                 "All Entities": "Entities", "All Scenarios": "Scenarios", "All Time Periods": "TimePeriods"}),
    ("FinPlan", {"Account": "Account", "All Balance Types": "BalanceTypes", "Entity": "Entity", "Period": "Period",
                 "Scenario": "Scenario", "Version": "Version", "Years": "Years"}),
])
def test_each_applications_dimensions_map_by_their_listed_alias_or_name(conn, want):
    assert rest.map_dimensions(conn, list(want), LISTED[conn]) == want


def test_live_shaped_lists_map_every_atlas_dimension():
    """The lists as the pods answered them on 2026-11-21 (names and Default aliases, 14 / 10 / 8 dimensions)."""
    fm = [("BalanceTypes", "All Balance Types"), ("TimePeriods", "All Time Periods"), ("Accounts", "All Accounts"),
          ("Counterparts", "All Counterparts"), ("Currencies", "All Currencies"), ("Entities", "All Entities"),
          ("Functions", "All Functions"), ("Products", "All Products"), ("Projects", "All Projects"),
          ("Reserves", "All Reserves"), ("Scenarios", "All Scenarios"), ("Systems", "All Systems"),
          ("GLSources", "All GL Sources"), ("LegalEntities", "All Legal Entities")]
    got = rest.map_dimensions("FinMax", [a for _, a in fm], fm)
    assert got["All GL Sources"] == "GLSources" and got["All Legal Entities"] == "LegalEntities" and len(got) == 14
    fp = [("Account", None), ("Period", None), ("Years", None), ("Scenario", None), ("Version", None),
          ("Entity", None), ("Products", "All Products"), ("BalanceTypes", "All Balance Types")]
    got = rest.map_dimensions("FinPlan", ["Account", "All Balance Types", "All Products", "Entity", "Period",
                                          "Scenario", "Version", "Years"], fp)
    assert got["All Products"] == "Products" and got["Years"] == "Years"


def test_a_dimension_the_list_does_not_have_is_an_error_naming_it_never_a_guess():
    listed = [("Accounts", "All Accounts"), ("Time Periods", None)]          # 'All Time Periods' is no alias here
    with pytest.raises(rest.RestError, match="'All Time Periods'") as e:
        rest.map_dimensions("FinMax", ["All Accounts", "All Time Periods"], listed)
    assert "TimePeriods" not in str(e.value)
    with pytest.raises(rest.RestError, match="'All Accounts'"):                 # two hits: no choice
        rest.map_dimensions("FinMax", ["All Accounts"], [("Accounts", "All Accounts"), ("Accounts2", "All Accounts")])


def test_a_dimension_rest_lists_and_the_atlas_lacks_is_a_stale_atlas():
    """The provider route finds this as POV members Smart View appends; REST would just leave the dimension out."""
    with pytest.raises(rest.RestError, match="atlas is incomplete or stale for FinPlan.*Products"):
        rest.map_dimensions("FinPlan", ["Account"], [("Account", None), ("Products", "All Products")])
    assert rest.map_dimensions("FinPlan", ["Account"], [("Account", None), ("Colour", None, "Attribute Dimension")]) \
        == {"Account": "Account"}


# ---------------------------------------------------------------- members: names in the answer, labels from the atlas
def test_a_member_the_atlas_keeps_under_its_label_gets_that_label_back(db):
    """FinPlan: REST takes the label the atlas keeps as the member's name ('1TEC Tech') and answers the real name
    ('1TEC'), which the atlas does not know: the label is the atlas's for the member asked for at that place."""
    epm = _Epm(cell=lambda row, col, pov: 2.5)
    f = rest.fetch_grid(_session(epm), "FinPlan", {"Scenario": "Target", "Version": "Final", "Years": "FY26",
                                                   "All Balance Types": "Periodic"},
                        [Axis("Account", ["EBX Ebitda X"]), Axis("Entity", ["1TEC Tech", "2BUS Biz"])],
                        [Axis("Period", ["Dec", "Jan"])], db)
    assert f.values[0] == [None, None, "Periodic", "Target", "Final", "FY26"]
    assert f.values[2:] == [["EBX Ebitda X", "1TEC Tech", 2.5, 2.5, None, None],
                            ["EBX Ebitda X", "2BUS Biz", 2.5, 2.5, None, None]]
    assert f.warnings == [] and epm.sent[0][1].endswith("/applications/FinPlan/plantypes/FinPlan/dimensions"
                                                        "?aliasTableName=Default")


def test_the_atlas_is_read_by_name_without_a_pass_over_a_dimension_when_every_name_is_known(db):
    """A top member's REST name ('Currencies') and a real name answered as asked need no pass over the dimension
    (FinMax's Projects has some 140 000 members and is in every grid's POV)."""
    statements = []
    db.set_trace_callback(statements.append)
    _grid(_Epm(), db, entities=("un.SY.E1", "un.SY.E2"), periods=("Aug 26",), pov={"All Scenarios": "Actual"})
    db.set_trace_callback(None)
    scans = [s for s in statements if "SELECT name, label FROM member" in s and " IN (" not in s]
    assert len(scans) == 1 and "'All Time Periods'" in scans[0]                # 'Aug 26' answered as Y26M08 only


def test_a_member_kept_under_its_label_keeps_that_label_when_its_name_is_another_atlas_members(db):
    """Live 2026-11-22: REST answered a FinPlan account asked by its label under a name the atlas holds for a member
    with another label (the atlas derived it from that label); the provider labels the place with the asked one."""
    atlas.store(db, "FinPlan", "Account", [{"name": "Account", "label": "Account", "parent": None, "depth": 0,
                                            "order": 0},
                                           {"name": "EBX Ebitda X", "label": "EBX Ebitda X", "parent": "Account",
                                            "depth": 1, "order": 1},
                                           {"name": "EBX", "label": "EBX Percent (pc)", "parent": "Account",
                                            "depth": 1, "order": 2}], 1.0)
    f = rest.fetch_grid(_session(_Epm()), "FinPlan", {"Scenario": "Target", "Version": "Final", "Years": "FY26"},
                        [Axis("Account", ["EBX Ebitda X"]), Axis("Entity", ["1TEC Tech"])], [Axis("Period", ["Dec"])],
                        db)
    assert f.values[2][:2] == ["EBX Ebitda X", "1TEC Tech"] and f.warnings == []


def test_a_member_asked_by_its_label_is_answered_under_its_name(db):
    """REST takes the label (an alias) and answers the name: that is the member asked, as the provider route sees it."""
    f = _grid(_Epm(), db, accounts=("C30 REVENUES (ac)", "ac.C70"))
    assert [r[0] for r in f.values[2:]] == ["C30 REVENUES (ac)", "C70 COSTS (ac)"] and f.warnings == []
    with pytest.raises(rest.RestError, match="different layout.*answered as ac.C70"):
        _grid(_Epm(answer=_answer_with(rows=lambda rs: [{**r, "headers": ["ac.C70", r["headers"][1]]} for r in rs])),
              db, accounts=("C30 REVENUES (ac)",))


def test_a_member_the_atlas_does_not_know_keeps_its_name_and_adds_one_warning(db, monkeypatch):
    monkeypatch.setattr(rest, "MAX_CELLS", 4)                                 # the member in two requests
    f = _grid(_Epm(), db, accounts=("ac.C30",), entities=("un.SY.E1", "un.SY.E9"))
    assert [r[:2] for r in f.values[2:]] == [["C30 REVENUES (ac)", "E1 Entity one (un.SY)"], ["C30 REVENUES (ac)",
                                                                                            "un.SY.E9"]]
    assert f.warnings == ["FinMax: the atlas has no label for All Entities member un.SY.E9 (kept as its name)"]


# ---------------------------------------------------------------- numbers and missing cells
@pytest.mark.parametrize("cell, want", [
    ("", None), (" ", None), (None, None),       # live: REST's "" where the provider's Fetched holds None
    ("#Missing", "#Missing"), ("#missing", "#Missing"),
    ("#NoAccess", "#NoAccess"), ("#Invalid", "#Invalid"),
    ("1234.5", 1234.5), ("-0.1", -0.1), ("1e3", 1000.0), (" 42 ", 42.0), (7, 7.0), (0.30000000000000004, 0.30000000000000004),
    ("123456789.12345678", 123456789.12345678), ("0.1", 0.1),
])
def test_a_cell_is_read_as_the_provider_routes_fetched_holds_it(db, cell, want):
    f = _grid(_Epm(cell=lambda *a: cell), db, accounts=("ac.C30",), periods=("Sep 26",))
    got = f.values[2][2]
    assert got == want and type(got) is type(want)


@pytest.mark.parametrize("cell", ["Other", "1,234.5", "NaN", "Infinity", True, [1]])
def test_a_cell_that_is_not_a_number_or_a_label_fails_the_grid(db, cell):
    with pytest.raises(svhttp.SvHttpError, match="non-numeric data"):
        _grid(_Epm(cell=lambda *a: cell), db, accounts=("ac.C30",), periods=("Sep 26",))


# ---------------------------------------------------------------- the cell budget: split along rows, stitched
def test_a_grid_over_the_cell_budget_is_split_along_rows_and_stitched(db, monkeypatch):
    """Cells as EPM counts them (live 2026-11-22: 'Error Threshold: 500000, Current Value: 588028' for 42 000 rows of
    2 member headers and 12 cells): a row is its member headers and its data cells - here 2 + 2."""
    accounts, entities = ("ac.C30", "ac.C70", "ac.C90"), ("un.SY.E1", "un.SY.E2")
    whole = _grid(_Epm(), db, accounts=accounts, entities=entities)
    for cells, requests in ((8, 3), (4, 6), (24, 1), (16, 2)):
        monkeypatch.setattr(rest, "MAX_CELLS", cells)
        epm = _Epm()
        f = _grid(epm, db, accounts=accounts, entities=entities)
        sent = _slices(epm)
        assert f.values == whole.values and f.layout.row_tuples == whole.layout.row_tuples
        assert len(sent) == requests
        for b in sent:
            (seg,) = b["gridDefinition"]["rows"]
            assert len(list(product(*seg["members"]))) * 4 <= cells


def test_a_grid_of_exactly_the_cell_budget_is_one_request_and_one_row_more_is_two(db, monkeypatch):
    accounts, entities = ("ac.C30", "ac.C70", "ac.C90"), ("un.SY.E1", "un.SY.E2")
    monkeypatch.setattr(rest, "MAX_CELLS", 24)                     # 6 rows of 2 headers and 2 cells
    epm = _Epm()
    _grid(epm, db, accounts=accounts, entities=entities)
    assert len(_slices(epm)) == 1
    monkeypatch.setattr(rest, "MAX_CELLS", 23)
    epm = _Epm()
    _grid(epm, db, accounts=accounts, entities=entities)
    assert len(_slices(epm)) == 2


def test_a_split_grid_whose_middle_request_fails_is_no_grid_at_all(db, monkeypatch):
    monkeypatch.setattr(rest, "MAX_CELLS", 4)                      # a row is 2 headers and 2 cells
    n = []

    def answer(path, body):
        if path.endswith("exportdataslice"):
            n.append(1)
            if len(n) == 2:
                return 500, "{}"
    epm = _Epm(answer=answer)
    with pytest.raises(rest.RestError, match="HTTP 500"):
        _grid(epm, db, accounts=("ac.C30", "ac.C70", "ac.C90"))
    assert len(_slices(epm)) == 2                                   # the third is never asked


def test_a_request_body_over_the_size_cap_is_split_further_not_refused(db, monkeypatch):
    accounts, entities = ("ac.C30", "ac.C70", "ac.C90"), ("un.SY.E1", "un.SY.E2")
    whole = _grid(_Epm(), db, accounts=accounts, entities=entities)
    one = _Epm()
    _grid(one, db, accounts=accounts[:1], entities=entities)
    cap = len(json.dumps(_slices(one)[0]).encode()) + 5             # one account's request, not two
    monkeypatch.setattr(epmhttp, "MAX_BODY_BYTES", cap)
    epm = _Epm()
    f = _grid(epm, db, accounts=accounts, entities=entities)
    assert f.values == whole.values and len(_slices(epm)) == 3
    assert all(len(json.dumps(b).encode()) <= cap for b in _slices(epm))


def test_one_row_whose_request_is_over_the_size_cap_is_an_error_before_anything_is_sent(db, monkeypatch):
    monkeypatch.setattr(epmhttp, "MAX_BODY_BYTES", 100)
    epm = _Epm()
    with pytest.raises(rest.RestError, match="over the .* bytes a request takes"):
        _grid(epm, db, accounts=("ac.C30",), entities=("un.SY.E1",))
    assert _slices(epm) == []


def test_the_cell_budget_is_half_epms_threshold():
    assert rest.MAX_CELLS * 2 == rest.EPM_CELL_THRESHOLD == 500_000


def test_one_row_over_the_cell_budget_is_an_error_not_a_bigger_request(db, monkeypatch):
    monkeypatch.setattr(rest, "MAX_CELLS", 3)                      # a row is 4
    epm = _Epm()
    with pytest.raises(rest.RestError, match="cell budget"):
        _grid(epm, db)
    assert _slices(epm) == []


def test_split_requests_that_disagree_on_the_columns_fail_the_grid(db, monkeypatch):
    monkeypatch.setattr(rest, "MAX_CELLS", 4)
    n = []

    def answer(path, body):
        if path.endswith("exportdataslice"):
            n.append(1)
            if len(n) == 2:
                return 200, json.dumps({"pov": ["BalanceTypes", "NOK", "Actual"], "columns": [["Y26M07", "Y26M09"]],
                                        "rows": [{"headers": ["ac.C70", "un.SY.E1"], "data": ["1", "2"]}]})
    with pytest.raises(rest.RestError, match="answered different POV or column members"):
        _grid(_Epm(answer=answer), db, pov={"All Scenarios": "Actual", "All Currencies": "NOK"})


# ---------------------------------------------------------------- the layout check
def _answer_with(**over):
    def answer(path, body):
        if path.endswith("exportdataslice"):
            status, text = _Epm()._slice(body)
            j = json.loads(text)
            for k, f in over.items():
                j[k] = f(j[k])
            return status, json.dumps(j)
    return answer


@pytest.mark.parametrize("over, match", [
    ({"rows": lambda rs: rs[::-1]}, "different layout"),                                 # rows moved
    ({"rows": lambda rs: rs[:-1]}, "2 rows asked, 1"),                                  # a row went missing
    ({"rows": lambda rs: [{**r, "data": r["data"][:-1]} for r in rs]}, "cells"),          # a cell went missing
    ({"rows": lambda rs: [{**r, "headers": r["headers"][:1]} for r in rs]}, "headers"),
    ({"columns": lambda cs: [cs[0][:1]]}, "column"),
    ({"columns": lambda cs: cs + cs}, "column"),
    ({"pov": lambda p: p[:-1]}, "POV"),
    ({"rows": lambda rs: "x"}, "not a Planning REST grid"),
])
def test_an_answer_that_is_not_the_grid_asked_for_fails(db, over, match):
    with pytest.raises(svhttp.SvHttpError, match=match):
        _grid(_Epm(answer=_answer_with(**over)), db)


@pytest.mark.parametrize("pov, at", [({"All Scenarios": "Actual"}, 0),                     # a top member
                                     ({"All Scenarios": "Actual", "All Currencies": "NOK"}, 0)])
def test_a_pov_member_answered_as_another_is_caught_where_it_can_be(db, pov, at):
    """A top member is answered under the name it was sent as; a member the atlas keeps under its real name ('code
    alias (prefix)' label) under that very name. ('Actual', kept under its label, could be an alias: not checked.)"""
    def answer(path, body):
        if path.endswith("exportdataslice"):
            status, text = _Epm()._slice(body)
            j = json.loads(text)
            j["pov"][at] = "Bogus"
            return status, json.dumps(j)
    with pytest.raises(rest.RestError, match="different layout.*answered as Bogus"):
        _grid(_Epm(answer=answer), db, pov=pov)


def test_a_pov_member_the_atlas_keeps_under_its_real_name_answered_as_another_is_caught(db):
    def answer(path, body):
        if path.endswith("exportdataslice"):
            status, text = _Epm()._slice(body)
            j = json.loads(text)
            j["pov"] = ["un.SY.E2" if m == "un.SY.E1" else m for m in j["pov"]]
            return status, json.dumps(j)
    with pytest.raises(rest.RestError, match="different layout.*un.SY.E1 answered as un.SY.E2"):
        rest.fetch_grid(_session(_Epm(answer=answer)), "FinMax", {"All Entities": "un.SY.E1"},
                        [Axis("All Accounts", ["ac.C30"])], [Axis("All Time Periods", ["Sep 26"])], db)


def test_columns_moved_among_real_names_are_caught(db):
    def answer(path, body):
        if path.endswith("exportdataslice"):
            status, text = _Epm()._slice(body)
            j = json.loads(text)
            j["columns"] = [j["columns"][0][::-1]]
            return status, json.dumps(j)
    with pytest.raises(rest.RestError, match="different layout"):
        rest.fetch_grid(_session(_Epm(answer=answer)), "FinMax", {}, [Axis("All Accounts", ["ac.C30"])],
                        [Axis("All Entities", ["un.SY.E1", "un.SY.E2"])], db)


def test_rows_moved_among_members_the_atlas_keeps_under_their_labels_are_caught_where_they_can_be(db):
    """Two different members asked must come back as two different names, each always the same one."""
    def answer(path, body):
        if path.endswith("exportdataslice"):
            return 200, json.dumps({"pov": ["Target", "Final", "FY26", "BalanceTypes"], "columns": [["Dec"]],
                                    "rows": [{"headers": ["EBX", "1TEC"], "data": ["1"]},
                                             {"headers": ["EBX", "1TEC"], "data": ["2"]}]})
    with pytest.raises(rest.RestError, match="different layout"):
        rest.fetch_grid(_session(_Epm(answer=answer)), "FinPlan", {"Scenario": "Target", "Version": "Final",
                                                                   "Years": "FY26"},
                        [Axis("Account", ["EBX Ebitda X"]), Axis("Entity", ["1TEC Tech", "2BUS Biz"])],
                        [Axis("Period", ["Dec"])], db)


@pytest.mark.parametrize("accounts, periods", [(("ac.C30", "ac.C70", "ac.C30"), ("Sep 26",)),
                                               (("ac.C30",), ("Sep 26", "Aug 26", "sep 26"))])
def test_a_member_asked_twice_on_an_axis_is_refused_before_anything_is_sent(db, accounts, periods):
    """Live 2026-11-22: REST answers a member asked twice once (984 rows for 1 000 asked, 16 of them twice), where the
    provider repeats it; rest.py refuses such a grid before it is sent rather than answer another shape."""
    epm = _Epm()
    with pytest.raises(rest.RestError, match="asked twice"):
        _grid(epm, db, accounts=accounts, periods=periods)
    assert _slices(epm) == []


def test_more_than_one_column_dimension_is_refused_until_its_answer_shape_is_seen(db):
    epm = _Epm()
    with pytest.raises(rest.RestError, match="one column dimension"):
        rest.fetch_grid(_session(epm), "FinMax", {}, [Axis("All Accounts", ["ac.C30"])],
                        [Axis("All Time Periods", ["Sep 26"]), Axis("All Entities", ["un.SY.E1"])], db)
    assert _slices(epm) == []


def test_a_query_that_is_wrong_on_every_route_is_an_input_error(db):
    with pytest.raises(InputError, match="used more than once"):
        rest.fetch_grid(_session(_Epm()), "FinMax", {"All Accounts": "ac.C30"}, [Axis("All Accounts", ["ac.C30"])],
                        [Axis("All Time Periods", ["Sep 26"])], db)
    with pytest.raises(rest.RestError, match="no REST route"):
        rest.fetch_grid(_session(_Epm()), "Vision", {}, [], [], db)


def test_a_stale_atlas_is_an_svhttp_error(db):
    atlas.store(db, "FinMax", "All Bogus", [], 1.0, "failed")
    with pytest.raises(svhttp.SvHttpError, match="atlas is incomplete or stale for FinMax.*All Bogus"):
        _grid(_Epm(), db)


# ---------------------------------------------------------------- errors
def _fails(status, text):
    def answer(path, body):
        if path.endswith("exportdataslice"):
            return status, text
    return answer


def test_a_400_is_a_refused_grid_with_oracles_message_scrubbed(db):
    jwt = ".".join(("h" * 20, "p" * 40, "s" * 30))                            # a token's shape, nothing of one
    text = json.dumps({"status": 400, "detail": f"Member ac.C99 not found.\n See https://x.example/a?token={jwt} "
                                                f"Bearer {jwt}" + " pad" * 200})
    with pytest.raises(rest.RestError, match="refused the grid \\(HTTP 400\\): Member ac.C99 not found") as e:
        _grid(_Epm(answer=_fails(400, text)), db)
    assert isinstance(e.value, svhttp.SvHttpError) and not isinstance(e.value, svhttp.RequestLost)
    assert jwt not in str(e.value) and "token=" not in str(e.value) and "\n" not in str(e.value)
    assert len(str(e.value)) < 400
    with pytest.raises(rest.RestError, match="HTTP 400\\): not JSON"):
        _grid(_Epm(answer=_fails(400, "<html>not JSON</html>")), db)


def test_a_401_takes_the_sessions_one_new_token(db):
    tokens = ["at-1", "at-1", "at-2"]                                          # the list's, the grid's, the new one
    calls = []

    def answer(path, body):
        if path.endswith("exportdataslice"):
            calls.append(1)
            return (401, "") if len(calls) == 1 else None
    epm = _Epm(answer=answer)
    s = _session(epm, *tokens)
    f = _grid(epm, db, s=s)
    assert f.values[2][0] == "C30 REVENUES (ac)" and s._tokens.asked == [None, None, "at-1"]


@pytest.mark.parametrize("status", [302, 403])
def test_a_gates_answer_takes_the_sessions_one_new_token_as_on_the_provider_route(db, status):
    calls = []

    def answer(path, body):
        if path.endswith("exportdataslice"):
            calls.append(1)
            return (status, "<html/>") if len(calls) == 1 else None
    epm = _Epm(answer=answer)
    s = _session(epm, "at-1", "at-1", "at-2")
    assert _grid(epm, db, s=s).values[2][0] == "C30 REVENUES (ac)" and s._tokens.asked == [None, None, "at-1"]


@pytest.mark.parametrize("status", [401, 302, 403])
def test_a_token_a_gate_refuses_twice_is_the_token_refused_error(db, status):
    epm = _Epm(answer=_fails(status, ""))
    with pytest.raises(svhttp.TokenRefused, match=f"refused the access token twice \\(HTTP {status}\\)") as e:
        _grid(epm, db, s=_session(epm, "at-1", "at-2", "at-3"))
    assert isinstance(e.value, rest.RestError) and len(_slices(epm)) == 2


def test_a_token_refused_twice_is_an_error_after_two_sends(db):
    epm = _Epm(answer=_fails(401, ""))
    with pytest.raises(rest.RestError, match="refused the access token twice \\(HTTP 401\\)"):
        _grid(epm, db, s=_session(epm, "at-1", "at-2", "at-3"))
    assert len(_slices(epm)) == 2


def test_a_404_on_the_plan_type_is_a_configuration_error(db):
    with pytest.raises(rest.NotConfigured, match="no plan type FinMax in application FinMax"):
        _grid(_Epm(answer=lambda path, body: (404, "{}") if path.endswith("/dimensions") else None), db)
    with pytest.raises(rest.NotConfigured, match="no plan type"):
        _grid(_Epm(answer=_fails(404, "")), db)


@pytest.mark.parametrize("status", [403, 302, 500, 503])
def test_another_status_is_an_error_naming_it(db, status):     # a gate's 302 / 403 after its one new token too
    with pytest.raises(rest.RestError, match=f"HTTP {status}"):
        _grid(_Epm(answer=_fails(status, "<html/>")), db)


def test_a_lost_request_is_the_provider_routes_lost_request(db):
    with pytest.raises(svhttp.RequestLost, match="lost the request \\(ConnectionResetError\\)") as e:
        _grid(_Epm(answer=_fails(ConnectionResetError("at-1 reset"), "")), db)
    assert isinstance(e.value, rest.RestError) and "at-1" not in str(e.value)


@pytest.mark.parametrize("text", ["not json", "[]", '{"pov": []}', json.dumps({"items": "x"}), "{}"])
def test_an_answer_that_is_not_json_of_the_shape_is_an_error(db, text):
    with pytest.raises(rest.RestError, match="not a Planning REST"):
        _grid(_Epm(answer=lambda path, body: (200, text)), db)


def test_a_dimension_list_with_more_pages_is_an_error(db):
    def answer(path, body):
        if path.endswith("/dimensions"):
            return 200, json.dumps({"items": [], "hasMore": True, "totalResults": 99})
    with pytest.raises(rest.RestError, match="more"):
        _grid(_Epm(answer=answer), db)


def test_no_sign_in_reaches_the_caller_as_it_is(db):
    epm = _Epm()
    with pytest.raises(oauth.NotSignedIn):
        _grid(epm, db, s=_session(epm, oauth.NotSignedIn("sign in")))
    assert epm.sent == []


# ---------------------------------------------------------------- a wrong member (Oracle's HTTP 400 names it)
def oracle_400(member, dimension):
    """Oracle's answer for a member the plan type does not have there (shape live 2026-11-22: the same text under
    detail, message and localizedMessage; the member as sent, in double quotes when it has a space; the dimension's
    REST name). The members here are invented."""
    shown = f'"{member}"' if " " in member else member
    msg = (f"The member {shown} does not exist in the {dimension} dimension for the specified cube or you do not have "
           "access to it.")
    return 400, json.dumps({"status": 400, "detail": msg, "message": msg, "localizedMessage": msg})


def refusing(*wrong):
    """exportdataslice answers HTTP 400 naming the first member of 'wrong' it was sent (POV, columns, rows)."""
    def answer(path, body):
        if not path.endswith("exportdataslice"):
            return None
        g = body["gridDefinition"]
        for seg in (g["pov"], *g["columns"], *g["rows"]):
            for dim, ms in zip(seg["dimensions"], seg["members"]):
                for m in ms:
                    if m in wrong:
                        return oracle_400(m, dim)
    return answer


def test_a_400_naming_a_member_asked_on_an_axis_of_several_is_a_member_that_can_be_left_out(db):
    with pytest.raises(rest.MemberRefused) as e:
        _grid(_Epm(answer=refusing("ac.C70")), db, accounts=("ac.C30", "ac.C70", "ac.C90"))
    assert (e.value.dimension, e.value.member, e.value.droppable) == ("All Accounts", "ac.C70", True)
    assert isinstance(e.value, svhttp.SvHttpError) and not isinstance(e.value, svhttp.NoFallback)
    assert "All Accounts member ac.C70" in str(e.value) and "FinMax" in str(e.value)


def test_a_member_with_a_space_is_read_from_oracles_quotes(db):
    with pytest.raises(rest.MemberRefused) as e:
        rest.fetch_grid(_session(_Epm(answer=refusing("2BUS Biz"))), "FinPlan",
                        {"Scenario": "Target", "Version": "Final", "Years": "FY26", "All Balance Types": "Periodic"},
                        [Axis("Account", ["EBX Ebitda X"]), Axis("Entity", ["1TEC Tech", "2BUS Biz"])],
                        [Axis("Period", ["Jan"])], db)
    assert (e.value.dimension, e.value.member, e.value.droppable) == ("Entity", "2BUS Biz", True)


@pytest.mark.parametrize("wrong, where", [("NOK", "a POV member"), ("un.SY.E1", "the only All Entities member")])
def test_a_member_that_cannot_be_left_out_is_a_clear_error_naming_it(db, wrong, where):
    """A POV member, or the only member on its axis: leaving it out would change what every remaining cell means or
    leave no cell - the error names the member, the grid and what to do."""
    epm = _Epm(answer=refusing(wrong))
    with pytest.raises(rest.MemberRefused) as e:
        _grid(epm, db, accounts=("ac.C30", "ac.C70"))
    assert e.value.member == wrong and not e.value.droppable
    assert where in str(e.value) and "cannot be left out" in str(e.value) and "smartass atlas build" in str(e.value)
    assert len(_slices(epm)) == 1                                                  # no reduced request was sent


@pytest.mark.parametrize("named", ["ac.C7", "ac.C300", "AC.C70", "ac.ZZ9"])
def test_a_400_naming_a_member_that_was_not_asked_is_never_acted_on(db, named):
    """Matched against the members asked exactly: a part, a longer name or another case of one is not that member."""
    with pytest.raises(rest.RestError) as e:
        _grid(_Epm(answer=_fails(*oracle_400(named, "Accounts"))), db, accounts=("ac.C30", "ac.C70"))
    assert not isinstance(e.value, rest.MemberRefused)
    assert "not asked" in str(e.value) and "smartass atlas build" in str(e.value)


def test_a_member_named_under_another_dimension_than_the_one_asked_is_not_acted_on(db):
    with pytest.raises(rest.RestError) as e:
        _grid(_Epm(answer=_fails(*oracle_400("ac.C70", "Entities"))), db, accounts=("ac.C30", "ac.C70"))
    assert not isinstance(e.value, rest.MemberRefused)


@pytest.mark.parametrize("detail", [
    "Error: The member ac.C70 does not exist in the Accounts dimension for the specified cube or you do not have "
    "access to it.",
    "The member ac.C70 does not exist in the Accounts dimension for the specified cube or you do not have access to "
    "it. Retry later.",
    "The member ac.C70 does not exist in the Accounts dimension"])
def test_only_oracles_whole_wrong_member_sentence_names_a_member(db, detail):
    with pytest.raises(rest.RestError) as e:
        _grid(_Epm(answer=_fails(400, json.dumps({"status": 400, "detail": detail}))), db, accounts=("ac.C30", "ac.C70"))
    assert not isinstance(e.value, rest.MemberRefused)
