"""The atlas over the token session (epm/members.py, ops/atlas_build.build_rest / build_into): a dimension's member
tree from Oracle's Get Member, turned into the rows the hidden Excel's zoom gave (ops.dump_dimension +
tree_from_labels), the build into a new file swapped in only when it completed. EPM is faked (an opener under a real
EpmSession, or a source object); every member, alias and code here is invented."""
import contextlib
import http.client
import io
import json
import sqlite3
import urllib.error

import pytest

from smartass.epm import epmhttp, members, oauth
from smartass.epm.grid import tree_from_labels
from smartass.knowledge import atlas
from smartass.ops import atlas_build, cli

POD = "https://plan1.example"


def node(name, alias=None, *children, storage="Store Data", used=("FinMax", "FinLite"), **extra):
    n = {"name": name, "dataStorage": storage, "usedIn": list(used), **extra}
    if alias is not None:
        n["alias"] = alias
    if children:
        n["children"] = list(children)
    return n


# An invented account tree as Get Member answers it (outline order), with a shared member under two parents and a
# label without the "code alias (prefix)" form (kept under its label).
ACCOUNTS = node("Accounts", "All Accounts",
                node("ac.R", "R Revenue total (ac)",
                     node("ac.R1", "R1 Sales one (ac)"),
                     node("ac.R2", "R2 Sales two (ac)")),
                node("ac.K", "K Key figures (ac)",
                     node("ac.R1", "R1 Sales one (ac)", storage="Shared"),
                     node("KPI1", None)))


def smart_view_display(top, bottom=True, indent=5):
    """What the zoom showed in the hidden Excel: indented labels, ancestors at the bottom (today's atlas)."""
    out = []

    def go(n, d):
        if not bottom:
            out.append(" " * indent * d + (n.get("alias") or n["name"]))
        for c in n.get("children") or []:
            go(c, d + 1)
        if bottom:
            out.append(" " * indent * d + (n.get("alias") or n["name"]))
    go(top, 0)
    return out


KEYS = ("label", "name", "depth", "parent", "order")


def test_the_rows_are_the_rows_the_zoom_gave():
    """label = the Default alias, else the name; the name derived from the label as before (ac.R1 from its label,
    KPI1 kept under its label); depth, parent and order as tree_from_labels reads an ancestors-at-the-bottom display:
    siblings in reverse outline order; a shared member under each of its parents."""
    got = members.rows(ACCOUNTS, "FinMax")
    want = tree_from_labels(smart_view_display(ACCOUNTS))
    assert [{k: r[k] for k in KEYS} for r in got] == want
    assert [(r["name"], r["parent"], r["depth"]) for r in got] == [
        ("All Accounts", None, 0), ("ac.K", "All Accounts", 1), ("KPI1", "ac.K", 2), ("ac.R1", "ac.K", 2),
        ("ac.R", "All Accounts", 1), ("ac.R2", "ac.R", 2), ("ac.R1", "ac.R", 2)]
    assert [r["order"] for r in got] == list(range(7))


def test_each_row_carries_the_members_real_name_alias_storage_and_types():
    tree = node("Account", None, node("EBX", "EBX Ebitda X", accountType="Revenue", timeBalance="None", used=("FinPlan",)),
                node("HC", None, storage="Never Share", accountType="Expense", timeBalance="Average",
                     used=("FinPlan",)),
                used=("FinPlan",))
    got = {r["label"]: r for r in members.rows(tree, "FinPlan")}
    assert {k: got["EBX Ebitda X"][k] for k in members.EXTRA} == {
        "real_name": "EBX", "alias": "EBX Ebitda X", "storage": "Store Data", "account_type": "Revenue",
        "time_balance": "None"}
    assert got["EBX Ebitda X"]["name"] == "EBX Ebitda X"            # the key stays the label, as today
    assert {k: got["HC"][k] for k in members.EXTRA} == {
        "real_name": "HC", "alias": None, "storage": "Never Share", "account_type": "Expense", "time_balance": "Average"}
    assert got["Account"]["account_type"] is None and got["Account"]["real_name"] == "Account"


def test_dynamic_time_series_and_members_of_another_plan_type_are_left_out():
    """The zoom showed neither: the X-T-D members under the top of a period dimension (Get Member lists them first),
    and members whose usedIn does not name the plan type (with their subtrees)."""
    tree = node("TimePeriods", "All Time Periods",
                *(node(f"{x}-T-D", None, used=()) for x in "HYSPQMWD"),
                node("Y1", "Year one", node("Q1", "Quarter one"), node("Q9", "Lite only", node("M9", "m"),
                                                                       used=("FinLite",))))
    assert [r["label"] for r in members.rows(tree, "FinMax")] == ["All Time Periods", "Year one", "Quarter one"]
    assert [r["label"] for r in members.rows(tree, "FinLite")] == ["All Time Periods", "Year one", "Lite only", "m",
                                                                  "Quarter one"]
    deeper = node("X", None, node("Y", None, node("H-T-D", None)))      # a name like it below the top is kept
    assert [r["label"] for r in members.rows(deeper, "FinMax")] == ["X", "Y", "H-T-D"]


def test_a_dimension_of_one_member_is_one_row():
    assert [{k: r[k] for k in KEYS} for r in members.rows(node("Version", None), "FinMax")] == [
        {"label": "Version", "name": "Version", "depth": 0, "parent": None, "order": 0}]


def test_a_name_used_twice_under_one_parent_is_kept_once_as_the_atlas_stores_it(tmp_path):
    """Two members whose labels give one name under one parent (store's INSERT OR IGNORE keeps the first); the
    build counts both, as the hidden Excel's build did."""
    tree = node("Accounts", "All Accounts", node("ac.A", "A First (ac)"), node("ac.A2", "A Second (ac)"))
    rows = members.rows(tree, "FinMax")
    assert [r["name"] for r in rows] == ["All Accounts", "ac.A", "ac.A"]
    with contextlib.closing(atlas.connect(tmp_path / "a.db")) as db:
        atlas.add_member_columns(db)
        atlas.store(db, "FinMax", "All Accounts", rows, 1.0)
        assert db.execute("SELECT count(*) FROM member").fetchone()[0] == 2
        assert db.execute("SELECT label, real_name FROM member WHERE name='ac.A'").fetchone() == ("A Second (ac)",
                                                                                                   "ac.A2")
        assert atlas.status(db)["built"][0]["members"] == 3


@pytest.mark.parametrize("answer", [
    [], "x", {"children": []}, {"name": 3}, {"name": "A", "children": {}}, {"name": "A", "children": ["B"]},
    {"name": "A", "children": [{"name": "B", "alias": 5}]}, {"name": "A", "usedIn": "FinMax"},
    {"name": "A", "children": [{"name": "B", "children": [{"alias": "no name"}]}]},
])
def test_an_answer_that_is_not_a_member_tree_is_refused(answer):
    with pytest.raises(members.TreeError, match="not a Planning REST member tree"):
        members.check_tree(answer)


# ---------------------------------------------------------------- the request (a real EpmSession, a fake wire)
class _Tokens:
    def scope_for(self, base):
        return "urn:scope"

    def access_token(self, base, reject=None):
        if isinstance(self, Exception):
            raise self
        return "at-1"


class _Answer(io.BytesIO):
    def __init__(self, status, text):
        super().__init__(text.encode())
        self.status, self.headers = status, http.client.HTTPMessage()

    def info(self):
        return self.headers


class _Opener:
    def __init__(self, status, text):
        self.status, self.text, self.sent = status, text, []

    def open(self, req, timeout):
        self.sent.append(req)
        if 200 <= self.status < 300:
            return _Answer(self.status, self.text)
        raise urllib.error.HTTPError(req.full_url, self.status, "no", http.client.HTTPMessage(),
                                     io.BytesIO(self.text.encode()))


def _session(status, text):
    op = _Opener(status, text)
    return epmhttp.EpmSession(POD, tokens=_Tokens(), opener=op), op


def test_the_tree_is_read_with_the_allowlisted_request_and_201_is_its_answer():
    s, op = _session(201, json.dumps(ACCOUNTS))
    assert members.read_tree(s, "FinMax", "Accounts") == ACCOUNTS
    assert [(r.get_method(), r.full_url, r.data) for r in op.sent] == [
        ("GET", POD + "/HyperionPlanning/rest/v3/applications/FinMax/dimensions/Accounts/members/Accounts?fields="
         + epmhttp.MEMBER_FIELDS, None)]


@pytest.mark.parametrize("status, text, match", [
    (200, json.dumps(ACCOUNTS), "HTTP 200"),          # exactly the status seen live, not a range
    (202, json.dumps(ACCOUNTS), "HTTP 202"),
    (400, '{"detail": "The member X does not exist"}', "HTTP 400"),
    (404, "{}", "HTTP 404"), (500, "{}", "HTTP 500"), (401, "{}", "HTTP 401"),
    (201, "<html>", "not JSON"),
])
def test_any_other_answer_is_an_error_naming_its_status_never_its_text(status, text, match):
    s, _ = _session(status, text)
    with pytest.raises(members.TreeError, match=match) as e:
        members.read_tree(s, "FinMax", "Accounts")
    assert "does not exist" not in str(e.value) and "<html>" not in str(e.value)


def test_a_dimension_without_a_rule_is_refused_before_anything_is_sent():
    s, op = _session(201, "{}")
    with pytest.raises(members.TreeError, match="refused"):
        members.read_tree(s, "FinMax", "Attributes")
    assert op.sent == []


def test_a_missing_sign_in_reaches_the_caller_as_it_is():
    class NoToken(_Tokens):
        def access_token(self, base, reject=None):
            raise oauth.NotSignedIn("sign in")
    s = epmhttp.EpmSession(POD, tokens=NoToken(), opener=_Opener(201, "{}"))
    with pytest.raises(oauth.NotSignedIn):
        members.read_tree(s, "FinMax", "Accounts")


# ---------------------------------------------------------------- the build
class _Source:
    """members.Source as the build sees it: the plan types' dimensions (REST name, atlas name) and the trees."""

    def __init__(self, trees, dims, fail=()):
        self.trees, self.dims, self.fail, self.read = trees, dims, set(fail), []

    def dimensions(self, conn):
        if conn in self.fail:
            raise members.TreeError(f"{conn}: HTTP 500")
        return self.dims[conn]

    def tree(self, conn, dim):
        self.read.append((conn, dim))
        if dim in self.fail:
            raise members.TreeError(f"{dim}: HTTP 500")
        return self.trees[dim]


ENTITIES = node("Entities", "All Entities", node("un.SY", "SY Group (un)", node("un.SY.E1", "E1 One (un.SY)")))
TREES = {"Accounts": ACCOUNTS, "Entities": ENTITIES,
         "Account": node("Account", None, node("EBX", "EBX Ebitda X", accountType="Revenue", timeBalance="None",
                                               used=("FinPlan",)), used=("FinPlan",))}
DIMS = {"FinMax": [("Accounts", "All Accounts"), ("Entities", "All Entities")],
        "FinLite": [("Accounts", "All Accounts")], "FinPlan": [("Account", "Account")]}


def _rows(db, conn, dim):
    return db.execute("SELECT name, label, parent, depth, ord FROM member WHERE connection=? AND dimension=? ORDER BY ord",
                      (conn, dim)).fetchall()


def test_build_rest_stores_every_dimension_of_every_cube(tmp_path):
    src = _Source(TREES, DIMS)
    with contextlib.closing(atlas.connect(tmp_path / "a.db")) as db:
        atlas.add_member_columns(db)
        done = atlas_build.build_rest(src, db=db, progress=lambda m: None)
        assert [(d["connection"], d["dimension"], d["members"]) for d in done] == [
            ("FinMax", "All Accounts", 7), ("FinMax", "All Entities", 3), ("FinLite", "All Accounts", 7),
            ("FinPlan", "Account", 2)]
        assert src.read == [("FinMax", "Accounts"), ("FinMax", "Entities"), ("FinLite", "Accounts"),
                            ("FinPlan", "Account")]
        want = [(r["name"], r["label"], r["parent"], r["depth"], r["order"])
                for r in tree_from_labels(smart_view_display(ACCOUNTS))]
        assert _rows(db, "FinMax", "All Accounts") == want == _rows(db, "FinLite", "All Accounts")
        assert atlas.has_member(db, "FinMax", "All Entities", "un.SY.E1")


def test_the_source_reads_finlites_trees_once_with_finmaxs(monkeypatch):
    """FinLite's dimensions are FinMax's application dimensions: one Get Member per (application, dimension) a build;
    the plan types' lists leave attribute dimensions out and name each dimension as the atlas does."""
    class Sv:
        def rest_session(self, conn):
            return conn
    reads = []
    monkeypatch.setattr(members, "read_tree", lambda s, app, dim: reads.append((s, app, dim)) or {"name": dim})
    monkeypatch.setattr(members.rest, "_read_list", lambda s, app, cube: [
        ("Accounts", "All Accounts", "Account"), ("Version", None, "Version"), ("Colour", None, "Attribute Text")])
    src = members.Source(Sv())
    assert src.dimensions("FinLite") == [("Accounts", "All Accounts"), ("Version", "Version")]
    assert src.tree("FinMax", "Accounts") is src.tree("FinLite", "Accounts")
    src.tree("FinPlan", "Accounts")
    assert reads == [("FinMax", "FinMax", "Accounts"), ("FinPlan", "FinPlan", "Accounts")]


def test_a_full_build_purges_the_cube_and_a_partial_one_does_not(tmp_path):
    with contextlib.closing(atlas.connect(tmp_path / "a.db")) as db:
        atlas.store(db, "FinMax", "All Gone", [{"name": "g", "label": "g", "parent": None, "depth": 0, "order": 0}],
                    1.0)
        atlas.add_member_columns(db)
        atlas_build.build_rest(_Source(TREES, DIMS), ["FinMax"], ["Entities"], db=db, progress=lambda m: None)
        assert atlas.has_member(db, "FinMax", "All Gone", "g")                 # a partial build keeps the cube
        done = atlas_build.build_rest(_Source(TREES, DIMS), ["FinMax"], ["accounts"], db=db, progress=lambda m: None)
        assert [d["dimension"] for d in done] == ["All Accounts"]              # any spelling of the dimension
        atlas_build.build_rest(_Source(TREES, DIMS), ["FinMax"], db=db, progress=lambda m: None)
        assert not atlas.has_member(db, "FinMax", "All Gone", "g")             # a full build purges it


def test_errors_are_recorded_per_dimension_and_never_stop_the_build(tmp_path):
    src = _Source(TREES, DIMS, fail=("Entities", "FinLite"))
    with contextlib.closing(atlas.connect(tmp_path / "a.db")) as db:
        atlas.add_member_columns(db)
        done = atlas_build.build_rest(src, db=db, progress=lambda m: None)
        assert [(d["connection"], d.get("dimension"), bool(d.get("error"))) for d in done] == [
            ("FinMax", "All Accounts", False), ("FinMax", "All Entities", True), ("FinLite", None, True),
            ("FinPlan", "Account", False)]
        st = {(b["connection"], b["dimension"]): b for b in atlas.status(db)["built"]}
        assert st["FinMax", "All Entities"]["error"] == "Entities: HTTP 500"


def _today(path):
    """Today's atlas: the old schema (no extra columns), a member, a card and a table smartass no longer writes."""
    db = sqlite3.connect(path)
    db.executescript("""
        CREATE TABLE member (connection TEXT, dimension TEXT, name TEXT, label TEXT, parent TEXT, depth INT, ord INT,
                             PRIMARY KEY (connection, dimension, name, parent));
        CREATE TABLE build (connection TEXT, dimension TEXT, built_at REAL, n INT, seconds REAL, error TEXT,
                            PRIMARY KEY (connection, dimension));
        CREATE TABLE card (connection TEXT, dimension TEXT, name TEXT, parent TEXT, operator TEXT, time_balance TEXT,
                           account_type TEXT, shared INT, derived_at REAL, PRIMARY KEY (connection, dimension, name, parent));
        CREATE TABLE snapshot (x);
        INSERT INTO member VALUES ('FinMax', 'All Accounts', 'ac.OLD', 'OLD Old (ac)', 'All Accounts', 1, 1);
        INSERT INTO member VALUES ('FinPlan', 'Account', 'Old', 'Old', 'Account', 1, 1);
        INSERT INTO build VALUES ('FinMax', 'All Accounts', 1.0, 2, 1.0, NULL);
        INSERT INTO card VALUES ('FinMax', 'All Accounts', 'ac.R1', 'ac.R', '+', 'flow', 'revenue', 0, 1.0);
        INSERT INTO snapshot VALUES (1);""")
    db.commit()
    db.close()


def test_a_completed_build_replaces_the_atlas_and_keeps_what_the_build_does_not_write(tmp_path):
    target = tmp_path / "atlas.db"
    _today(target)
    done, swapped = atlas_build.build_into(target, lambda db: atlas_build.build_rest(
        _Source(TREES, DIMS), ["FinMax"], db=db, progress=lambda m: None), progress=lambda m: None)
    assert swapped and not (tmp_path / "atlas.db.new").exists()
    with contextlib.closing(sqlite3.connect(target)) as db:
        assert db.execute("SELECT count(*) FROM member WHERE name='ac.OLD'").fetchone()[0] == 0     # purged
        assert db.execute("SELECT count(*) FROM member WHERE connection='FinPlan'").fetchone()[0] == 1  # not built
        assert db.execute("SELECT operator FROM card").fetchall() == [("+",)]                         # cards kept
        assert db.execute("SELECT x FROM snapshot").fetchall() == [(1,)]
        assert db.execute("SELECT real_name FROM member WHERE name='KPI1'").fetchone() == ("KPI1",)


def test_a_failed_build_leaves_todays_atlas_untouched_and_says_so(tmp_path):
    target = tmp_path / "atlas.db"
    _today(target)
    before = target.read_bytes()
    said = []
    done, swapped = atlas_build.build_into(target, lambda db: atlas_build.build_rest(
        _Source(TREES, DIMS, fail=("Entities",)), ["FinMax"], db=db, progress=lambda m: None), progress=said.append)
    assert not swapped and target.read_bytes() == before and not (tmp_path / "atlas.db.new").exists()
    assert "not replaced" in said[-1] and "1 dimension" in said[-1]
    with contextlib.closing(sqlite3.connect(target)) as db:                  # the old schema, as it was
        assert [r[1] for r in db.execute("PRAGMA table_info(member)")] == [
            "connection", "dimension", "name", "label", "parent", "depth", "ord"]


def test_a_build_that_raises_leaves_todays_atlas_untouched(tmp_path):
    target = tmp_path / "atlas.db"
    _today(target)
    before = target.read_bytes()

    def run(db):
        atlas_build.build_rest(_Source(TREES, DIMS), ["FinMax"], db=db, progress=lambda m: None)
        raise oauth.NotSignedIn("sign in")
    with pytest.raises(oauth.NotSignedIn):
        atlas_build.build_into(target, run, progress=lambda m: None)
    assert target.read_bytes() == before and not (tmp_path / "atlas.db.new").exists()


def test_a_first_build_creates_the_atlas(tmp_path):
    target = tmp_path / "home" / "atlas.db"
    done, swapped = atlas_build.build_into(target, lambda db: atlas_build.build_rest(
        _Source(TREES, DIMS), db=db, progress=lambda m: None), progress=lambda m: None)
    assert swapped
    with contextlib.closing(atlas.connect(target)) as db:
        assert atlas.known_dimensions(db, "FinPlan") == ["Account"]


def test_reading_the_atlas_never_adds_the_columns(tmp_path):
    """Only a build migrates (its new file): opening today's atlas to read it writes nothing - an installed older
    release keeps writing it as before."""
    target = tmp_path / "atlas.db"
    _today(target)
    with contextlib.closing(atlas.connect(target)) as db:
        assert "real_name" not in [r[1] for r in db.execute("PRAGMA table_info(member)")]
        assert atlas.canonical(db, "FinMax", "All Accounts", "ac.old") == "ac.OLD"


def test_the_cli_builds_over_rest_and_the_old_route_stays_behind_a_flag(monkeypatch, capsys):
    """The gate passed (docs/audit/2026-11-22-atlas-rest.md): REST is the default; --excel is the hidden Excel's zoom
    until Task 6."""
    seen = []
    monkeypatch.setattr(atlas_build, "run_build", lambda cubes, dims, excel: seen.append(excel) or [
        {"connection": "FinMax", "dimension": "x", "members": 3}])
    assert cli.main(["atlas", "build"]) == 0 and cli.main(["atlas", "build", "--excel"]) == 0
    assert seen == [False, True]


def test_run_build_goes_over_rest_into_the_atlas_and_derives_the_cards_after_the_swap(tmp_path, monkeypatch):
    target = tmp_path / "atlas.db"
    _today(target)
    monkeypatch.setattr(atlas_build.config, "ATLAS_DB", target)
    monkeypatch.setattr(members, "Source", lambda: _Source(TREES, DIMS))
    cards = []
    monkeypatch.setattr(atlas_build, "cards_after_build", lambda done: cards.append(len(done)))
    done = atlas_build.run_build()
    assert [d.get("error") for d in done] == [None] * 4 and cards == [4]
    with contextlib.closing(sqlite3.connect(target)) as db:
        assert db.execute("SELECT count(*) FROM member WHERE name='ac.OLD'").fetchone()[0] == 0
    atlas_build.run_build(["FinPlan"], ["Account"])
    assert cards == [4]                                     # a partial build derives no cards


def test_run_build_without_a_sign_in_changes_nothing(tmp_path, monkeypatch):
    target = tmp_path / "atlas.db"
    _today(target)
    before = target.read_bytes()
    monkeypatch.setattr(atlas_build.config, "ATLAS_DB", target)

    class NoSignIn(_Source):
        def dimensions(self, conn):
            raise oauth.NotSignedIn("sign in")
    monkeypatch.setattr(members, "Source", lambda: NoSignIn(TREES, DIMS))
    done = atlas_build.run_build()
    assert len(done) == 1 and "smartass login" in done[0]["error"] and target.read_bytes() == before
