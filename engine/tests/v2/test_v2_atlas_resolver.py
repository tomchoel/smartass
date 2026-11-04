import pytest

from smartass.knowledge.atlas_resolver import AtlasResolver
from smartass.model.resolver import MetricDef, UnknownMetric
from smartass.knowledge import atlas


@pytest.fixture
def db(tmp_path):
    con = atlas.connect(tmp_path / "atlas.db")
    rows = [("FinMax", "All Entities", "un.SY.2BUS", "2BUS Business Norway", "un.SY", 2, 1),
            ("FinMax", "All Entities", "un.SY.4BUS_B", "4BUS_B Business", "un.SY.2BUS", 3, 1),
            ("FinMax", "All Entities", "un.SY.G910", "G910 Talkmore Business", "un.SY.2BUS", 3, 2),
            ("FinMax", "All Accounts", "ac.C30", "C30 REVENUES", "ac.C48", 3, 1),
            ("FinMax", "All Accounts", "ac.X1", "X1 child of 2BUS name clash", "un.SY.2BUS", 4, 1),
            ("FinPlan", "Scenario", "Forecast Q226", "Forecast Q226", "Forecasts", 2, 1),
            ("FinPlan", "Scenario", "Forecast Q326", "Forecast Q326", "Forecasts", 2, 2)]
    con.executemany("INSERT INTO member (connection, dimension, name, label, parent, depth, ord) "
                    "VALUES (?,?,?,?,?,?,?)", rows)
    con.commit()
    return con


def res(db, yy=26):
    return AtlasResolver(db, [MetricDef("Revenue", {"FinMax": ("ac.C30",)}, sign="income")], forecast_year=yy)


def test_children_stay_in_the_members_dimension(db):
    assert res(db).children("un.SY.2BUS", "FinMax") == ("un.SY.4BUS_B", "un.SY.G910")
    assert res(db).children("un.SY.G910", "FinMax") == ()


def test_exists_label_and_metrics(db):
    r = res(db)
    assert r.exists("un.sy.2bus", "FinMax") and not r.exists("un.SY.2BUS", "FinPlan")
    assert r.label("un.SY.2BUS", "FinMax") == "2BUS Business Norway"
    assert r.label("un.SY.ZZ", "FinMax") == "un.SY.ZZ"
    assert r.members("revenue", "FinMax") == ("ac.C30",)
    assert r.members("ac.C30", "FinMax") == ("ac.C30",)
    with pytest.raises(UnknownMetric):
        r.members("ac.NOPE", "FinMax")
    assert r.weight("un.SY.G910", "FinMax") == 1.0


def test_latest_forecast_of_the_year(db):
    assert res(db).latest_forecast() == "Forecast Q326"
    with pytest.raises(UnknownMetric, match="Forecast Q<q>25"):
        res(db, 25).latest_forecast()


def test_label_drops_a_trailing_parenthesis_suffix(db):
    db.execute("INSERT INTO member (connection, dimension, name, label, parent, depth, ord) "
               "VALUES ('FinMax','All Entities','un.SY.2WHS','2WHS Wholesale (un.SY)','un.SY',2,9)")
    db.commit()
    r = res(db)
    assert r.label("un.SY.2WHS", "FinMax") == "2WHS Wholesale"
    assert r.label("un.SY.2BUS", "FinMax") == "2BUS Business Norway"


def test_canonical_spelling(db):
    r = res(db)
    assert r.canonical("un.sy.2bus", "FinMax") == "un.SY.2BUS"
    assert r.canonical("un.SY.ZZ", "FinMax") == "un.SY.ZZ"
