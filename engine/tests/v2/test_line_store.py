import asyncio
from types import SimpleNamespace

import pytest

from smartass import server
from smartass.core.errors import InputError
from smartass.epm import line_store

CELL = {"All Accounts": "ac.C40S", "All Entities": "un.SY.5BMT_T", "All Time Periods": "Yr 26"}

SAMPLE_ROWS = [
    {"Period": "Jan 26", "Entity": "un.SY.VB20", "CC_DESC": "Wholesale Core", "ACCOUNT": "43200",
     "ACCOUNT_NAME": "Power COGS", "AP_CUSTOMER_NAME": "Statkraft AS", "JE_SOURCE": "Payables",
     "GL_DESCRIPTION_LINE": "Power invoice Jan", "AMOUNT": -500.0},
    {"Period": "Feb 26", "Entity": "un.SY.VB20", "CC_DESC": "Wholesale Core", "ACCOUNT": "43200",
     "ACCOUNT_NAME": "Power COGS", "AP_CUSTOMER_NAME": "Statkraft AS", "JE_SOURCE": "Payables",
     "GL_DESCRIPTION_LINE": "Accrual reversal pair", "AMOUNT": -100.0},
    {"Period": "Feb 26", "Entity": "un.SY.VB20", "CC_DESC": "Wholesale Core", "ACCOUNT": "43200",
     "ACCOUNT_NAME": "Power COGS", "AP_CUSTOMER_NAME": "Statkraft AS", "JE_SOURCE": "Payables",
     "GL_DESCRIPTION_LINE": "Accrual reversal pair", "AMOUNT": 100.0},
    {"Period": "Nov 26", "Entity": "un.SY.VB50", "CC_DESC": "Roaming", "ACCOUNT": "41000",
     "ACCOUNT_NAME": "Interconnect", "AP_CUSTOMER_NAME": "Telia Carrier", "JE_SOURCE": "Spreadsheet",
     "GL_DESCRIPTION_LINE": "Employee Ola Nordmann bonus", "AMOUNT": -250.0},
]
SAMPLE_COLS = list(SAMPLE_ROWS[0].keys())


def _seed(tmp_path):
    db_path = tmp_path / "lines_test.db"
    did = line_store.store_drill(
        CELL, SAMPLE_ROWS, SAMPLE_COLS,
        cube_value=-750.0,
        reconciliation={"cube_value": -750.0, "lines_sum": -750.0, "gap": 0.0, "status": "exact", "lines": 4},
        warnings=[], db_path=db_path,
    )
    return db_path, did


def test_structured_filtering_period_range_reversals_and_multi_group(tmp_path):
    db_path, did = _seed(tmp_path)
    out = line_store.query_lines(
        did,
        where={"period_from": "Jan 26", "period_to": "Oct 26"},
        group_by=["ACCOUNT", "Entity"],
        exclude_reversals=True,
        db_path=db_path,
    )
    assert out["reconciliation"]["status"] == "exact"
    assert out["filter_summary"]["matched_sum"] == -500.0
    assert out["filter_summary"]["excluded_reversal_pairs"]["lines"] == 2
    assert len(out["rows"]) == 1
    r = out["rows"][0]
    assert (r["ACCOUNT"], r["Entity"], r["ACCOUNT_NAME"], r["CC_DESC"], r["net_amount"]) == (
        "43200", "un.SY.VB20", "Power COGS", "Wholesale Core", -500.0,
    )


def test_read_only_sql_and_authorizer_blocks_mutations(tmp_path):
    db_path, did = _seed(tmp_path)
    res = line_store.query_lines(
        did,
        sql="SELECT Period, COUNT(*) AS n, SUM(AMOUNT) AS net FROM lines WHERE period_ym <= '2026-10' GROUP BY Period",
        db_path=db_path,
    )
    assert res["mode"] == "sql"
    assert {(r["Period"], r["net"]) for r in res["rows"]} == {("Jan 26", -500.0), ("Feb 26", 0.0)}

    pii = line_store.query_lines(did, sql="SELECT GL_DESCRIPTION_LINE, AMOUNT FROM lines", db_path=db_path)
    assert any("aggregate personal" in w.lower() for w in pii["warnings"])

    for bad_sql in (
        "DELETE FROM lines",
        "DROP TABLE drills",
        "PRAGMA table_info(drills)",
        "SELECT * FROM drills",
        "SELECT 1; DROP TABLE drills",
    ):
        with pytest.raises(InputError):
            line_store.query_lines(did, sql=bad_sql, db_path=db_path)


def test_fin_drill_explore_mode_caches_without_excel(tmp_path, monkeypatch):
    monkeypatch.setattr("smartass.epm.line_store._session_db_path", lambda: tmp_path / "sess.db")
    drills = []

    async def fake_drill(cell):
        drills.append(cell)
        return SimpleNamespace(
            cell=cell, rows=SAMPLE_ROWS, columns=SAMPLE_COLS, cell_value=-750.0,
            reconciliation={"cube_value": -750.0, "lines_sum": -750.0, "gap": 0.0, "status": "exact", "lines": 4},
            profile=None, warnings=[], source="FinLite",
        )

    monkeypatch.setattr(server, "_drill", fake_drill)
    out1 = asyncio.run(server.fin_drill(CELL))
    assert out1["cached"] is False and len(drills) == 1
    out2 = asyncio.run(server.fin_drill(CELL))
    assert out2["cached"] is True and out2["drill_id"] == out1["drill_id"] and len(drills) == 1

    queried = server.fin_lines(out1["drill_id"], group_by="account")
    assert queried["total_results"] == 2 and queried["filter_summary"]["matched_sum"] == -750.0
