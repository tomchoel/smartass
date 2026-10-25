"""templates.py / ops / answer: FinMax entity mapping via the atlas, pull-member validation, error-cell scan."""
import sys
from types import SimpleNamespace as NS

import pytest

sys.path.insert(0, __file__.rsplit("\\tests", 1)[0])
from smartview_mcp import answer, atlas, ops, templates  # noqa: E402
from smartview_mcp.grid import Axis  # noqa: E402
from smartview_mcp.host import InputError  # noqa: E402

VALUE_ERR = -2146826273                       # Excel #VALUE! as COM returns it


def _m(name, parent, depth=1):
    return {"name": name, "label": name, "parent": parent, "depth": depth, "order": 0}


@pytest.fixture
def fake_atlas(tmp_path, monkeypatch):
    db = atlas.connect(tmp_path / "atlas.db")
    atlas.store(db, "FinMax", "All Entities", [_m("All Entities", None, 0), _m("un.SY", "All Entities"),
                                               _m("un.SY.2TEC", "un.SY", 2)], 1.0)
    atlas.store(db, "FinMax", "All Accounts", [_m("ac.EBITDA", "All Accounts")], 1.0)
    atlas.store(db, "FinPlan", "Entity", [_m("BANORWAY+TIPO", "Entity"), _m("BANORWAY", "BANORWAY+TIPO", 2),
                                          _m("2TEC Technology", "BANORWAY", 3)], 1.0)
    monkeypatch.setattr(atlas, "connect", lambda *a, **k: db)
    return db


def fake_pull(sheet, rows, cols, values):
    """values[i][j] for rows[i] x cols[j] (single row and column dimension)."""
    lay = NS(row_tuples=[(r,) for r in rows], col_tuples=[(c,) for c in cols], nr=1, nc=1)
    parsed = NS(rows=[[r, *values[i]] for i, r in enumerate(rows)], pov_problems=[], unknown_members=[])
    return NS(sheet=sheet, layout=lay, parsed=parsed, row_labels={(r,): (r,) for r in rows},
              col_labels={(c,): (c,) for c in cols}, positional=False)


# ------------------------------------------------------------------ FinMax entity mapping
def test_finmax_entity_maps_division_top_company_and_unknown(fake_atlas):
    assert templates.map_finmax_entity("2TEC Technology") == ("un.SY.2TEC", None)
    assert templates.map_finmax_entity("BANORWAY") == ("un.SY", None)
    assert templates.map_finmax_entity("BANORWAY", "un.SY.2TEC") == ("un.SY.2TEC", None)
    ent, note = templates.map_finmax_entity("9XYZ Nowhere")
    assert ent is None
    assert note == "No FinMax entity matches 9XYZ Nowhere; pass finmax_entity to reconcile."


def _plan(monkeypatch, entity, **kw):
    monkeypatch.setattr(templates, "closed_month", lambda h, c="FinMax": "Sep 26")
    return templates.actual_vs_plan(None, entity, month="Sep", year="FY26", **kw)


def _fp_pulls(ebitda=100.0):
    accts = [a for _, a, _ in templates.fp_lines()]
    act = fake_pull("Q6 FinPlan actual", accts, ["Actual"], [[ebitda]] * len(accts))
    plan = fake_pull("Q6 FinPlan plan", accts, ["Target"], [[90.0]] * len(accts))
    return {"act": act, "plan": plan}


def test_actual_vs_plan_drops_recon_for_unknown_entity(fake_atlas, monkeypatch):
    plan = _plan(monkeypatch, "9XYZ Nowhere")
    assert [p.id for p in plan.pulls] == ["act", "plan"]
    spec = plan.answer(_fp_pulls())
    keys = [ln["key"] for sec in spec["sections"] for ln in sec["lines"]]
    assert "fm_ebitda" not in keys and "fm_diff" not in keys
    assert "No FinMax entity matches 9XYZ Nowhere; pass finmax_entity to reconcile." in spec["notes"]
    assert not any("n/a" in kp for kp in spec["key_points"])


def test_actual_vs_plan_banorway_reconciles_to_un_sy(fake_atlas, monkeypatch):
    plan = _plan(monkeypatch, "BANORWAY")
    fm = next(p for p in plan.pulls if p.id == "fm")
    assert fm.pov["All Entities"] == "un.SY"


def test_recon_key_point_without_finmax_value_never_says_na(fake_atlas, monkeypatch):
    plan = _plan(monkeypatch, "BANORWAY")
    P = {**_fp_pulls(), "fm": fake_pull("Q6 FinMax recon", ["ac.EBITDA"], ["Sep 26"], [["#NoAccess"]])}
    kp = " ".join(plan.answer(P)["key_points"])
    assert "n/a" not in kp
    assert "FinMax reconciliation unavailable (" in kp and "#NoAccess" in kp


def test_cost_base_share_not_claimed_without_total(fake_atlas, monkeypatch):
    monkeypatch.setattr(templates, "closed_month", lambda h, c="FinMax": "Sep 26")
    plan = templates.cost_base(None, parent="un.SY", accounts=["ac.C70"], period="Sep 26", top=1)
    p = fake_pull("Q1 FinMax", ["un.SY.2TEC", "un.SY"], ["ac.C70"], [[-5.0], ["#NoAccess"]])
    kp = " ".join(plan.answer({"cb": p})["key_points"])
    assert "% of the total" not in kp


# ------------------------------------------------------------------ pull member validation
def test_unknown_pull_member_raises_input_error(fake_atlas):
    ps = templates.PullSpec("fm", "FinMax", [Axis("All Accounts", ["ac.EBITDA"])],
                            [Axis("All Time Periods", ["Sep 26"])],          # dimension not in the atlas: skipped
                            {"All Entities": "un.SY.BANORWAY"})
    with pytest.raises(InputError) as e:
        templates.check_pulls([ps])
    msg = str(e.value)
    assert "un.SY.BANORWAY" in msg and "FinMax" in msg and "All Entities" in msg and "sv_find_member" in msg


def test_known_members_and_missing_atlas_pass(fake_atlas):
    ok = templates.PullSpec("fm", "FinMax", [Axis("All Accounts", ["ac.EBITDA"])],
                            [Axis("All Time Periods", ["Sep 26"])], {"All Entities": "un.SY.2TEC"})
    no_atlas = templates.PullSpec("x", "FinLite", [Axis("All Accounts", ["ac.NOPE"])],
                                  [Axis("All Time Periods", ["Sep 26"])], {})
    templates.check_pulls([ok, no_atlas])


def test_stale_atlas_is_skipped(fake_atlas):
    fake_atlas.execute("UPDATE build SET built_at = built_at - 40 * 86400")
    assert ops.unknown_members("FinMax", {"All Entities": "un.SY.NOPE"}, [], []) == []


# ------------------------------------------------------------------ error-cell scan
def test_trace_maps_pull_cells_and_dependencies():
    fm = fake_pull("Q6 FinMax recon", ["ac.EBITDA"], ["Sep 26"], [[None]])
    plan = [("line", {"key": "ebitda", "cells": [{"value": 1}]}, 19),
            ("line", {"key": "fm_ebitda", "cells": [{"pull": fm.sheet, "row": ["ac.EBITDA"], "col": ["Sep 26"]}]}, 20),
            ("line", {"key": "fm_diff", "cells": [{"expr": "ebitda-fm_ebitda"}]}, 21)]
    tr = answer.cell_trace(plan, {"ebitda": 19, "fm_ebitda": 20, "fm_diff": 21}, {fm.sheet: fm}, 1)
    assert tr[(20, 3)] == {"key": "fm_ebitda", "pull": "'Q6 FinMax recon'!B3"}
    assert tr[(21, 3)]["key"] == "fm_diff" and set(tr[(21, 3)]["deps"]) == {(19, 3), (20, 3)}


def test_scan_flags_noaccess_pull_and_value_error_on_answer_line():
    ans = [[None] * 3 for _ in range(21)]
    ans[18][2], ans[19][2], ans[20][2] = 100.0, "#NoAccess", VALUE_ERR
    trace = {(19, 3): {"key": "ebitda"}, (20, 3): {"key": "fm_ebitda", "pull": "'Q6 FinMax recon'!B3"},
             (21, 3): {"key": "fm_diff", "deps": [(19, 3), (20, 3)]}}
    pull_grid = ((None, "All Entities"), (None, "Sep 26"), ("EBITDA", "#NoAccess"))
    w = answer.error_warnings("Q6 Svar", ans, trace, {"Q6 FinMax recon": pull_grid})
    text = "\n".join(w)
    assert "'Q6 FinMax recon'!B3 is #NoAccess: no read access or member not valid in this cube" in text
    assert "formula error on line 'fm_diff': 'Q6 Svar'!C21 is #VALUE! (depends on 'Q6 FinMax recon'!B3)" in text
    assert "formula error on line 'fm_ebitda': 'Q6 Svar'!C20 is #NoAccess" in text
    assert not any("line 'ebitda'" in x for x in w)


# ------------------------------------------------------------------ review of 4c9ca14
def test_member_check_is_case_insensitive_and_pov_gets_atlas_spelling(fake_atlas):
    assert atlas.has_member(fake_atlas, "FinMax", "All Entities", "un.sy.2tec")
    ps = templates.PullSpec("fm", "FinMax", [Axis("All Accounts", ["AC.ebitda"])],
                            [Axis("All Time Periods", ["Sep 26"])], {"All Entities": "UN.SY.2TEC"})
    templates.check_pulls([ps])
    assert ps.pov["All Entities"] == "un.SY.2TEC"


def test_atlas_miss_is_confirmed_live_before_rejecting(fake_atlas):
    ps = templates.PullSpec("fm", "FinMax", [Axis("All Accounts", ["ac.NEW1"])],
                            [Axis("All Time Periods", ["Jan 27"])], {"All Entities": "un.SY.NOPE"})
    asked = []

    def confirm(conn, dim, member):
        asked.append((conn, dim, member))
        return member == "ac.NEW1"               # new account exists in the cube, the entity does not
    with pytest.raises(InputError) as e:
        templates.check_pulls([ps], confirm=confirm)
    assert "un.SY.NOPE" in str(e.value) and "ac.NEW1" not in str(e.value)
    assert set(asked) == {("FinMax", "All Entities", "un.SY.NOPE"), ("FinMax", "All Accounts", "ac.NEW1")}
    ok = templates.PullSpec("fm", "FinMax", [Axis("All Accounts", ["ac.NEW1"])], [Axis("All Time Periods", ["x"])], {})
    templates.check_pulls([ok], confirm=lambda c, d, m: True)


def test_failed_live_confirmation_does_not_reject(fake_atlas):
    ps = templates.PullSpec("fm", "FinMax", [Axis("All Accounts", ["ac.NEW1"])], [Axis("All Time Periods", ["x"])], {})

    def broken(conn, dim, member):
        raise RuntimeError("Smart View busy")
    templates.check_pulls([ps], confirm=broken)  # unverified, not rejected: the error-cell scan still runs


class FakeFind:
    """ops.live_dimension stand-in: {member: dimension}; everything else is not found."""
    def __init__(self, known):
        self.known, self.calls = known, []

    def __call__(self, h, conn, member):
        self.calls.append(member)
        return self.known.get(member)


def test_live_confirm_checks_dimension_and_that_the_check_works(fake_atlas, monkeypatch):
    find = FakeFind({"ac.NEW1": "All Accounts", "All Entities": "All Entities", "un.SY.2TEC": "All Entities"})
    monkeypatch.setattr(ops, "live_dimension", find)
    confirm = ops.live_confirm(None)
    assert confirm("FinMax", "All Accounts", "ac.NEW1")
    assert not confirm("FinMax", "All Accounts", "un.SY.2TEC")          # exists, but in another dimension
    assert not confirm("FinMax", "All Entities", "un.SY.NOPE")          # probe (atlas top) found: a real miss
    assert find.calls[-1] == "All Entities"
    find.known.clear()                                                  # session broke: nothing is found now
    with pytest.raises(Exception):
        confirm("FinMax", "All Entities", "Jan 27")                     # probed again, not trusted from cache
    find.known["All Entities"] = "All Entities"                         # works again
    assert not confirm("FinMax", "All Entities", "un.SY.NOPE")


def test_error_scan_groups_pull_cells_and_caps_answer_lines():
    pull = [[None, "Sep 26"]] + [[f"cc{i}", "#NoAccess"] for i in range(400)]
    ans = [[None] * 3 for _ in range(60)]
    trace = {}
    for r in range(10, 60):
        ans[r - 1][2] = VALUE_ERR
        trace[(r, 3)] = {"key": f"k{r}"}
    w = answer.error_warnings("Q6 Svar", ans, trace, {"Q6 FinMax cb": pull})
    assert w[0] == ("error cells on 'Q6 FinMax cb': 400 x #NoAccess (first: B2, B3, B4, ...): "
                    "no read access or member not valid in this cube")
    assert len(w) <= 1 + answer.MAX_LINE_WARNINGS + 1
    assert w[-1] == f"... and {50 - answer.MAX_LINE_WARNINGS} more answer lines with error cells on 'Q6 Svar'"


def test_member_check_folds_norwegian_letters(fake_atlas):
    atlas.store(fake_atlas, "FinMax", "All Accounts", [_m("ac.STA20Y", "All Accounts") | {"label": "Årsverk"}], 1.0)
    assert atlas.canonical(fake_atlas, "FinMax", "All Accounts", "årsverk") == "ac.STA20Y"

# ------------------------------------------------------------------ closed month: HTTP first, COM fallback, TTL cache
class _Http:
    def __init__(self, answer):
        self.answer, self.calls = answer, 0

    def closed_month(self, conn="FinMax"):
        self.calls += 1
        if isinstance(self.answer, Exception):
            raise self.answer
        return self.answer


class _Host:
    def __init__(self):
        self.calls = 0

    def on_host(self, fn, timeout=None):
        self.calls += 1
        return fn()


def _com(monkeypatch, month="08"):
    monkeypatch.setattr(ops, "find_member", lambda h, c, n: {"info": {"ChildMbrName": f"[CurrRepMonth].[Y26M{month}]"}})


def _setup(tmp_path, monkeypatch, http):
    from smartview_mcp import config, svhttp
    monkeypatch.setattr(config, "SVKIT", tmp_path)
    monkeypatch.setattr(svhttp, "shared", lambda: http)
    return tmp_path / "closed_month.json"


def _write(path, **entry):
    import json
    path.write_text(json.dumps({"FinMax": entry}), encoding="utf-8")


def test_closed_month_reads_http_first_and_reuses_it_within_the_ttl(tmp_path, monkeypatch):
    http = _Http("Sep 26")
    _setup(tmp_path, monkeypatch, http)
    h = _Host()
    assert templates.closed_month(h) == "Sep 26" and templates.closed_month(h) == "Sep 26"
    assert http.calls == 1 and h.calls == 0                       # second call from the cache, no Excel at all


def test_an_expired_http_entry_is_asked_again(tmp_path, monkeypatch):
    import time
    http = _Http("Sep 26")
    path = _setup(tmp_path, monkeypatch, http)
    _write(path, month="Aug 26", at=time.time() - 3600, source="http")
    assert templates.closed_month(_Host()) == "Sep 26" and http.calls == 1


def test_yesterdays_com_entry_is_asked_again_but_todays_is_kept(tmp_path, monkeypatch):
    import time
    http = _Http("Sep 26")
    path = _setup(tmp_path, monkeypatch, http)
    _write(path, month="Aug 26", at=time.time() - 2 * 86400, source="com")
    assert templates.closed_month(_Host()) == "Sep 26" and http.calls == 1
    _write(path, month="Aug 26", at=time.time(), source="com")    # slow COM answer: kept for the day
    assert templates.closed_month(_Host()) == "Aug 26" and http.calls == 1


def test_the_old_cache_format_is_tolerated(tmp_path, monkeypatch):
    import datetime as dt
    import json
    http = _Http("Sep 26")
    path = _setup(tmp_path, monkeypatch, http)
    path.write_text(json.dumps({"date": dt.date.today().isoformat(), "FinMax": "Aug 26"}), encoding="utf-8")
    assert templates.closed_month(_Host()) == "Aug 26" and http.calls == 0
    path.write_text(json.dumps({"date": "2020-01-01", "FinMax": "Aug 20"}), encoding="utf-8")
    assert templates.closed_month(_Host()) == "Sep 26"
    path.write_text("[1, 2]", encoding="utf-8")                   # not even a dict
    assert templates.closed_month(_Host()) == "Sep 26"


def test_closed_month_falls_back_to_com_and_caches_it_for_the_day(tmp_path, monkeypatch):
    import json
    from smartview_mcp import svhttp
    http = _Http(svhttp.SvHttpError("signed out"))
    path = _setup(tmp_path, monkeypatch, http)
    _com(monkeypatch)
    h = _Host()
    assert templates.closed_month(h) == "Aug 26" and h.calls == 1 and http.calls == 1
    assert json.loads(path.read_text(encoding="utf-8"))["FinMax"]["source"] == "com"
    assert templates.closed_month(h) == "Aug 26" and h.calls == 1 and http.calls == 1   # neither route again


def test_http_and_com_both_failing_raises_and_caches_nothing(tmp_path, monkeypatch):
    from smartview_mcp import svhttp
    from smartview_mcp.host import HostError
    path = _setup(tmp_path, monkeypatch, _Http(svhttp.SvHttpError("signed out")))
    monkeypatch.setattr(ops, "find_member", lambda h, c, n: {"info": {}})
    with pytest.raises(HostError) as e:
        templates.closed_month(_Host())
    assert e.value.__cause__ is None and not path.exists()


def test_an_unexpected_http_error_is_logged_as_a_warning_then_com(tmp_path, monkeypatch, caplog):
    _setup(tmp_path, monkeypatch, _Http(KeyError("boom")))
    _com(monkeypatch)
    with caplog.at_level("INFO"):
        assert templates.closed_month(_Host()) == "Aug 26"
    assert [r.levelname for r in caplog.records if "closed month" in r.getMessage()] == ["WARNING"]


def test_fin_pnl_is_announced_and_division_pnl_marked_for_retirement():
    from smartview_mcp import server, tools_ws
    assert "fin_pnl(area)" in server.INSTRUCTIONS
    assert "fin_pnl" in templates.division_pnl.__doc__ and "retired" in templates.division_pnl.__doc__
    assert "to be retired: prefer fin_pnl" in tools_ws.sv_analyse.__doc__
    assert "division_pnl" in templates.TEMPLATES                      # still works until the retirement
