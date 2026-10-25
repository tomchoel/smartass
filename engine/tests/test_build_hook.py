"""tools_ws._build_job: the before_answer hook sees the pulls by id and can change the spec (fin_pnl's checks)."""
from types import SimpleNamespace as NS

from smartview_mcp import tools_ws, userexcel


def test_before_answer_gets_the_pulls_by_id_and_its_spec_is_built(monkeypatch):
    t = NS(path="C:/x/a.xlsx", created=False)
    monkeypatch.setattr(userexcel, "resolve_target", lambda h, target, allow, title=None: t)
    monkeypatch.setattr(userexcel, "next_question", lambda t: 3)
    monkeypatch.setattr(userexcel, "pull", lambda h, t, q, conn, pov, rows, cols, title=None: NS(
        sheet=title, parsed=NS(pov_problems=[], unknown_members=[])))
    monkeypatch.setattr(userexcel, "log_source", lambda *a, **k: None)
    monkeypatch.setattr(tools_ws, "_tidy", lambda *a, **k: None)
    built = {}

    def answer_job(t, q, spec, ids, sess, tm, warns):
        built.update(spec=spec, ids=ids)
        return {"answer_sheet": f"Q{q} Svar", "warnings": []}
    monkeypatch.setattr(tools_ws, "_answer_job", answer_job)
    seen = {}

    def hook(spec, pulls):
        seen.update(pulls)
        return {**spec, "checks": [{"label": "warehouse = sheet", "static": "OK"}]}
    pull = tools_ws.PullIn(id="m", connection="FinMax", rows=[{"dimension": "All Accounts", "members": ["ac.C30"]}],
                           columns=[{"dimension": "All Time Periods", "members": ["Sep 26"]}], title="Q3 FinMax")
    tm = NS(summary=lambda: {})
    tools_ws._build_job(None, tm, "q", [pull], {"columns": []}, "session", False, None, before_answer=hook)
    assert set(seen) == {"m"} and seen["m"].sheet == "Q3 FinMax"
    assert built["spec"]["checks"] == [{"label": "warehouse = sheet", "static": "OK"}]
    assert built["ids"] == {"m": "Q3 FinMax"}
