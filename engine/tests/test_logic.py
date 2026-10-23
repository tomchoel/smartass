"""Unit tests for the analysis logic that needs no Excel."""
from types import SimpleNamespace

from smartview_mcp.grid import Axis, Layout
from smartview_mcp.tools_ws import fin
from smartview_mcp.templates import _clean, _fmt, _pct, months_to, same_period_ly, split_period
from smartview_mcp.userexcel import _verify_labels, is_archive, local_path


def test_fin_matches_sheet_format():
    assert fin(-1112834010.34) == "(1 112 834 010)"
    assert fin(59752923.67) == "59 752 924"
    assert fin(0) == "–" and fin(0.2) == "–"
    assert fin(544.2, "tall1") == "544,2"
    assert fin(0.13, "prosent1") == "+13,0 %" and fin(-0.044, "prosent1") == "-4,4 %"


def test_text_amounts_use_minus_never_double_parentheses():
    assert _fmt(-151328997) == "−151 328 997"
    assert "(" not in _fmt(-1) and _fmt(1234) == "1 234"
    assert _pct(954.0, 1000.0) == "−4,6 %"          # (current − comparison) / |comparison|
    assert _pct(-1112.8, -1066.4) == "−4,4 %"       # costs/EBITDA negative: worse = negative


def test_periods():
    assert split_period("Sep 26") == (9, 26)
    assert same_period_ly("Jan 26") == "Jan 25"
    assert months_to("Mar 26") == ["Jan 26", "Feb 26", "Mar 26"]


def test_clean_labels():
    assert _clean("C30 REVENUES (ac)") == "C30 REVENUES"
    assert _clean("EBITDA EBITDA (ac)") == "EBITDA"
    assert _clean("2TEC Technology (un.SY)") == "2TEC Technology"


def test_verify_labels_catches_shifted_rows():
    lay = Layout([], [Axis("All Accounts", ["ac.C30", "ac.C40S"])], [Axis("All Time Periods", ["Sep 26"])])
    ok = SimpleNamespace(row_labels=[["C30 REVENUES (ac)"], ["C40S Total Cost (ac)"]], col_labels=[["Sep 26"]])
    shifted = SimpleNamespace(row_labels=[["C40S Total Cost (ac)"], ["C48 Contribution (ac)"]], col_labels=[["Sep 26"]])
    assert _verify_labels(lay, ok) == []
    assert len(_verify_labels(lay, shifted)) == 2


def test_archive_guard_local_url_and_unsaved(monkeypatch, tmp_path):
    from smartview_mcp import config
    root = tmp_path / "Analyse (Telenor Norge AS) - Produktregnskapet"
    monkeypatch.setattr(config, "ARCHIVE_ROOT", root)
    assert is_archive(str(root / "2025" / "x.xlsx"))
    assert is_archive("https://telenorgroup.sharepoint.com/sites/AnalyseTelenorNorgeAS/Shared Documents/2025/x.xlsx")
    assert not is_archive(str(tmp_path / "analyser" / "a.xlsx"))
    assert not is_archive("Book1")


def test_local_path_maps_personal_onedrive(monkeypatch):
    from smartview_mcp import userexcel
    monkeypatch.setattr(userexcel, "_onedrive_mounts",
                        lambda: [("https://contoso-my.sharepoint.com/personal/a_b/Documents/", r"C:\OD")])
    url = "https://contoso-my.sharepoint.com/personal/a_b/Documents/smartass/analyser/a%20b.xlsx"
    assert local_path(url) == r"C:\OD\smartass\analyser\a b.xlsx"


def test_axes_in_accepts_dict_list_and_models():
    from smartview_mcp.server import AxisIn, axes_in
    d = axes_in({"All Accounts": ["ac.C30"], "All Time Periods": "Sep 26"})
    assert [(a.dimension, a.members) for a in d] == [("All Accounts", ["ac.C30"]), ("All Time Periods", ["Sep 26"])]
    li = axes_in([{"dimension": "All Accounts", "members": ["x"]}, AxisIn(dimension="E", members=["y"])])
    assert [a.dimension for a in li] == ["All Accounts", "E"]


def test_pull_in_accepts_dict_rows():
    from smartview_mcp.tools_ws import PullIn
    p = PullIn(id="m", connection="FinMax", rows={"All Accounts": ["ac.C30"]}, columns={"All Time Periods": "Sep 26"})
    assert p.rows[0].members == ["ac.C30"] and p.columns[0].dimension == "All Time Periods"


def test_axes_in_wraps_a_single_axis_object():
    from smartview_mcp.server import axes_in
    a = axes_in({"dimension": "All Accounts", "members": ["ac.C30"]})
    assert [(x.dimension, x.members) for x in a] == [("All Accounts", ["ac.C30"])]


def test_slow_stage_sends_progress_notes_and_stops_after(monkeypatch):
    import time as _t

    from smartview_mcp import timing
    monkeypatch.setattr(timing, "PROGRESS_AFTER_S", 0.05)
    notes = []
    timing.Timer.note_notify(notes.append)
    timing.Timer.start()
    with timing.stage("open target"):
        with timing.stage("slow"):
            _t.sleep(0.13)
    n = len(notes)
    _t.sleep(0.12)
    assert n >= 1 and "open target.slow" in notes[0] and notes[0].startswith("still working: ")
    assert len(notes) == n                       # no notes once the stages are done
    timing.Timer.start()                         # the next timer on this thread has no notify of its own
    assert timing.Timer.current().notify is None


def test_progress_job_clears_its_notify_after_the_job():
    import asyncio

    from smartview_mcp import server, timing

    async def run():
        server._ctx.set(object())                 # any ctx: notify is only called by a ticker
        job = server._with_progress(lambda h: timing.Timer.start())
        t = job(None)
        return t
    t = asyncio.run(run())
    assert t.notify is None and getattr(timing._local, "pending_notify", None) is None


def test_wait_login_recovers_an_ie_fallback_login_by_itself(monkeypatch):
    from smartview_mcp import dialogs, userexcel
    from smartview_mcp.dialogs import Dialog
    state = {"ie": True, "recovered": 0}
    clock = {"t": 0.0}
    monkeypatch.setattr(userexcel.time, "monotonic", lambda: clock["t"])

    def scan(pid):
        return [Dialog(2, "#32770", "Login", "", "login", ie=True)] if state["ie"] else []

    def recover(pid, hwnds):
        state["recovered"] += 1
        state["ie"] = False                       # Smart View reopens the login in WebView2; the watcher clicks it
        return {"ended": [10], "closed": hwnds}
    monkeypatch.setattr(dialogs, "scan", scan)
    monkeypatch.setattr(userexcel, "recover_ie_login", recover)
    monkeypatch.setattr(userexcel, "login_open", lambda pid: state["ie"])
    monkeypatch.setattr(userexcel.time, "sleep", lambda s: clock.__setitem__("t", clock["t"] + s))
    userexcel.wait_login_cleared(1, limit=60)
    assert state["recovered"] == 1


def test_wait_login_after_recovery_waits_for_the_new_login_and_spaces_retries(monkeypatch):
    from smartview_mcp import dialogs, userexcel
    from smartview_mcp.dialogs import Dialog
    clock = {"t": 0.0}
    monkeypatch.setattr(userexcel.time, "monotonic", lambda: clock["t"])
    monkeypatch.setattr(userexcel.time, "sleep", lambda s: clock.__setitem__("t", clock["t"] + s))
    # IE login until 1 s; then nothing for 5 s (Smart View has not reopened yet); a new login 6-9 s; then signed in
    def login_at(t):
        return t < 1 or 6 <= t < 9
    monkeypatch.setattr(userexcel, "login_open", lambda pid: login_at(clock["t"]))
    monkeypatch.setattr(dialogs, "scan", lambda pid: [Dialog(2, "#32770", "Login", "", "login", ie=True)]
                        if clock["t"] < 1 else [])
    rec = []
    monkeypatch.setattr(userexcel, "recover_ie_login", lambda pid, h: rec.append(clock["t"]) or {"ended": [], "closed": h})
    userexcel.wait_login_cleared(1, limit=60)
    assert clock["t"] >= 9 and len(rec) == 1      # did not return in the gap before the new login


def test_smart_view_box_guard_dismisses_each_box_once_and_keeps_its_text(monkeypatch):
    import time as _t

    from smartview_mcp import dialogs, userexcel
    from smartview_mcp.dialogs import Dialog
    box = Dialog(5, "#32770", "Oracle Smart View for Office", "Empty response from provider.", "message")
    pressed = []
    monkeypatch.setattr(dialogs, "scan", lambda pid: [box] if 5 not in pressed else [])
    monkeypatch.setattr(dialogs, "press_ok", lambda h: pressed.append(h) or True)
    monkeypatch.setattr(userexcel, "BOX_POLL_S", 0.01)
    with userexcel.SmartViewBoxes(1) as g:
        _t.sleep(0.1)
    assert pressed == [5] and g.texts == ["Empty response from provider."]
