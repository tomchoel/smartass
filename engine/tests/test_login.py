"""`smartass login` (ops/login.py), the doctor's oauth check and the servers' background renewal. epm/oauth.py is
faked."""
import pytest

from smartass.core import config
from smartass.epm import oauth
from smartass.ops import cli, doctor, login

STRAT, PLAN1 = "https://strat.example", "https://plan1.example"


@pytest.fixture
def fake(monkeypatch):
    """oauth with two pods; state is a dict pod -> 'none' | 'signed_in' | 'ended' the test edits."""
    state = {PLAN1: "signed_in", STRAT: "none"}
    calls = []
    monkeypatch.setattr(oauth, "configured", lambda: True)
    monkeypatch.setattr(oauth, "pods", lambda: sorted(state))
    monkeypatch.setattr(oauth, "status", lambda: [
        {"pod": p, "state": s, **({"refreshed_days_ago": 2.0, "days_left_unused": 5.0} if s != "none" else {})}
        for p, s in sorted(state.items())])
    monkeypatch.setattr(oauth, "start_login", lambda pod: calls.append(("start", pod)) or {
        "device_code": "dc", "user_code": "ABCD-1234", "verification_uri": "https://idcs.example/ui/v1/device",
        "expires_in": 300})

    def finish(pod, device):
        calls.append(("finish", pod))
        state[pod] = "signed_in"
    monkeypatch.setattr(oauth, "finish_login", finish)
    monkeypatch.setattr(login, "_no_atlas", lambda: False)        # an atlas is there (the build is tested below)
    return state, calls


def test_login_signs_in_only_where_needed_opens_the_page_and_copies_the_code(fake):
    state, calls = fake
    out, opened, copied = [], [], []
    assert login.run(out=out.append, open_page=opened.append, copy=copied.append) == 0
    assert calls == [("start", STRAT), ("finish", STRAT)]
    assert opened == ["https://idcs.example/ui/v1/device"] and copied == ["ABCD-1234"]
    text = "\n".join(out)
    assert "code: ABCD-1234" in text and "Signed in to strat.example." in text and "plan1" not in text
    out.clear()
    assert login.run(out=out.append, open_page=opened.append, copy=copied.append) == 0
    assert len(calls) == 2 and "signed in (renewed 2 days ago; lasts 5 more days without use)" in out[0]


def test_login_force_signs_in_everywhere_and_a_failure_is_exit_1(fake, monkeypatch):
    state, calls = fake

    def refuse(pod, device):
        raise oauth.OAuthError("the sign-in code ended before it was approved")
    monkeypatch.setattr(oauth, "finish_login", refuse)
    out = []
    assert login.run(force=True, out=out.append, open_page=lambda u: None, copy=lambda c: None) == 1
    assert [c for c in calls if c[0] == "start"] == [("start", PLAN1), ("start", STRAT)]
    assert sum("failed: the sign-in code ended" in line for line in out) == 2


def test_login_says_how_to_set_it_up_when_it_is_not(monkeypatch):
    monkeypatch.setattr(oauth, "configured", lambda: False)
    for fn in (login.run, login.show, login.forget):
        out = []
        assert fn(out=out.append) == 2 and "[oauth]" in out[0]


def test_status_and_forget_lines(fake, monkeypatch):
    state, _ = fake
    state[STRAT] = "ended"
    out = []
    assert login.show(out=out.append) == 0
    assert out == ["plan1.example: signed in (renewed 2 days ago; lasts 5 more days without use)",
                   "strat.example: the sign-in has ended (not used for a week)"]
    monkeypatch.setattr(oauth, "forget", lambda pod: pod == PLAN1)
    out.clear()
    assert login.forget(out=out.append) == 0
    assert out == ["plan1.example: sign-in deleted from this PC", "strat.example: no sign-in stored"]


def test_the_cli_routes_login_and_its_flags(monkeypatch):
    seen = []
    monkeypatch.setattr(login, "run", lambda force=False, atlas=True: seen.append(("run", force, atlas)) or 0)
    monkeypatch.setattr(login, "show", lambda: seen.append("show") or 0)
    monkeypatch.setattr(login, "forget", lambda: seen.append("forget") or 0)
    argvs = (["login"], ["login", "--force"], ["login", "--status"], ["login", "--forget"], ["login", "--no-atlas"])
    assert [cli.main(a) for a in argvs] == [0] * 5
    assert seen == [("run", False, True), ("run", True, True), "show", "forget", ("run", False, False)]


def test_doctor_oauth_fails_when_not_set_up_or_the_home_scope_is_missing_and_warns_when_a_sign_in_is(fake, monkeypatch):
    state, _ = fake
    monkeypatch.setattr(config, "EPM_BASE", "https://plan1.example")
    c = doctor._oauth()
    assert (c.status, c.detail) == ("WARN", "plan1: signed in; strat: not signed in") and "smartass login" in c.fix
    state[STRAT] = "signed_in"
    assert doctor._oauth().status == "OK"
    state[STRAT] = "ended"
    assert "strat: sign-in ended" in doctor._oauth().detail and doctor._oauth().status == "WARN"
    del state[PLAN1]                       # the home pod has no scope: no answer at all (the closed month is read there)
    state[STRAT] = "signed_in"
    c = doctor._oauth()
    assert (c.status, c.detail) == ("FAIL", "strat: signed in; plan1: no scope (epm.base: no answers)")
    assert "[oauth.scopes]" in c.fix and "https://plan1.example" in c.fix and "smartass login" in c.fix
    assert "administrator" in c.fix
    state[STRAT] = "not_signed_in"                 # a missing sign-in elsewhere does not hide the FAIL
    assert doctor._oauth().status == "FAIL"
    monkeypatch.setattr(oauth, "configured", lambda: False)
    c = doctor._oauth()
    assert c.status == "FAIL" and "not set up" in c.detail and c.fix == oauth.NOT_SET_UP
    assert doctor._oauth in doctor.CHECKS


def test_the_servers_renew_the_sign_ins_now_and_every_six_hours_and_never_raise(monkeypatch):
    rounds, slept = [], []

    def keep_fresh():
        rounds.append(1)
        if len(rounds) == 2:
            raise RuntimeError("boom")
        return ["https://strat.example: renewed"]
    monkeypatch.setattr(oauth, "keep_fresh", keep_fresh)
    oauth.run_keep_fresh(sleep=slept.append, rounds=3)
    assert len(rounds) == 3 and slept == [6 * 3600] * 3


def test_the_renewal_thread_only_starts_when_the_route_is_set_up(monkeypatch):
    from smartass.core import serverkit
    started = []
    monkeypatch.setattr(oauth, "run_keep_fresh", lambda: started.append(1))
    monkeypatch.setattr(oauth, "configured", lambda: False)
    assert oauth.start_keep_fresh() is None and started == []
    monkeypatch.setattr(oauth, "configured", lambda: True)
    t = oauth.start_keep_fresh()
    t.join(5)
    assert started == [1] and t.daemon and t.name == "oauth-keep-fresh"
    serverkit._keep_signed_in()                                    # the MCP server's start-up hook uses the same
    assert len(started) == 2 or t.join(1) is None

def test_doctor_says_when_the_identity_domain_is_refused_for_not_being_https(fake, monkeypatch):
    monkeypatch.setattr(config, "OAUTH_IDCS", "")
    monkeypatch.setattr(config, "OAUTH_IDCS_REFUSED", True)
    monkeypatch.setattr(oauth, "configured", lambda: False)
    c = doctor._oauth()
    assert (c.status, c.detail, c.fix) == ("FAIL", "not set up: [oauth] idcs is not an https URL", oauth.NOT_SET_UP)


# ---------------------------------------------------------------- the atlas built at the first login (Task 2)


@pytest.fixture
def no_atlas(fake, monkeypatch):
    """A PC with no atlas; the build is a fake (conftest makes the real one a violation), its rows as run_build's."""
    have = {"atlas": False}
    monkeypatch.setattr(login, "_no_atlas", lambda: not have["atlas"])
    built = []

    def build():
        built.append(1)
        have["atlas"] = True
        return [{"connection": "FinMax", "dimension": "All Accounts", "members": 10, "seconds": 1.0}]
    return have, built, build


def test_login_builds_the_atlas_when_there_is_none_saying_what_and_how_long(no_atlas):
    have, built, build = no_atlas
    out = []
    assert login.run(out=out.append, open_page=lambda u: None, copy=lambda c: None, build=build) == 0
    text = "\n".join(out)
    assert built == [1] and "Signed in to strat.example." in text
    assert out.index("Signed in to strat.example.") < next(i for i, s in enumerate(out) if "No atlas" in s)
    assert "half a minute" in text and "--no-atlas" in text and "atlas: 1 built, 0 failed" in text
    assert any(s.startswith("Atlas built in ") and s.endswith(" s.") for s in out)


def test_login_builds_when_every_pod_is_already_signed_in_and_never_over_an_atlas(no_atlas, fake):
    have, built, build = no_atlas
    state, calls = fake
    state["https://strat.example"] = "signed_in"
    assert login.run(out=lambda s: None, build=build) == 0 and built == [1] and calls == []
    assert login.run(out=lambda s: None, build=build) == 0 and built == [1]          # there now: never rebuilt


def test_login_no_atlas_skips_the_build(no_atlas):
    have, built, build = no_atlas
    out = []
    assert login.run(atlas=False, out=out.append, open_page=lambda u: None, copy=lambda c: None, build=build) == 0
    assert built == [] and not any("No atlas" in s for s in out)


def test_a_failed_sign_in_builds_nothing_and_says_when_the_atlas_comes(no_atlas, monkeypatch):
    have, built, build = no_atlas

    def refuse(pod, device):
        raise oauth.OAuthError("the sign-in code ended before it was approved")
    monkeypatch.setattr(oauth, "finish_login", refuse)
    out = []
    assert login.run(out=out.append, open_page=lambda u: None, copy=lambda c: None, build=build) == 1
    assert built == [] and "built after every sign-in has succeeded" in out[-1] and "smartass login" in out[-1]


@pytest.mark.parametrize("failure", ["raises", "not replaced"])
def test_a_failed_build_keeps_the_sign_in_and_says_which_command_to_run_later(no_atlas, failure):
    def build():
        if failure == "raises":
            raise RuntimeError("EPM did not answer")
        return [{"connection": "-", "error": "atlas not replaced: 1 dimension or cube failed"}]
    out = []
    assert login.run(out=out.append, open_page=lambda u: None, copy=lambda c: None, build=build) == 0
    assert "Signed in to strat.example." in out
    assert "your sign-in is kept" in out[-1] and "`smartass atlas build`" in out[-1]


def test_status_and_forget_never_build(no_atlas):
    have, built, build = no_atlas
    login.show(out=lambda s: None)
    login.forget(out=lambda s: None)
    assert built == []
