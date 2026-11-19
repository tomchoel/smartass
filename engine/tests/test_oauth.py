"""epm/oauth.py: the device-code sign-in, the token file and the refresh. The identity domain is faked
(oauth._post_form); DPAPI and the lock file are real."""
import threading

import pytest

from smartass.core import config
from smartass.epm import oauth

POD = "https://strat.example"
PAIR = {"access_token": "at-1", "refresh_token": "rt-1", "expires_in": 3600, "token_type": "Bearer"}   # allow-secret


class Idcs:
    """Scripted answers per path (the last one repeats); every call is recorded."""

    def __init__(self, **answers):
        self.answers = {k: list(v) for k, v in answers.items()}
        self.calls = []

    def __call__(self, path, fields, timeout=30):
        self.calls.append((path, dict(fields)))
        q = self.answers[path.rsplit("/", 1)[1]]
        a = q.pop(0) if len(q) > 1 else q[0]
        if isinstance(a, Exception):
            raise a
        return a

    def to(self, name):
        return [f for p, f in self.calls if p.endswith(name)]


@pytest.fixture
def oa(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "OAUTH_IDCS", "https://idcs.example")
    monkeypatch.setattr(config, "OAUTH_CLIENT_ID", "cid")
    monkeypatch.setattr(config, "OAUTH_SCOPES", {POD: "urn:scope"})
    monkeypatch.setattr(config, "OAUTH_DIR", tmp_path / "oauth")
    oauth._mem.clear()

    def use(**answers):
        idcs = Idcs(**answers)
        monkeypatch.setattr(oauth, "_post_form", idcs)
        return idcs
    return use


def _signed_in(now=None, **over):
    import time
    now = time.time() if now is None else now
    oauth._write(POD, {"access_token": "at-1", "exp": now + 3600, "refresh_token": "rt-1", "signed_in": now,
                       "refreshed": now, **over})
    oauth._mem.clear()


def test_off_until_idcs_client_and_scopes_are_all_set(monkeypatch, oa):
    assert oauth.configured() and oauth.pods() == [POD] and oauth.scope_for(POD + ":443/HyperionPlanning") == "urn:scope"
    assert oauth.scope_for("https://plan1.example") is None
    monkeypatch.setattr(config, "OAUTH_CLIENT_ID", "")
    assert not oauth.configured() and oauth.pods() == [] and oauth.scope_for(POD) is None
    with pytest.raises(oauth.NotSignedIn):
        oauth.access_token(POD)
    assert oauth.status() == []


def test_scopes_are_keyed_by_the_pod_origin_and_junk_is_left_out():
    got = config.oauth_scopes({"HTTPS://Strat.Example:443/HyperionPlanning": " urn:a ", "not a url": "urn:b",
                               "https://empty.example": "", "https://n.example": 5})
    assert got == {"https://strat.example": "urn:a"}
    assert config.oauth_scopes(None) == {} and config.oauth_scopes("x") == {}


def test_login_asks_for_offline_access_polls_until_approved_and_stores_the_pair(oa, tmp_path):
    idcs = oa(device=[(200, {"device_code": "dc", "user_code": "ABCD", "verification_uri": "https://idcs.example:443/ui",
                             "expires_in": 300, "interval": 5})],
              token=[(400, {"error": "authorization_pending"}), (400, {"error": "slow_down"}), (200, PAIR)])
    d = oauth.start_login(POD)
    assert idcs.to("device") == [{"response_type": "device_code", "scope": "urn:scope offline_access", "client_id": "cid"}]
    slept = []
    oauth.finish_login(POD, d, sleep=slept.append)
    assert slept == [5, 5, 10]                                    # slow_down: wait longer
    assert idcs.to("token")[0] == {"grant_type": oauth.DEVICE_GRANT, "device_code": "dc", "client_id": "cid"}
    assert oauth.status()[0]["state"] == "signed_in"
    assert oauth.access_token(POD) == "at-1" and len(idcs.to("token")) == 3      # no refresh needed
    raw = (tmp_path / "oauth" / "strat.example.bin").read_bytes()
    assert b"at-1" not in raw and b"rt-1" not in raw             # DPAPI, not plain text


def test_a_sign_in_page_on_another_host_is_never_handed_out(oa):
    oa(device=[(200, {"device_code": "dc", "user_code": "ABCD", "verification_uri": "https://evil.example/ui"})])
    with pytest.raises(oauth.OAuthError, match="not on the configured identity domain"):
        oauth.start_login(POD)


@pytest.mark.parametrize("answers, msg", [
    ([(400, {"error": "access_denied", "error_description": "no"})], "refused .access_denied - no"),
    ([(400, {"error": "authorization_pending"})], "ended before it was approved"),
    ([(200, {"access_token": "at-1", "expires_in": 3600})], "no refresh token"),   # allow-secret
])
def test_a_login_that_does_not_complete_stores_nothing(oa, answers, msg):
    oa(token=answers)
    t = iter(range(0, 10_000, 100))
    with pytest.raises(oauth.OAuthError, match=msg):
        oauth.finish_login(POD, {"device_code": "dc", "expires_in": 300}, sleep=lambda s: None, clock=lambda: next(t))
    assert oauth.status()[0]["state"] == "none"


def test_no_code_from_the_identity_domain_is_an_oauth_error(oa):
    oa(device=[(400, {"error": "invalid_scope", "error_description": "Invalid scope."})])
    with pytest.raises(oauth.OAuthError, match="invalid_scope - Invalid scope"):
        oauth.start_login(POD)
    with pytest.raises(oauth.OAuthError, match="no OAuth scope"):
        oauth.start_login("https://plan1.example")


def test_a_token_near_its_end_is_refreshed_and_the_new_refresh_token_is_the_one_used_next(oa):
    import time
    idcs = oa(token=[(200, {"access_token": "at-2", "refresh_token": "rt-2", "expires_in": 100}),   # allow-secret
                     (200, {"access_token": "at-3", "refresh_token": "rt-3", "expires_in": 3600})])  # allow-secret
    _signed_in(exp=time.time() + 60)                               # inside the margin
    assert oauth.access_token(POD) == "at-2"
    assert idcs.to("token") == [{"grant_type": "refresh_token", "client_id": "cid", "refresh_token": "rt-1"}]
    assert oauth.access_token(POD) == "at-3"                      # 100 s is inside the margin too
    assert idcs.to("token")[1]["refresh_token"] == "rt-2"         # single use: never rt-1 again
    assert oauth.access_token(POD) == "at-3" and len(idcs.calls) == 2      # held in memory now


def test_a_refused_token_is_never_handed_out_again(oa):
    idcs = oa(token=[(200, {"access_token": "at-2", "refresh_token": "rt-2", "expires_in": 3600})])   # allow-secret
    _signed_in()
    assert oauth.access_token(POD) == "at-1"
    assert oauth.access_token(POD, reject="at-1") == "at-2" and len(idcs.calls) == 1
    # another process refreshed meanwhile: its token is taken from the file, without a second refresh
    oauth._mem[POD] = ("at-old", oauth._mem[POD][1])
    assert oauth.access_token(POD, reject="at-old") == "at-2" and len(idcs.calls) == 1


def test_a_refused_refresh_ends_the_sign_in_and_is_not_asked_again(oa):
    import time
    idcs = oa(token=[(400, {"error": "invalid_grant", "error_description": "The token has already been consumed"})],
              device=[(200, {"device_code": "dc", "user_code": "ABCD", "verification_uri": "https://idcs.example/ui"})])
    _signed_in(exp=time.time() - 1)
    with pytest.raises(oauth.NotSignedIn, match="already been consumed.*smartass login"):
        oauth.access_token(POD)
    with pytest.raises(oauth.NotSignedIn, match="has ended"):
        oauth.access_token(POD)
    assert len(idcs.calls) == 1
    st = oauth.status()[0]
    assert st["state"] == "ended" and "consumed" in st["why"]
    idcs.answers["token"] = [(200, PAIR)]                          # signing in again clears it
    oauth.finish_login(POD, oauth.start_login(POD), sleep=lambda s: None)
    assert oauth.status()[0]["state"] == "signed_in" and oauth.access_token(POD) == "at-1"


@pytest.mark.parametrize("answer", [(503, {}), (429, {"error": "too_many_requests"}), (400, {}),
                                    oauth.OAuthError("the identity domain did not answer (URLError)")])
def test_a_refresh_that_could_not_be_done_keeps_the_sign_in(oa, answer):
    import time
    oa(token=[answer])
    _signed_in(exp=time.time() - 1)
    with pytest.raises(oauth.OAuthError) as e:
        oauth.access_token(POD)
    assert not isinstance(e.value, oauth.NotSignedIn)
    assert oauth.status()[0]["state"] == "signed_in"


def test_many_threads_refresh_once(oa, monkeypatch):
    import time
    gate = threading.Event()

    def slow(path, fields, timeout=30):
        calls.append(fields["refresh_token"])
        gate.wait(5)
        return 200, {"access_token": "at-2", "refresh_token": "rt-2", "expires_in": 3600}   # allow-secret
    calls, got = [], []
    oa(token=[(500, {})])
    monkeypatch.setattr(oauth, "_post_form", slow)
    try:
        _signed_in(exp=time.time() - 1)
        ts = [threading.Thread(target=lambda: got.append(oauth.access_token(POD))) for _ in range(8)]
        [t.start() for t in ts]
        time.sleep(0.2)
        gate.set()
        [t.join(10) for t in ts]
    finally:
        gate.set()
    assert calls == ["rt-1"] and got == ["at-2"] * 8


def test_a_lock_that_cannot_be_had_is_an_oauth_error_not_an_unlocked_refresh(oa, monkeypatch):
    import time
    idcs = oa(token=[(200, PAIR)])
    _signed_in(exp=time.time() - 1)
    monkeypatch.setattr(oauth, "LOCK_WAIT_S", 0.2)
    held, release = threading.Event(), threading.Event()

    def hold():
        with oauth._locked(POD):
            held.set()
            release.wait(10)
    t = threading.Thread(target=hold)
    t.start()
    held.wait(5)
    try:
        with pytest.raises(oauth.OAuthError, match="still renewing"):
            oauth.access_token(POD)
    finally:
        release.set()
        t.join(10)
    assert idcs.calls == []


def test_keep_fresh_renews_only_sign_ins_older_than_a_day(oa):
    import time
    idcs = oa(token=[(200, {"access_token": "at-2", "refresh_token": "rt-2", "expires_in": 3600})])   # allow-secret
    _signed_in()
    assert oauth.keep_fresh() == [] and idcs.calls == []
    _signed_in(refreshed=time.time() - 2 * 86400)                  # the access token is still good; the age decides
    assert oauth.keep_fresh() == [f"{POD}: renewed"] and idcs.to("token")[0]["refresh_token"] == "rt-1"
    assert oauth.status()[0]["refreshed_days_ago"] == 0.0
    oauth._write(POD, {"refresh_token": "rt-9", "dead": "ended", "refreshed": 0})
    assert oauth.keep_fresh() == [] and len(idcs.calls) == 1       # an ended sign-in is left to `smartass login`


def test_status_counts_down_the_week_and_never_shows_a_token(oa):
    import time
    oa(token=[(500, {})])
    assert oauth.status() == [{"pod": POD, "state": "none"}]
    _signed_in(refreshed=time.time() - 2.5 * 86400)
    st = oauth.status()[0]
    assert (st["state"], st["refreshed_days_ago"], st["days_left_unused"]) == ("signed_in", 2.5, 4.5)
    assert "at-1" not in str(st) and "rt-1" not in str(st)
    _signed_in(refreshed=time.time() - 8 * 86400)
    assert oauth.status()[0]["state"] == "ended"


def test_forget_deletes_the_sign_in(oa):
    oa(token=[(500, {})])
    _signed_in()
    assert oauth.access_token(POD) == "at-1"
    assert oauth.forget(POD) is True and oauth.forget(POD) is False
    with pytest.raises(oauth.NotSignedIn, match="not signed in"):
        oauth.access_token(POD)
