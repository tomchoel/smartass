"""OAuth 2 sign-in to EPM Cloud for the bearer route (svhttp.BearerTransport): Oracle's device-code flow against the
identity domain, with a public client an administrator registered there (Oracle: 'Authentication with OAuth 2 - Only
for OCI (Gen 2) Environments', REST APIs for Oracle Fusion Cloud EPM). Live spike 2026-11-19 on the FinPlan pod: the
access token is accepted by /HyperionPlanning/SmartView as well as by the REST API; grids (svhttp.BearerTransport)
and drill lines (drill_token over epmhttp) go over it, with no browser.

What the code rests on (Oracle's page and the spike):
- a token covers ONE environment (its scope names the service instance), so there is one sign-in and one token file
  per pod;
- the access token lasts an hour. The refresh token lasts 7 days, works once and is replaced by every refresh: the
  refresh runs under a cross-process lock and the new token is on disk before anything else happens (two smartass
  servers run at once);
- the user signs in in their own browser (Telenor SSO); smartass never sees a password.

A tool call that meets NotSignedIn answers with sign_in_message(): the page and a code per pod (pending_logins starts
the device login and finishes it in the background), so nothing waits for a browser.

Needs [oauth] idcs, client_id and [oauth.scopes] (pod base URL = scope) in config.toml; without them a tool call
answers with NOT_SET_UP. Tokens are DPAPI-protected files under HOME\\oauth (this Windows user) and are never logged or
returned by status()."""
from __future__ import annotations

import contextlib
import json
import logging
import math
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from urllib.parse import urlsplit

from ..core import config
from . import drillcore

log = logging.getLogger(__name__)

DEVICE_GRANT = "urn:ietf:params:oauth:grant-type:device_code"
REFRESH_DAYS = 7              # Oracle's default refresh-token lifetime (an identity-domain admin can change it)
MARGIN_S = 120                # an access token this close to its end is replaced
LOCK_WAIT_S = 40.0            # a refresh in another process is one request: longer than this means it is stuck
ENDED = ("invalid_grant", "invalid_client", "unauthorized_client")   # refusals that end a sign-in (RFC 6749 5.2)
_FORM = {"Content-Type": "application/x-www-form-urlencoded;charset=utf-8", "Accept": "application/json"}

_mem: dict[str, tuple[str, float]] = {}           # pod -> (access token, its end as time.time()); this process only
_tlocks: dict[str, threading.Lock] = {}
_tlocks_guard = threading.Lock()


class OAuthError(Exception):
    """The bearer route cannot authenticate right now (svhttp.BearerTransport makes it an SvHttpError)."""


class NotSignedIn(OAuthError):
    """This pod has no usable sign-in. It reaches the tool call as it is, which answers with sign_in_reply() (the
    page and a code); `smartass login` does the same in a terminal."""


NOT_SET_UP = ("EPM access is not set up on this PC: config.toml needs [oauth] idcs (the identity domain's URL) and "
              "client_id, and [oauth.scopes] with one scope per EPM pod; the identity domain's administrator gives "
              "these values (docs\\troubleshooting.md, 'oauth'). Next step: add them to config.toml, then run "
              "`smartass login`.")


class NotSetUp(NotSignedIn):
    """No [oauth] set-up at all on this PC: no code can be offered, the reply is what config.toml needs."""

    def __init__(self, msg: str = NOT_SET_UP):
        super().__init__(msg)


class LoginDeclined(OAuthError):
    """The user (or the identity domain) declined the device code."""


class LoginExpired(OAuthError):
    """The device code ended before it was approved."""


def origin(base: str) -> str:
    """'https://host' of a pod base URL (lower case, no port or path); '' when it has no host."""
    host = (urlsplit(str(base or "").strip()).hostname or "").lower()
    return f"https://{host}" if host else ""


def configured() -> bool:
    return bool(config.OAUTH_IDCS and config.OAUTH_CLIENT_ID and config.OAUTH_SCOPES)


def pods() -> list[str]:
    """The pods a scope is configured for."""
    return sorted(config.OAUTH_SCOPES) if configured() else []


def scope_for(base: str) -> str | None:
    return config.OAUTH_SCOPES.get(origin(base)) if configured() else None


# ---------------------------------------------------------------- the identity domain
class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None


def _post_form(path: str, fields: dict, timeout: float = 30) -> tuple[int, dict]:
    """POST a form to the identity domain: (HTTP status, the JSON object or {}). No answer at all is an OAuthError."""
    req = urllib.request.Request(config.OAUTH_IDCS + path, data=urllib.parse.urlencode(fields).encode(),
                                 headers=_FORM, method="POST")
    try:
        with urllib.request.build_opener(_NoRedirect).open(req, timeout=timeout) as r:
            status, text = r.status, r.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        status, text = e.code, e.read().decode("utf-8", "replace")
    except OSError as e:                                    # URLError, timeouts, TLS
        raise OAuthError(f"the identity domain did not answer ({type(e).__name__})") from None
    try:
        j = json.loads(text)
    except ValueError:
        j = {}
    return status, j if isinstance(j, dict) else {}


def _why(status: int, j: dict) -> str:
    return " - ".join(str(j[k]) for k in ("error", "error_description") if j.get(k))[:200] or f"HTTP {status}"


# ---------------------------------------------------------------- the token file
def _file(pod: str) -> Path:
    return Path(config.OAUTH_DIR) / f"{urlsplit(pod).hostname}.bin"


def _read(pod: str) -> dict | None:
    """The pod's stored sign-in; None when there is none or it cannot be unprotected (another user or machine)."""
    try:
        blob = json.loads(config._dpapi().CryptUnprotectData(_file(pod).read_bytes(), None, None, None, 0)[1])
    except Exception:   # noqa: BLE001 - OSError, pywintypes.error, bad bytes: all mean "no usable sign-in"
        return None
    return blob if isinstance(blob, dict) else None


def _write(pod: str, blob: dict) -> None:
    f = _file(pod)
    f.parent.mkdir(parents=True, exist_ok=True)
    tmp = f.with_name(f.name + ".tmp")
    tmp.write_bytes(config._dpapi().CryptProtectData(json.dumps(blob).encode("utf-8"), "smartass epm sign-in",
                                                     None, None, None, 0))
    tmp.replace(f)


@contextlib.contextmanager
def _locked(pod: str, wait_s: float | None = None):
    """One reader-refresher-writer of the pod's token file at a time, across threads and processes (msvcrt byte
    lock on <file>.lock; Windows frees it when the holder dies). Unlike a PID file, a refresh must never run
    without the lock - the refresh token works once - so a lock that cannot be had is an OAuthError."""
    import msvcrt
    wait_s = LOCK_WAIT_S if wait_s is None else wait_s
    with _tlocks_guard:
        tlock = _tlocks.setdefault(pod, threading.Lock())
    if not tlock.acquire(timeout=wait_s):
        raise OAuthError("another query is still renewing the EPM sign-in")
    try:
        f = _file(pod)
        f.parent.mkdir(parents=True, exist_ok=True)
        with open(f.with_name(f.name + ".lock"), "a+b") as fh:
            t0 = time.monotonic()
            while True:
                try:
                    fh.seek(0)
                    msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
                    break
                except OSError:
                    if time.monotonic() - t0 > wait_s:
                        raise OAuthError("another smartass process is still renewing the EPM sign-in") from None
                    time.sleep(0.05)
            try:
                yield
            finally:
                try:
                    fh.seek(0)
                    msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
                except OSError:
                    pass
    finally:
        tlock.release()


def _blob(tok: dict, old: dict | None, now: float) -> dict:
    return {"access_token": tok["access_token"], "exp": now + int(tok.get("expires_in") or 0),
            "refresh_token": tok.get("refresh_token") or (old or {}).get("refresh_token"),
            "signed_in": (old or {}).get("signed_in") or now, "refreshed": now}


# ---------------------------------------------------------------- access tokens
def access_token(base: str, reject: str | None = None) -> str:
    """A valid access token for the pod, refreshed when it is about to end. reject is a token EPM just refused:
    it is never handed out again (a newer one another process stored is used, else a refresh is done).
    NotSignedIn when the pod needs `smartass login`; OAuthError when the refresh could not be done right now."""
    pod = origin(base)
    if scope_for(pod) is None:
        raise NotSignedIn(f"no OAuth scope is configured for {pod or base!r}")
    held = _mem.get(pod)
    if held and held[0] != reject and held[1] - time.time() > MARGIN_S:
        return held[0]
    with _locked(pod):
        blob = _read(pod)
        if not blob or not blob.get("refresh_token"):
            raise NotSignedIn(f"not signed in to {pod}: run `smartass login`")
        if blob.get("dead"):
            raise NotSignedIn(f"the sign-in to {pod} has ended ({blob['dead']}): run `smartass login`")
        fresh = blob.get("access_token") and blob["access_token"] != reject
        if not fresh or blob.get("exp", 0) - time.time() <= MARGIN_S:
            blob = _refresh(pod, blob)
        _mem[pod] = (blob["access_token"], blob["exp"])
        return blob["access_token"]


def _refresh(pod: str, blob: dict) -> dict:
    """Trade the refresh token for a new pair and store it (the caller holds the lock). A refusal of the token or
    the client ends the sign-in: it is marked in the file, so later calls say so at once instead of asking the
    identity domain again. Any other failure (429, 5xx, no answer) keeps the sign-in for the next try."""
    status, j = _post_form("/oauth2/v1/token", {"grant_type": "refresh_token", "client_id": config.OAUTH_CLIENT_ID,
                                                "refresh_token": blob["refresh_token"]})
    if status == 200 and j.get("access_token"):
        new = _blob(j, blob, time.time())
        _write(pod, new)
        return new
    why = _why(status, j)
    if j.get("error") in ENDED:
        _write(pod, {**blob, "access_token": None, "exp": 0, "dead": why})
        _mem.pop(pod, None)
        log.warning("oauth: the sign-in to %s has ended (%s)", pod, why)
        raise NotSignedIn(f"the sign-in to {pod} has ended ({why}): run `smartass login`")
    raise OAuthError(f"the identity domain could not renew the sign-in to {pod} ({why})")


def keep_fresh(max_age_s: float = 86400.0) -> list[str]:
    """Renew every sign-in whose last refresh is older than max_age_s, so the 7-day refresh token is replaced long
    before it ends (run_keep_fresh calls it at server start and every 6 hours). Returns what was done, per pod."""
    done = []
    for pod in pods():
        try:
            with _locked(pod):
                blob = _read(pod)
            if not blob or blob.get("dead") or not blob.get("refresh_token"):
                continue
            if time.time() - float(blob.get("refreshed") or 0) < max_age_s:
                continue
            access_token(pod, reject=blob.get("access_token"))
            done.append(f"{pod}: renewed")
        except OAuthError as e:
            done.append(f"{pod}: {e}")
    return done


def run_keep_fresh(every_s: float = 6 * 3600, sleep=time.sleep, rounds: int | None = None) -> None:
    """keep_fresh now and every 6 hours, for as long as the process lives. Never raises."""
    n = 0
    while rounds is None or n < rounds:
        n += 1
        try:
            for line in keep_fresh():
                log.info("oauth: %s", line)
        except Exception as e:   # noqa: BLE001 - a background nicety: the next query renews what it needs anyway
            log.info("oauth keep-fresh failed: %s", str(e)[:200])
        sleep(every_s)


def start_keep_fresh() -> threading.Thread | None:
    """Start run_keep_fresh in a daemon thread (the MCP server and `smartass serve` do, so a PC that is on once a
    week never has to sign in again); None when the bearer route is not set up."""
    if not configured():
        return None
    t = threading.Thread(target=run_keep_fresh, daemon=True, name="oauth-keep-fresh")
    t.start()
    return t


# ---------------------------------------------------------------- sign-in (device code)
def start_login(base: str) -> dict:
    """Ask the identity domain for a device code: {device_code, user_code, verification_uri, expires_in, interval}.
    Nothing is signed in until the user enters user_code at verification_uri."""
    scope = scope_for(base)
    if scope is None:
        raise OAuthError(f"no OAuth scope is configured for {origin(base) or base!r}")
    status, j = _post_form("/oauth2/v1/device", {"response_type": "device_code", "scope": f"{scope} offline_access",
                                                 "client_id": config.OAUTH_CLIENT_ID})
    if status != 200 or not all(j.get(k) for k in ("device_code", "user_code", "verification_uri")):
        raise OAuthError(f"the identity domain gave no sign-in code ({_why(status, j)})")
    if origin(j["verification_uri"]) != origin(config.OAUTH_IDCS):
        raise OAuthError("the sign-in page is not on the configured identity domain; nothing was opened")
    if urlsplit(j["verification_uri"]).scheme.lower() != "https":
        raise OAuthError("the sign-in page is not https; nothing was opened")
    return j


TRANSIENT = (429, 500, 502, 503, 504)   # the token endpoint is busy or down for a moment: the code is still good


def finish_login(base: str, device: dict, sleep=time.sleep, clock=time.monotonic) -> None:
    """Wait until the user has approved the code, then store the sign-in. Polls every interval (longer after
    slow_down or 429) until the code ends; no answer at all, a 429 and a 5xx are waited out like
    authorization_pending, so a network blip does not lose an approval. LoginDeclined (access_denied),
    LoginExpired (expired_token, or the code's time is up) and OAuthError for any other refusal end it."""
    pod = origin(base)
    interval = max(1, int(device.get("interval") or 5))
    deadline = clock() + int(device.get("expires_in") or 300)
    while clock() < deadline:
        sleep(interval)
        try:
            status, j = _post_form("/oauth2/v1/token", {"grant_type": DEVICE_GRANT,
                                                        "device_code": device["device_code"],
                                                        "client_id": config.OAUTH_CLIENT_ID})
        except OAuthError as e:                     # no answer (network, VPN, sleep): the code is still good
            log.info("oauth: no answer while waiting for the sign-in to %s (%s); still waiting", pod, e)
            continue
        if status == 200 and j.get("access_token"):
            if not j.get("refresh_token"):
                raise OAuthError("the identity domain gave no refresh token (the client lacks offline access)")
            with _locked(pod):
                _write(pod, _blob(j, None, time.time()))
            _mem.pop(pod, None)
            return
        err = j.get("error")
        if err == "slow_down" or status == 429:
            interval += 5
        elif err == "authorization_pending" or status in TRANSIENT:
            continue
        elif err == "access_denied":
            raise LoginDeclined(f"the sign-in was declined ({_why(status, j)})")
        elif err == "expired_token":
            raise LoginExpired("the sign-in code ended before it was approved")
        else:
            raise OAuthError(f"the sign-in was refused ({_why(status, j)})")
    raise LoginExpired("the sign-in code ended before it was approved")


# ---------------------------------------------------------------- sign-in from the tool call
_pending: dict[str, tuple[dict, float]] = {}      # pod -> (device answer, its end as time.monotonic()); this process
_starting: dict[str, threading.Event] = {}        # pod -> set when the start_login under way has ended
_ended: dict[str, str] = {}                       # pod -> why its last code ended without a sign-in (said once)
_pending_lock = threading.Lock()                  # guards the three dicts; never held over a network call
START_WAIT_S = 40.0                               # a start_login takes at most its 30 s timeout


def pending_logins() -> list[dict]:
    """For every configured pod that is not signed in: the sign-in the user has to approve, {pod, verification_uri,
    user_code, minutes_left} (plus note: why the earlier code was replaced, once) - a device login started here
    (start_login) and finished in a daemon thread (finish_login), or the one started earlier until it ends, so a
    second tool call shows the same code. A pod whose login could not be started is {pod, error} instead. Empty when
    every pod is signed in. The device code (what the thread polls with) never leaves this module."""
    out = []
    for s in status():
        if s["state"] == "signed_in":
            with _pending_lock:
                _ended.pop(s["pod"], None)
            continue
        out.append(_pending_for(s["pod"]))
    return out


def _pending_for(pod: str) -> dict:
    """The pod's waiting code, or a new one. One start per pod at a time: a second caller waits for the first one's
    code instead of asking for another (Review Focus 4), without the lock being held over the network."""
    while True:
        with _pending_lock:
            device, deadline = _pending.get(pod, (None, 0.0))
            if device is not None and deadline > time.monotonic():
                return _shown(pod, device, deadline)
            starting = _starting.get(pod)
            mine = starting is None
            if mine:
                starting = _starting[pod] = threading.Event()
        if not mine:
            starting.wait(START_WAIT_S)
            continue
        try:
            try:
                device = start_login(pod)
            except OAuthError as e:                     # the identity domain's own words: no URL query in a reply
                return {"pod": pod, "error": drillcore.redact(str(e))}
            deadline = time.monotonic() + int(device.get("expires_in") or 300)
            with _pending_lock:
                _pending[pod] = (device, deadline)
                shown = _shown(pod, device, deadline)
            _finish_in_background(pod, device)
            return shown
        finally:
            with _pending_lock:
                _starting.pop(pod, None)
            starting.set()


def _shown(pod: str, device: dict, deadline: float) -> dict:
    """What the user is shown for a waiting code (the caller holds the lock)."""
    page = urlsplit(device["verification_uri"])._replace(query="", fragment="").geturl()
    out = {"pod": pod, "verification_uri": page, "user_code": device["user_code"],
           "minutes_left": max(1, math.ceil((deadline - time.monotonic()) / 60))}
    note = _ended.pop(pod, None)
    return {**out, "note": note} if note else out


def _finish_in_background(pod: str, device: dict) -> threading.Thread:
    """finish_login in a daemon thread; when it ends (approved, declined, expired or refused) the pod's pending code
    is dropped - unless a newer one has replaced it - so the next tool call either works or gets a new code, with a
    note saying why the earlier one was replaced."""
    def run():
        why = None
        try:
            finish_login(pod, device)
            log.info("oauth: signed in to %s from a tool call's code", pod)
        except LoginDeclined:
            why = "The earlier sign-in was declined."
        except LoginExpired:
            why = "The earlier code expired before it was approved."
        except Exception as e:   # noqa: BLE001 - a daemon thread: the next tool call says what is needed
            why = f"The earlier sign-in did not complete ({drillcore.redact(str(e))[:200]})."
        finally:
            if why:
                log.warning("oauth: the sign-in to %s did not complete: %s", pod, why)
            with _pending_lock:
                if _pending.get(pod, (None,))[0] is device:
                    _pending.pop(pod, None)
                    if why:
                        _ended[pod] = why
    t = threading.Thread(target=run, daemon=True, name="oauth-sign-in")
    t.start()
    return t


def sign_in_message() -> str | None:
    """What a tool call that met NotSignedIn answers: one sentence per pod with the page and the code to enter (or why
    no code could be had); None when every pod is signed in."""
    said = []
    for p in pending_logins():
        host = urlsplit(p["pod"]).hostname or p["pod"]
        if "error" in p:
            said.append(f"EPM sign-in needed for {host}, but it could not be started ({p['error']}); "
                        "run `smartass login` in a terminal to see more.")
        else:
            said.append((f"{p['note']} " if p.get("note") else "")
                        + f"EPM sign-in needed. Open {p['verification_uri']} and enter the code {p['user_code']} "
                        f"({host}, valid {p['minutes_left']} min), approve with Telenor SSO, then ask again.")
    return " ".join(said) or None


def sign_in_reply(e: Exception) -> str:
    """The reply to a call that met NotSignedIn (MCP tools and the add-in's HTTP API alike): sign_in_message(), else
    the error's own text (NotSetUp: what config.toml needs). Redacted; never raises."""
    from ..core import logsetup
    try:
        msg = sign_in_message()
    except Exception as why:   # noqa: BLE001 - the reply must still say what is wrong
        log.info("oauth: sign-in message not made: %s", logsetup.redact(why)[:200])
        msg = None
    return logsetup.redact(msg or e)


def forget(base: str) -> bool:
    """Delete the pod's stored sign-in (this PC only); True when there was one."""
    pod = origin(base)
    _mem.pop(pod, None)
    with _locked(pod):
        try:
            _file(pod).unlink()
            return True
        except OSError:
            return False


def status() -> list[dict]:
    """Per configured pod, without any token: state 'signed_in' | 'ended' | 'none', days since the last refresh and
    the days left before an unused sign-in ends (by Oracle's default lifetime). Reads files only."""
    out = []
    for pod in pods():
        blob = _read(pod)
        if not blob or not blob.get("refresh_token"):
            out.append({"pod": pod, "state": "none"})
            continue
        age = max(0.0, (time.time() - float(blob.get("refreshed") or 0)) / 86400)
        state = "ended" if blob.get("dead") or age >= REFRESH_DAYS else "signed_in"
        out.append({"pod": pod, "state": state, "refreshed_days_ago": round(age, 1),
                    "days_left_unused": round(max(0.0, REFRESH_DAYS - age), 1),
                    **({"why": str(blob["dead"])} if blob.get("dead") else {})})
    return out
