"""`smartass login`: the EPM sign-in for the bearer route (epm/oauth.py). One browser approval per pod with a
scope (device code: the page opens, the code is on the clipboard); after that grids need neither a browser nor a
sign-in page for as long as smartass is used at least once a week."""
from __future__ import annotations

import subprocess
import webbrowser
from urllib.parse import urlsplit

from ..epm import oauth

NOT_SET_UP = ("The fast EPM route is not set up: config.toml needs [oauth] idcs and client_id, and [oauth.scopes] "
              "with one scope per pod (docs\\troubleshooting.md, 'oauth'). Grids use the hidden browser meanwhile.")


def _copy(text: str) -> None:
    try:
        subprocess.run(["clip"], input=text, text=True, timeout=5, check=False, creationflags=0x08000000)
    except (OSError, subprocess.SubprocessError):
        pass                                                # the code is printed as well


def _host(pod: str) -> str:
    return urlsplit(pod).hostname or pod


def _line(s: dict) -> str:
    if s["state"] == "none":
        return f"{_host(s['pod'])}: not signed in"
    if s["state"] == "ended":
        return f"{_host(s['pod'])}: the sign-in has ended ({s.get('why') or 'not used for a week'})"
    return (f"{_host(s['pod'])}: signed in (renewed {s['refreshed_days_ago']:g} days ago; "
            f"lasts {s['days_left_unused']:g} more days without use)")


def show(out=print) -> int:
    if not oauth.configured():
        out(NOT_SET_UP)
        return 2
    for s in oauth.status():
        out(_line(s))
    return 0


def forget(out=print) -> int:
    if not oauth.configured():
        out(NOT_SET_UP)
        return 2
    for pod in oauth.pods():
        out(f"{_host(pod)}: {'sign-in deleted from this PC' if oauth.forget(pod) else 'no sign-in stored'}")
    return 0


def run(force: bool = False, out=print, open_page=webbrowser.open, copy=_copy) -> int:
    """Sign in to every pod that needs it (all of them with force). 0 = all signed in, 1 = a sign-in failed,
    2 = not set up."""
    if not oauth.configured():
        out(NOT_SET_UP)
        return 2
    todo = [s["pod"] for s in oauth.status() if force or s["state"] != "signed_in"]
    if not todo:
        show(out)
        return 0
    failed = 0
    for pod in todo:
        try:
            d = oauth.start_login(pod)
            out(f"Sign in to {_host(pod)} (valid for {int(d.get('expires_in') or 300) // 60} minutes):\n"
                f"  page: {d['verification_uri']}\n"
                f"  code: {d['user_code']}\n"
                "The page is opening in your browser and the code is on the clipboard: paste it and approve.")
            copy(d["user_code"])
            open_page(d["verification_uri"])
            oauth.finish_login(pod, d)
            out(f"Signed in to {_host(pod)}.")
        except oauth.OAuthError as e:
            out(f"Sign-in to {_host(pod)} failed: {e}")
            failed += 1
    return 1 if failed else 0
