"""`smartass login`: the EPM sign-in for the bearer route (epm/oauth.py). One browser approval per pod with a
scope (device code: the page opens, the code is on the clipboard); after that grids and drill lines need no sign-in
page for as long as smartass is used at least once a week. On a PC with no atlas it then builds one (ops/atlas_build:
read-only, over the sign-in it just made), unless --no-atlas; it never replaces an atlas that is there."""
from __future__ import annotations

import subprocess
import time
import webbrowser
from urllib.parse import urlsplit

from ..epm import oauth

NOT_SET_UP = oauth.NOT_SET_UP


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


BUILDING = ("No atlas on this PC yet (the member catalogue every answer is built on): building it now over EPM, "
            "read-only - the member trees take about half a minute, then the behaviour cards about two minutes more. "
            "(`smartass login --no-atlas` skips this; `smartass atlas build` builds it at any time.)")
BUILD_LATER = "your sign-in is kept. Run `smartass atlas build` later to build the atlas."
ATLAS_NOT_BUILT = 3            # login's exit code: every sign-in worked, the atlas the PC lacks was not built


def _no_atlas() -> bool:
    from ..knowledge import atlas
    return not atlas.present()


def _build() -> list[dict]:
    """The atlas build of `smartass atlas build` (its log file too)."""
    from ..core import config, logsetup
    from . import atlas_build
    logsetup.configure(config.HOME / "logs" / "atlas.log")
    return atlas_build.run_build()


def build_when_none(out=print, build=None) -> bool:
    """The atlas built when the PC has none (never over one that is there), said before it starts, with how long it
    took after. A build that fails or replaces nothing leaves the sign-in as it is and says which command to run.
    False when it was needed and not built."""
    if not _no_atlas():
        return True
    from . import atlas_build
    out(BUILDING)
    t0 = time.monotonic()
    try:
        done = (build or _build)()
    except Exception as e:   # noqa: BLE001 - the sign-in stands whatever the build does
        from ..core import logsetup
        out(f"The atlas was not built ({type(e).__name__}: {logsetup.redact(e)}); {BUILD_LATER}")
        return False
    line, ok = atlas_build.outcome(done)
    out(line)
    out(f"Atlas built in {time.monotonic() - t0:.0f} s." if ok else f"The atlas was not built (above); {BUILD_LATER}")
    return ok


def run(force: bool = False, out=print, open_page=webbrowser.open, copy=_copy, atlas: bool = True,
        build=None) -> int:
    """Sign in to every pod that needs it (all of them with force), then build the atlas when there is none (atlas:
    False skips it; build: the build, a test's fake). 0 = all signed in (and the atlas there, or skipped), 1 = a sign-in
    failed (no build), 2 = not set up, 3 = all signed in but the atlas was not built (the text says which command to
    run later): `smartass login && ...` in a script does not go on as if the PC were ready."""
    if not oauth.configured():
        out(NOT_SET_UP)
        return 2
    todo = [s["pod"] for s in oauth.status() if force or s["state"] != "signed_in"]
    if not todo:
        show(out)
        return 0 if not atlas or build_when_none(out, build) else ATLAS_NOT_BUILT
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
    if failed:
        if atlas and _no_atlas():
            out("No atlas on this PC yet: it is built after every sign-in has succeeded - run `smartass login` again.")
        return 1
    return 0 if not atlas or build_when_none(out, build) else ATLAS_NOT_BUILT
