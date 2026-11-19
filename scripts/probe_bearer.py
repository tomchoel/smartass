"""Probe what the OAuth access token opens on a pod (read-only GETs and the Smart View connect): run it when a pod
gets a scope in config.toml [oauth.scopes] and `smartass login` has been done, before trusting the bearer route there.

Usage (from engine\\):
  .venv\\Scripts\\python.exe ..\\scripts\\probe_bearer.py              every pod with a scope
  .venv\\Scripts\\python.exe ..\\scripts\\probe_bearer.py <pod base URL>
Prints status codes, application and data source names and timings; never a token or a figure.

First run 2026-11-19 (FinPlan pod): REST 200, Smart View connect 200, the pod's own data source list, /aif/rest 200.
A pod whose scope is missing or wrong answers 401 'Token Audience'. Not probed here: drill-through over the token
(the SUI query and getData are UI endpoints; getData from a session that never loaded a drill page has signed the
whole EPM session out before, see epm/drill_direct.py)."""
from __future__ import annotations

import json
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "engine"))
from smartass.epm import oauth, svxml  # noqa: E402


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None


def call(method: str, url: str, token: str, body: str | None = None) -> tuple[int, str, float]:
    headers = {"Authorization": "Bearer " + token, "Accept": "application/json"}
    if body is not None:
        svxml.check_read_only(body)
        headers["Content-Type"] = "text/xml"
    req = urllib.request.Request(url, data=None if body is None else body.encode("utf-8"), headers=headers, method=method)
    t0 = time.monotonic()
    try:
        with urllib.request.build_opener(_NoRedirect).open(req, timeout=60) as r:
            return r.status, r.read().decode("utf-8", "replace"), time.monotonic() - t0
    except urllib.error.HTTPError as e:
        why = e.headers.get("WWW-Authenticate") or ""
        return e.code, why or e.read().decode("utf-8", "replace")[:200], time.monotonic() - t0


def probe(pod: str) -> bool:
    print(f"== {pod}")
    try:
        token = oauth.access_token(pod)
    except oauth.OAuthError as e:
        print(f"  no token: {e}")
        return False
    ok = True
    status, text, s = call("GET", pod + "/HyperionPlanning/rest/v3/applications", token)
    apps = [a.get("name") for a in json.loads(text).get("items", [])] if status == 200 else text[:160]
    print(f"  Planning REST applications: HTTP {status} in {s:.2f}s: {apps}")
    ok &= status == 200
    status, text, s = call("POST", pod + "/HyperionPlanning/SmartView", token, svxml.connect())
    connected = status == 200 and "<res_ConnectToProvider>" in text
    print(f"  Smart View connect: HTTP {status} in {s:.2f}s: {'ok' if connected else text[:160]}")
    ok &= connected
    if connected:
        status, text, s = call("POST", pod + "/HyperionPlanning/SmartView", token, svxml.provisioned_data_sources())
        try:
            sources = {name: [f"{app} at {base}" for app, base in entries]
                       for name, entries in svxml.data_sources(text).items()}
        except svxml.SvXmlError as e:
            sources, ok = f"error: {e}", False
        print(f"  Smart View data sources: HTTP {status} in {s:.2f}s: {sources}")
    status, text, s = call("GET", pod + "/aif/rest/V1", token)
    print(f"  Data Integration REST (/aif/rest/V1): HTTP {status} in {s:.2f}s")
    return bool(ok)


def main(argv: list[str]) -> int:
    if not oauth.configured():
        print("The bearer route is not set up: config.toml needs [oauth] idcs, client_id and [oauth.scopes].")
        return 2
    pods = [oauth.origin(a) for a in argv] or oauth.pods()
    return 0 if all([probe(p) for p in pods]) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
