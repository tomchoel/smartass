"""The atlas over the token session: one dimension's member tree from Oracle's Planning REST "Get Member" - GET
epmhttp.member_tree(app, dim) (.../applications/{app}/dimensions/{dim}/members/{dim}) with ?fields=MEMBER_FIELDS, the
top member and its whole visible subtree, nested - turned into the rows the hidden Excel's zoom gave the atlas
(ops.dump_dimension: Smart View zooms the top member to all levels and grid.tree_from_labels reads the indented
labels). Live 2026-11-22 (docs/audit/2026-11-22-atlas-rest.md): 201 on every dimension of both pods, all of them in
about 8 s; today's rows reproduced, key for key.

How the zoom's display is rebuilt from the tree (each rule found in today's atlas):
- a member's label is its Default alias, else its name (`alias` is absent when it has none);
- ancestors at the bottom: the zoom showed every dimension so (siblings in reverse outline order in the atlas, which
  tree_from_labels reads bottom-up), so the display is written so and read by the same function - the key derived
  from the label (grid.name_from_label), a member kept under its label, a shared member under each parent, a name
  given twice under one parent (atlas.store keeps the first) all come out as before;
- the zoom shows neither the Dynamic Time Series members of a period dimension (`H-T-D` .. `D-T-D`, the first
  children of its top in Get Member's answer) nor a member whose `usedIn` does not name the plan type (with its
  subtree): FinLite's dimensions are FinMax's application dimensions, read once for both.
Each row also carries the member's real name, alias, data storage, account type and time balance (EXTRA: the atlas's
side table, atlas.store; nothing reads the types yet).

Errors are TreeError (its text names the status, never the answer's text); oauth.NotSignedIn reaches the caller as
it is (the build stops, the atlas is not replaced)."""
from __future__ import annotations

import json
import re

from . import epmhttp, oauth, rest, svhttp
from .grid import tree_from_labels

OK_STATUS = 201                 # Get Member's success status (Oracle: "historical"; live: 201 on every dimension)
DTS = re.compile(r"[HYSPQMWD]-T-D")
INDENT = 3                      # spaces a level in the rebuilt display (tree_from_labels reads any step)
# (atlas column, Get Member property)
EXTRA_FIELDS = (("real_name", "name"), ("alias", "alias"), ("storage", "dataStorage"), ("account_type", "accountType"),
                ("time_balance", "timeBalance"))
EXTRA = tuple(c for c, _ in EXTRA_FIELDS)


class TreeError(Exception):
    """The member tree could not be read: refused, lost, or an answer that is not one."""


def read_tree(session: epmhttp.EpmSession, app: str, dim: str) -> dict:
    """Dimension dim of application app as Get Member answers it from its top member (checked: check_tree)."""
    where = f"Planning REST at {session.origin}: {app}/{dim}"
    try:
        r = session.request("GET", f"{epmhttp.member_tree(app, dim)}?fields={epmhttp.MEMBER_FIELDS}",
                            timeout=svhttp.timeout_s(), renew_on=epmhttp.GATE_STATUSES)
    except oauth.NotSignedIn:
        raise
    except oauth.OAuthError as e:
        raise TreeError(f"no EPM access token for {session.origin}: {e}") from None
    except (epmhttp.Refused, epmhttp.TooLarge) as e:   # their texts quote no body, token or query
        raise TreeError(f"{where}: {e}") from None
    except OSError as e:
        raise TreeError(f"{where}: lost the request ({type(e).__name__})") from None
    if r.status != OK_STATUS:
        raise TreeError(f"{where}: HTTP {r.status}")
    try:
        tree = json.loads(r.text)
    except ValueError:
        raise TreeError(f"{where}: not JSON") from None
    check_tree(tree)
    return tree


def check_tree(tree) -> None:
    """A member is an object with a name; alias a text when present; usedIn a list of names when present; children a
    list of members when present - at every level. Else TreeError (the answer is never quoted)."""
    todo = [tree]
    while todo:
        n = todo.pop()
        ok = (isinstance(n, dict) and isinstance(n.get("name"), str) and n["name"]
              and isinstance(n.get("alias", ""), str) and isinstance(n.get("children", []), list)
              and isinstance(n.get("usedIn", []), list) and all(isinstance(x, str) for x in n.get("usedIn", [])))
        if not ok:
            raise TreeError("not a Planning REST member tree")
        todo.extend(n.get("children", []))


def label(n: dict) -> str:
    """What Smart View displays for a member with the Default alias table: its alias, else its name."""
    return n.get("alias") or n["name"]


def rows(tree: dict, cube: str, period: bool = False) -> list[dict]:
    """The atlas rows of one dimension for plan type cube: [{label, name, depth, parent, order, *EXTRA}], as
    ops.dump_dimension + tree_from_labels gave them (see the module's text). period: a period dimension (the plan
    type's list says dimType 'Period'), the only kind whose Dynamic Time Series members are left out."""
    labels: list[str] = []
    extras: list[dict] = []

    def shown(n: dict, depth: int) -> bool:
        if period and depth == 1 and DTS.fullmatch(n["name"]):
            return False
        return "usedIn" not in n or cube in n["usedIn"]

    def go(n: dict, depth: int) -> None:                 # ancestors at the bottom: children first
        for c in n.get("children", []):
            if shown(c, depth + 1):
                go(c, depth + 1)
        labels.append(" " * (INDENT * depth) + label(n))
        extras.append({col: n.get(key) for col, key in EXTRA_FIELDS})
    go(tree, 0)
    return tree_from_labels(labels, extras)


def readable(app: str, dim: str) -> bool:
    """The allowlist has a member-tree rule for this dimension (epmhttp.MEMBER_TREES, fixed: a dimension EPM adds
    later is not read until a release adds its rule)."""
    return (app, dim) in epmhttp.MEMBER_TREES


class Source:
    """The plan types' dimensions and their trees on the pods' sessions (svhttp.SvHttp.rest_session), each tree read
    once per build (FinLite's are FinMax's)."""

    def __init__(self, sv=None):
        self.sv = sv or svhttp.shared()
        self._trees: dict[tuple[str, str], dict] = {}

    def dimensions(self, conn: str) -> list[tuple[str, str, str | None]]:
        """(REST name, atlas name, dimType) of every dimension of the connection's plan type, attribute dimensions
        aside: the atlas names a dimension by its Default alias, else its name (rest.map_dimensions)."""
        _, app, cube = svhttp.CUBES[conn]
        listed = rest.read_dimensions(self.sv.rest_session(conn), app, cube)
        return [(n, a or n, t[0] if t else None) for n, a, *t in listed
                if not (t and str(t[0]).startswith("Attribute"))]

    def tree(self, conn: str, dim: str) -> dict:
        _, app, _ = svhttp.CUBES[conn]
        if (app, dim) not in self._trees:
            self._trees[app, dim] = read_tree(self.sv.rest_session(conn), app, dim)
        return self._trees[app, dim]
