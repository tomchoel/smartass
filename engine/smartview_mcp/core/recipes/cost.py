"""Recipe cost_base (B06): cost base per leaf cost centre under a parent, top N by |C70| ranked after the pull
(specs repeat.over.top), the rest as one line, the parent's total, and the check top + other = total."""
from __future__ import annotations

from ... import templates
from ...grid import Axis
from ...host import InputError
from . import Draft, RecipeContext, pull_dict

ACCOUNTS = ["ac.C50S", "ac.C60", "ac.C70", "ac.STA20Y"]


def cost_base(params: dict, ctx: RecipeContext) -> Draft:
    parent, conn = params.get("parent", "un.SY"), params.get("connection", "FinMax")
    accounts, period = params.get("accounts") or ACCOUNTS, params.get("period") or ctx.closed
    top = int(params.get("top", 25))
    members, leaves = templates.cost_base_members(conn, parent)
    if len(members) * len(accounts) > 50000:
        raise InputError(f"{len(members)} cost centres × {len(accounts)} accounts is too large; narrow 'parent'.")
    pov = {"All Scenarios": "Actual", "All Currencies": "NOK", "All Time Periods": period,
           "All Balance Types": "Accumulated"}
    pull = templates.PullSpec("cb", conn, [Axis("All Entities", members)], [Axis("All Accounts", accounts)], pov)
    main = "ac.C70" if "ac.C70" in accounts else accounts[0]

    def cells(row):
        return [{"pull": "cb", "row": [row], "col": [a]} for a in accounts]
    name = ctx.label(conn, parent)
    cols = [{"header": ctx.label(conn, a), "kind": "tall1" if "STA" in a else "heltall"} for a in accounts]
    lay = {"title": f"Cost base per cost centre — {name} — YTD {period}",
           "subtitle": f"NOK · {conn} · Actual · Accumulated {period}",
           "columns": cols,
           "sections": [{"title": f"Top {top} cost centres by {ctx.label(conn, main)}",
                         "repeat": {"over": {"top": top, "pull": "cb", "by": [main],
                                             "among": [m for m in members if m in leaves], "abs": True},
                                    "line": {"key": "c{i}", "label": "{label}", "kind": "line",
                                             "cells": cells("{m}")}},
                         "lines": [{"key": "topsum", "label": f"Top {top}", "aux": True,
                                    "cells": [{"sum": "c*"} for _ in accounts]},
                                   {"key": "rest", "label": "Other cost centres", "kind": "line",
                                    "cells": [{"expr": "tot-topsum"} for _ in accounts]},
                                   {"key": "tot", "label": f"Total {name}", "kind": "grand",
                                    "cells": cells(parent)}]}],
           "checks": [{"label": "Top + other = total", "expr": "topsum+rest-tot",
                       "columns": list(range(len(accounts)))}],
           "notes": ["FTE (STA20Y) is the end-of-period headcount.",
                     "Parent nodes are restated to today's organisation.",
                     f"Ranked by |{ctx.label(conn, main)}| among {len(leaves)} leaf cost centres; the full tree "
                     "is on the data sheet."],
           "source": f"Source: Oracle EPM Cloud ({conn}) via Smart View, retrieved {ctx.now}."}
    return Draft([pull_dict(pull)], lay, assumptions=[f"Actual, NOK, YTD {period}, leaf cost centres"],
                 result={"period": period})
