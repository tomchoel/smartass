"""Recipe cost_base (B06): cost base per leaf cost centre under a parent, the top 10 by |C70| ranked after the pull
(specs repeat.over.top), the rest as one 'Other' line, the parent's total, the main account last year and each
line's share of the total; the check top + other = total."""
from __future__ import annotations

from ... import templates
from ...grid import Axis
from ...host import InputError
from . import Draft, RecipeContext, pull_dict

ACCOUNTS = ["ac.C50S", "ac.C60", "ac.C70", "ac.STA20Y"]
TOP = 10
DEFAULTS = {"top": (lambda v: str(v).strip() == str(TOP), "the standard ranking is the top 10 + Other")}
FOLLOWUPS = ["Monthly trend of the largest cost centre through {period}",
             "Which suppliers make up the costs of {entity}, YTD {period}",
             "The cost base of {entity} vs budget, YTD {period}"]


def cost_base(params: dict, ctx: RecipeContext) -> Draft:
    parent, conn = params.get("parent", "un.SY"), params.get("connection", "FinMax")
    accounts, period = params.get("accounts") or ACCOUNTS, params.get("period") or ctx.closed
    top, ly = int(params.get("top", TOP)), templates.same_period_ly(period)
    members, leaves = templates.cost_base_members(conn, parent)
    if len(members) * (len(accounts) + 1) > 50000:
        raise InputError(f"{len(members)} cost centres × {len(accounts)} accounts is too large; narrow 'parent'.")
    pov = {"All Scenarios": "Actual", "All Currencies": "NOK", "All Time Periods": period,
           "All Balance Types": "Accumulated"}
    main = "ac.C70" if "ac.C70" in accounts else accounts[0]
    pull = templates.PullSpec("cb", conn, [Axis("All Entities", members)], [Axis("All Accounts", accounts)], pov)
    pull_ly = templates.PullSpec("ly", conn, [Axis("All Entities", members)], [Axis("All Accounts", [main])],
                                 {**pov, "All Time Periods": ly})
    n, share = len(accounts), {"share": {"col": accounts.index(main), "of": "tot"}}

    def cells(row):
        return [{"pull": "cb", "row": [row], "col": [a]} for a in accounts] + \
            [{"pull": "ly", "row": [row], "col": [main]}, share]
    name, main_name = ctx.label(conn, parent), ctx.label(conn, main)
    cols = [{"header": ctx.label(conn, a), "kind": "tall1" if "STA" in a else "heltall"} for a in accounts] + \
        [{"header": f"{main_name} {ly}", "kind": "heltall"}, {"header": "Share %", "kind": "prosent1"}]
    lay = {"title": f"Cost base per cost centre — {name} — YTD {period}",
           "subtitle": f"NOK · {conn} · Actual · Accumulated {period} and {ly}",
           "columns": cols,
           "sections": [{"title": f"Top {top} cost centres by {main_name}",
                         "repeat": {"over": {"top": top, "pull": "cb", "by": [main],
                                             "among": [m for m in members if m in leaves], "abs": True},
                                    "line": {"key": "c{i}", "label": "{label}", "kind": "line",
                                             "cells": cells("{m}")}},
                         "lines": [{"key": "topsum", "label": f"Top {top}", "aux": True,
                                    "cells": [{"sum": "c*"} for _ in range(n + 1)] + [share]},
                                   {"key": "rest", "label": "Other cost centres", "kind": "line",
                                    "cells": [{"expr": "tot-topsum"} for _ in range(n + 1)] + [share]},
                                   {"key": "tot", "label": f"Total {name}", "kind": "grand",
                                    "cells": cells(parent)}]}],
           "checks": [{"label": "Top + other = total", "expr": "topsum+rest-tot", "columns": list(range(n + 1))}],
           "notes": ["FTE (STA20Y) is the end-of-period headcount.",
                     "Parent nodes are restated to today's organisation.",
                     f"Ranked by |{main_name}| among {len(leaves)} leaf cost centres; the full tree is on the data "
                     f"sheet. Share % = the line's {main_name} / the total's."],
           "source": f"Source: Oracle EPM Cloud ({conn}) via Smart View, retrieved {ctx.now}."}
    return Draft([pull_dict(pull), pull_dict(pull_ly)], lay,
                 assumptions=[f"Actual, NOK, YTD {period} vs YTD {ly}, leaf cost centres"],
                 result={"period": period})
