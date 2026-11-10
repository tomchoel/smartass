"""Generate the skill's readable cube atlas (reference/cubes.md) from atlas.db.
Run after `smartass atlas build`:  python gen_atlas_md.py [output.md]"""
import os
import sqlite3
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from smartass.core import config  # noqa: E402

DB = str(config.ATLAS_DB)
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.expanduser(r"~\.claude\skills\smartass\reference\cubes.md")
# extra depth for the dimensions people navigate most
DEEP = {("FinMax", "All Entities"): "un.SY", ("FinLite", "All Entities"): "un.SY", ("FinMax", "All Accounts"): None,
        ("FinPlan", "Entity"): None, ("FinPlan", "Account"): None}
MAX_CHILDREN = 40
# nodes worth a two-level drill-down of their own (where Produktregnskapet work starts)
EXTRA = [("FinMax", "All Entities", "un.SY"), ("FinMax", "All Accounts", "ac.C3_C8"),
         ("FinPlan", "Entity", "Entity")]

db = sqlite3.connect(DB)


def kids(conn, dim, parent):
    return db.execute("SELECT DISTINCT name, label FROM member WHERE connection=? AND dimension=? AND parent=? "
                      "ORDER BY ord", (conn, dim, parent)).fetchall()


def fmt(name, label):
    return f"`{name}`" + (f" {label}" if label and label != name else "")


lines = ["# Cube atlas (generated)", "",
         f"Generated {time.strftime('%Y-%m-%d %H:%M')} from `%LOCALAPPDATA%\\smartass\\atlas.db` by `gen_atlas_md.py`. "
         "Top levels only; search everything with `fin_lookup`. Periods (current month, closed months, quarters) are not "
         "listed here: ask `fin_status`, or `fin_lookup` (kind 'member').", ""]
for (conn,) in db.execute("SELECT DISTINCT connection FROM build ORDER BY connection"):
    lines += [f"## {conn}", "", "| Dimension | Members | Built |", "|---|---:|---|"]
    builds = db.execute("SELECT dimension, n, built_at, error FROM build WHERE connection=? ORDER BY dimension",
                        (conn,)).fetchall()
    for dim, n, b, err in builds:
        lines.append(f"| {dim} | {n:,} | {time.strftime('%Y-%m-%d', time.localtime(b))}{' ERROR' if err else ''} |")
    lines.append("")
    for dim, n, _, err in builds:
        if err or not n:
            continue
        roots = db.execute("SELECT DISTINCT name, label FROM member WHERE connection=? AND dimension=? AND parent IS NULL",
                           (conn, dim)).fetchall()
        lines.append(f"### {conn} · {dim}")
        for rn, rl in roots:
            ch = kids(conn, dim, rn)
            lines.append(f"- {fmt(rn, rl)} — {len(ch)} children")
            for cn, cl in ch[:MAX_CHILDREN]:
                lines.append(f"  - {fmt(cn, cl)}")
                deep_root = DEEP.get((conn, dim), "skip")
                if deep_root != "skip" and (deep_root is None or cn == deep_root):
                    for gn, gl in kids(conn, dim, cn)[:MAX_CHILDREN]:
                        lines.append(f"    - {fmt(gn, gl)}")
            if len(ch) > MAX_CHILDREN:
                lines.append(f"  - … {len(ch) - MAX_CHILDREN} more")
        lines.append("")
lines += ["## Key drill-downs", ""]
for conn, dim, node in EXTRA:
    lab = db.execute("SELECT label FROM member WHERE connection=? AND name=? LIMIT 1", (conn, node)).fetchone()
    ch = kids(conn, dim, node)
    lines.append(f"### {conn} · {fmt(node, lab[0] if lab else None)} — {len(ch)} children")
    for cn, cl in ch[:60]:
        gk = kids(conn, dim, cn)
        lines.append(f"- {fmt(cn, cl)}" + (f" — {len(gk)} children" if gk else ""))
        for gn, gl in gk[:25]:
            lines.append(f"  - {fmt(gn, gl)}")
        if len(gk) > 25:
            lines.append(f"  - … {len(gk) - 25} more")
    lines.append("")
os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(lines) + "\n")
print(f"wrote {OUT} ({len(lines)} lines)")
