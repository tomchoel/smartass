"""One-off (spec 2026-11-10 §2): fold knowledge/catalog.toml, capabilities.toml and the 15.10 business maps into
knowledge/semantic/{rules,entities,kpis,vocab,notes} and bundle metrics. Run once from engine/:
    uv run python -B ../scripts/fold_knowledge.py < /dev/null
Refuses to overwrite an existing file. Deleted after the fold (Task 11); git history keeps it."""
import json
import pathlib
import sys
import tomllib

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "engine"))
from smartass.knowledge.semantic import Entity, Kpi, Metric, Rule, Vocab, to_toml  # noqa: E402

K = pathlib.Path(__file__).resolve().parents[1] / "knowledge"
S = K / "semantic"
TOM = {"by": "tom", "on": "2026-10-15", "via": "catalog.toml"}
LINE_METRIC = {"rev": "revenue", "cogs": "cogs", "gm": "gross_margin", "gp": "gross_margin", "pers": "personnel",
               "oth": "other_opex", "opex": "opex", "boi": "ebitda_boi", "c73": "other_items",
               "ebitda": "ebitda", "da": "da", "fte": "fte"}
cat = tomllib.loads((K / "catalog.toml").read_text(encoding="utf-8"))
cap = tomllib.loads((K / "capabilities.toml").read_text(encoding="utf-8"))
written = []


def put(folder: str, item_id: str, text: str) -> None:
    p = S / folder / f"{item_id}.toml"
    if p.exists():
        raise SystemExit(f"refusing to overwrite {p}")
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding="utf-8")
    written.append(p.relative_to(S).as_posix())



# rules
fm = cat["pnl"]["FinMax"]
fp = cat["pnl"]["FinPlan"]
put("rules", "pnl", to_toml(Rule(
    id="pnl", title="P&L chain", kind="pnl", source="catalog.toml [pnl] + [checks]", status="approved", approved=TOM,
    text="FinMax: C30 + C40S = C48; C50S + C60 = C70; C48 + C70 = EBITDA_BOI; + C73 = EBITDA; D&A C75 below. "
         "FinPlan: Revenue + COGS = Gross_Profit; + OPEX = EBITDA before Other Items; + Total Other Items = EBITDA "
         "(FinPlan stops at EBITDA). Join actual vs budget on the metric, never on the line key.",
    data={"FinMax": [LINE_METRIC[x["key"]] for x in fm], "FinPlan": [{**LINE_METRIC, "oth": "other_items"}[x["key"]] for x in fp],
          "checks": {c: [{"label": x["label"], "expr": x["expr"]} for x in cat["checks"][c]] for c in cat["checks"]}})))
sc = cat["scenarios"]
put("rules", "scenarios", to_toml(Rule(
    id="scenarios", title="Scenarios per cube", kind="scenarios", source="catalog.toml [scenarios] + series.py",
    status="approved", approved=TOM,
    text="Actuals: FinMax/FinLite scenario Actual. Plan lives only in FinPlan: actual = Actual/Working, budget = "
         "Target/Final, forecast = the latest Forecast Q<q><yy>/Final.",
    data={"FinMax": {"value": ["Actual", ""], "ly": ["Actual", ""]},
          "FinLite": {"value": ["Actual", ""], "ly": ["Actual", ""]},
          "FinPlan": {"value": [sc["actual_finplan"]["scenario"], sc["actual_finplan"]["version"]],
                      "ly": [sc["actual_finplan"]["scenario"], sc["actual_finplan"]["version"]],
                      "budget": [sc["budget"]["scenario"], sc["budget"]["version"]],
                      "forecast": ["latest_forecast", sc["forecast"]["version"]]}})))
mt = cat["materiality"]
put("rules", "materiality", to_toml(Rule(
    id="materiality", title="Materiality", kind="materiality", source="catalog.toml [materiality]", status="approved", approved=TOM,
    text=f"A variance is material at {mt['pct']:.0%} of |comparison| (from NOK {mt['pct_floor_nok']:,}) or at NOK "
         f"{mt['abs_nok']:,}; revenue, personnel and EBITDA are always commented.".replace(",", " "),
    data={"abs_nok": float(mt["abs_nok"]), "pct": mt["pct"], "pct_floor_nok": float(mt["pct_floor_nok"]),
          "rule": mt["rule"], "always": [LINE_METRIC[x] for x in mt["always"]]})))
gc = cat["group_companies"]
put("rules", "group_companies", to_toml(Rule(
    id="group_companies", title="Group companies", kind="group_companies", source="catalog.toml [group_companies]",
    status="approved", approved=TOM, text="Group companies get their own block in a supplier ranking, never in its top 10.",
    data={"column": gc["column"], "patterns": gc["patterns"]})))
put("rules", "signs", to_toml(Rule(
    id="signs", title="Signs and variance", source="catalog.toml header", status="approved", approved=TOM,
    text="FinMax signs: revenue positive, costs negative, so a higher value is better on every P&L line. Variance = "
         "current - comparison (positive = better); variance % is empty on a tiny base.")))
put("rules", "cube_choice", to_toml(Rule(
    id="cube_choice", title="Which cube", source="skill smartass §6", status="approved", approved=TOM,
    text="Actuals: FinMax (FTE ac.STA20Y). FinLite is lighter (no FTE) and the drill source for ledger lines. Plan "
         "(budget, forecast) only in FinPlan. History is restated to today's organisation.")))

# entities
for e in cat["entity"]:
    put("entities", e["key"], to_toml(Entity(
        id=e["key"], label=e["label"], aliases=e.get("aliases", []), generic_aliases=e.get("generic_aliases", []),
        finmax=e.get("finmax"), finplan=e.get("finplan"), abolished=e.get("abolished", ""),
        successors=e.get("successors", []), note=e.get("note", ""), source="catalog.toml [[entity]]",
        status="approved", approved=TOM)))

# kpis (synonyms from capabilities [[measure]] kpi = ...)
syn = {m["kpi"]: m["synonyms"] for m in cap["measure"] if "kpi" in m}
note = {m["kpi"]: m.get("note", "") for m in cap["measure"] if "kpi" in m}
for kp in cat["kpi"]:
    put("kpis", kp["key"], to_toml(Kpi(
        id=kp["key"], name=kp["label"], aliases=[s for s in syn.get(kp["key"], []) if s.casefold() !=
                                                 kp["label"].casefold()],
        description=note.get(kp["key"], ""), cube=kp["connection"], numerator=LINE_METRIC[kp["numerator"]],
        denominator=LINE_METRIC[kp["denominator"]], unit=kp["unit"], basis=kp.get("basis", "period"),
        annualise=kp.get("annualise", False), sign=kp.get("sign", 1), better=kp.get("better", "higher"),
        headline=kp.get("headline", False), source="catalog.toml [[kpi]]", status="approved", approved=TOM)))

# bundles -> metrics; vague words; grains; currency
for m in cap["measure"]:
    if "accounts" in m:
        put("metrics", m["key"], to_toml(Metric(
            id=m["key"], name=m["label"], aliases=m["synonyms"], description=m.get("note", "") or m["label"],
            sign="cost", status="approved", approved=TOM | {"via": "capabilities.toml"},
            finmax={"members": m["accounts"]}, finlite={"members": m["accounts"]})))
    elif m.get("vague"):
        put("vocab", m["key"], to_toml(Vocab(
            id=m["key"], kind="vague", words=m["synonyms"], options=m["options"],
            source="capabilities.toml [[measure]]", status="approved", approved=TOM | {"via": "capabilities.toml"})))
for g in cap["grain"]:
    put("vocab", f"grain_{g['key']}", to_toml(Vocab(
        id=f"grain_{g['key']}", kind="grain", words=g["synonyms"],
        feasible=[LINE_METRIC.get(x, x) if x != "*" else x for x in g["feasible"]], proxy=g.get("proxy", ""),
        note=g.get("note", ""), source="capabilities.toml [[grain]]", status="approved", approved=TOM | {"via": "capabilities.toml"})))
cur = cap["currency"]
put("vocab", "currency", to_toml(Vocab(
    id="currency", kind="currency", words=sorted(cur["words"]), codes=cur["words"], reporting=cur["reporting"],
    proxy=cur["proxy"], source="capabilities.toml [currency]", status="approved", approved=TOM | {"via": "capabilities.toml"})))

# member notes from the maps (meaning + booking pattern / pitfalls; Norwegian as approved)
def notes(file, rows_key, cube, fmt):
    d = json.loads((K / file).read_text(encoding="utf-8"))
    lines = []
    for r in d[rows_key]:
        text = fmt(r).strip()
        if text:
            lines.append("[[note]]\n" + f"cube = {json.dumps(cube)}\nmember = {json.dumps(r['member'], ensure_ascii=False)}\n"
                         f"text = {json.dumps(text, ensure_ascii=False)}\n")
    return "\n".join(lines)


def join(*parts):
    return " ".join(p.rstrip(".") + "." for p in parts if p)


put("notes", "finmax_accounts", notes("accounts_map.json", "accounts", "FinMax",
                                      lambda r: join(r.get("meaning"), r.get("booking_pattern"), r.get("pitfalls"))))
put("notes", "finmax_entities", notes("org_map.json", "entities", "FinMax",
                                      lambda r: join(r.get("meaning"), r.get("kind") and f"Type: {r['kind']}",
                                                     r.get("moved") and f"Moved: {r['moved']}", r.get("pitfalls"))))
put("notes", "finplan_accounts", notes("finplan_map.json", "accounts", "FinPlan",
                                       lambda r: join(r.get("meaning"),
                                                      r.get("finmax_equivalent") and
                                                      f"FinMax equivalent: {r['finmax_equivalent']}",
                                                      r.get("pitfalls"))))
print("\n".join(written))
print(len(written), "files")
