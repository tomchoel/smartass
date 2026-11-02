"""Commentary (design §3.4; procedure knowledge/advisor/variance-commentary.md, adapted from Anthropic
financial-services variance-commentary, Apache-2.0): the engine flags material variances by the catalog's
materiality rules; the agent writes the drivers (why, not what). A material line without a driver is a Warning in
the review; drivers are rendered on the answer sheet, hypotheses visibly labelled."""
from __future__ import annotations

HYPOTHESIS = "Hypothesis — not verified in data"
STOCK_KEYS = {"fte"}          # levels: a NOK amount test makes no sense


def _num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _always(cat) -> set[str]:
    keys = set(cat.materiality.always)
    concepts = {ln.concept for ln in cat.pnl.get("FinMax", ()) if ln.key in keys and ln.concept}
    return keys | concepts


def material(prepared: dict, values: dict, cat) -> list[dict]:
    """Lines whose first variance cell is material, or which are always commented. Each:
    {line, label, cur, cmp, delta, pct, material, always}. Order = the answer's line order.
    A line may limit this (a recipe's choice): driver='never' (subtotals, reconciling lines) is never flagged;
    driver={'top': n, 'rest': 'supplier changes'} flags only the n largest material changes of all lines sharing
    that rest text - the others come back with optional=True (one Info in the review, no driver needed)."""
    m = cat.materiality
    always = _always(cat)
    out = []
    for sec in prepared.get("sections", []):
        for ln in sec.get("lines", []):
            key = ln.get("key")
            if not key or ln.get("aux") or ln.get("driver") == "never":
                continue
            var = next((c["var"] for c in ln.get("cells", []) if isinstance(c, dict) and "var" in c), None)
            v = values.get(key) or []
            if not var or max(var) >= len(v) or not (_num(v[var[0]]) and _num(v[var[1]])):
                continue
            cur, cmp = float(v[var[0]]), float(v[var[1]])
            delta = cur - cmp
            pct = delta / abs(cmp) if cmp else None
            stock = key in STOCK_KEYS or ln.get("concept") in STOCK_KEYS or ln.get("measure") == "stock"
            pct_hit = pct is not None and abs(pct) >= m.pct and (stock or abs(delta) >= m.pct_floor_nok)
            abs_hit = not stock and abs(delta) >= m.abs_nok
            if pct is None and not stock:
                pct_hit = abs_hit             # a zero comparison is material when the amount threshold is passed
            is_material = (pct_hit or abs_hit) if m.rule == "either" else (pct_hit and abs_hit)
            is_always = key in always or ln.get("concept") in always
            if is_material or is_always:
                f = {"line": key, "label": ln.get("label") or key, "cur": cur, "cmp": cmp, "delta": delta,
                     "pct": pct, "material": is_material, "always": is_always}
                if isinstance(ln.get("driver"), dict):
                    f["_top"] = ln["driver"]
                out.append(f)
    return _top_only(out)


def _top_only(flagged: list[dict]) -> list[dict]:
    """Per driver={'top', 'rest'} group: the n largest material changes stay; the rest become optional."""
    groups: dict[str, list[dict]] = {}
    for f in flagged:
        top = f.get("_top")
        if top and f["material"] and not f["always"]:
            groups.setdefault(str(top.get("rest") or "line changes"), []).append(f)
    for rest, fs in groups.items():
        n = int(fs[0]["_top"].get("top", 3))
        for f in sorted(fs, key=lambda x: -abs(x["delta"]))[n:]:
            f.update(optional=True, rest=rest)
    for f in flagged:
        f.pop("_top", None)
    return flagged


def _amount(f: dict) -> str:
    pct = f", {f['pct']:+.0%}" if f["pct"] is not None else ""
    return f"{f['delta']:+,.0f}{pct}"


def findings(flagged: list[dict], drivers: list[dict]) -> list[dict]:
    """Review findings for flagged lines that no driver explains. drivers: [{line, text, basis, evidence?}]."""
    explained = {d.get("line") for d in drivers if d.get("line")}
    out, rest = [], {}
    for f in flagged:
        if f["line"] in explained:
            continue
        if f.get("optional"):
            rest[f["rest"]] = rest.get(f["rest"], 0) + 1
        elif f["material"]:
            out.append({"severity": "Warning",
                        "text": f"material variance on '{f['label']}' ({_amount(f)}) has no driver - write one "
                                f"(data-backed with evidence, or a labelled hypothesis) or 'driver unclear - flag "
                                f"for controller'"})
        else:
            out.append({"severity": "Info", "text": f"always-commented line '{f['label']}' ({_amount(f)}) has no "
                                                    f"driver"})
    out += [{"severity": "Info", "text": f"{n} other {what} above materiality — see sheet"} for what, n in rest.items()]
    return out


def sheet_lines(drivers: list[dict]) -> list[str]:
    """Drivers as answer-sheet key points; a hypothesis is labelled as one."""
    out = []
    for d in drivers:
        if d.get("basis") == "hypothesis":
            out.append(f"{HYPOTHESIS}: {d['text']}")
        else:
            ev = f" (data: {d['evidence']})" if d.get("evidence") else ""
            out.append(f"Driver: {d['text']}{ev}")
    return out
