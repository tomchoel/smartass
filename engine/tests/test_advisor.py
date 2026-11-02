"""Phase 4: semantic layer, period parser, learned defaults, the advisor gate (benchmark table) and commentary."""
from __future__ import annotations

import pytest

from smartview_mcp import catalog
from smartview_mcp.core import semantic


def test_measure_synonyms_resolve_english_and_norwegian():
    assert [m.key for m in semantic.measures("EBITDA for Business")] == ["ebitda"]
    assert [m.key for m in semantic.measures("driftskostnader i IT")] == ["opex"]
    assert [m.key for m in semantic.measures("Cost per FTE in IT")] == ["opex_per_fte"]   # longest match wins
    vague = semantic.measures("How are we doing on costs?")
    assert {m.key for m in vague} == {"costs", "numbers"} and all(m.vague for m in vague)
    assert [m.key for m in semantic.measures("How much did we spend on consultants?")] == ["consultants"]


def test_grain_feasibility_and_proxy():
    g = semantic.grains("EBITDA per product")[0]
    ok, proxy = semantic.feasible("ebitda", g)
    assert g.key == "product" and not ok and "contribution margin" in proxy
    assert semantic.feasible("revenue", g) == (True, None)
    assert not semantic.feasible("ebitda_margin", semantic.grains("per customer")[0])[0]


def test_currency_only_nok():
    assert semantic.currency("Revenue in EUR for Technology") == ("EUR", False)
    assert semantic.currency("Revenue for Technology") == ("NOK", True)


def _capability_problems() -> list[str]:
    """Broken references in capabilities.toml: unknown catalog concept/KPI, a synonym used by two measures, a grain
    naming an unknown measure, a vague measure without options."""
    cat = catalog.load()
    concepts = {ln.concept for lines in cat.pnl.values() for ln in lines if ln.concept}
    kpis = {k.key for k in cat.kpis}
    out: list[str] = []
    owner: dict[str, str] = {}
    keys = {m.key for m in semantic.all_measures()}
    for m in semantic.all_measures():
        if sum(bool(x) for x in (m.concept, m.kpi, m.account, m.accounts)) != (0 if m.vague else 1):
            out.append(f"measure {m.key}: exactly one of concept/kpi/account/accounts (none when vague)")
        if m.concept and m.concept not in concepts:
            out.append(f"measure {m.key}: unknown catalog concept {m.concept}")
        if m.kpi and m.kpi not in kpis:
            out.append(f"measure {m.key}: unknown catalog kpi {m.kpi}")
        if m.vague and not m.options:
            out.append(f"measure {m.key}: a vague measure lists options")
        for s in m.synonyms:
            if s.lower() in owner:
                out.append(f"synonym '{s}' used by {owner[s.lower()]} and {m.key}")
            owner[s.lower()] = m.key
    for g in semantic.all_grains():
        out += [f"grain {g.key}: unknown measure {k}" for k in g.feasible if k != "*" and k not in keys]
        if "*" not in g.feasible and not g.proxy:
            out.append(f"grain {g.key}: a restricted grain names its proxy")
    return out


def test_capabilities_reference_only_known_concepts():
    assert _capability_problems() == []


def test_describe_states_the_definition():
    assert "ac.C60080 + ac.C60090" in semantic.describe(semantic.measure("consultants"))
    assert semantic.describe(semantic.measure("opex_per_fte")).startswith("Opex per FTE")


@pytest.mark.parametrize("q,key", [
    ("Hvilke konsulentselskaper fakturerte mest til Telenor Norge i 2026?", "consultants"),
    ("Which consulting firms invoiced us most this year?", "consultants"),
    ("Hvor mye brukte vi på innleie i september?", "contractors"),
    ("Travel costs YTD for Business", "travel"), ("advokatkostnader 2026", "legal_fees"),
    ("Legal costs this year", "legal"), ("legal fees 2026", "legal_fees"), ("legal disputes Q3", "legal_disputes")])
def test_cost_type_measures_match_en_and_no(q, key):
    assert semantic.measures(q)[0].key == key


def test_consultants_sum_c60080_and_c60090():
    assert semantic.measure("consultants").accounts == ("ac.C60080", "ac.C60090")
    assert semantic.measure("consultants_and_temps").accounts == ("ac.C693",)

from smartview_mcp.core import period  # noqa: E402


@pytest.mark.parametrize("text,kind,label,gap", [
    ("Business EBITDA Sep 26 vs Sep 25", "month", "Sep 26", None),
    ("EBITDA YTD September 2026", "ytd", "YTD Sep 26", None),
    ("Q3 numbers for Consumer", "quarter", "Q3 26", None),
    ("full-year EBITDA latest forecast", "fy", "FY 26", None),
    ("Opex for Wholesale this year", "this_year", None, "YTD to Sep 26, or full year (actual + forecast)?"),
    ("Cost per FTE in IT", "none", "YTD Sep 26", None),
    ("Forecast for October 26 actuals", "month", "Oct 26", None),
    ("Driftskostnader hittil i år", "ytd", "YTD Sep 26", None),
    ("What may we expect in mars 2026?", "month", "Mar 26", None),
    ("EBITDA per month 2026 vs 2025", "year", "2026", None),
])
def test_period_parse(text, kind, label, gap):
    p = period.parse(text, "Sep 26")
    assert (p.kind, p.label, p.gap) == (kind, label, gap)


def test_open_month_is_flagged():
    assert period.parse("Oct 26 actuals", "Sep 26").open and not period.parse("Sep 26", "Sep 26").open
    assert period.parse("Cost per FTE", "Sep 26").convention


def test_remember_then_lookup_roundtrip(tmp_path, monkeypatch):
    from smartview_mcp.core import prefs
    monkeypatch.setattr(prefs.config, "PREFERENCES", tmp_path / "p.toml")
    assert prefs.lookup("measure", "costs") is None
    prefs.remember("measure", "Costs", "opex")
    assert prefs.lookup("measure", "costs") == "opex"
    prefs.remember("measure", "costs", "personnel")          # replaces, never duplicates
    assert prefs.lookup("measure", "COSTS") == "personnel"
    assert (tmp_path / "p.toml").read_text(encoding="utf-8").count("[[default]]") == 1


def test_remember_rejects_unknown_item_or_value(tmp_path, monkeypatch):
    from smartview_mcp.core import prefs
    from smartview_mcp.host import InputError
    monkeypatch.setattr(prefs.config, "PREFERENCES", tmp_path / "p.toml")
    with pytest.raises(InputError):
        prefs.remember("colour", "x", "y")
    with pytest.raises(InputError):
        prefs.remember("measure", "costs", "not_a_measure")
    with pytest.raises(InputError):
        prefs.remember("measure", "costs", "numbers")          # a vague measure is no answer
    with pytest.raises(InputError):
        prefs.remember("period", "this year", "sometimes")
    assert not (tmp_path / "p.toml").exists()


from smartview_mcp.core import scope  # noqa: E402

GATE_CASES = [
    ("B09", "What do we earn on mobile subscriptions?", "clear"),
    ("B10", "EBITDA per product for Mobile Consumer", "infeasible"),
    ("B11", "How are we doing on costs?", "ambiguous"),
    ("B13", "EBITDA YTD for unit un.SY.5OK_W", "clear"),
    ("B14", "Business full-year EBITDA: latest forecast vs budget", "clear"),
    ("T01", "What's our budget for Business?", "ambiguous"),
    ("T02", "Opex for Wholesale this year", "ambiguous"),
    ("T03", "EBITDA margin per customer", "infeasible"),
    ("T04", "Revenue in EUR for Technology", "infeasible"),
    ("T05", "Profit per product", "infeasible"),
    ("T06", "Business EBITDA Sep 26 vs Sep 25", "clear"),
    ("T07", "Q3 numbers for Consumer", "ambiguous"),
    ("T08", "Cost per FTE in IT", "clear"),
    ("T09", "How much did we spend on consultants?", "clear"),
    ("T10", "Forecast for October 26 actuals", "infeasible"),
    ("T11", "Hvilke konsulentselskaper fakturerte mest til Telenor Norge i 2026? og på hvilke initiativer / prosjekt",
     "clear"),
    ("T12", "Which suppliers invoiced most on travel in 2026?", "clear"),
]
_NONE = lambda i, p: None  # noqa: E731


@pytest.mark.parametrize("cid,q,want", GATE_CASES, ids=[c[0] for c in GATE_CASES])
def test_gate_outcome(cid, q, want):
    """The advisor benchmark: if a case fails, fix the rule or capabilities.toml, never the expected outcome."""
    g = scope.gate(q, "Sep 26", learned=_NONE)
    assert g["suggested_outcome"] == want, g


def test_gate_learned_default_turns_an_ask_into_a_stated_assumption():
    d = {("measure", "costs"): "opex", ("period", "this year"): "ytd"}
    g = scope.gate("How are we doing on costs this year?", "Sep 26", learned=lambda i, p: d.get((i, p)))
    assert g["suggested_outcome"] == "clear"
    assert g["checklist"]["measure"]["source"] == g["checklist"]["period"]["source"] == "learned"
    assert g["checklist"]["period"]["value"] == "YTD Sep 26"


def test_gate_details():
    t02 = scope.gate("Opex for Wholesale this year", "Sep 26", learned=_NONE)
    assert len(t02["questions"]) == 1 and "full year" in t02["questions"][0]
    b13 = scope.gate("EBITDA YTD for unit un.SY.5OK_W", "Sep 26", learned=_NONE)
    assert any("last year only" in w for w in b13["warnings"]) and b13["entities"][0]["finplan"] is None
    t10 = scope.gate("Forecast for October 26 actuals", "Sep 26", learned=_NONE)
    assert "Oct 26" in t10["reason"] and "forecast" in t10["proxy"]
    t08 = scope.gate("Cost per FTE in IT", "Sep 26", learned=_NONE)
    assert t08["checklist"]["measure"]["value"].startswith("Opex per FTE")     # the definition is stated
    t09 = scope.gate("How much did we spend on consultants?", "Sep 26", learned=_NONE)
    assert "C60080" in t09["checklist"]["measure"]["value"] and "temporary" in t09["checklist"]["measure"]["value"]
    b09 = scope.gate("What do we earn on mobile subscriptions?", "Sep 26", learned=_NONE)
    assert b09["checklist"]["measure"]["source"] == "catalog" and b09["concepts"]
    assert b09["checklist"]["entity"]["source"] == b09["checklist"]["period"]["source"] == "convention"
    t06 = scope.gate("Business EBITDA Sep 26 vs Sep 25", "Sep 26", learned=_NONE)
    assert "last year" in t06["checklist"]["scenario"]["value"]
    b14 = scope.gate("Business full-year EBITDA: latest forecast vs budget", "Sep 26", learned=_NONE)
    assert b14["checklist"]["scenario"]["value"] == "Latest forecast vs Target (budget)"


def test_gate_supplier_question_states_definition_grain_and_ytd():
    g = scope.gate(GATE_CASES[-2][1], "Sep 26", learned=_NONE)
    c = g["checklist"]
    assert "ac.C60080 + ac.C60090" in c["measure"]["value"]
    assert "fakturerte mest" in c["granularity"]["value"] and "initiativer" in c["granularity"]["value"]
    assert c["period"]["value"].startswith("YTD Sep 26") and c["period"]["source"] == "convention"
    assert any("ledger lines" in w for w in g["warnings"])
    assert g["measures"][0]["accounts"] == ["ac.C60080", "ac.C60090"]
    last = scope.gate("Travel costs 2025", "Sep 26", learned=_NONE)            # a past year stays the year
    assert last["checklist"]["period"]["value"] == "2025"


@pytest.mark.anyio
async def test_sv_gate_tool_uses_the_closed_month(monkeypatch):
    from smartview_mcp import tools_core

    async def closed():
        return "Sep 26", None
    monkeypatch.setattr(tools_core, "_closed_month", closed)
    monkeypatch.setattr(tools_core.prefs, "lookup", lambda i, p: None)
    out = await tools_core.sv_gate("Forecast for October 26 actuals")
    assert out["suggested_outcome"] == "infeasible" and out["closed_month"] == "Sep 26"


@pytest.mark.anyio
async def test_sv_gate_tool_falls_back_when_epm_is_down(monkeypatch):
    from smartview_mcp import tools_core

    def down(h):
        raise RuntimeError("no session")
    monkeypatch.setattr(tools_core.templates, "closed_month", down)
    monkeypatch.setattr(tools_core.prefs, "lookup", lambda i, p: None)
    out = await tools_core.sv_gate("Business EBITDA Sep 26 vs Sep 25")
    assert out["suggested_outcome"] == "clear" and any("assumed last month" in w for w in out["warnings"])


@pytest.mark.anyio
async def test_sv_remember_tool(tmp_path, monkeypatch):
    from mcp.server.mcpserver.exceptions import ToolError

    from smartview_mcp import tools_core
    monkeypatch.setattr(tools_core.config, "PREFERENCES", tmp_path / "p.toml")
    out = await tools_core.sv_remember("measure", "costs", "opex")
    assert out["saved"]["value"] == "opex" and out["file"].endswith("p.toml")
    with pytest.raises(ToolError):
        await tools_core.sv_remember("measure", "costs", "nonsense")


from smartview_mcp.core import commentary, contract  # noqa: E402

PREP = {"columns": [{"header": "Sep 26"}, {"header": "Sep 25"}, {"header": "Var"}],
        "sections": [{"lines": [
            {"key": "rev", "label": "Revenue", "cells": [{"pull": "m"}, {"pull": "m"}, {"var": [0, 1]}]},
            {"key": "small", "label": "Small", "cells": [{"pull": "m"}, {"pull": "m"}, {"var": [0, 1]}]},
            {"key": "big", "label": "Big cost", "cells": [{"pull": "m"}, {"pull": "m"}, {"var": [0, 1]}]},
            {"key": "fte", "label": "FTE", "cells": [{"pull": "m"}, {"pull": "m"}, {"var": [0, 1]}]},
            {"key": "basis", "aux": True, "cells": [{"pull": "m"}, {"pull": "m"}, {"var": [0, 1]}]}]}]}
VALUES = {"rev": [100.0, 99.0, 1.0], "small": [1_000.0, 2_000.0, -1_000.0], "big": [-20e6, -10e6, -10e6],
          "fte": [900.0, 1000.0, -100.0], "basis": [-9e9, 0.0, -9e9]}


def test_material_uses_catalog_thresholds_and_always_list():
    flagged = commentary.material(PREP, VALUES, catalog.load())
    # rev: always list; big: >= NOK 5m; small: under the 500k floor; fte: a level, 10 % without NOK floor; aux skipped
    assert [f["line"] for f in flagged] == ["rev", "big", "fte"]
    assert not flagged[0]["material"] and flagged[0]["always"]
    assert flagged[1]["delta"] == -10e6 and flagged[1]["pct"] == pytest.approx(-1.0)


def test_material_line_without_driver_is_a_warning_always_line_is_info():
    flagged = commentary.material(PREP, VALUES, catalog.load())
    f = commentary.findings(flagged, [{"line": "big", "text": "x", "basis": "hypothesis"}])
    assert [(x["severity"], x["text"].split("'")[1]) for x in f] == [("Info", "Revenue"), ("Warning", "FTE")]


def test_driver_lines_label_hypotheses():
    out = commentary.sheet_lines([{"line": "big", "text": "More licences", "basis": "data", "evidence": "Q1 FinMax C60"},
                                  {"line": "rev", "text": "Price rise", "basis": "hypothesis"}])
    assert out == ["Driver: More licences (data: Q1 FinMax C60)", "Hypothesis — not verified in data: Price rise"]


_SPEC = {"version": 1, "question": "EBITDA Sep 26", "gate": {"outcome": "clear"},
         "pulls": [{"id": "m", "connection": "FinMax", "rows": [{"dimension": "All Accounts", "members": ["ac.EBITDA"]}],
                    "columns": [{"dimension": "All Time Periods", "members": ["Sep 26"]}]}],
         "layout": {"title": "t", "columns": [{"header": "Sep 26"}],
                    "sections": [{"lines": [{"key": "e", "label": "EBITDA",
                                             "cells": [{"pull": "m", "row": ["ac.EBITDA"], "col": ["Sep 26"]}]}]}]}}


def test_contract_driver_line_must_be_a_layout_line():
    ok = {**_SPEC, "commentary": {"drivers": [{"line": "e", "text": "x", "basis": "hypothesis"}]}}
    contract.validate(ok)
    bad = {**_SPEC, "commentary": {"drivers": [{"line": "nope", "text": "x", "basis": "hypothesis"}]}}
    with pytest.raises(Exception, match="unknown line 'nope'"):
        contract.validate(bad)


def test_drivers_join_the_key_points_on_the_sheet():
    from smartview_mcp.core import pipeline
    s = {**_SPEC, "commentary": {"drivers": [{"line": "e", "text": "Lower churn", "basis": "hypothesis"}]}}
    out = pipeline.answer_layout(s["layout"], contract.validate(s))
    assert out["key_points"] == ["Hypothesis — not verified in data: Lower churn"]
