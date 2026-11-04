# engine/tests/test_catalog.py
import pytest

from smartview_mcp import catalog


def _entity(cat, text):
    found = catalog.resolve_entities(cat, text)
    return found[0] if len(found) == 1 else None


OLD_PNL = [
    ("rev", "ac.C30", "line", None), ("cogs", "ac.C40S", "line", None), ("gm", "ac.C48", "sum", None),
    ("pers", "ac.C50S", "line", None), ("oth", "ac.C60", "line", None), ("opex", "ac.C70", "sum", None),
    ("boi", "ac.EBITDA_BOI", "sum", None), ("c73", "ac.C73", "line", None), ("ebitda", "ac.EBITDA", "grand", None),
    ("da", "ac.C75", "line", None), ("fte", "ac.STA20Y", "line", "tall1"),
]
OLD_PNL_CHECKS = [("C30 + C40S = C48", "rev+cogs-gm"), ("C50S + C60 = C70", "pers+oth-opex"),
                  ("C48 + C70 = EBITDA_BOI", "gm+opex-boi"), ("EBITDA_BOI + C73 = EBITDA", "boi+c73-ebitda")]
OLD_FP_PNL = [("rev", "Revenue", "line"), ("cogs", "COGS", "line"), ("gp", "Gross_Profit", "sum"),
              ("opex", "OPEX", "line"), ("boi", "EBITDA before Other Items", "sum"), ("oth", "Total Other Items", "line"),
              ("ebitda", "EBITDA", "grand")]
OLD_FP_CHECKS = [("Revenue + COGS = Gross profit", "rev+cogs-gp"),
                 ("Gross profit + OPEX = EBITDA before other items", "gp+opex-boi"),
                 ("EBITDA before other items + Other items = EBITDA", "boi+oth-ebitda")]


def test_repo_catalogue_carries_the_template_lists_unchanged():
    cat = catalog.load()
    assert [(x.key, x.account, x.kind, x.num) for x in cat.pnl["FinMax"]] == OLD_PNL
    assert [(c.label, c.expr) for c in cat.checks["FinMax"]] == OLD_PNL_CHECKS
    assert [(x.key, x.account, x.kind) for x in cat.pnl["FinPlan"]] == OLD_FP_PNL
    assert [(c.label, c.expr) for c in cat.checks["FinPlan"]] == OLD_FP_CHECKS


def test_repo_catalogue_has_areas_scenarios_kpis_materiality():
    cat = catalog.load()
    keys = {e.key for e in cat.entities}
    assert {"company", "wholesale", "business", "ssm", "ei", "technology", "it", "adm", "consumer"} <= keys
    assert cat.scenarios["budget"].scenario == "Target" and cat.scenarios["forecast"].rule == "latest_forecast"
    assert {k.key for k in cat.kpis} >= {"ebitda_margin", "opex_per_fte"}
    assert cat.materiality.pct == 0.10 and "ebitda" in cat.materiality.always


BASE = {
    "pnl": {"FinMax": [{"key": "rev", "account": "ac.C30", "kind": "line", "label": "Revenue"},
                       {"key": "ebitda", "account": "ac.EBITDA", "kind": "grand", "label": "EBITDA"}],
            "FinPlan": [{"key": "rev", "account": "Revenue", "kind": "line", "label": "Revenue"},
                        {"key": "gp", "account": "Gross_Profit", "kind": "grand", "label": "Gross profit"}]},
    "checks": {"FinMax": [{"label": "x", "expr": "rev-ebitda"}], "FinPlan": [{"label": "y", "expr": "rev-gp"}]},
    "scenarios": {"actual": {"connection": "FinMax", "scenario": "Actual", "label": "Actual"}},
    "materiality": {"pct": 0.1, "abs_nok": 1, "always": ["rev"]},
    "entity": [{"key": "w", "label": "W", "finmax": "un.SY.2WHS", "aliases": ["wholesale"]}],
    "kpi": [{"key": "m", "label": "M", "connection": "FinMax", "numerator": "ebitda", "denominator": "rev",
             "unit": "pct"}],
}


def _with(path, value):
    import copy
    d = copy.deepcopy(BASE)
    node = d
    for k in path[:-1]:
        node = node[k]
    node[path[-1]] = value
    return d


@pytest.mark.parametrize("data, msg", [
    (_with(["pnl", "FinMax", 1, "kind"], "total"), "kind"),
    (_with(["pnl", "FinMax", 1, "key"], "rev"), "duplicate"),
    (_with(["checks", "FinMax", 0, "expr"], "rev-nope"), "nope"),
    (_with(["kpi", 0, "denominator"], "nope"), "nope"),
    (_with(["materiality", "always"], ["nope"]), "nope"),
    (_with(["entity"], BASE["entity"] * 2), "duplicate"),
    (_with(["scenarios", "actual"], {"connection": "FinMax", "label": "A"}), "scenario or rule"),
    (_with(["entity", 0, "successors"], ["nope"]), "successor 'nope'"),
    (_with(["entity", 0, "generic_aliases"], "total"), r"entity\[0\]: generic_aliases must be a list"),
    (_with(["entity", 0, "generic_aliases"], ["  --  "]), r"entity 'w'.*empty"),
    (_with(["entity", 0, "aliases"], ["wholesale", "!!"]), r"entity 'w'.*empty"),
    (_with(["entity", 0, "aliases"], "wholesale"), r"entity\[0\]: aliases must be a list"),
    (_with(["materiality", "always"], "rev"), "materiality: always must be a list"),
    (_with(["scenarios", "actual", "connection"], "Nope"), "scenario 'actual': connection 'Nope' is not a pnl cube"),
    (_with(["kpi", 0, "connection"], "Nope"), "kpi 'm': unknown connection 'Nope'"),
    (_with(["kpi", 0, "unit"], "eur"), "kpi 'm': unit 'eur'"),
    (_with(["kpi"], BASE["kpi"] * 2), "duplicate kpi key 'm'"),
    (_with(["pnl", "FinMax", 0, "better"], "up"), "FinMax line 'rev': better 'up'"),
    (_with(["entity", 0, "finmax"], None), "entity 'w' needs finmax or finplan"),
    (_with(["entity", 0, "abolished"], "1 June 2026"), "entity 'w': abolished '1 June 2026' is not an ISO date"),
    (_with(["pnl", "FinMax", 0, "kind"], "grand"), "FinMax has 2 grand lines"),
    (_with(["scenarios", "actual", "rule"], "latest_forecast"), "scenario 'actual' has both scenario and rule"),
    (_with(["materiality", "pct"], 0), r"materiality: pct 0.0 is not in \(0, 1\]"),
    (_with(["materiality", "pct"], 10), r"materiality: pct 10.0 is not in \(0, 1\]"),
    (_with(["materiality", "abs_nok"], -1), "materiality: abs_nok -1.0 is negative"),
    (_with(["materiality", "pct_floor_nok"], -1), "materiality: pct_floor_nok -1.0 is negative"),
    (_with(["materiality", "pct"], "ten"), "materiality: pct 'ten' is not a number"),
    (_with(["materiality", "abs_nok"], None), "materiality: abs_nok None is not a number"),
    (_with(["pnl", "FinMax", 0, "acount"], "x"), r"pnl.FinMax\[0\]: unknown field 'acount'"),
    (_with(["entity", 0, "alias"], ["x"]), r"entity\[0\]: unknown field 'alias'"),
    (_with(["kpi", 0, "nominator"], "rev"), r"kpi\[0\]: unknown field 'nominator'"),
    (_with(["scenarios", "actual", "versoin"], "x"), r"scenarios\[actual\]: unknown field 'versoin'"),
    (_with(["checks", "FinMax", 0, "exp"], "x"), r"checks.FinMax\[0\]: unknown field 'exp'"),
    (_with(["materiality", "pcts"], 1), "materiality: unknown field 'pcts'"),
    (_with(["kpi", 0], {"key": "m"}), r"kpi\[0\]: missing field"),
    (_with(["entity", 0], "w"), r"entity\[0\]: expected a table"),
    (_with(["checks", "FinMax", 0, "expr"], "rev*ebitda"), r"FinMax check 'x': expression 'rev\*ebitda'"),
    (_with(["checks", "FinMax", 0, "expr"], "rev--ebitda"), "FinMax check 'x': expression"),
    (_with(["checks", "FinMax", 0, "expr"], "rev+1"), "FinMax check 'x': expression"),
    (_with(["pnl", "FinMax"], [{**BASE["pnl"]["FinMax"][0], "concept": "x"},
                              {**BASE["pnl"]["FinMax"][1], "concept": "x"}]), "duplicate FinMax concept 'x'"),
    (_with(["pnl", "FinMax", 0, "measure"], "level"), "FinMax line 'rev': measure 'level'"),
    (_with(["kpi", 0, "basis"], "avg"), "kpi 'm': basis 'avg'"),
    (_with(["kpi", 0, "annualise"], "yes"), "kpi 'm': annualise must be true or false"),
    (_with(["kpi", 0, "basis"], "average_stock"), "kpi 'm': basis average_stock needs a stock denominator"),
    (_with(["materiality", "rule"], "and"), "materiality: rule 'and'"),
    (_with(["pnl", "Finmax"], BASE["pnl"]["FinMax"]), r"\[pnl\]: unknown cube 'Finmax'"),
    (_with(["checks", "FinLite"], []), r"\[checks\]: unknown cube 'FinLite'"),
    (_with(["pnl"], {"FinMax": BASE["pnl"]["FinMax"]}), r"\[pnl\]: missing cube 'FinPlan'"),
    (_with(["checks"], {"FinPlan": BASE["checks"]["FinPlan"]}), r"\[checks\]: missing cube 'FinMax'"),
    ({k: v for k, v in BASE.items() if k != "checks"}, r"\[checks\]: missing cube"),
])
def test_parse_rejects_broken_catalogues(data, msg):
    with pytest.raises(catalog.CatalogError, match=msg):
        catalog.parse(data)


@pytest.mark.parametrize("text, key", [
    ("P&L for Wholesale Sep 26", "wholesale"),
    ("hvordan går det med bedriftsmarkedet?", "business"),
    ("EBITDA i TV og internett hittil i år", "ei"),
    ("un.SY.2TEC", "technology"),
    ("2SSM Security Solutions & Mobile", "ssm"),
    ("it", "it"),
    ("Telenor Norge total", "company"),
])
def test_resolve_entity(text, key):
    assert _entity(catalog.load(), text).key == key


@pytest.mark.parametrize("text", ["is it good?", "what about ei?", "something else entirely"])
def test_resolve_entity_needs_a_real_match(text):
    assert _entity(catalog.load(), text) is None


@pytest.mark.parametrize("text, key", [
    ("total opex for SSM", "ssm"),
    ("Total for IT", "it"),
    ("total opex for Telenor Norge", "company"),
    ("Telenor Norge", "company"),
    ("kostnader i IT", "it"),
    ("kostnad per årsverk i IT", "it"),
    ("hva bruker bedriftsmarked på konsulenter?", "business"),
    ("EBITDA i grossistmarkedet", "wholesale"),
    ("kostnader i administrasjonen", "adm"),
    ("stabsenheter Q3", "adm"),
    ("admin opex", "adm"),
    ("tv & internett hittil i år", "ei"),
    ("opex 2NRD_N Sep 26", "nordics"),
    ("un.SY.2NRD_N", "nordics"),
    ("Consumer Q3 mot i fjor", "consumer"),
])
def test_resolve_entity_in_a_sentence(text, key):
    assert _entity(catalog.load(), text).key == key


@pytest.mark.parametrize("text", ["mobil", "mobilkostnader i fjor", "Wholesale vs Business",
                                  "Telenor Norge vs Wholesale"])
def test_resolve_entity_none_when_not_one_area(text):
    assert _entity(catalog.load(), text) is None


@pytest.mark.parametrize("text, keys", [
    ("Wholesale vs Business", ("wholesale", "business")),
    ("Business vs Wholesale vs Business", ("business", "wholesale")),
    ("IT-divisjonen mot Technology", ("it", "technology")),
    ("Security Solutions & Mobile og E&I", ("ssm", "ei")),
    ("Total for IT og SSM", ("it", "ssm")),
    ("Telenor Norge vs Wholesale", ("company", "wholesale")),
    ("Wholesale mot TN", ("wholesale", "company")),
    ("telenor norge total mot Business", ("company", "business")),
    ("selskapet og Wholesale", ("wholesale",)),
    ("company total vs IT", ("it",)),
    ("total opex for SSM", ("ssm",)),
    ("total opex", ("company",)),
    ("nothing here", ()),
])
def test_resolve_entities_in_order_of_appearance(text, keys):
    assert tuple(e.key for e in catalog.resolve_entities(catalog.load(), text)) == keys


def test_abolished_consumer_names_its_successors():
    cat = catalog.load()
    con = _entity(cat, "privatmarkedet")
    assert con.key == "consumer" and con.abolished == "2026-06-01" and con.successors == ("ssm", "ei")


def test_catalogue_tables_are_read_only():
    cat = catalog.parse(BASE)
    with pytest.raises(TypeError):
        cat.pnl["FinMax"] = ()
    with pytest.raises(TypeError):
        cat.scenarios["x"] = cat.scenarios["actual"]
    with pytest.raises(TypeError):
        cat.checks["FinMax"] = ()


def test_parse_accepts_the_base_and_a_toml_date():
    assert catalog.parse(BASE).entities[0].key == "w"
    import datetime
    cat = catalog.parse(_with(["entity", 0, "abolished"], datetime.date(2026, 6, 1)))
    assert cat.entities[0].abolished == "2026-06-01"


CONCEPTS = {
    "FinMax": {"rev": "revenue", "cogs": "cogs", "gm": "gross_margin", "pers": "personnel", "oth": "other_opex",
               "opex": "opex", "boi": "ebitda_boi", "c73": "other_items", "ebitda": "ebitda", "da": "da", "fte": "fte"},
    "FinPlan": {"rev": "revenue", "cogs": "cogs", "gp": "gross_margin", "opex": "opex", "boi": "ebitda_boi",
                "oth": "other_items", "ebitda": "ebitda"},
}


def test_lines_carry_their_cross_cube_concept():
    cat = catalog.load()
    assert {cube: {x.key: x.concept for x in lines} for cube, lines in cat.pnl.items()} == CONCEPTS
    assert cat.line_by_concept("FinMax", "other_items").key == "c73"
    assert cat.line_by_concept("FinPlan", "other_items").key == "oth"
    assert cat.line_by_concept("FinMax", "other_opex").key == "oth"
    assert cat.line_by_concept("FinPlan", "gross_margin").key == "gp"
    assert cat.line_by_concept("FinPlan", "fte") is None and cat.line_by_concept("Nope", "ebitda") is None


def test_materiality_is_either_threshold():
    m = catalog.load().materiality
    assert (m.pct, m.abs_nok, m.rule) == (0.10, 5_000_000, "either")
    assert m.always == ("rev", "pers", "ebitda")
    assert catalog.parse(BASE).materiality.rule == "either"   # the default
    assert m.pct_floor_nok == 500_000                         # owner decision 2026-10-25
    assert catalog.parse(BASE).materiality.pct_floor_nok == 0  # the default: no floor


def test_ig_successors_are_the_five_destinations():
    ig = next(e for e in catalog.load().entities if e.key == "ig")
    assert ig.successors == ("technology", "adm", "ei", "ssm", "business")   # TEC first: the largest destination


def test_fte_is_a_stock_and_per_fte_kpis_use_the_average_level():
    cat = catalog.load()
    assert {x.key: x.measure for x in cat.pnl["FinMax"]}["fte"] == "stock"
    assert {x.measure for x in cat.pnl["FinMax"] if x.key != "fte"} == {"flow"}
    kpis = {k.key: k for k in cat.kpis}
    for key in ("opex_per_fte", "personnel_per_fte"):
        assert (kpis[key].basis, kpis[key].annualise) == ("average_stock", True)
    assert (kpis["ebitda_margin"].basis, kpis["ebitda_margin"].annualise) == ("period", False)


def test_only_generic_company_words_give_way():
    company = {e.key: e for e in catalog.load().entities}["company"]
    assert set(company.generic_aliases) == {"total", "company", "selskapet"}
    assert "telenor norge" in company.aliases and "total" not in company.aliases
    assert _entity(catalog.load(), "total").key == "company"
    assert _entity(catalog.load(), "selskapet hittil i år").key == "company"


def _two_areas(**changes):
    import copy
    d = _with(["entity"], [*copy.deepcopy(BASE["entity"]), {"key": "b", "label": "B", "finmax": "un.SY.2BUS", "aliases": ["bedrift"]}])
    d["entity"][0].update(changes)
    return d


def test_a_name_may_not_belong_to_two_areas():
    with pytest.raises(catalog.CatalogError, match=r"'un sy 2bus'.*'w'.*'b'|'un sy 2bus'.*'b'.*'w'"):
        catalog.parse(_two_areas(aliases=["wholesale", "un.SY.2BUS"]))
    with pytest.raises(catalog.CatalogError, match="'bedrift'"):   # a generic alias, normalised
        catalog.parse(_two_areas(generic_aliases=["Bedrift"]))
    with pytest.raises(catalog.CatalogError, match="'b'"):         # an alias equal to another area's label
        catalog.parse(_two_areas(aliases=["B"]))


def test_a_name_may_repeat_within_one_area():
    cat = catalog.parse(_two_areas(label="Wholesale", generic_aliases=["Wholesale"]))
    assert cat.entities[0].key == "w"


def _atlas_with(tmp_path, members):
    from smartview_mcp import atlas
    db = atlas.connect(tmp_path / "a.db")
    for (conn, dim), names in members.items():
        atlas.store(db, conn, dim, [{"name": n, "label": n, "parent": None, "depth": 1, "order": i}
                                    for i, n in enumerate(names)], 0.1)
    return db


def test_validate_names_missing_members_and_unchecked_dimensions(tmp_path):
    cat = catalog.parse(BASE)
    db = _atlas_with(tmp_path, {("FinMax", "All Accounts"): ["ac.C30"],
                                ("FinMax", "All Entities"): ["un.SY.2WHS"],
                                ("FinPlan", "Account"): ["Revenue", "Gross_Profit"]})
    problems, unchecked, notes = catalog.validate(cat, db)
    assert notes == [] and problems == ["FinMax line 'ebitda': 'ac.EBITDA' is not in FinMax / All Accounts"]
    assert unchecked == ["FinMax / All Scenarios"]


def test_validate_passes_a_complete_atlas(tmp_path):
    cat = catalog.parse(BASE)
    db = _atlas_with(tmp_path, {("FinMax", "All Accounts"): ["ac.C30", "ac.EBITDA"],
                                ("FinMax", "All Entities"): ["un.SY.2WHS"],
                                ("FinMax", "All Scenarios"): ["Actual"],
                                ("FinPlan", "Account"): ["Revenue", "Gross_Profit"]})
    assert catalog.validate(cat, db) == ([], [], [])


def test_validate_reports_a_missing_entity_and_scenario(tmp_path):
    cat = catalog.parse(BASE)
    db = _atlas_with(tmp_path, {("FinMax", "All Accounts"): ["ac.C30", "ac.EBITDA"],
                                ("FinMax", "All Entities"): ["un.SY.2BUS"],
                                ("FinMax", "All Scenarios"): ["Budget"],
                                ("FinPlan", "Account"): ["Revenue", "Gross_Profit"]})
    problems, _, _ = catalog.validate(cat, db)
    assert problems == ["entity 'w': 'un.SY.2WHS' is not in FinMax / All Entities",
                        "scenario 'actual': 'Actual' is not in FinMax / All Scenarios"]


def test_doctor_reports_catalogue_problems(monkeypatch, tmp_path):
    from smartview_mcp import doctor
    monkeypatch.setattr(doctor.config, "ATLAS_DB", tmp_path / "missing.db")
    c = doctor._catalog()
    assert c.name == "catalog" and c.status == "WARN" and "no atlas" in c.detail
    assert "_catalog" in [f.__name__ for f in doctor.CHECKS]


def test_templates_read_the_catalogue():
    import inspect

    from smartview_mcp import templates
    assert templates.fp_lines() == OLD_FP_PNL and templates.fp_checks() == OLD_FP_CHECKS
    assert '("rev", "ac.C30"' not in inspect.getsource(templates)   # no second copy of the list


def test_a_broken_catalogue_fails_only_the_pnl_templates(monkeypatch, tmp_path):
    import importlib

    from smartass.core import config
    from smartview_mcp import templates
    (tmp_path / "catalog.toml").write_text("[materiality]\npct = 10\n", encoding="utf-8")
    monkeypatch.setattr(config, "KNOWLEDGE_DIR", tmp_path)
    monkeypatch.setattr(config, "PACKS", [])
    catalog._load.cache_clear()
    try:
        importlib.reload(templates)                 # import must not read the catalogue
        import smartview_mcp.server  # noqa: F401
        with pytest.raises(catalog.CatalogError):
            templates.fp_lines()
    finally:
        monkeypatch.undo()
        catalog._load.cache_clear()
        importlib.reload(templates)


def test_validate_notes_a_missing_member_of_an_abolished_area(tmp_path):
    cat = catalog.parse(_two_areas(abolished="2026-06-01"))
    db = _atlas_with(tmp_path, {("FinMax", "All Accounts"): ["ac.C30", "ac.EBITDA"],
                                ("FinMax", "All Entities"): ["un.SY.2BUS"],
                                ("FinMax", "All Scenarios"): ["Actual"],
                                ("FinPlan", "Account"): ["Revenue", "Gross_Profit"]})
    assert catalog.validate(cat, db) == (
        [], [], ["w: 'un.SY.2WHS' no longer in FinMax / All Entities (abolished 2026-06-01)"])


def test_doctor_warns_when_a_dimension_is_unchecked(monkeypatch, tmp_path):
    from smartview_mcp import atlas, doctor
    atlas.connect(tmp_path / "a.db").close()
    monkeypatch.setattr(doctor.config, "ATLAS_DB", tmp_path / "a.db")
    c = doctor._catalog()
    assert c.status == "WARN" and "not checked" in c.detail


def test_annualise_needs_a_flow_numerator_and_per_fte_unit_is_allowed():
    kpi = {"key": "m", "label": "M", "connection": "FinMax", "numerator": "fte", "denominator": "rev",
           "unit": "nok", "annualise": True}
    d = _with(["pnl", "FinMax"], [*BASE["pnl"]["FinMax"], {"key": "fte", "account": "ac.F", "kind": "line",
                                                           "label": "FTE", "measure": "stock"}])
    d["kpi"] = [kpi]
    with pytest.raises(catalog.CatalogError, match="annualise needs a flow numerator"):
        catalog.parse(d)
    d["kpi"] = [{**kpi, "numerator": "ebitda", "denominator": "fte", "unit": "nok_per_fte",
                 "basis": "average_stock", "annualise": True}]
    assert catalog.parse(d).kpis[0].unit == "nok_per_fte"
    assert {k.key: k.unit for k in catalog.load().kpis}["opex_per_fte"] == "nok_per_fte"


def test_kpi_sign_better_and_headline_are_validated():
    def with_kpi(**changes):
        d = _with(["kpi", 0], {**BASE["kpi"][0], **changes})
        return d
    assert catalog.parse(with_kpi(sign=-1, better="lower", headline=True)).kpis[0].headline is True
    for bad, msg in ((dict(sign=2), "sign must be 1 or -1"), (dict(sign=True), "sign must be 1 or -1"),
                     (dict(better="up"), "better 'up'"), (dict(headline="yes"), "headline must be true or false")):
        with pytest.raises(catalog.CatalogError, match=msg):
            catalog.parse(with_kpi(**bad))
    two = _with(["kpi"], [{**BASE["kpi"][0], "headline": True}, {**BASE["kpi"][0], "key": "m2", "headline": True}])
    with pytest.raises(catalog.CatalogError, match="at most one headline kpi per connection"):
        catalog.parse(two)
    cat = catalog.load()
    assert [k.key for k in cat.kpis if k.headline] == ["personnel_per_fte"]
    assert {k.key: (k.sign, k.better) for k in cat.kpis if k.sign == -1} == {
        "opex_per_fte": (-1, "lower"), "personnel_per_fte": (-1, "lower")}
