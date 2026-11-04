import pytest
from pydantic import ValidationError
from v2util import spec_dict

from smartass.model.spec import AnalysisSpec, DriversBlock, json_schema


def test_base_spec_is_valid():
    s = AnalysisSpec.model_validate(spec_dict())
    assert isinstance(s.blocks[4], DriversBlock) and s.blocks[4].depth == 2


def test_unknown_fields_are_refused():
    with pytest.raises(ValidationError, match="extra"):
        AnalysisSpec.model_validate(spec_dict(colour="red"))


def test_purpose_required_unless_quick():
    with pytest.raises(ValidationError, match="purpose is required"):
        AnalysisSpec.model_validate(spec_dict(plan={"depth": "standard"}))
    AnalysisSpec.model_validate(spec_dict(plan={"depth": "quick"}))


def test_budget_only_in_finplan():
    d = spec_dict()
    d["queries"][0]["compare"] = ["ly", "budget"]
    with pytest.raises(ValidationError, match="only in FinPlan"):
        AnalysisSpec.model_validate(d)


def test_block_must_reference_a_known_query_and_its_metrics_and_series():
    d = spec_dict(blocks=[{"type": "ranking", "query": "xx", "metric": "Opex"}])
    with pytest.raises(ValidationError, match="unknown query 'xx'"):
        AnalysisSpec.model_validate(d)
    d = spec_dict(blocks=[{"type": "ranking", "query": "fm", "metric": "Capex"}])
    with pytest.raises(ValidationError, match="not in query 'fm'"):
        AnalysisSpec.model_validate(d)
    d = spec_dict()
    d["queries"][0]["compare"] = []
    with pytest.raises(ValidationError, match="compare"):
        AnalysisSpec.model_validate(d)


def test_table_shape_rules():
    bad = [({"type": "table", "query": "fm", "columns": ["value", "var"]}, "var and var_pct need"),
           ({"type": "table", "query": "fm", "columns": ["value", "share"]}, "share needs rows"),
           ({"type": "table", "query": "fm", "rows": "entities", "columns": ["value"]},
            "rows 'entities' needs")]
    for b, msg in bad:
        with pytest.raises(ValidationError, match=msg):
            AnalysisSpec.model_validate(spec_dict(blocks=[b]))


def test_data_driver_needs_evidence():
    b = {"type": "commentary", "drivers": [{"line": "Opex", "text": "x", "basis": "data"}]}
    with pytest.raises(ValidationError, match="needs evidence"):
        AnalysisSpec.model_validate(spec_dict(blocks=[b]))


def test_duplicate_query_ids_refused():
    d = spec_dict()
    d["queries"].append(dict(d["queries"][0]))
    with pytest.raises(ValidationError, match="unique"):
        AnalysisSpec.model_validate(d)


def test_json_schema_lists_every_block_type():
    text = str(json_schema())
    for t in ("table", "kpi_tiles", "trend", "bridge", "ranking", "matrix", "drivers", "lines", "commentary"):
        assert f"'{t}'" in text


def test_unparseable_periods_refused_but_relative_ok():
    d = spec_dict()
    d["queries"][0]["periods"] = "September 2026"
    with pytest.raises(ValidationError, match="periods"):
        AnalysisSpec.model_validate(d)
    for ok in ("closed", "closed-ytd", "closed-12m", "fy", "Sep 26", "Jan 26..Sep 26"):
        d["queries"][0]["periods"] = ok
        AnalysisSpec.model_validate(d)


def test_table_variance_columns_need_a_base_to_the_left_and_duplicates_are_limited():
    from pydantic import ValidationError

    from smartass.model.spec import TableBlock
    ok = {"type": "table", "query": "fp", "columns": ["value", "budget", "var", "forecast", "var", "var_pct"]}
    TableBlock.model_validate(ok)
    for cols, msg in ((["value", "var", "ly"], "to their left"),
                      (["value", "ly", "ly"], "more than once"),
                      (["value", "ly", "var", "var"], "same base"),
                      (["value", "ly", "var", "var_pct", "var_pct"], "same base")):
        with pytest.raises(ValidationError, match=msg):
            TableBlock.model_validate({**ok, "columns": cols})


def test_plan_periods_validated_like_periods():
    d = spec_dict()
    d["queries"][0]["plan_periods"] = "next year"
    with pytest.raises(ValidationError, match="periods"):
        AnalysisSpec.model_validate(d)
    d["queries"][0]["plan_periods"] = "fy"
    AnalysisSpec.model_validate(d)
