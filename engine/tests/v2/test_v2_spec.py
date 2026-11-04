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
    bad = [{"type": "table", "query": "fm", "columns": ["value", "var"]},
           {"type": "table", "query": "fm", "columns": ["value", "share"]},
           {"type": "table", "query": "fm", "rows": "entities", "columns": ["value"]}]
    for b in bad:
        with pytest.raises(ValidationError):
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
