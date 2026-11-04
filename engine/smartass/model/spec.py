# engine/smartass/model/spec.py
"""Analysis spec v2: the contract the finance agent writes. Typed blocks, relative periods, strict (unknown
fields refused). Cross-references (block -> query, metrics, series) are checked here; members and metrics are
checked against knowledge in validate.py."""
from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from .periods import resolve

Compare = Literal["ly", "budget", "forecast"]
Series = Literal["value", "ly", "budget", "forecast"]
Column = Literal["value", "ly", "budget", "forecast", "var", "var_pct", "share"]
BASE_SERIES = ("ly", "budget", "forecast")
PLAN_ONLY = ("budget", "forecast")


class _M(BaseModel):
    model_config = ConfigDict(extra="forbid")


class Plan(_M):
    depth: Literal["quick", "standard", "deep"]
    purpose: str = ""
    hypotheses: list[str] = []
    audience: Literal["controller", "management", "analyst"] = "controller"

    @model_validator(mode="after")
    def _purpose(self) -> "Plan":
        if self.depth != "quick" and not self.purpose.strip():
            raise ValueError("plan.purpose is required unless plan.depth is 'quick'")
        return self


class Query(_M):
    id: str = Field(pattern=r"^[a-z][a-z0-9_]{0,15}$")
    cube: Literal["FinMax", "FinLite", "FinPlan"]
    metrics: list[str] = Field(min_length=1)
    entity: str
    periods: str
    compare: list[Compare] = []
    preliminary: bool = False

    @field_validator("periods")
    @classmethod
    def _periods_parse(cls, v: str) -> str:
        try:
            resolve(v, (2000, 1))
        except ValueError as e:
            raise ValueError(f"periods {v!r}: {e}") from e
        return v

    @model_validator(mode="after")
    def _plan_only(self) -> "Query":
        bad = [c for c in self.compare if c in PLAN_ONLY]
        if bad and self.cube != "FinPlan":
            raise ValueError(f"query {self.id}: {', '.join(bad)} exist only in FinPlan "
                             "(actual vs plan is done in FinPlan)")
        return self


class TableBlock(_M):
    type: Literal["table"]
    query: str
    title: str = ""
    rows: Literal["metrics", "entities"] = "metrics"
    metrics: list[str] = []
    metric: str | None = None
    columns: list[Column] = ["value", "ly", "var", "var_pct"]
    top: int | None = Field(None, ge=1)
    other: bool = True
    sort: Literal["none", "desc", "asc", "abs"] = "none"

    @model_validator(mode="after")
    def _shape(self) -> "TableBlock":
        cols = self.columns
        has_base = any(c in cols for c in BASE_SERIES)
        if ("var" in cols or "var_pct" in cols) and not ("value" in cols and has_base):
            raise ValueError("table: var and var_pct need 'value' and one of ly/budget/forecast in columns")
        if "share" in cols and (self.rows != "entities" or "value" not in cols):
            raise ValueError("table: share needs rows 'entities' and a 'value' column")
        if self.rows == "entities" and (not self.metric or "value" not in cols):
            raise ValueError("table: rows 'entities' needs 'metric' and a 'value' column")
        return self


class KpiBlock(_M):
    type: Literal["kpi_tiles"]
    query: str
    title: str = ""
    metrics: list[str] = Field(min_length=1, max_length=6)
    compare: Compare = "ly"


class TrendBlock(_M):
    type: Literal["trend"]
    query: str
    title: str = ""
    metrics: list[str] = Field(min_length=1)
    series: list[Series] = ["value", "ly"]
    cumulative: bool = False
    chart: bool = True


class BridgeBlock(_M):
    type: Literal["bridge"]
    query: str
    title: str = ""
    total: str
    steps: list[str] = Field(min_length=1)
    base: Compare = "ly"


class RankingBlock(_M):
    type: Literal["ranking"]
    query: str
    title: str = ""
    metric: str
    top: int = Field(10, ge=1)
    compare: Compare | None = "ly"


class MatrixBlock(_M):
    type: Literal["matrix"]
    query: str
    title: str = ""
    metric: str


class DriversBlock(_M):
    type: Literal["drivers"]
    query: str
    title: str = ""
    metric: str
    base: Compare = "ly"
    depth: int = Field(2, ge=1, le=4)
    drill: bool = False


class LinesBlock(_M):
    type: Literal["lines"]
    title: str = ""
    cell: dict[str, str]
    group_by: Literal["supplier", "project", "account"] = "supplier"


class Driver(_M):
    line: str
    text: str
    basis: Literal["data", "hypothesis"]
    evidence: str = ""

    @model_validator(mode="after")
    def _evidence(self) -> "Driver":
        if self.basis == "data" and not self.evidence.strip():
            raise ValueError(f"driver '{self.line}': basis 'data' needs evidence (a cell key)")
        return self


class CommentaryBlock(_M):
    type: Literal["commentary"]
    title: str = "Commentary"
    key_points: list[str] = []
    drivers: list[Driver] = []


Block = Annotated[Union[TableBlock, KpiBlock, TrendBlock, BridgeBlock, RankingBlock, MatrixBlock, DriversBlock,
                        LinesBlock, CommentaryBlock], Field(discriminator="type")]


def block_metrics(b) -> list[str]:
    if isinstance(b, TableBlock):
        return list(b.metrics) + ([b.metric] if b.metric else [])
    if isinstance(b, (KpiBlock, TrendBlock)):
        return list(b.metrics)
    if isinstance(b, BridgeBlock):
        return [b.total, *b.steps]
    if isinstance(b, (RankingBlock, MatrixBlock, DriversBlock)):
        return [b.metric]
    return []


def block_series(b) -> list[str]:
    """Comparison series a block needs from its query (besides 'value')."""
    if isinstance(b, TableBlock):
        return [c for c in b.columns if c in BASE_SERIES]
    if isinstance(b, KpiBlock):
        return [b.compare]
    if isinstance(b, TrendBlock):
        return [s for s in b.series if s != "value"]
    if isinstance(b, (BridgeBlock, DriversBlock)):
        return [b.base]
    if isinstance(b, RankingBlock):
        return [b.compare] if b.compare else []
    return []


class AnalysisSpec(_M):
    version: Literal[2]
    question: str = Field(min_length=1)
    plan: Plan
    queries: list[Query] = Field(min_length=1)
    blocks: list[Block] = Field(min_length=1)

    @model_validator(mode="after")
    def _refs(self) -> "AnalysisSpec":
        ids = [q.id for q in self.queries]
        if len(ids) != len(set(ids)):
            raise ValueError("query ids must be unique")
        by_id = {q.id: q for q in self.queries}
        for i, b in enumerate(self.blocks):
            qid = getattr(b, "query", None)
            if qid is None:
                continue
            q = by_id.get(qid)
            if q is None:
                raise ValueError(f"block {i} ({b.type}): unknown query '{qid}'")
            missing = [m for m in block_metrics(b) if m not in q.metrics]
            if missing:
                raise ValueError(f"block {i} ({b.type}): metrics {missing} are not in query '{qid}' "
                                 "(add them to its metrics)")
            absent = [s for s in block_series(b) if s not in q.compare]
            if absent:
                raise ValueError(f"block {i} ({b.type}): {absent} not in query '{qid}' compare")
        return self


def json_schema() -> dict:
    return AnalysisSpec.model_json_schema()
