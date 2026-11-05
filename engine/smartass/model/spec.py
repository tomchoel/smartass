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
    metrics: list[str]                # may be empty only for a query that serves concept blocks alone
    entity: str
    periods: str
    plan_periods: str | None = None
    compare: list[Compare] = []
    preliminary: bool = False
    products: list[str] = []          # empty: the cube's product total (no product dimension in the grid)

    @field_validator("periods", "plan_periods")
    @classmethod
    def _periods_parse(cls, v: str | None) -> str | None:
        if v is None:
            return v
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
    rows: Literal["metrics", "entities", "products"] = "metrics"
    metrics: list[str] = []
    metric: str | None = None
    columns: list[Column] = ["value", "ly", "var", "var_pct"]
    top: int | None = Field(None, ge=1)
    other: bool = True
    sort: Literal["none", "desc", "asc", "abs"] = "none"
    periods: list[str] | None = None   # column groups: one per period expression (e.g. closed, closed-ytd)
    chain: bool = True                 # metric rows with EBITDA: the whole P&L chain, in chain order
    chain_added: list[str] = []        # the chain metrics complete_chain added (quoted after the rest)
    material: bool = True              # with a var column: a last Material marker column (1/0, live formula)

    @field_validator("periods")
    @classmethod
    def _groups_parse(cls, v: list[str] | None) -> list[str] | None:
        if v is None:
            return v
        if not v:
            raise ValueError("table: periods needs at least one period expression (or leave it out)")
        for x in v:
            try:
                resolve(x, (2000, 1))
            except ValueError as e:
                raise ValueError(f"table periods {x!r}: {e}") from e
        return v

    @model_validator(mode="after")
    def _shape(self) -> "TableBlock":
        cols = self.columns
        if "value" not in cols and ("var" in cols or "var_pct" in cols):
            raise ValueError("table: var and var_pct need 'value' and one of ly/budget/forecast in columns")
        for i, c in enumerate(cols):
            if c in ("var", "var_pct") and not any(x in BASE_SERIES for x in cols[:i]):
                raise ValueError("table: var and var_pct need 'value' and one of ly/budget/forecast in columns "
                                 "to their left")
        for c in dict.fromkeys(cols):
            if cols.count(c) > 1 and c not in ("var", "var_pct"):
                raise ValueError(f"table: column {c!r} appears more than once (only var and var_pct may repeat, "
                                 "once per base column)")
        for c in ("var", "var_pct"):
            bases = [next(x for x in reversed(cols[:i]) if x in BASE_SERIES)
                     for i, x in enumerate(cols) if x == c and any(y in BASE_SERIES for y in cols[:i])]
            if len(bases) != len(set(bases)):
                raise ValueError(f"table: {c} appears twice against the same base column")
        if "share" in cols and (self.rows == "metrics" or "value" not in cols):
            raise ValueError("table: share needs rows 'entities' or 'products' and a 'value' column")
        if self.rows != "metrics" and (not self.metric or "value" not in cols):
            raise ValueError(f"table: rows '{self.rows}' needs 'metric' and a 'value' column")
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
    rows: Literal["entities", "products"] = "entities"


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
    top: int = Field(10, ge=1, le=50)

    @model_validator(mode="after")
    def _cell(self) -> "LinesBlock":
        missing = [k for k in ("All Accounts", "All Time Periods") if k not in self.cell]
        if missing:
            raise ValueError(f"lines: cell needs {' and '.join(repr(k) for k in missing)} "
                             "(plus 'All Entities' for the unit)")
        bal = self.cell.get("All Balance Types")
        if bal is not None and bal != "Periodic":
            raise ValueError("lines: drill is Periodic only; leave 'All Balance Types' out or set 'Periodic'")
        return self


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


class ConceptBlock(_M):
    """A product concept (fin_lookup kind 'concept') on its query: the query gives the cube (FinMax or FinLite), the
    entity, the periods and the compare; the concept gives the accounts and products."""
    type: Literal["concept"]
    query: str
    concept: str
    title: str = ""


Block = Annotated[Union[TableBlock, KpiBlock, TrendBlock, BridgeBlock, RankingBlock, MatrixBlock, DriversBlock,
                        LinesBlock, CommentaryBlock, ConceptBlock], Field(discriminator="type")]


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
    queries: list[Query] = []
    blocks: list[Block] = Field(min_length=1)

    @model_validator(mode="after")
    def _refs(self) -> "AnalysisSpec":
        if not self.queries and not any(isinstance(b, LinesBlock) for b in self.blocks):
            raise ValueError("give at least one query, or a lines block (the only block that needs no query)")
        ids = [q.id for q in self.queries]
        if len(ids) != len(set(ids)):
            raise ValueError("query ids must be unique")
        by_id = {q.id: q for q in self.queries}
        for q in self.queries:
            users = [b for b in self.blocks if getattr(b, "query", None) == q.id]
            if not q.metrics and (not users or not all(isinstance(b, ConceptBlock) for b in users)):
                raise ValueError(f"query {q.id} needs metrics (only a query that serves concept blocks alone may "
                                 "have none)")
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
            if isinstance(b, ConceptBlock) and q.cube not in ("FinMax", "FinLite"):
                raise ValueError(f"block {i} (concept): a concept reads FinMax members; give it a FinMax or FinLite "
                                 f"query, not {q.cube}")
            if getattr(b, "rows", None) == "products" and not q.products:
                raise ValueError(f"block {i} ({b.type}): rows 'products' needs products in query '{qid}'")
            absent = [s for s in block_series(b) if s not in q.compare]
            if absent:
                raise ValueError(f"block {i} ({b.type}): {absent} not in query '{qid}' compare")
        return self


def json_schema() -> dict:
    return AnalysisSpec.model_json_schema()
