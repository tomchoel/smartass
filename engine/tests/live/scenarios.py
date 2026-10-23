"""Live scenarios S1-S10. Each is fn(ctx) -> dict(notes=..., budget=...) and raises AssertionError/ToolFailed on failure."""
from __future__ import annotations

from dataclasses import dataclass, field

REGISTRY: dict[str, "Scenario"] = {}


@dataclass
class Scenario:
    id: str
    fn: object
    tags: list[str]
    budget: float            # seconds; over it -> SLOW
    limit: str | None = None # set when the failure is an accepted, documented limit
    fresh: bool = False      # only runs with --fresh


@dataclass
class Ctx:
    client: object           # McpClient
    scratch: object          # Path to <analyser>\_matrix
    keep: bool
    state: dict = field(default_factory=dict)   # shared between scenarios (e.g. S2's workbook for S7)


def scenario(id: str, tags: list[str], budget: float = 60, limit: str | None = None, fresh: bool = False):
    def deco(fn):
        REGISTRY[id] = Scenario(id, fn, tags, budget, limit, fresh)
        return fn
    return deco


@scenario("S0.status", ["smoke"], budget=10)
def s0(ctx):
    b, _ = ctx.client.must("sv_status", {})
    assert b
