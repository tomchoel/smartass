"""Record a run's grids (and drills) to a local JSON file and replay them without EPM: the test set's offline
mode. Files live outside git (%LOCALAPPDATA%\smartass\v2\testset): they hold real financial data."""
from __future__ import annotations

import dataclasses
import json
from pathlib import Path

from ..model.frame import Fact
from ..model.lines import LineSet
from ..model.requests import GridRequest
from .provider import FetchResult


def _key(req: GridRequest) -> str:
    """The request as JSON; no products leaves the field out, so recordings made before products keep their keys."""
    d = dataclasses.asdict(req)
    if not d["products"]:
        del d["products"]
    return json.dumps(d, sort_keys=True)


class RecordingProvider:
    def __init__(self, primary):
        self.primary = primary
        self.grids: dict[str, list[dict]] = {}
        self.closed = None

    def closed_months(self):
        self.closed = self.primary.closed_months()
        return self.closed

    def fetch(self, req):
        res = self.primary.fetch(req)
        self.grids[_key(req)] = [dataclasses.asdict(f) for f in res.facts]
        return res

    def save(self, path: Path) -> None:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        closed = {k: list(v) for k, v in (self.closed or {}).items()}
        path.write_text(json.dumps({"closed": closed, "grids": self.grids}, ensure_ascii=False), encoding="utf-8")


class ReplayProvider:
    def __init__(self, path: Path):
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        self.closed = {k: tuple(v) for k, v in data["closed"].items()}
        self.grids = data["grids"]

    def closed_months(self):
        if not self.closed:
            raise KeyError("closed months not recorded - record the case again")
        return self.closed

    def fetch(self, req):
        k = _key(req)
        if k not in self.grids:
            raise KeyError(f"grid not recorded: {k} - record the case again")
        return FetchResult(req, [Fact(**f) for f in self.grids[k]], route="replay")   # older facts: no product ('')


def save_lines(path: Path, sets: dict) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({k: dataclasses.asdict(v) for k, v in sets.items()}, ensure_ascii=False, default=str),
                    encoding="utf-8")


def load_lines(path: Path) -> dict:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    return {k: LineSet(**v) for k, v in data.items()}
