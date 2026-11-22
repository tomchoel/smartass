r"""Record a run's grids (and drills) to a local JSON file and replay them without EPM: the test set's offline
mode. Files live outside git (%LOCALAPPDATA%\smartass\v2\testset): they hold real financial data. A run where a
member was left out over Planning REST (FetchResult.dropped) is never stored: its totals are incomplete, and a
recording keeps facts only - a golden made from it would hold the partial total without the warning."""
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


class LeftOut(ValueError):
    """The run left a member out over Planning REST: it is not stored (record the case again once it is fixed)."""


class RecordingProvider:
    def __init__(self, primary):
        self.primary = primary
        self.grids: dict[str, list[dict]] = {}
        self.closed = None
        self.dropped: list[tuple[str, str, str]] = []    # (cube, dimension, member) left out over REST in this run

    def closed_months(self):
        self.closed = self.primary.closed_months()
        return self.closed

    def fetch(self, req):
        res = self.primary.fetch(req)
        self.grids[_key(req)] = [dataclasses.asdict(f) for f in res.facts]
        self.dropped += [d for d in getattr(res, "dropped", ()) if d not in self.dropped]
        return res

    def save(self, path: Path) -> None:
        """The recording to path; LeftOut (nothing written) when the run left a member out."""
        if self.dropped:
            raise LeftOut(f"{len(self.dropped)} member(s) left out over Planning REST in this run ("
                          + ", ".join(" ".join(d) for d in self.dropped) + "): their totals are incomplete, so the run "
                          "is not stored - check the member or rebuild the atlas (`smartass atlas build`), then record "
                          "again")
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
