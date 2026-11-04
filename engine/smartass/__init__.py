"""smartass engine: read-only Smart View (FinMax / FinLite / FinPlan) for AI agents and Excel."""
import re as _re
from importlib import metadata as _metadata
from pathlib import Path as _Path


def _version() -> str:
    """engine/pyproject.toml is the single source; the installed metadata is read first, the file as fallback."""
    try:
        return _metadata.version("smartass")
    except _metadata.PackageNotFoundError:
        pass
    try:
        text = (_Path(__file__).resolve().parents[1] / "pyproject.toml").read_text(encoding="utf-8")
        m = _re.search(r'(?m)^version\s*=\s*"([^"]+)"', text)
        return m.group(1) if m else "0.0.0"
    except OSError:
        return "0.0.0"


__version__ = _version()
