"""smartass engine: read-only Smart View (FinMax / FinLite / FinPlan) for AI agents and Excel."""
from pathlib import Path as _Path


def _version() -> str:
    """engine/VERSION is the single source (scripts/bump.ps1 keeps pyproject.toml equal to it; the relay
    run_server.py reloads the server when it changes). Installed metadata is the fallback for a wheel install."""
    try:
        return (_Path(__file__).resolve().parents[1] / "VERSION").read_text(encoding="utf-8").strip()
    except OSError:
        pass
    try:
        from importlib import metadata
        return metadata.version("smartass")
    except Exception:   # noqa: BLE001
        return "0.0.0"


__version__ = _version()
