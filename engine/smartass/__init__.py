"""smartass engine: read-only Smart View (FinMax / FinLite / FinPlan) for AI agents and Excel."""
from datetime import datetime as _dt
import os as _os
from pathlib import Path as _Path
import struct as _struct


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


def _git_dir(repo: _Path) -> _Path | None:
    """Resolve the repository's .git directory (handling worktree gitdir files)."""
    git_dir = repo / ".git"
    if git_dir.is_file():
        line = git_dir.read_text(encoding="utf-8", errors="ignore").strip()
        if line.startswith("gitdir:"):
            p = _Path(line.split(":", 1)[1].strip())
            git_dir = p if p.is_absolute() else (repo / p)
    return git_dir if git_dir.exists() else None


def _git_head_sha(repo: _Path) -> str:
    """Read the short git commit hash directly from .git/HEAD without spawning subprocesses."""
    git_dir = _git_dir(repo)
    head_file = (git_dir / "HEAD") if git_dir else None
    if not head_file or not head_file.exists():
        return "unknown"
    head = head_file.read_text(encoding="utf-8", errors="ignore").strip()
    if head.startswith("ref:"):
        ref = head.split(":", 1)[1].strip()
        ref_file = git_dir / ref
        if ref_file.exists():
            return ref_file.read_text(encoding="utf-8", errors="ignore").strip()[:7]
        packed = git_dir / "packed-refs"
        if packed.exists():
            for row in packed.read_text(encoding="utf-8", errors="ignore").splitlines():
                if row and not row.startswith("#") and not row.startswith("^"):
                    parts = row.split(" ", 1)
                    if len(parts) == 2 and parts[1].strip() == ref:
                        return parts[0][:7]
        return ref.rsplit("/", 1)[-1]
    return head[:7] if head else "unknown"


def _git_dirty(repo: _Path) -> bool | None:
    """Detect uncommitted tracked edits or untracked engine .py files via .git/index without subprocesses."""
    if _os.environ.get("PYTEST_CURRENT_TEST"):
        return None
    git_dir = _git_dir(repo)
    index_path = (git_dir / "index") if git_dir else None
    if not index_path or not index_path.is_file():
        return None
    try:
        data = index_path.read_bytes()
        if len(data) < 12:
            return None
        sig, ver, count = _struct.unpack(">4sII", data[:12])
        if sig != b"DIRC" or ver not in (2, 3):
            return None
        pos = 12
        tracked: set[str] = set()
        for _ in range(count):
            start = pos
            mtime_s, mtime_ns = _struct.unpack(">II", data[pos + 8:pos + 16])
            size = _struct.unpack(">I", data[pos + 36:pos + 40])[0]
            flags = _struct.unpack(">H", data[pos + 60:pos + 62])[0]
            pos += 62
            if ver == 3 and (flags & 0x4000):
                pos += 2
            end = data.index(b"\x00", pos)
            rel = data[pos:end].decode("utf-8", "surrogateescape")
            pos = start + ((end - start + 8) & ~7)
            tracked.add(rel)
            try:
                st = _os.stat(repo / rel)
            except OSError:
                return True
            if st.st_size != size or st.st_mtime_ns != (mtime_s * 1_000_000_000 + mtime_ns):
                return True
        src_root = repo / "engine" / "smartass"
        if src_root.is_dir():
            for p in src_root.rglob("*.py"):
                if "__pycache__" not in p.parts and p.relative_to(repo).as_posix() not in tracked:
                    return True
        return False
    except Exception:   # noqa: BLE001
        return None


_REPO_ROOT = _Path(__file__).resolve().parents[2]
__version__ = _version()
__commit__ = _git_head_sha(_REPO_ROOT)
__dirty__ = _git_dirty(_REPO_ROOT)
__started_at__ = _dt.now().isoformat(timespec="seconds")
__build__ = f"{__version__}+{__commit__}{'-dirty' if __dirty__ else ''}"
