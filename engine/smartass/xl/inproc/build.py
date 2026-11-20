"""The add-in, built on first use: Runner.cs compiled with the .NET Framework compiler that ships with Windows, next to
a renamed copy of Excel-DNA's 64-bit loader (vendor/, zlib licence: vendor/LICENSE-ExcelDna.txt).

As the Smart View bridge (xl/instances.py bridge_path): a loaded add-in's file stays locked and its function names
stay taken in that Excel, so every build has its own. hash8 is the sha1 of everything that defines a build: Runner.cs
(line endings normalised), the vendored files, the .dna and .xll.config texts, the compiler flags, the jobs folder
(compiled into the add-in: it runs job files from that one folder only) and the batch version. The build is config.INPROC_DIR / smartass_run_<hash8> holding smartass_run_<hash8>.xll (the loader),
.dna (which dll to load, from bytes: the dll is not locked), .xll.config (Excel-DNA's diagnostics off: no log window
in the user's Excel on a load problem) and .dll (Runner.cs with the hash in its function names). It is built in a
temporary folder and renamed into place, so two jobs can build at once. A compile error is kept in
smartass_run_<hash8>.failed and not retried for that hash for a day. Older builds are removed unless their add-in is
loaded somewhere (its file locked)."""
from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

from ...core import config
from . import batch

SOURCE = Path(__file__).resolve().parent
RUNNER_CS = SOURCE / "Runner.cs"
XLL = SOURCE / "vendor" / "ExcelDna64.xll"
INTEGRATION = SOURCE / "vendor" / "ExcelDna.Integration.dll"     # only to compile against: the loader carries its own
CSC = Path(os.environ.get("SystemRoot") or r"C:\Windows") / "Microsoft.NET" / "Framework64" / "v4.0.30319" / "csc.exe"
PREFIX = "smartass_run_"
COMPILE_LIMIT_S = 120
STALE_S = 24 * 3600                     # leftover temporary build folders older than this are removed
MARKER_S = 24 * 3600                    # a compile error is not retried for this long
FLAGS = ("-nologo", "-target:library", "-langversion:5", "-optimize+")
FRAMEWORK_REFS = ("System.Web.Extensions.dll",)      # named by their path in the compiler's own Framework folder

DNA = """<?xml version="1.0" encoding="utf-8"?>
<DnaLibrary Name="{name}" RuntimeVersion="v4.0">
  <ExternalLibrary Path="{name}.dll" ExplicitExports="false" LoadFromBytes="true" Pack="false" />
</DnaLibrary>
"""

# Excel-DNA's diagnostic logging off: the source switched Off, and its LogDisplay listener (the window it opens on a
# load problem) configured with an Off filter, which is how Excel-DNA is told to remove it.
CONFIG = """<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <system.diagnostics>
    <sources>
      <source name="ExcelDna.Integration" switchValue="Off">
        <listeners>
          <remove name="Default" />
          <add name="LogDisplay" type="ExcelDna.Logging.LogDisplayTraceListener,ExcelDna.Integration">
            <filter type="System.Diagnostics.EventTypeFilter" initializeData="Off" />
          </add>
        </listeners>
      </source>
    </sources>
  </system.diagnostics>
</configuration>
"""

PARTS = (".xll", ".dna", ".xll.config", ".dll")


class BuildError(Exception):
    """The add-in cannot be built or loaded: the COM executor does the work."""


def jobs_folder() -> Path:
    """The one folder the add-in runs job files from (absolute, as the add-in compares it)."""
    return Path(os.path.abspath(config.INPROC_DIR / "jobs"))


def _template() -> str:
    return RUNNER_CS.read_bytes().decode("utf-8").replace("\r\n", "\n")


def hash8() -> str:
    h = hashlib.sha1()
    flags = "\0".join((*FLAGS, *FRAMEWORK_REFS))
    for part in (_template().encode(), XLL.read_bytes(), INTEGRATION.read_bytes(), DNA.encode(), CONFIG.encode(),
                 flags.encode(), str(jobs_folder()).encode(), f"batch {batch.VERSION}".encode()):
        h.update(len(part).to_bytes(8, "little"))
        h.update(part)
    return h.hexdigest()[:8]


def ensure() -> tuple[Path, str]:
    """(the add-in's .xll, hash8): built now unless it is there; older builds pruned. BuildError when it cannot be."""
    h = hash8()
    dest = config.INPROC_DIR / f"{PREFIX}{h}"
    if not _complete(dest):
        _build(h, dest)
    prune(h)
    return dest / f"{dest.name}.xll", h


def _complete(folder: Path) -> bool:
    return all((folder / f"{folder.name}{ext}").is_file() for ext in PARTS)


def _build(h: str, dest: Path) -> None:
    name = dest.name
    marker = config.INPROC_DIR / f"{name}.failed"
    if marker.exists() and time.time() - marker.stat().st_mtime > MARKER_S:
        marker.unlink(missing_ok=True)                    # a day old: the compiler is tried again
    if marker.exists():
        raise BuildError(f"the add-in did not compile before (kept in {marker}): "
                         f"{marker.read_text(encoding='utf-8', errors='replace').strip()[:500]}")
    if not CSC.is_file():
        raise BuildError(f"no C# compiler at {CSC}")
    config.INPROC_DIR.mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp(prefix=f".build_{h}_", dir=config.INPROC_DIR))
    try:
        cs = tmp / "Runner.cs"
        jobs = str(jobs_folder()).replace('"', '""')       # inside a C# verbatim string
        cs.write_text(_template().replace("__HASH__", h).replace("__JOBS__", jobs), encoding="utf-8")
        shutil.copyfile(XLL, tmp / f"{name}.xll")
        (tmp / f"{name}.dna").write_text(DNA.format(name=name), encoding="utf-8")
        (tmp / f"{name}.xll.config").write_text(CONFIG, encoding="utf-8")
        dll = tmp / f"{name}.dll"
        code, text = _compile(_command(cs, dll), tmp)
        if code != 0 or not dll.is_file():
            marker.write_text(text or f"csc exit code {code}", encoding="utf-8")
            raise BuildError(f"the add-in did not compile: {text.strip()[:500]}")
        if dest.exists() and not _complete(dest):
            _remove(dest)                                # a broken leftover (not loaded anywhere: else it stays)
        try:
            os.rename(tmp, dest)
        except OSError:
            if not _complete(dest):                     # else another job put the same build in place meanwhile
                raise BuildError(f"the build could not be put in place at {dest}") from None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _command(cs: Path, dll: Path) -> list[str]:
    """csc's command line: every reference by its full path (the Framework's from the compiler's own folder, Excel-DNA's
    from vendor/), so none is picked up from the working directory or LIB."""
    refs = [CSC.parent / r for r in FRAMEWORK_REFS] + [INTEGRATION]
    return [str(CSC), *FLAGS, f"-out:{dll}", *(f"-r:{r}" for r in refs), str(cs)]


def _compile(cmd: list[str], cwd: Path) -> tuple[int, str]:
    """csc's exit code and output (no window), run in cwd (the temporary build folder). BuildError when it did not
    run."""
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, errors="replace", timeout=COMPILE_LIMIT_S,
                           creationflags=subprocess.CREATE_NO_WINDOW, cwd=cwd)
    except (OSError, subprocess.TimeoutExpired) as e:
        raise BuildError(f"the C# compiler did not run: {e}") from None
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def _remove(folder: Path) -> bool:
    """Remove a build unless its add-in is loaded somewhere (the .xll is the one file Excel keeps locked)."""
    try:
        (folder / f"{folder.name}.xll").unlink(missing_ok=True)
    except OSError:
        return False
    shutil.rmtree(folder, ignore_errors=True)
    return True


def prune(keep: str) -> None:
    """Remove other builds and their markers, and temporary build folders a day old (job files: runner.py)."""
    root = config.INPROC_DIR
    if not root.is_dir():
        return
    now = time.time()
    for p in root.iterdir():
        try:
            if p.name.startswith(PREFIX + keep):
                continue
            if p.name.startswith(PREFIX):
                if p.is_dir():
                    _remove(p)
                else:
                    p.unlink()
            elif p.name.startswith(".build_") and now - p.stat().st_mtime > STALE_S:
                shutil.rmtree(p, ignore_errors=True)
        except OSError:
            pass                                         # in use: next time
