"""The Excel settings a file job changes, recorded so that what a job that did not finish left changed is put back by
the next job in that same Excel process (final review M3).

A quiet batch turns ScreenUpdating, EnableEvents and DisplayAlerts off and Calculation to manual and puts them back
itself (runner._quiet / _restore, Runner.cs); the look switches AutoSave off in our own file and the save switches it
back on. A server killed in the middle of a batch over COM or between the look and the save, or a put-back Excel
refused, leaves them so, and the next job's quiet batch then reads them as the user's own. So right after its look,
before its first quiet batch, a job writes config.HOME / "jobstate" / "<Excel's pid>.json": that Excel's pid and
process start time (key_of), the four settings as the look found them, what its quiet batches set (each
quiet value that differs from what was found) and the workbook whose AutoSave the look switched off (our own file
only: the user's workbook keeps it off, with a note). The job's end removes it, unless a batch of the job reported a
setting it could not put back (runner.unrestored): then it stays for the next job.

The next job's first contact with that Excel - right before its look, holding that Excel's turn, so no other job
runs there - reads the file of its own Excel: one written for another process with that pid (another start time) is
removed unread and nothing is touched; else each setting that still reads what the dead job set gets back what it
found (one the user changed since is left alone), AutoSave goes back on in that workbook when it is still open under
that full name, the file is removed and the job's notes say what was put back. An Excel whose process cannot be told
gets no file. No figure is ever in the file."""
from __future__ import annotations

import json
import logging
from pathlib import Path

from ..core import config
from .host import HostError
from .inproc import runner
from .inproc.batch import Batch

log = logging.getLogger("smartview.jobstate")

SAME_START_S = 1.0                      # a process start time read again differs by less than this
NOTE = "an earlier smartass job in this Excel did not finish and left {}: put back now"
SAID = {"ScreenUpdating": "ScreenUpdating off", "EnableEvents": "EnableEvents off",
        "DisplayAlerts": "DisplayAlerts off", "Calculation": "Calculation manual"}


def folder() -> Path:
    return config.HOME / "jobstate"


def key_of(t) -> tuple | None:
    """(pid, process start time) of the target's Excel (its pid as the target resolved it; no call to Excel); None
    when it cannot be told (no state is kept for it)."""
    pid = getattr(t, "pid", 0)
    if not isinstance(pid, int) or pid <= 0:
        return None
    try:
        return pid, _start_time(pid)
    except Exception as e:   # noqa: BLE001 - no file for an Excel that cannot be told apart
        log.debug("the start time of Excel pid %s cannot be read: %s", pid, e)
        return None


def _start_time(pid: int) -> float:
    import psutil
    return psutil.Process(pid).create_time()


def record_found(b: Batch, key: str = "settings"):
    """The four quiet settings read in b (each an operation that may fail); the reader gives {name: value} of those
    read."""
    with b.attempt():
        for name, _ in runner.QUIET:
            b.ret(getattr(b.app, name), f"{key}.{name}")

    def found(res) -> dict:
        return {name: res.values[f"{key}.{name}"] for name, _ in runner.QUIET if f"{key}.{name}" in res.values}
    return found


def begin(key: tuple | None, found: dict, autosave: tuple[str, str] | None) -> tuple | None:
    """The job's state written before its first quiet batch: key - its Excel (key_of); found - the settings as its
    look read them; autosave - (name, full name) of our own file whose AutoSave the look switched off. Returns the
    token end() takes (None: nothing written - no Excel told apart, or nothing the job could leave changed)."""
    if key is None:
        return None
    sets = {name: value for name, value in runner.QUIET if name in found and found[name] != value}
    if not sets and autosave is None:
        return None
    state = {"pid": key[0], "started": key[1], "found": found, "set": sets,
             "autosave": None if autosave is None else {"name": autosave[0], "full_name": autosave[1]}}
    path = folder() / f"{key[0]}.json"
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(state), encoding="utf-8")
    except OSError as e:
        log.warning("could not record the job's Excel settings in %s: %s", path, e)
        return None
    return path, runner.unrestored()


def end(token: tuple | None, keep: bool = False) -> None:
    """At the job's end: its state file removed - unless one of its batches could not put a setting back, or the job
    could not switch AutoSave back on in our own file (keep), which the next job then puts right."""
    if token is None:
        return
    path, before = token
    if keep or runner.unrestored() != before:
        log.warning("a setting of Excel could not be put back; the next job in that Excel puts it back (%s)", path)
        return
    _remove(path)


def put_back(app, key: tuple | None) -> str | None:
    """What a job that did not finish left changed in app's Excel (key: key_of) put back; the note that says so, or
    None (no state of this Excel, or nothing still as that job set it). Removes the file - kept when the put-back
    failed or Excel refused a setting of it, for the next job. The states of Excel processes that are gone are pruned
    first. Never raises but for a busy Excel or an abandoned job (HostError): the job then fails at its start as for
    any busy Excel."""
    if key is None:
        return None
    _prune(key)
    path = folder() / f"{key[0]}.json"
    if not path.exists():
        return None
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
        started, found, sets = float(state["started"]), dict(state["found"]), dict(state["set"])
        autosave = state.get("autosave")
    except (OSError, ValueError, KeyError, TypeError) as e:
        log.warning("the Excel settings an earlier job recorded could not be read (%s): left as they are", e)
        _remove(path)
        return None
    if abs(started - key[1]) >= SAME_START_S:          # another Excel process that had this pid: never touched
        _remove(path)
        return None
    try:
        said, whole = _put_back(app, found, sets, autosave)
    except HostError:
        raise
    except Exception as e:   # noqa: BLE001 - the job goes on; its own batches put back what they change
        log.warning("could not put back the Excel settings an earlier job left: %s", e)
        said, whole = [], False
    if whole:
        _remove(path)
    else:
        log.warning("the Excel settings an earlier job left are not all put back; the next job tries again (%s)", path)
    return NOTE.format(", ".join(said)) if said else None


def _prune(key: tuple) -> None:
    """The state files of Excel processes that are gone - no such pid, or one started at another time - removed
    (this Excel's own is read by put_back; a live one of another Excel is left alone)."""
    try:
        paths = [p for p in folder().glob("*.json") if p.stem != str(key[0])]
    except OSError:
        return
    for p in paths:
        try:
            pid, started = int(p.stem), float(json.loads(p.read_text(encoding="utf-8"))["started"])
        except (OSError, ValueError, KeyError, TypeError):
            continue                                    # being written, or not ours to judge
        try:
            gone = abs(_start_time(pid) - started) >= SAME_START_S
        except Exception:   # noqa: BLE001 - no such process
            gone = True
        if gone:
            _remove(p)


def _put_back(app, found: dict, sets: dict, autosave: dict | None) -> tuple[list[str], bool]:
    """One batch reads what the dead job set (and its workbook's AutoSave), one puts back what still reads so: (the
    texts of what was put back, whether all of it was)."""
    from .inproc.runner import run
    b = Batch(quiet=False)
    with b.attempt():
        for name in sets:
            b.ret(getattr(b.app, name), f"now.{name}")
    if autosave is not None:
        with b.attempt():
            wb = b.app.Workbooks(autosave["name"])
            b.ret(wb.FullName, "now.full name")
            b.ret(wb.AutoSaveOn, "now.autosave")
    now = run(app, b, limit=30).values
    b = Batch(quiet=False)
    spans: list[tuple[str, range]] = []
    with b.attempt():
        for name, value in sets.items():
            if f"now.{name}" in now and now[f"now.{name}"] == value and name in found:
                start = b.position
                setattr(b.app, name, found[name])
                spans.append((SAID[name], range(start, b.position)))
        if (autosave is not None and now.get("now.autosave") is False
                and str(now.get("now.full name") or "").casefold() == autosave["full_name"].casefold()):
            start = b.position
            b.app.Workbooks(autosave["name"]).AutoSaveOn = True
            spans.append((f"AutoSave off in {autosave['name']}", range(start, b.position)))
    if not spans:
        return [], True
    skipped = {i for i, _, _ in run(app, b, limit=30).skipped}
    said = [text for text, span in spans if not skipped & set(span)]
    return said, len(said) == len(spans)


def _remove(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError as e:
        log.warning("could not remove %s: %s", path, e)
