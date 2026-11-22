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
setting it could not put back (runner.unrestored): then it stays whole for the next job; when only AutoSave of our
file could not be switched back on, only that is kept - the settings the job restored itself are not (final review
M2: the next job would "put back" what the user chose since).

The next job's first contact with that Excel - right before its look, holding that Excel's turn, so no other job
runs there - reads the file of its own Excel: one written for another process with that pid (another start time) is
removed unread and nothing is touched; else each setting that still reads what the dead job set gets back what it
found (one the user changed since is left alone), AutoSave goes back on in that workbook when it is still open under
that full name, the file is removed and the job's notes say what was put back. What Excel refused to put back stays
in the file, and only that (one the user changed since is dropped); the same job's own record then carries it, with
what the DEAD job found, and its end leaves it in the file, so it is put back by the first job whose put-back goes
through (final review M1). An Excel whose process cannot be told gets no file; a state is pruned only when its Excel
is known to be gone (M12). No figure is ever in the file."""
from __future__ import annotations

import json
import logging
from pathlib import Path

from ..core import config
from ..core.errors import HostError
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
    look read them; autosave - (name, full name) of our own file whose AutoSave the look switched off. What this
    Excel's file still holds (a put-back that failed right before: put_back) is carried into it, with what the dead
    job found and its workbook when this job has none of its own (one AutoSave per file: this job's own goes first).
    Returns the token end() takes (None: nothing written - no Excel told apart, or nothing the job could leave changed;
    a carried state alone is left in its file as it is)."""
    if key is None:
        return None
    path = folder() / f"{key[0]}.json"
    carried = _carried(path, key)
    sets = {name: value for name, value in runner.QUIET if name in found and found[name] != value}
    if not sets and autosave is None:
        return None
    own = None if autosave is None else {"name": autosave[0], "full_name": autosave[1]}
    state = {"pid": key[0], "started": key[1], "found": {**found, **carried["found"]},
             "set": {**sets, **carried["set"]}, "autosave": own or carried["autosave"]}
    if not _write(path, state):
        return None
    return path, runner.unrestored(), state, carried


def end(token: tuple | None, keep: bool = False) -> None:
    """At the job's end: its state file removed - kept whole when one of its batches could not put a setting back;
    when the job could not switch AutoSave back on in our own file (keep) only that AutoSave is kept; and what an
    earlier put-back could not put back (carried) stays in either case - the next job puts it right."""
    if token is None:
        return
    path, before, state, carried = token
    if runner.unrestored() != before:
        log.warning("a setting of Excel could not be put back; the next job in that Excel puts it back (%s)", path)
        return
    left = dict(state, found=carried["found"], set=carried["set"],
                autosave=state["autosave"] if keep else carried["autosave"])
    if left["set"] or left["autosave"] is not None:
        log.warning("a setting of Excel is still left changed; the next job in that Excel puts it back (%s)", path)
        _write(path, left)
        return
    _remove(path)


def records_autosave(token: tuple | None, full_name: str) -> bool:
    """The state begin() wrote (its token) records the workbook of that full name for AutoSave: a job that cannot
    switch it back on leaves it to the next job."""
    if token is None or not full_name:
        return False
    autosave = token[2]["autosave"]
    return autosave is not None and str(autosave["full_name"]).casefold() == full_name.casefold()


def _carried(path: Path, key: tuple) -> dict:
    """What this Excel's state file holds when a job begins - left there by a put-back that failed ({"found", "set",
    "autosave"}, empty when there is none, or it is another process's, or it cannot be read)."""
    none = {"found": {}, "set": {}, "autosave": None}
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
        if abs(float(state["started"]) - key[1]) >= SAME_START_S:
            return none
        sets = dict(state["set"])
        return {"found": {k: v for k, v in dict(state["found"]).items() if k in sets}, "set": sets,
                "autosave": state.get("autosave")}
    except (OSError, ValueError, KeyError, TypeError):
        return none


def _write(path: Path, state: dict) -> bool:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(state), encoding="utf-8")
    except OSError as e:
        log.warning("could not record the job's Excel settings in %s: %s", path, e)
        return False
    return True


def put_back(app, key: tuple | None) -> str | None:
    """What a job that did not finish left changed in app's Excel (key: key_of) put back; the note that says so, or
    None (no state of this Excel, or nothing still as that job set it). Removes the file - when the put-back failed it
    is kept, when Excel refused a part of it only that part is, for the next job. The states of Excel processes that
    are gone are pruned first. Never raises but for a busy Excel or an abandoned job (HostError): the job then fails at
    its start as for any busy Excel."""
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
        said, left, left_autosave = _put_back(app, found, sets, autosave)
    except HostError:
        raise
    except Exception as e:   # noqa: BLE001 - the job goes on; its own batches put back what they change
        log.warning("could not put back the Excel settings an earlier job left: %s", e)
        said, left, left_autosave = [], sets, autosave
    if not left and left_autosave is None:
        _remove(path)
    else:
        log.warning("the Excel settings an earlier job left are not all put back; the next job tries again (%s)", path)
        _write(path, dict(state, found={k: v for k, v in found.items() if k in left}, set=left,
                          autosave=left_autosave))
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
        except Exception as e:   # noqa: BLE001 - gone only when there is no such process (not: access denied)
            gone = _no_such_process(e)
        if gone:
            _remove(p)


def _no_such_process(e: Exception) -> bool:
    import psutil
    return isinstance(e, psutil.NoSuchProcess)


def _put_back(app, found: dict, sets: dict, autosave: dict | None) -> tuple[list[str], dict, dict | None]:
    """One batch reads what the dead job set (and its workbook's AutoSave), one puts back what still reads so: (the
    texts of what was put back, the settings and the AutoSave Excel refused to put back - still to do)."""
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
    spans: list[tuple[str, str | None, range]] = []        # (text, the setting - None: the AutoSave, its operations)
    with b.attempt():
        for name, value in sets.items():
            if f"now.{name}" in now and now[f"now.{name}"] == value and name in found:
                start = b.position
                setattr(b.app, name, found[name])
                spans.append((SAID[name], name, range(start, b.position)))
        if (autosave is not None and now.get("now.autosave") is False
                and str(now.get("now.full name") or "").casefold() == autosave["full_name"].casefold()):
            start = b.position
            b.app.Workbooks(autosave["name"]).AutoSaveOn = True
            spans.append((f"AutoSave off in {autosave['name']}", None, range(start, b.position)))
    if not spans:
        return [], {}, None
    skipped = {i for i, _, _ in run(app, b, limit=30).skipped}
    refused = [name for _, name, span in spans if skipped & set(span)]
    said = [text for text, _, span in spans if not skipped & set(span)]
    return said, {n: sets[n] for n in refused if n is not None}, autosave if None in refused else None


def _remove(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError as e:
        log.warning("could not remove %s: %s", path, e)
