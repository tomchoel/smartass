"""The save after the reply: our own analysis file is saved by Excel itself right after the job has answered, so the
user does not wait for it (0.4-1.1 s of every answer before; live, 20.11.2026).

How: the job's save step asks the add-in (Runner.cs SmartassSaveLater_<hash>: one Application.Run, a few ms) to save
the workbook. The add-in waits config.SAVE_DELAY_S on a timer of its own, so Excel paints the answer first (a save
blocks painting: one that starts at once leaves the old or a half-painted answer on screen while it runs), then
queues the save as a macro (ExcelAsyncUtil.QueueAsMacro), which runs when Excel is ready: Workbook.Save, then AutoSave
back on when the look switched it off (also after a failed save, as a failed job's clean-up does), then the outcome
is written. A batch's Run (or another SaveLater) that comes before the save ran runs it first, before anything else:
the next job never finds a save waiting between its batches, and no save runs twice or is lost. A job that comes
while the save runs waits for Excel as for any busy Excel (its call is rejected and asked again: jobs.com). Without
the add-in in that Excel (the COM executor), with SV_SAVE=sync, or when the add-in refuses or cannot be asked, the
caller saves in the job as before (pipeline._save). Only our own file is ever asked for: never the user's workbook.

The outcome is config.INPROC_DIR / "saved" / "<Excel's pid>.json", written by the add-in, with no figure in it: the
token the request carried, the workbook's name, ok, Excel's error text, whether the workbook was closed first,
whether AutoSave was to go back on and its own error, how the save ran ("after the reply", "before a batch", "before
the next"), when it was queued and started (seconds since 1970) and its seconds. The next job in that Excel reads it
right after its look (take: moved to "last.json", so it is read once, whichever server process reads it); a failed
one is a note of that job's reply and that job saves in the foreground. One the add-in could not write comes back with
the next clear and is read the same way; one that cannot be read is an unknown outcome (a note, the job saves in the
foreground). Outcomes of an Excel that is gone, or a week old, are removed when the next save is asked for. The scripts wait for the saves they caused
(wait); fin_status shows the newest outcome (status)."""
from __future__ import annotations

import functools
import json
import logging
import os
import threading
import time
import uuid
from datetime import datetime
from pathlib import Path

from ...core import config, timing
from ..host import HostError
from . import build, runner

log = logging.getLogger("smartview.inproc")
LAST = "last.json"
QUEUE_LIMIT_S = 120                     # a busy Excel is waited for this long when the save is asked for (as a batch)
KEEP = 50                               # outcomes a job read that the scripts have not waited for, at most
POLL_S = 0.02

_lock = threading.Lock()
_queued: dict[str, tuple[int, float]] = {}      # token -> the Excel's key: asked for here, its outcome not read yet
_asked_at: dict[str, float] = {}                # token -> when it was asked for (status: how many are to come)
FRESH_S = 60                                    # a save asked for longer ago with no outcome is not "to come" (its
                                                # Excel quit first, say)
_done: dict[str, dict] = {}                     # token -> the outcome a job read, for wait()
_unwritten: dict[int, dict] = {}                # pid -> an outcome the add-in could not write (told by clear)
UNWRITTEN = "unwritten: "                       # clear's answer then, before the outcome's record
CLEAR_LIMIT_S = 120                             # a save that runs is waited for this long (as a batch)
KEEP_DAYS = 7                                   # an outcome older than this is removed (with those of gone Excels)
UNKNOWN = "its outcome file could not be read"


def folder() -> Path:
    return config.INPROC_DIR / "saved"


def available(app) -> bool:
    """The save can be asked for after the reply in app's Excel: SV_SAVE is not "sync" and the add-in is loaded there
    (loading it the first time this Excel process is asked)."""
    return config.SAVE == "later" and runner.available(app)


def clear(app, limit: float = CLEAR_LIMIT_S) -> None:
    """A job's first contact with app's Excel (jobs.hold_turn calls it once it holds the turn there), before any plain
    COM call of the job writes in it: one add-in call (Runner.cs SmartassClear) runs a save after the reply still
    waiting there, and while one runs the add-in answers busy and the call is asked again (runner.asked: bounded by
    limit, abortable: HostError). An add-in this server process has not loaded in that Excel is called by this
    build's name (another server may have loaded it); no such macro there: no add-in, so no save of ours can be
    waiting. An outcome the add-in could not write comes back here and is kept for the next look (take)."""
    if not config.INPROC:
        return
    try:
        key = runner._identify(app)
    except HostError:
        raise
    except Exception:   # noqa: BLE001 - which Excel this is cannot be told: no add-in of ours found there
        return
    ex = runner._EXCELS.get(key)
    if ex is not None and ex.ready is False:
        return                                              # the add-in does not load in that Excel
    macro = ex.clear_macro if ex is not None and ex.ready else f"SmartassClear_{_build_hash()}"
    try:
        got = runner.asked(lambda: app.Run(macro), limit)
    except HostError:
        raise
    except Exception:   # noqa: BLE001 - the macro is not there: no add-in in that Excel
        return
    if isinstance(got, str) and got.startswith(UNWRITTEN):
        try:
            out = json.loads(got[len(UNWRITTEN):])
        except ValueError:
            out = None
        out = out if _valid(out) else _unknown()
        with _lock:
            _unwritten[key[0]] = out
        log.warning("the outcome of the save after the reply in Excel pid %s could not be written (%s); the next "
                    "answer there reads it from memory", key[0], out.get("unwritten"))


@functools.lru_cache(maxsize=1)
def _build_hash() -> str:
    return build.hash8()


def queue(app, workbook: str, autosave: bool) -> str | None:
    """Ask app's Excel to save workbook (our own file) after the reply, then switch its AutoSave back on when
    autosave; the request's token, or None when it was not asked for - no add-in there, SV_SAVE=sync, the add-in
    refused it or its answer cannot be read (then a save it may have queued runs at the start of the caller's own
    save batch, before that batch saves the file again): the caller saves in the foreground. One Application.Run,
    counted as one batch of the current stage. A busy Excel beyond QUEUE_LIMIT_S or an abandoned job raises
    HostError, as a batch does."""
    if config.SAVE != "later":
        return None
    ex = runner._ready(app)
    if ex is None:
        return None
    key = runner._identify(app)
    _prune()
    token = uuid.uuid4().hex
    timing.count_batch()
    try:
        got = runner.asked(lambda: app.Run(ex.later_macro, workbook, bool(autosave), token,
                                           float(config.SAVE_DELAY_S * 1000)), QUEUE_LIMIT_S)
    except HostError:
        raise
    except Exception as e:   # noqa: BLE001 - not asked for (or not known to be): saved in the foreground
        log.warning("the save of %s after the reply could not be asked for (%s); it is saved now", workbook,
                    runner._message(e))
        return None
    if got != "queued":
        log.warning("the add-in did not take the save of %s after the reply (%s); it is saved now", workbook,
                    str(got)[:200])
        return None
    with _lock:
        _queued[token] = key
        _asked_at[token] = time.time()
        for t in [t for t, at in _asked_at.items() if time.time() - at > FRESH_S and t not in _queued]:
            _asked_at.pop(t)
    return token


def take(app) -> dict | None:
    """The outcome of the save last asked for in app's Excel (by any server process), once: moved to LAST, so no
    other job reads it again; None when there is none to read (none asked for, already read, an earlier Excel's
    with the same pid, or no add-in in use). Read right after a job's look: that Run ran any save still waiting."""
    if not config.INPROC:
        return None
    try:
        pid, started = runner._identify(app)
    except Exception:   # noqa: BLE001 - which Excel this is cannot be told: no outcome of its
        return None
    with _lock:
        kept = _unwritten.pop(pid, None)
    if kept is not None:
        _seen(kept)
        return kept
    path = folder() / f"{pid}.json"
    out = _read(path)
    if out is None:
        if not path.exists():
            return None
        time.sleep(POLL_S)                                # being replaced: read once more
        out = _read(path)
        if out is None and path.exists():
            try:
                path.unlink()                             # unreadable: read once, as an unknown outcome
            except OSError:
                pass
            return _unknown()
        if out is None:
            return None
    try:
        os.replace(path, folder() / LAST)
    except FileNotFoundError:
        return None                                       # another job read it meanwhile
    except OSError:
        try:                                              # LAST is held open (fin_status reading it): read anyway
            path.unlink()
        except OSError:
            pass
    if out["started"] < started - 1:                      # written by an Excel that had this pid before
        return None
    _seen(out)
    return out


def _seen(out: dict) -> None:
    """A job read the outcome of a save asked for here: kept for wait()."""
    with _lock:
        if _queued.pop(out.get("token"), None) is not None:
            _done[out["token"]] = out
            while len(_done) > KEEP:
                _done.pop(next(iter(_done)))


def _unknown() -> dict:
    """An outcome that cannot be read: which save it was is not known."""
    return {"token": "", "workbook": "", "full_name": "", "ok": False, "unknown": True, "error": UNKNOWN,
            "closed": False, "seconds": 0.0, "started": 0.0}


def _alive(pid: int) -> bool:
    import psutil
    return psutil.pid_exists(pid)


def _prune() -> None:
    """Outcome files of an Excel that is gone, or older than KEEP_DAYS, removed (LAST is kept: fin_status's); a
    temporary file an hour old too. Never raises."""
    now = time.time()
    for p in _files(all_=True):
        try:
            age = now - p.stat().st_mtime
            if p.name == LAST:
                continue
            if p.suffix == ".tmp":
                if age > 3600:
                    p.unlink()
                continue
            stem = p.name.split(".")[0]
            if age > KEEP_DAYS * 86400 or (stem.isdigit() and not _alive(int(stem))):
                p.unlink()
        except OSError:
            pass


def give_back(app, out: dict) -> None:
    """A failed outcome a job took (take) but could not act on - it failed before its own save - put back in place,
    so the next job in that Excel says it and saves; not when a newer outcome is there. Never raises."""
    try:
        pid, _ = runner._identify(app)
        path = folder() / f"{pid}.json"
        if path.exists():
            return
        tmp = path.with_name(f"{path.name}.back.tmp")
        tmp.write_text(json.dumps(out), encoding="utf-8")
        os.replace(tmp, path)
    except Exception as e:   # noqa: BLE001 - the failing job's error says it already
        log.warning("a failed save after the reply could not be put back for the next job: %s", e)


def wait(limit: float = 60.0) -> list[dict]:
    """For the scripts: the outcome of every save asked for in this process that has not been waited for, in the
    order asked for, each once; waited for up to limit seconds in all (the outcome files are read, not taken: the
    next job still reads the newest). One that does not come in time is a failure "no outcome within ..."."""
    with _lock:
        tokens = list(_done) + [t for t in _queued if t not in _done]
    found: dict[str, dict] = {}
    end = time.monotonic() + limit
    while True:
        with _lock:
            found.update({t: _done[t] for t in tokens if t in _done})
            pending = {t: _queued[t] for t in tokens if t not in found and t in _queued}
        for token, (pid, _) in pending.items():
            for path in (folder() / f"{pid}.json", folder() / LAST):
                out = _read(path)
                if out is not None and out["token"] == token:
                    found[token] = out
                    break
        if len(found) == len(tokens) or time.monotonic() >= end:
            break
        time.sleep(POLL_S)
    with _lock:
        for t in tokens:
            _queued.pop(t, None)
            _done.pop(t, None)
    return [found.get(t) or {"token": t, "ok": False, "error": f"no outcome within {limit:g} s", "seconds": None}
            for t in tokens]


def status() -> dict:
    """For fin_status: the newest outcome in the folder (workbook, ok, error, closed, seconds, at, how) and how many
    saves this server asked for in the last FRESH_S seconds whose outcome is not written yet."""
    outs = [o for o in (_read(p) for p in _files()) if o is not None]
    newest = max(outs, key=lambda o: o["started"], default=None)
    with _lock:
        asked = {t: at for t, at in _asked_at.items() if t in _queued and time.time() - at <= FRESH_S}
    written = {o["token"] for o in outs}
    last = None
    if newest is not None:
        last = {k: newest.get(k) for k in ("workbook", "ok", "error", "closed", "seconds", "how")}
        last["at"] = datetime.fromtimestamp(newest["started"]).astimezone().isoformat(timespec="seconds")
    return {"last": last, "waiting": len([t for t in asked if t not in written])}


def _files(all_: bool = False) -> list[Path]:
    try:
        return [p for p in folder().iterdir() if all_ or p.suffix == ".json"]
    except OSError:
        return []


def _read(path: Path) -> dict | None:
    """An outcome file, checked; None when it is not there or cannot be read (a file being replaced)."""
    try:
        out = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return out if _valid(out) else None


def _valid(out) -> bool:
    return (isinstance(out, dict) and isinstance(out.get("token"), str) and isinstance(out.get("ok"), bool)
            and isinstance(out.get("started"), (int, float)) and isinstance(out.get("workbook"), str))
