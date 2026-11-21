"""One user-Excel job: COM calls with retry/abort, the machine-wide per-Excel turn, fast mode."""
from __future__ import annotations

import contextlib
import logging
import threading
import time

import pywintypes

from ..core import config, timing
from .host import HostError
from .hostproc import is_busy, retry_busy

EDIT_WAIT_S = 30
START_WAIT_S = EDIT_WAIT_S          # a job's start waits this long in all for an Excel that rejects calls (starting)
BUSY = "Excel is busy - probably a cell is being edited or a dialog is open. Press Enter or Esc in Excel and ask again."

log = logging.getLogger("smartview.excel")

# ------------------------------------------------------------------ low level

# Every user-Excel job runs on its own STA thread with its own token. When the server abandons a job stuck in the
# user's Excel, only that job's token is aborted: its next COM call raises (so it rolls back instead of silently
# finishing and duplicating a retried analysis), and jobs running in parallel in other Excels are not affected.
_job = threading.local()


class JobAborted(HostError):
    pass


class JobToken:
    def __init__(self, what: str = "a request"):
        self.aborted = False
        self.what = what                     # the request, as a job waiting for a turn this job holds names it


def bind_job(token: "JobToken | None" = None) -> None:
    _job.token = token or JobToken()
    _job.turns = []
    _job.start_left = None


def unbind_job() -> None:
    """The job's end on its thread (host._call_once, after its turns are released): nothing of it stays bound there."""
    for name in ("token", "turns", "start_left"):
        _job.__dict__.pop(name, None)


def give_up(token: JobToken) -> None:
    """The host gives up on the job of token (an Excel that does not answer: host._call_once): its next COM call
    raises JobAborted, and a job waiting for a turn it still holds is refused (Turn.acquire) - in this process and in
    another one (Turn.mark_given_up)."""
    token.aborted = True
    with _holders_lock:
        held = [tt for tt in _holders.values() if tt.token is token]
    for tt in held:
        tt.mark_given_up()


def _check_alive():
    tk = getattr(_job, "token", None)
    if tk is not None and tk.aborted:
        raise JobAborted("This request was abandoned after a timeout; its partial sheets are being removed.")


def com(call, limit: float = 60.0, check: bool = True, charged: bool = True):
    """A COM call in the user's Excel, retried while Excel is busy - the job's abort checked before every try (check)
    -; then a readable HostError (BUSY). The seconds Excel rejected it count against the job's start (spent), unless
    charged is False (a save of ours is known to run: runner.asked)."""
    timing.count_com()
    try:
        return retry_busy(call, limit, _check_alive if check else None, spent if charged else None)
    except pywintypes.com_error as e:
        if is_busy(e):
            raise HostError(BUSY) from None
        raise


@contextlib.contextmanager
def starting():
    """A job's start: finding its workbook (the reads before its turn), the first contact with its Excel
    (inproc/later.clear) and the wait for Excel to be ready. Only the seconds Excel rejects calls there (a cell in
    edit mode, a dialog) and the wait for Ready count (spent), against one START_WAIT_S (start_limit): the start gives
    up with the busy message after 30 s of them in all, as before the save after the reply. A save of ours is waited
    for longer: the add-in's answer that one runs (runner.asked), a save the first contact runs itself (a call that
    answers is never counted), and Excel's rejections while a save this process asked for may still run
    (inproc/later.saving). Time the start spends otherwise (another agent's turn, a login, Excel starting or opening
    a file) is not counted either. A start inside a start (a session's workbook resolved by its path) is the outer
    one's."""
    if getattr(_job, "start_left", None) is not None:
        yield
        return
    _job.start_left = float(START_WAIT_S)
    try:
        yield
    finally:
        _job.start_left = None


def start_limit(limit: float) -> float:
    """limit, or what is left of the job's start's waits (starting) when that is less - not while a save after the
    reply this process asked for may still run (inproc/later.saving): Excel's rejections may be that save."""
    left = getattr(_job, "start_left", None)
    if left is None:
        return limit
    from .inproc import later
    return limit if later.saving() else max(min(limit, left), 0.0)


def spent(seconds: float) -> None:
    """seconds of a wait for a busy Excel counted against the job's start (starting); nothing outside one."""
    left = getattr(_job, "start_left", None)
    if left is not None:
        _job.start_left = max(left - seconds, 0.0)



def book_name(t) -> str:
    """The name of t's workbook, as a batch finds it (Workbooks(name)): read once per Target (a busy Excel waited
    for, also for an abandoned job's clean-up) and kept on it - a job never renames its workbook."""
    name = getattr(t, "book_name", None)
    if name is None:
        name = str(com(lambda: t.wb.Name, check=False))
        t.book_name = name
    return name


TURN_WAIT_S = 600
MAX_EXCELS = 4                     # the user's Excel(s) + extra windows started for parallel agents
ABANDONED_WAIT_S = 5               # a turn a job given up on still holds is waited for this long, then refused
UNMARK_TRIES = 10                  # a given-up mark another process is reading is removed again this often
ABANDONED = ("An earlier request in that Excel ({held}) was given up after Excel did not respond in time, and it still "
             "holds that Excel: {what} was not started. Check Excel for an open dialog or a cell in edit mode (press Esc "
             "or Enter); if Excel has stopped responding, it has to be closed. Then ask again.")
_holders: dict[int, "Turn"] = {}   # Excel pid -> the turn a job of this process holds there (Turn.given_up)
_holders_lock = threading.Lock()


class Turn:
    """Turn-taking for ONE Excel process, across threads, servers and agents (a lock file per Excel pid). One Excel
    runs all automation on its single UI thread and Smart View works on the active sheet, so two writers in the same
    Excel would only interleave and could put numbers on each other's sheets. Different Excels run in parallel.

    A turn belongs to the job whose thread took it (its token). When the host gives up on that job (an Excel that does
    not answer: host._call_once), its thread may stay blocked in Excel and keeps the turn: it never writes again (its
    next call raises JobAborted) and frees the turn when it ends. A job of this process that waits for such a turn is
    refused after ABANDONED_WAIT_S with what holds it (live, 21.11.2026: the next job waited silently until it was
    given up on too); one that ends meanwhile hands the turn over as any other. A job of ANOTHER process learns it from
    a file next to the lock (user_excel_<pid>.given_up.json: what holds the turn, and which process), written when the
    job is given up on and removed when its turn ends (Task D review, Minor 1); one left behind (its removal failed)
    expires with that Excel process and with the process that wrote it, and goes with the next turn taken there. The lock itself is never freed early:
    the given-up job's call may still complete in that Excel, and two writers there must never be possible."""

    def __init__(self, pid: int):
        import msvcrt
        self._m = msvcrt
        config.SVKIT.mkdir(parents=True, exist_ok=True)
        self.pid = pid
        self.f = open(config.SVKIT / f"user_excel_{pid}.lock", "a+b")
        self.held = False
        self.token = getattr(_job, "token", None)     # the job taking it (bind_job)
        self.marked = False                           # the given-up file is this turn's (mark_given_up)

    @property
    def given_up_file(self):
        return config.SVKIT / f"user_excel_{self.pid}.given_up.json"

    def try_acquire(self) -> bool:
        """The turn, if it is free now. A given-up mark found then is stale - only the job holding the turn marks it,
        and its close could not remove it (Task D2 review, Minor 4) - and is removed."""
        try:
            self.f.seek(0)
            self._m.locking(self.f.fileno(), self._m.LK_NBLCK, 1)
            self.held = True
            with _holders_lock:
                _holders[self.pid] = self
                try:
                    self.given_up_file.unlink(missing_ok=True)
                except OSError:
                    pass                          # it expires with its Excel or its process (given_up_elsewhere)
        except OSError:
            pass
        return self.held
    def given_up(self) -> bool:
        """Held by a job the host has given up on (its token aborted)."""
        return self.held and self.token is not None and self.token.aborted

    def mark_given_up(self) -> None:
        """This held turn's job was given up on: written for jobs of other processes (given_up_elsewhere). Under
        _holders_lock, as close's removal of the file: a turn whose close has begun is not marked (a file left behind
        by that race would have other processes' waiters refused while this server runs)."""
        import json
        import os

        import psutil
        me = psutil.Process()
        what = self.token.what if self.token is not None else "a request"
        try:
            excel_ctime = psutil.Process(self.pid).create_time()
        except psutil.Error:
            excel_ctime = None                        # that Excel is gone: the mark is stale from the start
        with _holders_lock:
            if not self.held or _holders.get(self.pid) is not self:
                return
            try:
                tmp = self.given_up_file.with_suffix(f".{os.getpid()}.tmp")
                tmp.write_text(json.dumps({"what": what, "owner": me.pid, "owner_ctime": me.create_time(),
                                           "excel_ctime": excel_ctime}), encoding="utf-8")
                os.replace(tmp, self.given_up_file)
                self.marked = True
            except OSError as e:
                log.warning("could not tell other processes that the job holding Excel pid %s was given up on: %s",
                            self.pid, e)

    def given_up_elsewhere(self) -> str | None:
        """What holds this Excel's turn in ANOTHER process, when that job was given up on; None otherwise. A mark
        expires with the process that wrote it and with the Excel process it was written for (pid and start time:
        another Excel may get that pid), so one left behind refuses no waiter for longer (Task D2 review, Minor 4)."""
        import json
        import os

        import psutil
        try:
            d = json.loads(self.given_up_file.read_text(encoding="utf-8"))
            owner = int(d["owner"])
            if owner == os.getpid() or abs(psutil.Process(owner).create_time() - float(d["owner_ctime"])) >= 1.0:
                return None
            if abs(psutil.Process(self.pid).create_time() - float(d["excel_ctime"])) >= 1.0:
                return None
            return str(d["what"])
        except (OSError, ValueError, KeyError, TypeError, psutil.Error):
            return None

    def _held_by_given_up(self) -> str | None:
        """What holds this Excel's turn when the host gave up on that job (in this process, or in another one); None
        when it is held by a job nobody gave up on."""
        holder = _holders.get(self.pid)
        if holder is not None:
            return holder.token.what if holder.given_up() else None
        return self.given_up_elsewhere()

    def _unmark(self) -> None:
        """The given-up mark removed, tried again for a moment while another process reads it (a sharing violation);
        one that stays is said, and expires (given_up_elsewhere) or goes with the next turn taken (try_acquire)."""
        for i in range(UNMARK_TRIES):
            try:
                self.given_up_file.unlink(missing_ok=True)
                return
            except OSError as e:
                if i == UNMARK_TRIES - 1:
                    log.warning("could not remove %s (%s): other processes' waiters may be told of a job given up on "
                                "until that Excel or this process ends, or the next turn there is taken",
                                self.given_up_file.name, e)
                else:
                    time.sleep(0.02)

    def acquire(self, what: str) -> float:
        t0 = time.monotonic()
        while not self.try_acquire():
            try:
                _check_alive()
            except JobAborted:
                self.close()
                raise
            held = self._held_by_given_up() if time.monotonic() - t0 >= ABANDONED_WAIT_S else None
            if held is not None:
                self.close()
                raise HostError(ABANDONED.format(held=held, what=what))
            if time.monotonic() - t0 > TURN_WAIT_S:
                self.close()
                raise HostError(f"Another agent has been working in that Excel for {TURN_WAIT_S // 60} min; "
                                f"{what} was not started. Try again shortly.")
            time.sleep(0.25)
        return time.monotonic() - t0

    def close(self):
        if self.held:
            with _holders_lock:
                if _holders.get(self.pid) is self:
                    del _holders[self.pid]
                if self.marked:
                    self._unmark()
                    self.marked = False
            try:
                self.f.seek(0)
                self._m.locking(self.f.fileno(), self._m.LK_UNLCK, 1)
            except OSError:
                pass
            self.held = False
        self.f.close()


def hold_turn(pid: int, wait: bool = True, what: str = "this request", app=None) -> bool:
    """Take this job's turn in Excel `pid` (kept until the job ends). wait=False: only if it is free right now. With
    app (that Excel), the job's first contact with it follows once the turn is held, every time it is asked for
    (inproc/later.clear: a save after the last reply still waiting there runs first, one that runs is waited for),
    before the caller's plain COM writes there."""
    turns = getattr(_job, "turns", None)
    if turns is None:
        turns = _job.turns = []
    if any(tt.pid == pid for tt in turns):
        _first_contact(app)
        return True
    tt = Turn(pid)
    if not wait:
        if not tt.try_acquire():
            tt.close()
            return False
    else:
        from ..core.timing import stage
        with stage("wait for that Excel (another agent)"):
            waited = tt.acquire(what)
        if waited > 1:
            log.info("waited %.1fs for Excel pid %s (another agent)", waited, pid)
    turns.append(tt)
    _first_contact(app)
    return True


def _first_contact(app) -> None:
    if app is not None:
        from .inproc import later
        later.clear(app)


def release_turns() -> None:
    for tt in getattr(_job, "turns", None) or []:
        tt.close()
    _job.turns = []
    _job.start_left = None


@contextlib.contextmanager
def fast_excel(app):
    """No screen redraw, no automatic recalculation and no Excel events while we write/format (the user's Excel
    otherwise repaints and recalculates after every single cell: measured 17 s vs 1.6 s for one answer sheet; and
    every add-in listening to sheet events - Smart View, Arixcel - runs its code on each write). The user's own
    settings are restored afterwards, also on failure; the caller recalculates explicitly when it needs results."""
    prev_su = prev_calc = prev_ev = None

    def quiet():
        app.ScreenUpdating = False
        app.EnableEvents = False
        if prev_calc != -4135:                                  # xlCalculationManual
            app.Calculation = -4135
    try:
        prev_su, prev_calc, prev_ev = com(lambda: (app.ScreenUpdating, app.Calculation, app.EnableEvents),
                                          check=False)
        com(quiet, check=False)
    except Exception:   # noqa: BLE001 - e.g. no workbook open: just run normally
        pass
    try:
        yield
    finally:
        for name, val in (("Calculation", prev_calc), ("EnableEvents", prev_ev), ("ScreenUpdating", prev_su)):
            if val is not None:
                try:
                    com(lambda: setattr(app, name, val), limit=30, check=False)
                except Exception:   # noqa: BLE001
                    log.warning("could not restore Application.%s=%s", name, val)
