"""One user-Excel job: COM calls with retry/abort, the machine-wide per-Excel turn, fast mode."""
from __future__ import annotations

import contextlib
import logging
import threading
import time

import pywintypes

from smartass.core import config, timing
from .host import HostError
from .hostproc import is_busy, retry_busy

EDIT_WAIT_S = 30
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
    def __init__(self):
        self.aborted = False


def bind_job(token: "JobToken | None" = None) -> None:
    _job.token = token or JobToken()
    _job.turns = []


def _check_alive():
    tk = getattr(_job, "token", None)
    if tk is not None and tk.aborted:
        raise JobAborted("This request was abandoned after a timeout; its partial sheets are being removed.")


def com(call, limit: float = 60.0, check: bool = True):
    """A COM call in the user's Excel, retried while Excel is busy; then a readable HostError (BUSY)."""
    timing.count_com()
    if check:
        _check_alive()
    try:
        return retry_busy(call, limit)
    except pywintypes.com_error as e:
        if is_busy(e):
            raise HostError(BUSY) from None
        raise


TURN_WAIT_S = 600
MAX_EXCELS = 4                     # the user's Excel(s) + extra windows started for parallel agents


class Turn:
    """Turn-taking for ONE Excel process, across threads, servers and agents (a lock file per Excel pid). One Excel
    runs all automation on its single UI thread and Smart View works on the active sheet, so two writers in the same
    Excel would only interleave and could put numbers on each other's sheets. Different Excels run in parallel."""

    def __init__(self, pid: int):
        import msvcrt
        self._m = msvcrt
        config.SVKIT.mkdir(parents=True, exist_ok=True)
        self.pid = pid
        self.f = open(config.SVKIT / f"user_excel_{pid}.lock", "a+b")
        self.held = False

    def try_acquire(self) -> bool:
        try:
            self.f.seek(0)
            self._m.locking(self.f.fileno(), self._m.LK_NBLCK, 1)
            self.held = True
        except OSError:
            pass
        return self.held

    def acquire(self, what: str) -> float:
        t0 = time.monotonic()
        while not self.try_acquire():
            if time.monotonic() - t0 > TURN_WAIT_S:
                self.close()
                raise HostError(f"Another agent has been working in that Excel for {TURN_WAIT_S // 60} min; "
                                f"{what} was not started. Try again shortly.")
            time.sleep(0.25)
        return time.monotonic() - t0

    def close(self):
        if self.held:
            try:
                self.f.seek(0)
                self._m.locking(self.f.fileno(), self._m.LK_UNLCK, 1)
            except OSError:
                pass
            self.held = False
        self.f.close()


def hold_turn(pid: int, wait: bool = True, what: str = "this request") -> bool:
    """Take this job's turn in Excel `pid` (kept until the job ends). wait=False: only if it is free right now."""
    turns = getattr(_job, "turns", None)
    if turns is None:
        turns = _job.turns = []
    if any(tt.pid == pid for tt in turns):
        return True
    tt = Turn(pid)
    if not wait:
        if not tt.try_acquire():
            tt.close()
            return False
    else:
        from smartass.core.timing import stage
        with stage("wait for that Excel (another agent)"):
            waited = tt.acquire(what)
        if waited > 1:
            log.info("waited %.1fs for Excel pid %s (another agent)", waited, pid)
    turns.append(tt)
    return True


def release_turns() -> None:
    for tt in getattr(_job, "turns", None) or []:
        tt.close()
    _job.turns = []


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
