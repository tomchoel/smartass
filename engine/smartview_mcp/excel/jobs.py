"""One user-Excel job: COM calls with retry/abort, the machine-wide per-Excel turn, fast mode."""
from __future__ import annotations

import contextlib
import logging
import threading
import time

import pythoncom
import pywintypes

from .. import config
from ..host import HostError

RETRY = {-2147418111, -2147417846}            # call rejected / retry later (edit mode, busy)
EDIT_WAIT_S = 30

log = logging.getLogger("smartview.excel")

# ------------------------------------------------------------------ low level
VBA_E_IGNORE = -2146777998                   # 0x800AC472: Excel busy, delivered inside DISP_E_EXCEPTION

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


def _busy(e: pywintypes.com_error) -> bool:
    inner = e.excepinfo[5] if e.excepinfo and len(e.excepinfo) > 5 else None
    return e.hresult in RETRY or e.hresult == VBA_E_IGNORE or inner == VBA_E_IGNORE


def com(call, limit: float = 60.0, check: bool = True):
    if check:
        _check_alive()
    t0, d = time.monotonic(), 0.1
    while True:
        try:
            return call()
        except pywintypes.com_error as e:
            if _busy(e) and time.monotonic() - t0 < limit:
                pythoncom.PumpWaitingMessages()
                time.sleep(d)
                d = min(d * 2, 2.0)
                continue
            if _busy(e):
                raise HostError("Excel is busy - probably a cell is being edited or a dialog is open. Press Enter or "
                                "Esc in Excel and ask again.") from None
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
        from ..timing import stage
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
    """No screen redraw and no automatic recalculation while we write/format (the user's Excel otherwise repaints
    and recalculates after every single cell: measured 17 s vs 1.6 s for one answer sheet). The user's own settings
    are restored afterwards, also on failure; the caller recalculates explicitly when it needs results."""
    prev_su = prev_calc = None
    try:
        prev_su = com(lambda: app.ScreenUpdating, check=False)
        prev_calc = com(lambda: app.Calculation, check=False)
        com(lambda: setattr(app, "ScreenUpdating", False), check=False)
        if prev_calc != -4135:                                  # xlCalculationManual
            com(lambda: setattr(app, "Calculation", -4135), check=False)
    except Exception:   # noqa: BLE001 - e.g. no workbook open: just run normally
        pass
    try:
        yield
    finally:
        for name, val in (("Calculation", prev_calc), ("ScreenUpdating", prev_su)):
            if val is not None:
                try:
                    com(lambda: setattr(app, name, val), limit=30, check=False)
                except Exception:   # noqa: BLE001
                    log.warning("could not restore Application.%s=%s", name, val)
