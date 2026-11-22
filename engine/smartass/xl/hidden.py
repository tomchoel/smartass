"""The engine's own hidden Excel on the host COM thread (HiddenExcel, ExcelHost's base in xl/host.py): its start
check, the VBA bridge, Smart View's global options, its sheets. Only the hidden Excel's code imports this module."""
from __future__ import annotations

import json
import logging
import threading
from typing import Any, Callable

import pywintypes

from ..core import config
from .hostproc import _is_ours, retry_busy

log = logging.getLogger("smartview.host")


class HiddenExcel:
    """ExcelHost's hidden Excel on the host COM thread: start check, VBA bridge, Smart View options, sheets."""

    def on_host(self, fn: Callable[[], Any], timeout: float | None = None):
        """Run fn on the host COM thread (where the hidden Excel's COM objects live), starting that Excel if needed.
        Called from a user-Excel job thread, it waits for the result; on the host thread it just runs."""
        if threading.current_thread() is self.thread:
            self._ensure_started()
            return fn()

        def job():
            self._wait_reset()
            self._ensure_started()
            return fn()
        released = threading.Event()
        fut = self.thread.submit(job, released.set)
        try:
            return fut.result(timeout or config.CALL_TIMEOUT_S)
        finally:
            if fut.done():                        # it answered (a result or its error): its release, as call's
                self._watch_release(released, "a job on the engine's hidden Excel", pid=self.pid)

    def _ensure_started(self):
        if self.app is not None:
            alive = self.pid is not None and self.ctime is not None and _is_ours(self.pid, self.ctime)
            try:
                if alive:
                    _ = self.app.Hwnd
                    return
            except Exception:   # noqa: BLE001 - a dead proxy raises com_error or AttributeError
                pass
            self._mark_dead("Excel was gone before the call (killed or crashed); restarting")
            self._drop_refs()
        self._start()

    def _load_bridge(self, wb):
        for path in config.BAS_FILES:
            mod = wb.VBProject.VBComponents.Add(1)
            mod.Name = "SmartViewVBA" if path.stem == "smartview_body" else path.stem
            mod.CodeModule.AddFromString(path.read_text(encoding="cp1252"))

    def _apply_global_options(self):
        """The user's own values are recorded ONCE in a file before we change them, so a crash (no clean shutdown)
        never turns our settings into the 'originals' of the next run. While the old svkit engine may still have its
        own values applied, its file holds the real originals: copy them (the legacy file stays for that engine)."""
        orig_file = config.SVKIT / "smartview_options_original.json"
        legacy = config.LEGACY_HOME / "smartview_options_original.json" if config.LEGACY_HOME else None
        saved, src = None, None
        for f in [p for p in (orig_file, legacy) if p]:
            try:
                saved, src = {int(k): v for k, v in json.loads(f.read_text()).items()}, f
                break
            except (OSError, ValueError):
                pass
        if saved is None:
            saved = {}
            for k in config.GLOBAL_OPTIONS:
                got = self._run("B_GetOption", k, "")
                if got and got[0] == "OK":
                    saved[k] = got[1]
        if src == orig_file:
            # A clean shutdown deletes our file, so it is left over from a crashed run: put the user's originals
            # back first (our values were never restored), then apply ours below as on any start.
            log.info("smartview options: restoring the originals left by a crashed run")
            for k, v in saved.items():
                try:
                    self._run("B_SetOption", k, v, "")
                except pywintypes.com_error as e:
                    log.info("restoring option %s failed: %s", k, str(e)[:200])
        else:
            orig_file.parent.mkdir(parents=True, exist_ok=True)
            orig_file.write_text(json.dumps({str(k): v for k, v in saved.items()}))
        self.saved_options = saved
        for k, v in config.GLOBAL_OPTIONS.items():
            self._run("B_SetOption", k, v, "")

    def _restore_global_options(self):
        ok = True
        for k, v in self.saved_options.items():
            try:
                self._run("B_SetOption", k, v, "")
            except pywintypes.com_error:
                ok = False
        if ok:
            try:
                (config.SVKIT / "smartview_options_original.json").unlink()
            except OSError:
                pass

    def _run(self, fn: str, *args):
        """Application.Run on the bridge, with retry on 'call rejected'."""
        return self.com(lambda: self.app.Run(f"'{self.wb.Name}'!modBridge.{fn}", *args), limit_s=60)

    com = staticmethod(retry_busy)

    def sheet(self, name: str, fresh: bool = False):
        ws = None
        for s in self.wb.Worksheets:
            if s.Name == name:
                ws = s
                break
        if ws is not None and fresh:
            self.app.DisplayAlerts = False
            ws.Delete()
            self.connected.pop(name, None)
            ws = None
        if ws is None:
            ws = self.wb.Worksheets.Add()
            ws.Name = name
        return ws

    def connect_sheet(self, name: str, connection: str):
        if self.connected.get(name) == connection:
            return
        self.wb.Worksheets(name).Activate()
        self.run("B_Connect", name, connection)
        self.connected[name] = connection

    def status(self) -> dict:
        return {
            "excel_running": self.app is not None,
            "pid": self.pid,
            "visible": self.visible,
            "started_at": self.started_at,
            "connected_sheets": dict(self.connected),
            "restarts": self.restarts,
            "smartview_version": self.sv_version,
            "last_error": self.last_error,
        }

    def _drop_refs(self):
        self.app = self.wb = None
        self.connected.clear()

    def _mark_dead(self, why: str):
        self.last_error = why
        self.restarts += 1
        log.warning("Excel marked dead: %s", why)
