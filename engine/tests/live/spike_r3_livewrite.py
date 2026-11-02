"""Risk spike R3 (docs/specs/2026-11-02-pipeline-architecture-design.md §7): is a one-block write into the session
workbook safe while the user works in the same Excel?

Cases, each writing a 70 x 9 block of synthetic numbers (no EPM data) into the session workbook WITHOUT activating it:
  baseline     nothing else going on
  other_active the user's own workbook is the active one -> it must stay active and untouched
  busy         Excel runs an 8 s VBA loop (the bridge's B_TestBusy) -> the write waits, then lands
  edit_mode    a cell in the user's workbook is being edited (keystroke posted to OUR Excel window only, never
               SendKeys) and Esc follows after 6 s -> the write waits, then lands; the user's book is untouched
Run from engine\\:  uv run python tests\\live\\spike_r3_livewrite.py      Writes docs\\audit\\<date>-spike-r3.json.
Go: every case lands exactly, the user's workbook stays active and unchanged, waits end within the retry limit."""
from __future__ import annotations

import datetime as dt
import json
import random
import sys
import threading
import time
from pathlib import Path

import pythoncom
import win32api
import win32con

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from smartview_mcp import userexcel  # noqa: E402
from smartview_mcp.host import ExcelHost  # noqa: E402
from smartview_mcp.userexcel import com  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from helpers import close_scratch  # noqa: E402

AUDIT = Path(__file__).resolve().parents[3] / "docs" / "audit"
ROWS, COLS = 70, 9
LIMIT = 30.0


def block(seed: int) -> list[list[float]]:
    rnd = random.Random(seed)
    return [[round(rnd.uniform(-1e7, 1e7), 2) for _ in range(COLS)] for _ in range(ROWS)]


def window_of(pid: int, caption: str) -> int:
    for p, _, ex7s in userexcel._excel_windows():
        if p == pid:
            for h, cap in ex7s:
                if caption.lower() in cap.lower():
                    return h
            if ex7s:
                return ex7s[0][0]
    raise RuntimeError(f"no workbook window for {caption} in Excel {pid}")


def busy_in_thread(ex7: int, macro: str, seconds: float) -> threading.Thread:
    def run():
        pythoncom.CoInitialize()
        try:
            app = userexcel._app_from_window(ex7)
            app.Run(macro, seconds)
        finally:
            pythoncom.CoUninitialize()
    th = threading.Thread(target=run, daemon=True)
    th.start()
    return th


def key(hwnd: int, vk: int | None = None, char: str | None = None) -> None:
    if char:
        win32api.PostMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
    else:
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, vk, 0)


def main() -> int:
    rep: dict = {"spike": "R3", "when": dt.datetime.now().isoformat(timespec="seconds"), "cases": {}}
    h = ExcelHost().ensure()
    t = ub = None
    try:
        t = userexcel.resolve_target(h, "new", title="spike-r3")
        ws = userexcel.add_sheet(t, "R3 write")
        ub = com(t.app.Workbooks.Add)                     # stands in for the user's own workbook
        com(ub.Activate)
        user_ex7 = window_of(t.pid, ub.Name)
        macro = f"'{t.bridge.Name}'!modBridge.B_TestBusy"
        addr = f"B5:{chr(65 + COLS)}{4 + ROWS}"

        def case(name: str, seed: int, before=None, after=None):
            data = block(seed)
            if before:
                before()
            t0 = time.monotonic()
            err = None
            try:
                com(lambda: setattr(ws.Range(addr), "Value2", data), limit=LIMIT)
            except Exception as e:  # noqa: BLE001 - the spike records the failure
                err = str(e)[:200]
            secs = time.monotonic() - t0
            if after:
                after()
            got = [list(r) for r in com(lambda: ws.Range(addr).Value2, limit=LIMIT)]
            active = com(lambda: t.app.ActiveWorkbook.Name, limit=LIMIT)
            used = com(lambda: ub.Worksheets(1).UsedRange.Address, limit=LIMIT)
            a1 = com(lambda: ub.Worksheets(1).Range("A1").Value2, limit=LIMIT)
            rep["cases"][name] = {"write_s": round(secs, 2), "error": err, "landed": got == data,
                                  "user_book_still_active": active == ub.Name,
                                  "user_book_untouched": used == "$A$1" and a1 is None}
            print(f"{name:13} {rep['cases'][name]}")

        case("baseline", 1)
        case("other_active", 2)

        busy: list[threading.Thread] = []
        case("busy", 3, before=lambda: (busy.append(busy_in_thread(user_ex7, macro, 8)), time.sleep(1.0)),
             after=lambda: busy[0].join(15))

        def start_edit():
            key(user_ex7, char="x")
            time.sleep(0.5)
            threading.Timer(6.0, lambda: key(user_ex7, vk=win32con.VK_ESCAPE)).start()
            try:
                rep["edit_mode_detected"] = not t.app.Ready
            except Exception:  # noqa: BLE001 - a rejected call also means edit mode
                rep["edit_mode_detected"] = True
        case("edit_mode", 4, before=start_edit)
    finally:
        try:
            if ub is not None:
                com(lambda: ub.Close(False), limit=LIMIT)
            if t is not None:
                close_scratch(t)
        except Exception as e:  # noqa: BLE001 - cleanup only
            print("cleanup:", e)
        h.shutdown()
    c = rep["cases"]
    rep["go"] = {"all_landed": all(x["landed"] for x in c.values()),
                 "user_book_safe": all(x["user_book_still_active"] and x["user_book_untouched"] for x in c.values()),
                 "no_errors": not any(x["error"] for x in c.values())}
    out = AUDIT / f"{dt.date.today().isoformat()}-spike-r3.json"
    out.write_text(json.dumps(rep, indent=1), encoding="utf-8")
    print("edit mode detected:", rep.get("edit_mode_detected"), "| go:", rep["go"], "->", out.name)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
