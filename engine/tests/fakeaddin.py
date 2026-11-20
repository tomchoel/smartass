"""A fake of the add-in's entry points (xl/inproc/Runner.cs) on a fake Excel's Application, for the tests: RegisterXLL
loads it; Run of SmartassPing_<HASH> answers the hash; SmartassRun_<HASH>(job path) reads, deletes and runs the job
through the COM executor and returns its record; SmartassSaveLater_<HASH>(workbook, autosave, token, delay ms) keeps
the save until the test lets the timer fire (release), SmartassClear_<HASH>() (a job's first contact: a waiting save
run first, "clear", or the outcome it could not write), as Runner.cs does: a save still waiting runs first at the start
of the next Run (of a batch or of another SaveLater), the save then AutoSave back on when asked (also after a failed
save), and the outcome written to config.INPROC_DIR / "saved" / "<pid>.json" (pid: the Excel's key, id(app) here). A
workbook of the request's name with another FullName at save time is not saved ("closed"), as in Runner.cs."""
from __future__ import annotations

import json
import time
from pathlib import Path

from smartass.core import config
from smartass.xl.inproc import build, later, runner

HASH = "abcd1234"


class AddIn:
    def __init__(self, app):
        self.app = app
        self.loaded = False
        self.pending: dict | None = None
        self.asked: list[tuple] = []                    # every SaveLater's arguments
        self.answer: str | None = None                  # SaveLater's answer instead of "queued" (a refusal)
        self.save_error: str | None = None              # the next save fails with this text
        self.log: list[str] = []                        # "save <name>", "autosave <name> True", "batch", "clear"
        self.busy = 0                                   # the next calls answer BUSY this many times (a save running)
        self.write_fails = False                        # the outcome cannot be written: kept for the next Clear
        self.unwritten: dict | None = None

    def RegisterXLL(self, path):
        self.loaded = True
        return True

    def Run(self, macro, *args):
        assert self.loaded and macro.endswith("_" + HASH), macro
        if macro.startswith("SmartassPing_"):
            return HASH
        if self.busy:
            self.busy -= 1
            return runner.BUSY + "the save after the reply is running"
        if macro.startswith("SmartassClear_"):
            self.log.append("clear")
            self.save_now("before a job")
            out, self.unwritten = self.unwritten, None
            return "clear" if out is None else "unwritten: " + json.dumps(out)
        self.save_now("before a batch" if macro.startswith("SmartassRun_") else "before the next")
        if macro.startswith("SmartassSaveLater_"):
            workbook, autosave, token, delay = args
            self.asked.append(tuple(args))
            if self.answer is not None:
                return self.answer
            try:
                full = self.app.Workbooks(workbook).FullName
            except Exception:   # noqa: BLE001
                return "not run: refused: no workbook of that name is open"
            self.pending = {"workbook": workbook, "full_name": full, "autosave": autosave, "token": token,
                            "queued": time.time()}
            return "queued"
        path = Path(args[0])
        job = json.loads(path.read_text(encoding="utf-8"))
        path.unlink()
        self.log.append("batch")
        return json.dumps(runner.execute(self.app, job, check=False))

    def release(self) -> dict | None:
        """The add-in's timer fired: the save waiting runs now; its outcome."""
        return self.save_now("after the reply")

    def save_now(self, how: str) -> dict | None:
        p, self.pending = self.pending, None
        if p is None:
            return None
        started = time.time()
        error, autosave_error, closed = None, None, False
        try:
            wb = self.app.Workbooks(p["workbook"])
        except Exception as e:   # noqa: BLE001
            wb, closed, error = None, True, str(e)
        if wb is not None and p.get("full_name") and wb.FullName != p["full_name"]:
            wb, closed, error = None, True, "another workbook of that name is open now"
        if wb is not None:
            try:
                if self.save_error is not None:
                    raise RuntimeError(self.save_error)
                wb.Save()
                self.log.append(f"save {p['workbook']}")
            except Exception as e:   # noqa: BLE001
                error = str(e)
            if p["autosave"]:
                wb.AutoSaveOn = True
                self.log.append(f"autosave {p['workbook']} True")
        out = {"token": p["token"], "workbook": p["workbook"], "full_name": p.get("full_name", ""), "ok": error is None, "error": error, "closed": closed,
               "autosave": p["autosave"], "autosave_error": autosave_error, "how": how, "queued": p["queued"],
               "started": started, "seconds": 0.25, "switched": 0.0}
        if self.write_fails:
            self.unwritten = {**out, "unwritten": "the disk is full"}
            return out
        folder = config.INPROC_DIR / "saved"
        folder.mkdir(parents=True, exist_ok=True)
        (folder / f"{id(self.app)}.json").write_text(json.dumps(out), encoding="utf-8")
        return out


def install(monkeypatch, app, folder: Path) -> AddIn:
    """The add-in on app (a fake Excel's Application): SV_INPROC on, its builds and job files under folder, each fake
    Excel its own process (id(app)), 64-bit, built without a compiler."""
    monkeypatch.setattr(config, "INPROC", True)
    monkeypatch.setattr(config, "INPROC_DIR", folder / "inproc")
    monkeypatch.setattr(runner, "_EXCELS", {})
    monkeypatch.setattr(runner, "_identify", lambda app, check=True: (id(app), 0.0))
    monkeypatch.setattr(runner, "_is_64bit", lambda pid: True)
    monkeypatch.setattr(build, "ensure", lambda: (folder / f"smartass_run_{HASH}.xll", HASH))
    monkeypatch.setattr(later, "_build_hash", lambda: HASH)       # a clear before this server loaded the add-in
    addin = AddIn(app)
    app.RegisterXLL, app.Run = addin.RegisterXLL, addin.Run
    return addin
